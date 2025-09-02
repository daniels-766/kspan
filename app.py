from flask import Flask, render_template, redirect, url_for, request, flash, send_file, send_from_directory
from flask_login import login_user, login_required, logout_user, current_user
from werkzeug.security import generate_password_hash, check_password_hash
import os
from extensions import db, migrate, login_manager
from datetime import datetime, timedelta
from models import Ticket, NomorTicket, User, Kontak, History, Catatan, db
from sqlalchemy.orm import joinedload, aliased
from sqlalchemy import func, or_, and_, asc, distinct, union
from werkzeug.utils import secure_filename
import pandas as pd
from io import BytesIO
from flask_apscheduler import APScheduler
from pytz import timezone
import pytz
from werkzeug.serving import is_running_from_reloader
from collections import defaultdict
from zoneinfo import ZoneInfo

JAKARTA_TZ = ZoneInfo("Asia/Jakarta")

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'mysql+pymysql://root:@localhost/dashboard-an'
# app.config['SQLALCHEMY_DATABASE_URI'] = 'mysql+pymysql://root:Ksp%40888@localhost/dashboard_an'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['SECRET_KEY'] = 'b35dfe6ce150230940bd145823034486'
app.config['UPLOAD_FOLDER'] = 'static/uploads'
app.config['MAX_CONTENT_LENGTH'] = 150 * 1024 * 1024 

UPLOAD_FOLDER = os.path.join('static', 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename  

db.init_app(app)
migrate.init_app(app, db)
login_manager.init_app(app)
login_manager.login_view = 'login'

@app.context_processor
def inject_sla_warning_tickets():
    subquery = (
        db.session.query(
            Ticket.nomor_ticket_id,
            func.min(Ticket.sla).label("min_sla")
        )
        .filter(Ticket.sla.between(1, 3))
        .group_by(Ticket.nomor_ticket_id)
    ).subquery()

    TicketAlias = aliased(Ticket)

    sla_warning_tickets = (
        db.session.query(TicketAlias)
        .join(subquery, and_(
            TicketAlias.nomor_ticket_id == subquery.c.nomor_ticket_id,
            TicketAlias.sla == subquery.c.min_sla
        ))
        .join(NomorTicket, NomorTicket.id == TicketAlias.nomor_ticket_id)
        .filter(NomorTicket.status.in_(['aktif', 'Reopen']))
        .order_by(TicketAlias.sla.asc())
        .all()
    )

    return {'sla_warning_tickets': sla_warning_tickets}

class Config:
    SCHEDULER_API_ENABLED = True

app.config.from_object(Config())

scheduler = APScheduler()
scheduler.init_app(app)

@scheduler.task('cron', id='decrease_sla_daily', hour=0, minute=0, timezone='Asia/Jakarta')
def decrease_sla():
    with app.app_context():
        tickets = Ticket.query.filter(Ticket.sla > 0).all()
        updated_count = 0
        for ticket in tickets:
            if ticket.status_ticket != '4':
                ticket.sla -= 1
                updated_count += 1
        db.session.commit()
        print(f"SLA updated at {datetime.now(timezone('Asia/Jakarta'))} — {updated_count} ticket(s) updated.")

def update_ticket_fields():
    with app.app_context():
        tickets = Ticket.query.filter(
            (Ticket.nama_os.in_([None, "-", "None"])) |
            (Ticket.nama_bucket.in_([None, "-", "None"]))
        ).all()

        for ticket in tickets:
            if ticket.nama_os in [None, "-", "None"]:
                ticket.nama_os = ""
            if ticket.nama_bucket in [None, "-", "None"]:
                ticket.nama_bucket = ""

        db.session.commit()
        print(f"Updated {len(tickets)} ticket(s).")

scheduler.add_job(
    id='update_none_fields',
    func=update_ticket_fields,
    trigger='interval',
    minutes=1
)

if not scheduler.running:
    scheduler.start()

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

@app.errorhandler(404)
def page_not_found(error):
    return render_template('404.html'), 404

from sqlalchemy import event

@event.listens_for(Ticket, "after_update")
def ticket_after_update(mapper, connection, target):
    connection.execute(
        NomorTicket.__table__.update().
        where(NomorTicket.id == target.nomor_ticket_id).
        values(change_date=datetime.now(JAKARTA_TZ))
    )

@event.listens_for(Ticket, "after_insert")
def ticket_after_insert(mapper, connection, target):
    connection.execute(
        NomorTicket.__table__.update().
        where(NomorTicket.id == target.nomor_ticket_id).
        values(change_date=datetime.now(JAKARTA_TZ))
    )

@app.route('/register', methods=['GET', 'POST'])
def register():
    if request.method == 'POST':
        username = request.form['username']
        email = request.form['email']
        phone = request.form['phone']
        password = request.form['password']
        hashed_pw = generate_password_hash(password)

        if User.query.filter_by(username=username).first():
            flash('Username sudah terdaftar.')
            return redirect(url_for('register'))
        if User.query.filter_by(email=email).first():
            flash('Email sudah terdaftar.')
            return redirect(url_for('register'))

        user = User(username=username, email=email, phone=phone, password=hashed_pw)
        db.session.add(user)
        db.session.commit()
        flash('Registrasi berhasil! Silakan login.')
        return redirect(url_for('login'))

    return render_template('register.html')

@app.route('/')
def home():
    return redirect(url_for('login'))

@app.route('/history')
@login_required
def history():
    page = request.args.get('page', 1, type=int)
    history_list = History.query.order_by(History.tanggal.desc()).paginate(page=page, per_page=10, error_out=False)
    return render_template('history.html', user=current_user, history_list=history_list)

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        login_input = request.form['username'] 
        password = request.form['password']

        user = User.query.filter(
            (User.username == login_input) | (User.email == login_input)
        ).first()

        if user and check_password_hash(user.password, password):
            login_user(user)
            flash('Login berhasil!')
            if user.role == 'admin':
                return redirect(url_for('admin_dashboard'))
            elif user.role == 'qc':
                return redirect(url_for('qc_dashboard'))
            else:
                return redirect(url_for('staff_dashboard'))
        else:
            flash('Username/email atau password salah.')

    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():
    logout_user()
    flash('Anda telah logout.')
    return redirect(url_for('login'))

@app.route('/admin_dashboard')
@login_required
def admin_dashboard():
    if current_user.role != 'admin':
        flash('Akses ditolak: Anda bukan admin.')
        return redirect(url_for('staff_dashboard'))
    
    staff_users = User.query.filter(User.role != 'admin').all()
    
    return render_template('admin_dashboard.html', user=current_user, users=staff_users)

@app.route('/qc-dashboard')
@login_required
def qc_dashboard():
    if current_user.role != 'qc':
        flash('Akses ditolak: Anda bukan QC!!!')
        return redirect(url_for('index'))

    page = request.args.get('page', 1, type=int)
    jenis = request.args.get('jenis')
    status = request.args.get('status')
    tanggal = request.args.get('tanggal')
    q = request.args.get('q')

    nomor_ticket_query = NomorTicket.query.filter(
        NomorTicket.id_qc == current_user.id,
        NomorTicket.label_case.is_(None),
        or_(
            NomorTicket.status == None,
            and_(
                NomorTicket.status != 'close',
                NomorTicket.status != 'reopen'
            )
        )
    )

    if q:
        matching_ids_from_nama = db.session.query(Ticket.nomor_ticket_id)\
            .filter(Ticket.nama_nasabah.ilike(f"%{q}%"))\
            .distinct().all()
        matching_ids_from_nama = [id[0] for id in matching_ids_from_nama]

        matching_ids_from_nomor = db.session.query(NomorTicket.id)\
            .filter(NomorTicket.nomor_ticket.ilike(f"%{q}%"))\
            .all()
        matching_ids_from_nomor = [id[0] for id in matching_ids_from_nomor]

        all_matching_ids = list(set(matching_ids_from_nama + matching_ids_from_nomor))
        nomor_ticket_query = nomor_ticket_query.filter(NomorTicket.id.in_(all_matching_ids))

    tickets_grouped = []

    for nt in nomor_ticket_query.all():
        query = Ticket.query.options(joinedload(Ticket.nomor_ticket))\
            .filter(Ticket.nomor_ticket_id == nt.id)

        if jenis:
            query = query.filter_by(jenis_pengaduan=jenis)
        if status:
            query = query.filter_by(status_ticket=status)
        if tanggal:
            try:
                tanggal_obj = datetime.strptime(tanggal, "%Y-%m-%d")
                query = query.filter(func.date(Ticket.tanggal) == tanggal_obj.date())
            except ValueError:
                pass

        first_ticket = query.order_by(Ticket.created_time.asc()).first()
        if first_ticket:
            tickets_grouped.append(first_ticket)

    tickets_grouped.sort(
        key=lambda x: x.created_time or datetime.min,
        reverse=True
    )

    per_page = 10
    total = len(tickets_grouped)
    start = (page - 1) * per_page
    end = start + per_page
    paginated_items = tickets_grouped[start:end]

    class Pagination:
        def __init__(self, items, page, per_page, total):
            self.items = items
            self.page = page
            self.per_page = per_page
            self.total = total
            self.pages = (total + per_page - 1) // per_page
            self.has_prev = page > 1
            self.has_next = page < self.pages
            self.prev_num = page - 1
            self.next_num = page + 1

    pagination = Pagination(paginated_items, page, per_page, total)

    count_by_nomor_ticket = dict(
        db.session.query(
            Ticket.nomor_ticket_id,
            func.count(Ticket.id)
        ).group_by(Ticket.nomor_ticket_id).all()
    )

    jumlah_tiket_aktif = db.session.query(NomorTicket)\
        .join(Ticket, Ticket.nomor_ticket_id == NomorTicket.id)\
        .filter(NomorTicket.label_case.is_(None), NomorTicket.id_qc == current_user.id)\
        .distinct()\
        .count()

    return render_template(
        'qc_dashboard.html',
        user=current_user,
        tickets=pagination,
        count_by_nomor_ticket=count_by_nomor_ticket,
        jumlah_tiket_aktif=jumlah_tiket_aktif
    )

@app.route('/qc/nomor-ticket/<int:nomor_ticket_id>')
@login_required
def list_ticket_by_nomor_qc(nomor_ticket_id):
    if current_user.role != 'qc':
        return redirect(request.referrer)
    
    nomor_ticket = NomorTicket.query.filter_by(id=nomor_ticket_id, id_qc=current_user.id).first_or_404()

    tickets = Ticket.query.filter_by(nomor_ticket_id=nomor_ticket_id)\
        .options(joinedload(Ticket.catatans))\
        .order_by(Ticket.created_time.asc()).all()
    
    catatan_list = Catatan.query.filter_by(nomor_ticket_id=nomor_ticket_id)\
    .order_by(Catatan.tanggal.asc()).all()
    
    ticket_ids = [t.id for t in tickets]
    catatan_map = {}
    if ticket_ids:
        all_catatan = Catatan.query.filter(Catatan.ticket_id.in_(ticket_ids))\
            .order_by(Catatan.tanggal.asc())
        for catatan in all_catatan:
            catatan_map.setdefault(catatan.ticket_id, []).append(catatan)
    else:
        all_catatan = []

    jenis_pengaduan_map = {
        1: "Informasi Pengajuan",
        2: "Permintaan Kode OTP",
        3: "Informasi Tenor",
        4: "Informasi Tagihan",
        5: "Informasi Denda",
        6: "Pembatalan Pinjaman",
        7: "Informasi Pencairan Dana",
        8: "Perilaku Petugas Penagihan",
        9: "Informasi Pembayaran",
        10: "Discount / Pemutihan"
    }

    detail_pengaduan_map = {
        1: [
            "Hasil Pengajuan",
            "Pengajuan Ditolak",
            "Status Pengajuan sedang ditransfer",
            "Tidak bisa pengajuan ulang karena keterlambatan",
            "Verifikasi Bank gagal",
            "Verifikasi KTP gagal",
            "Cara pengajuan",
            "Perubahan Nomor Handphone",
            "Perubahan Nomor Rekening"
        ],
        2: [
            "OTP Limit",
            "Tidak terima SMS OTP"
        ],
        3: [
            "Informasi Pinjaman"
        ],
        4: [
            "Konsultasi detail pinjaman saat ini",
            "Konsultasi Perpanjangan",
            "Bukti Transfer"
        ],
        5: [
            "Denda Keterlambatan"
        ],
        6: [
            "Hapus Data (Penutupan Akun)",
            "Pembatalan Pinjaman"
        ],
        7: [
            "Pencairan dana berhasil",
            "Status pencairan dana gagal",
            "Status pengajuan pencairan dana ulang",
            "Tidak terima dana",
            "Operasi Gagal (tidak bisa verifikasi wajah dan KTP)"
        ],
        8: [
            "Keluhan Penagihan",
            "Keluhan Reminder",
            "Penipuan"
        ],
        9: [
            "Konfirmasi Pembayaran",
            "Pembayaran belum masuk",
            "Pembayaran bukan ke VA UATAS",
            "Refund (pembayaran double)",
            "Meminta VA (cicilan)",
            "Meminta VA (pelunasan)",
            "Meminta VA (perpanjangan)",
            "Tidak bisa ambil VA"
        ],
        10: [
            "Meminta keringanan pembayaran (cicilan)",
            "Meminta keringanan pembayaran (potongan denda)",
            "Tidak ada dana"
        ]
    }

    qc_users = User.query.filter_by(role='qc').all()

    return render_template(
        'list_ticket_qc.html',
        nomor_ticket=nomor_ticket,
        tickets=tickets,
        user=current_user,
        jenis_pengaduan_map=jenis_pengaduan_map,
        detail_pengaduan_map=detail_pengaduan_map,
        qc_users=qc_users,
        catatan_map=catatan_map,
        catatan_list=catatan_list
    )

@app.route('/follow-up-pengaduan-qc/<int:nomor_ticket_id>', methods=['POST'])
@login_required
def follow_up_pengaduan_qc(nomor_ticket_id):
    if current_user.role != 'qc':
        return redirect(request.referrer)

    deskripsi_qc = request.form.get("deskripsi_qc")  
    uploaded_files = request.files.getlist("file_qc")  

    existing = request.form.getlist("existing_images")
    deleted = request.form.getlist("deleted_images")

    new_filenames = []
    for file in uploaded_files:
        if file and file.filename:
            filename = secure_filename(file.filename)
            save_path = os.path.join("static/uploads", filename)
            file.save(save_path)
            new_filenames.append(filename)

    final_files = [f for f in existing if f not in deleted] + new_filenames
    joined_filenames = ",".join(final_files)

    nomor_ticket = NomorTicket.query.get_or_404(nomor_ticket_id)
    tickets = Ticket.query.filter_by(nomor_ticket_id=nomor_ticket.id).all()

    for ticket in tickets:
        ticket.deskripsi_qc = deskripsi_qc
        ticket.file_qc = joined_filenames

    db.session.commit()
    flash("Data berhasil diupdate", "success")
    return redirect(url_for("list_ticket_by_nomor_qc", nomor_ticket_id=nomor_ticket_id))

@app.route('/list_user')
@login_required
def list_user():
    if current_user.role != 'admin':
        flash('Akses ditolak: Hanya admin yang bisa melihat daftar user.')
        return redirect(url_for('staff_dashboard'))

    staff_users = User.query.filter_by(role='staff').all()
    return render_template('list_user.html', user=current_user, users=staff_users)

@app.route('/add_user', methods=['POST'])
@login_required
def add_user():
    if current_user.role != 'admin':
        flash('Akses ditolak: Anda bukan admin.')
        return redirect(url_for('list_user'))

    username = request.form['username']
    email = request.form['email']
    phone = request.form['phone']
    password = request.form['password']
    role = request.form.get('role') 
    hashed_pw = generate_password_hash(password)

    if User.query.filter_by(username=username).first():
        flash('Username sudah terdaftar.')
        return redirect(url_for('list_user'))
    if User.query.filter_by(email=email).first():
        flash('Email sudah terdaftar.')
        return redirect(url_for('list_user'))

    user = User(username=username, email=email, phone=phone, password=hashed_pw, role=role)
    db.session.add(user)
    db.session.commit()

    flash(f'User {role} berhasil ditambahkan.')
    return redirect(url_for('admin_dashboard'))

@app.route('/delete_user/<int:user_id>', methods=['POST'])
@login_required
def delete_user(user_id):
    if current_user.role != 'admin':
        flash('Akses ditolak: Anda bukan admin.', 'error')
        return redirect(url_for('list_user'))

    user = User.query.get_or_404(user_id)

    if user.id == current_user.id:
        flash('Anda tidak dapat menghapus akun Anda sendiri.', 'error')
        return redirect(url_for('list_user'))

    db.session.delete(user)
    db.session.commit()
    flash(f'User {user.username} berhasil dihapus.', 'success')
    return redirect(url_for('list_user'))

@app.route('/filtering', methods=['GET'])
@login_required
def filtering():
    os_selected = request.args.getlist('os') or []
    bucket_selected = request.args.getlist('bucket') or []
    range1 = request.args.get('range1', '')
    range2 = request.args.get('range2', '')

    # --- Helper untuk parsing range dari request
    def parse_range(date_range):
        try:
            start, end = date_range.split(' - ')
            return start.strip(), end.strip()
        except ValueError:
            return None, None

    # --- Helper untuk format label range (Bahasa Indonesia)
    def format_range_label(start, end):
        bulan_id = {
            'January': 'Januari', 'February': 'Februari', 'March': 'Maret', 'April': 'April',
            'May': 'Mei', 'June': 'Juni', 'July': 'Juli', 'August': 'Agustus',
            'September': 'September', 'October': 'Oktober', 'November': 'November', 'December': 'Desember'
        }
        try:
            start_dt = datetime.strptime(start, '%Y-%m-%d')
            end_dt = datetime.strptime(end, '%Y-%m-%d')

            start_label = start_dt.strftime('%d %B %Y')
            end_label = end_dt.strftime('%d %B %Y')

            for en, idn in bulan_id.items():
                start_label = start_label.replace(en, idn)
                end_label = end_label.replace(en, idn)

            return f"{start_label} - {end_label}"
        except Exception:
            return "Tidak ada data"

    # --- Ambil range dari request
    range1_start, range1_end = parse_range(range1)
    range2_start, range2_end = parse_range(range2)

    # --- Kalau range1 kosong, fallback pakai min/max dari DB
    if not range1_start or not range1_end:
        min_tanggal = db.session.query(func.min(Ticket.tanggal)).scalar()
        max_tanggal = db.session.query(func.max(Ticket.tanggal)).scalar()

        if min_tanggal and max_tanggal:
            range1_start = min_tanggal.strftime('%Y-%m-%d')
            range1_end = max_tanggal.strftime('%Y-%m-%d')

    # --- Label untuk chart
    label_range1 = format_range_label(range1_start, range1_end) if range1_start and range1_end else "Tidak ada data"
    label_range2 = format_range_label(range2_start, range2_end) if range2_start and range2_end else None

    chart_title = f"Jumlah Perbandingan antar OS ({label_range1}" + (f" | {label_range2}" if label_range2 else "") + ")"

    # --- Query data sesuai filter
    def get_filtered_data(start_date, end_date):
        query = Ticket.query
        query = query.filter(Ticket.nama_os.isnot(None)).filter(Ticket.nama_os != '')

        if os_selected:
            query = query.filter(Ticket.nama_os.in_(os_selected))
        if bucket_selected:
            query = query.filter(Ticket.nama_bucket.in_(bucket_selected))
        if start_date and end_date:
            try:
                start = datetime.strptime(start_date, "%Y-%m-%d")
                end = datetime.strptime(end_date, "%Y-%m-%d")
                query = query.filter(Ticket.tanggal.between(start, end))
            except ValueError:
                pass

        data_grouped = query.with_entities(
            Ticket.nama_os,
            Ticket.nama_bucket,
            func.count(Ticket.id)
        ).group_by(Ticket.nama_os, Ticket.nama_bucket).all()

        os_totals = {}
        os_buckets = {}

        for os, bucket, count in data_grouped:
            if os:
                os_totals[os] = os_totals.get(os, 0) + count
                if bucket: 
                    if os not in os_buckets:
                        os_buckets[os] = []
                    os_buckets[os].append(f"{bucket}: {count}")

        return os_totals, os_buckets

    os_count1, bucket_info1 = get_filtered_data(range1_start, range1_end) if range1_start and range1_end else ({}, {})
    os_count2, bucket_info2 = get_filtered_data(range2_start, range2_end) if range2_start and range2_end else ({}, {})

    # --- Gabung label OS dari kedua range
    chart_labels = sorted(list(set(os_count1.keys()) | set(os_count2.keys())))

    # --- Siapkan data chart
    chart_series = []
    if os_count1:
        chart_series.append({
            "name": label_range1,
            "data": [os_count1.get(os, 0) for os in chart_labels],
            "bucket_info": [bucket_info1.get(os, []) for os in chart_labels]
        })

    if os_count2:
        chart_series.append({
            "name": label_range2,
            "data": [os_count2.get(os, 0) for os in chart_labels],
            "bucket_info": [bucket_info2.get(os, []) for os in chart_labels]
        })

    # --- Data tambahan untuk filter dropdown
    list_os = db.session.query(Ticket.nama_os).distinct().all()
    list_bucket = db.session.query(Ticket.nama_bucket).distinct().all()

    # --- Warna default chart
    default_colors = [
        "#1E90FF", "#28a745", "#ffc107", "#dc3545", "#6f42c1",
        "#20c997", "#fd7e14", "#6610f2", "#17a2b8", "#343a40"
    ]
    color_map = {label: default_colors[i % len(default_colors)] for i, label in enumerate(chart_labels)}
    chart_colors = [color_map[os] for os in chart_labels]

    return render_template(
        'filtering.html',
        user=current_user,
        chart_labels=chart_labels,
        chart_series=chart_series,
        chart_title=chart_title,
        chart_colors=chart_colors,
        list_os=[os[0] for os in list_os if os[0]],
        list_bucket=[b[0] for b in list_bucket if b[0]],
        os_selected=os_selected,
        bucket_selected=bucket_selected,
        range1=range1,
        range2=range2
    )

@app.route('/filtering-kanal', methods=['GET'])
@login_required
def filtering_kanal():
    range1 = request.args.get('range1', '')
    range2 = request.args.get('range2', '')

    def parse_range(date_range):
        try:
            start, end = date_range.split(' - ')
            start_dt = datetime.strptime(start.strip(), '%Y-%m-%d')
            end_dt = datetime.strptime(end.strip(), '%Y-%m-%d') + timedelta(days=1)
            return start_dt, end_dt
        except:
            return None, None

    start1, end1 = parse_range(range1)
    start2, end2 = parse_range(range2)

    kanal_raw = db.session.query(Ticket.kanal_pengaduan)\
        .filter(Ticket.kanal_pengaduan.isnot(None), Ticket.kanal_pengaduan != '')\
        .distinct().all()

    kanal_set = set()
    for k in kanal_raw:
        kanal_normalized = (k[0] or '').strip().lower()
        if kanal_normalized:
            kanal_set.add(kanal_normalized)

    kanal_list = sorted(kanal_set)
    chart_labels = [k.title() for k in kanal_list]

    def get_data_by_range(start, end):
        q = Ticket.query.join(NomorTicket)\
            .filter(Ticket.kanal_pengaduan.isnot(None), Ticket.kanal_pengaduan != '')
        if start and end:
            q = q.filter(Ticket.tanggal >= start, Ticket.tanggal < end)

        result = q.with_entities(
            func.lower(func.trim(Ticket.kanal_pengaduan)).label('kanal'),
            func.count(distinct(Ticket.nomor_ticket_id))
        ).group_by('kanal').all()

        data_dict = {kanal: count for kanal, count in result}
        return [data_dict.get(k, 0) for k in kanal_list]

    chart_series = []

    if not range1 and not range2:
        total_data = get_data_by_range(None, None)
        chart_series = [{
            "name": "Total",
            "data": total_data,
            "bucket_info": [[] for _ in total_data]
        }]
    else:
        data1 = get_data_by_range(start1, end1)
        data2 = get_data_by_range(start2, end2)
        if range1:
            chart_series.append({
                "name": f"Range {range1}",
                "data": data1,
                "bucket_info": [[] for _ in data1]
            })
        if range2:
            chart_series.append({
                "name": f"Range {range2}",
                "data": data2,
                "bucket_info": [[] for _ in data2]
            })

    kanal_colors = [
        "#3081D0", "#FF6768", "#00C49F", "#FFBB28", "#AF7AC5",
        "#2ECC71", "#F39C12", "#E74C3C", "#17A589", "#5D6D7E"
    ]
    default_colors = kanal_colors[:len(kanal_list)]

    return render_template(
        'filtering_kanal.html',
        chart_labels=chart_labels,  
        chart_series=chart_series,
        chart_colors=default_colors,
        range1=range1,
        range2=range2,
        user=current_user
    )

from sqlalchemy import func, distinct

@app.route('/staff_dashboard')
@login_required
def staff_dashboard():
    if current_user.role != 'staff':
        flash('Akses ditolak: Anda bukan staff.')
        return redirect(url_for('admin_dashboard'))

    total_ticket = (
        db.session.query(func.count(distinct(NomorTicket.id)))
        .join(Ticket, NomorTicket.id == Ticket.nomor_ticket_id)
        .filter(Ticket.input_by == current_user.id)
        .scalar()
    )

    open_ticket = (
        db.session.query(func.count(distinct(NomorTicket.id)))
        .join(Ticket, NomorTicket.id == Ticket.nomor_ticket_id)
        .filter(Ticket.input_by == current_user.id, Ticket.status_ticket == "1")
        .scalar()
    )

    in_progress = (
        db.session.query(func.count(distinct(NomorTicket.id)))
        .join(Ticket, NomorTicket.id == Ticket.nomor_ticket_id)
        .filter(Ticket.input_by == current_user.id, Ticket.status_ticket == "2")
        .scalar()
    )

    pending = (
        db.session.query(func.count(distinct(NomorTicket.id)))
        .join(Ticket, NomorTicket.id == Ticket.nomor_ticket_id)
        .filter(Ticket.input_by == current_user.id, Ticket.status_ticket == "3")
        .scalar()
    )
    
    resolved = (
        db.session.query(func.count(distinct(NomorTicket.id)))
        .join(Ticket, NomorTicket.id == Ticket.nomor_ticket_id)
        .filter(Ticket.input_by == current_user.id, Ticket.status_ticket == "5")
        .scalar()
    )

    closed = (
        db.session.query(func.count(distinct(NomorTicket.id)))
        .join(Ticket, NomorTicket.id == Ticket.nomor_ticket_id)
        .filter(Ticket.input_by == current_user.id, Ticket.status_ticket == "4")
        .scalar()
    )

    return render_template(
        'staff_dashboard.html',
        user=current_user,
        total_ticket=total_ticket,
        in_progress=in_progress,
        pending=pending,
        resolved=resolved,
        closed=closed,
        open_ticket=open_ticket
    )

@app.route('/pengaduan')
@login_required
def pengaduan():
    if current_user.role != 'staff':
        return redirect(request.referrer)

    page = request.args.get('page', 1, type=int)
    jenis = request.args.get('jenis')
    status = request.args.get('status')
    tanggal = request.args.get('tanggal')
    q = request.args.get('q')
    tahapan = request.args.get('tahapan')

    tahapan_options = [row[0] for row in db.session.query(Ticket.tahapan).distinct().all() if row[0]]

    nomor_ticket_query = NomorTicket.query\
        .join(Ticket, Ticket.nomor_ticket_id == NomorTicket.id)\
        .filter(
            NomorTicket.id_qc == None,
            NomorTicket.status == 'aktif',
            Ticket.status_ticket == '1',
            Ticket.input_by == current_user.id
        )\
        .distinct()

    if q:
        matching_ids_from_nama = db.session.query(Ticket.nomor_ticket_id)\
            .filter(Ticket.nama_nasabah.ilike(f"%{q}%"))\
            .distinct().all()
        matching_ids_from_nama = [id[0] for id in matching_ids_from_nama]

        matching_ids_from_nomor = db.session.query(NomorTicket.id)\
            .filter(NomorTicket.nomor_ticket.ilike(f"%{q}%"))\
            .all()
        matching_ids_from_nomor = [id[0] for id in matching_ids_from_nomor]

        all_matching_ids = list(set(matching_ids_from_nama + matching_ids_from_nomor))

        nomor_ticket_query = nomor_ticket_query.filter(NomorTicket.id.in_(all_matching_ids))

    tickets_grouped = []

    for nt in nomor_ticket_query.all():
        query = Ticket.query.options(joinedload(Ticket.nomor_ticket))\
            .filter(Ticket.nomor_ticket_id == nt.id, Ticket.sla != 0)

        if jenis:
            query = query.filter_by(jenis_pengaduan=jenis)
        if status:
            query = query.filter_by(status_ticket=status)
        if tanggal:
            try:
                tanggal_obj = datetime.strptime(tanggal, "%Y-%m-%d")
                query = query.filter(func.date(Ticket.tanggal) == tanggal_obj.date())
            except ValueError:
                pass
        if tahapan:
            query = query.filter_by(tahapan=tahapan)

        first_ticket = query.order_by(Ticket.created_time.asc()).first()
        if first_ticket:
            tickets_grouped.append(first_ticket)

    tickets_grouped.sort(
        key=lambda x: x.created_time or datetime.min,
        reverse=True
    )

    per_page = 10
    total = len(tickets_grouped)
    start = (page - 1) * per_page
    end = start + per_page
    paginated_items = tickets_grouped[start:end]
    class Pagination:
        def __init__(self, items, page, per_page, total):
            self.items = items
            self.page = page
            self.per_page = per_page
            self.total = total
            self.pages = (total + per_page - 1) // per_page
            self.has_prev = page > 1
            self.has_next = page < self.pages
            self.prev_num = page - 1
            self.next_num = page + 1

    pagination = Pagination(paginated_items, page, per_page, total)

    count_by_nomor_ticket = dict(
        db.session.query(
            Ticket.nomor_ticket_id,
            func.count(Ticket.id)
        ).group_by(Ticket.nomor_ticket_id).all()
    )

    jumlah_tiket_aktif = db.session.query(NomorTicket)\
        .join(Ticket, Ticket.nomor_ticket_id == NomorTicket.id)\
        .filter(
            NomorTicket.status == 'aktif',
            Ticket.status_ticket == '1',
            Ticket.sla != 0,
            NomorTicket.id_qc == None
        )\
        .distinct()\
        .count()

    return render_template(
        'pengaduan.html',
        user=current_user,
        tickets=pagination,
        count_by_nomor_ticket=count_by_nomor_ticket,
        jumlah_tiket_aktif=jumlah_tiket_aktif,
        tahapan_options=tahapan_options
    )

@app.route('/in-progress')
@login_required
def in_progress():
    if current_user.role != 'staff':
        return redirect(request.referrer)

    page = request.args.get('page', 1, type=int)
    jenis = request.args.get('jenis')
    status = request.args.get('status')
    tanggal = request.args.get('tanggal')
    q = request.args.get('q')
    tahapan = request.args.get('tahapan')

    tahapan_options = [row[0] for row in db.session.query(Ticket.tahapan).distinct().all() if row[0]]

    nomor_ticket_query = NomorTicket.query\
        .join(Ticket, Ticket.nomor_ticket_id == NomorTicket.id)\
        .filter(
            NomorTicket.id_qc == None,
            NomorTicket.status == 'aktif',
            Ticket.status_ticket == '2',
            Ticket.input_by == current_user.id
        )\
        .distinct()

    if q:
        matching_ids_from_nama = db.session.query(Ticket.nomor_ticket_id)\
            .filter(Ticket.nama_nasabah.ilike(f"%{q}%"))\
            .distinct().all()
        matching_ids_from_nama = [id[0] for id in matching_ids_from_nama]

        matching_ids_from_nomor = db.session.query(NomorTicket.id)\
            .filter(NomorTicket.nomor_ticket.ilike(f"%{q}%"))\
            .all()
        matching_ids_from_nomor = [id[0] for id in matching_ids_from_nomor]

        all_matching_ids = list(set(matching_ids_from_nama + matching_ids_from_nomor))

        nomor_ticket_query = nomor_ticket_query.filter(NomorTicket.id.in_(all_matching_ids))

    tickets_grouped = []

    for nt in nomor_ticket_query.all():
        query = Ticket.query.options(joinedload(Ticket.nomor_ticket))\
            .filter(Ticket.nomor_ticket_id == nt.id, Ticket.sla != 0)

        if jenis:
            query = query.filter_by(jenis_pengaduan=jenis)
        if status:
            query = query.filter_by(status_ticket=status)
        if tanggal:
            try:
                tanggal_obj = datetime.strptime(tanggal, "%Y-%m-%d")
                query = query.filter(func.date(Ticket.tanggal) == tanggal_obj.date())
            except ValueError:
                pass
        if tahapan:
            query = query.filter_by(tahapan=tahapan)

        first_ticket = query.order_by(Ticket.created_time.asc()).first()
        if first_ticket:
            tickets_grouped.append(first_ticket)

    tickets_grouped.sort(
        key=lambda x: x.created_time or datetime.min,
        reverse=True
    )

    per_page = 10
    total = len(tickets_grouped)
    start = (page - 1) * per_page
    end = start + per_page
    paginated_items = tickets_grouped[start:end]
    class Pagination:
        def __init__(self, items, page, per_page, total):
            self.items = items
            self.page = page
            self.per_page = per_page
            self.total = total
            self.pages = (total + per_page - 1) // per_page
            self.has_prev = page > 1
            self.has_next = page < self.pages
            self.prev_num = page - 1
            self.next_num = page + 1

    pagination = Pagination(paginated_items, page, per_page, total)

    count_by_nomor_ticket = dict(
        db.session.query(
            Ticket.nomor_ticket_id,
            func.count(Ticket.id)
        ).group_by(Ticket.nomor_ticket_id).all()
    )

    jumlah_tiket_aktif = db.session.query(NomorTicket)\
        .join(Ticket, Ticket.nomor_ticket_id == NomorTicket.id)\
        .filter(
            NomorTicket.status == 'aktif',
            Ticket.status_ticket == '2',
            Ticket.sla != 0,
            NomorTicket.id_qc == None
        )\
        .distinct()\
        .count()

    return render_template(
        'progress.html',
        user=current_user,
        tickets=pagination,
        count_by_nomor_ticket=count_by_nomor_ticket,
        jumlah_tiket_aktif=jumlah_tiket_aktif,
        tahapan_options=tahapan_options
    )

@app.route('/pending')
@login_required
def pending():
    if current_user.role != 'staff':
        return redirect(request.referrer)

    page = request.args.get('page', 1, type=int)
    jenis = request.args.get('jenis')
    status = request.args.get('status')
    tanggal = request.args.get('tanggal')
    q = request.args.get('q')
    tahapan = request.args.get('tahapan')

    tahapan_options = [row[0] for row in db.session.query(Ticket.tahapan).distinct().all() if row[0]]

    nomor_ticket_query = NomorTicket.query\
        .join(Ticket, Ticket.nomor_ticket_id == NomorTicket.id)\
        .filter(
            NomorTicket.id_qc == None,
            NomorTicket.status == 'aktif',
            Ticket.status_ticket == '3',
            Ticket.input_by == current_user.id
        )\
        .distinct()

    if q:
        matching_ids_from_nama = db.session.query(Ticket.nomor_ticket_id)\
            .filter(Ticket.nama_nasabah.ilike(f"%{q}%"))\
            .distinct().all()
        matching_ids_from_nama = [id[0] for id in matching_ids_from_nama]

        matching_ids_from_nomor = db.session.query(NomorTicket.id)\
            .filter(NomorTicket.nomor_ticket.ilike(f"%{q}%"))\
            .all()
        matching_ids_from_nomor = [id[0] for id in matching_ids_from_nomor]

        all_matching_ids = list(set(matching_ids_from_nama + matching_ids_from_nomor))

        nomor_ticket_query = nomor_ticket_query.filter(NomorTicket.id.in_(all_matching_ids))

    tickets_grouped = []

    for nt in nomor_ticket_query.all():
        query = Ticket.query.options(joinedload(Ticket.nomor_ticket))\
            .filter(Ticket.nomor_ticket_id == nt.id, Ticket.sla != 0)

        if jenis:
            query = query.filter_by(jenis_pengaduan=jenis)
        if status:
            query = query.filter_by(status_ticket=status)
        if tanggal:
            try:
                tanggal_obj = datetime.strptime(tanggal, "%Y-%m-%d")
                query = query.filter(func.date(Ticket.tanggal) == tanggal_obj.date())
            except ValueError:
                pass
        if tahapan:
            query = query.filter_by(tahapan=tahapan)

        first_ticket = query.order_by(Ticket.created_time.asc()).first()
        if first_ticket:
            tickets_grouped.append(first_ticket)

    tickets_grouped.sort(
        key=lambda x: x.created_time or datetime.min,
        reverse=True
    )

    per_page = 10
    total = len(tickets_grouped)
    start = (page - 1) * per_page
    end = start + per_page
    paginated_items = tickets_grouped[start:end]
    class Pagination:
        def __init__(self, items, page, per_page, total):
            self.items = items
            self.page = page
            self.per_page = per_page
            self.total = total
            self.pages = (total + per_page - 1) // per_page
            self.has_prev = page > 1
            self.has_next = page < self.pages
            self.prev_num = page - 1
            self.next_num = page + 1

    pagination = Pagination(paginated_items, page, per_page, total)

    count_by_nomor_ticket = dict(
        db.session.query(
            Ticket.nomor_ticket_id,
            func.count(Ticket.id)
        ).group_by(Ticket.nomor_ticket_id).all()
    )

    jumlah_tiket_aktif = db.session.query(NomorTicket)\
        .join(Ticket, Ticket.nomor_ticket_id == NomorTicket.id)\
        .filter(
            NomorTicket.status == 'aktif',
            Ticket.status_ticket == '3',
            Ticket.sla != 0,
            NomorTicket.id_qc == None
        )\
        .distinct()\
        .count()

    return render_template(
        'pending.html',
        user=current_user,
        tickets=pagination,
        count_by_nomor_ticket=count_by_nomor_ticket,
        jumlah_tiket_aktif=jumlah_tiket_aktif,
        tahapan_options=tahapan_options
    )

@app.route('/resolved')
@login_required
def resolved():
    if current_user.role != 'staff':
        return redirect(request.referrer)

    page = request.args.get('page', 1, type=int)
    jenis = request.args.get('jenis')
    status = request.args.get('status')
    tanggal = request.args.get('tanggal')
    q = request.args.get('q')
    tahapan = request.args.get('tahapan')

    tahapan_options = [row[0] for row in db.session.query(Ticket.tahapan).distinct().all() if row[0]]

    nomor_ticket_query = NomorTicket.query\
        .join(Ticket, Ticket.nomor_ticket_id == NomorTicket.id)\
        .filter(
            NomorTicket.id_qc == None,
            NomorTicket.status == 'aktif',
            Ticket.status_ticket == '5',
            Ticket.input_by == current_user.id
        )\
        .distinct()

    if q:
        matching_ids_from_nama = db.session.query(Ticket.nomor_ticket_id)\
            .filter(Ticket.nama_nasabah.ilike(f"%{q}%"))\
            .distinct().all()
        matching_ids_from_nama = [id[0] for id in matching_ids_from_nama]

        matching_ids_from_nomor = db.session.query(NomorTicket.id)\
            .filter(NomorTicket.nomor_ticket.ilike(f"%{q}%"))\
            .all()
        matching_ids_from_nomor = [id[0] for id in matching_ids_from_nomor]

        all_matching_ids = list(set(matching_ids_from_nama + matching_ids_from_nomor))

        nomor_ticket_query = nomor_ticket_query.filter(NomorTicket.id.in_(all_matching_ids))

    tickets_grouped = []

    for nt in nomor_ticket_query.all():
        query = Ticket.query.options(joinedload(Ticket.nomor_ticket))\
            .filter(Ticket.nomor_ticket_id == nt.id, Ticket.sla != 0)

        if jenis:
            query = query.filter_by(jenis_pengaduan=jenis)
        if status:
            query = query.filter_by(status_ticket=status)
        if tanggal:
            try:
                tanggal_obj = datetime.strptime(tanggal, "%Y-%m-%d")
                query = query.filter(func.date(Ticket.tanggal) == tanggal_obj.date())
            except ValueError:
                pass
        if tahapan:
            query = query.filter_by(tahapan=tahapan)

        first_ticket = query.order_by(Ticket.created_time.asc()).first()
        if first_ticket:
            tickets_grouped.append(first_ticket)

    tickets_grouped.sort(
        key=lambda x: x.created_time or datetime.min,
        reverse=True
    )

    per_page = 10
    total = len(tickets_grouped)
    start = (page - 1) * per_page
    end = start + per_page
    paginated_items = tickets_grouped[start:end]
    class Pagination:
        def __init__(self, items, page, per_page, total):
            self.items = items
            self.page = page
            self.per_page = per_page
            self.total = total
            self.pages = (total + per_page - 1) // per_page
            self.has_prev = page > 1
            self.has_next = page < self.pages
            self.prev_num = page - 1
            self.next_num = page + 1

    pagination = Pagination(paginated_items, page, per_page, total)

    count_by_nomor_ticket = dict(
        db.session.query(
            Ticket.nomor_ticket_id,
            func.count(Ticket.id)
        ).group_by(Ticket.nomor_ticket_id).all()
    )

    jumlah_tiket_aktif = db.session.query(NomorTicket)\
        .join(Ticket, Ticket.nomor_ticket_id == NomorTicket.id)\
        .filter(
            NomorTicket.status == 'aktif',
            Ticket.status_ticket == '5',
            Ticket.sla != 0,
            NomorTicket.id_qc == None
        )\
        .distinct()\
        .count()

    return render_template(
        'resolved.html',
        user=current_user,
        tickets=pagination,
        count_by_nomor_ticket=count_by_nomor_ticket,
        jumlah_tiket_aktif=jumlah_tiket_aktif,
        tahapan_options=tahapan_options
    )

@app.route('/closed')
@login_required
def closed():
    if current_user.role != 'staff':
        return redirect(request.referrer)

    page = request.args.get('page', 1, type=int)
    jenis = request.args.get('jenis')
    status = request.args.get('status')
    tanggal = request.args.get('tanggal')
    tanggal_tutup = request.args.get('tanggal_tutup')
    q = request.args.get('q')
    tahapan = request.args.get('tahapan')

    tahapan_options = [row[0] for row in db.session.query(Ticket.tahapan).distinct().all() if row[0]]

    nomor_ticket_query = NomorTicket.query\
        .join(Ticket, Ticket.nomor_ticket_id == NomorTicket.id)\
        .filter(
            NomorTicket.id_qc == None,
            NomorTicket.status == 'close',
            Ticket.status_ticket == '4',
            Ticket.input_by == current_user.id
        )\
        .distinct()

    if q:
        matching_ids_from_nama = db.session.query(Ticket.nomor_ticket_id)\
            .filter(Ticket.nama_nasabah.ilike(f"%{q}%"))\
            .distinct().all()
        matching_ids_from_nama = [id[0] for id in matching_ids_from_nama]

        matching_ids_from_nomor = db.session.query(NomorTicket.id)\
            .filter(NomorTicket.nomor_ticket.ilike(f"%{q}%"))\
            .all()
        matching_ids_from_nomor = [id[0] for id in matching_ids_from_nomor]

        all_matching_ids = list(set(matching_ids_from_nama + matching_ids_from_nomor))

        nomor_ticket_query = nomor_ticket_query.filter(NomorTicket.id.in_(all_matching_ids))

    tickets_grouped = []

    for nt in nomor_ticket_query.all():
        query = Ticket.query.options(joinedload(Ticket.nomor_ticket))\
            .join(NomorTicket, Ticket.nomor_ticket_id == NomorTicket.id)\
            .filter(Ticket.nomor_ticket_id == nt.id, Ticket.sla != 0)

        if jenis:
            query = query.filter_by(jenis_pengaduan=jenis)
        if status:
            query = query.filter_by(status_ticket=status)
        if tanggal:
            try:
                tanggal_obj = datetime.strptime(tanggal, "%Y-%m-%d")
                query = query.filter(func.date(Ticket.tanggal) == tanggal_obj.date())
            except ValueError:
                pass
        if tanggal_tutup:
            try:
                tanggal_tutup_obj = datetime.strptime(tanggal_tutup, "%Y-%m-%d")
                query = query.filter(func.date(NomorTicket.closed_ticket) == tanggal_tutup_obj.date())
            except ValueError:
                pass
        if tahapan:
            query = query.filter_by(tahapan=tahapan)

        first_ticket = query.order_by(Ticket.created_time.asc()).first()
        if first_ticket:
            tickets_grouped.append(first_ticket)

    tickets_grouped.sort(
        key=lambda x: x.created_time or datetime.min,
        reverse=True
    )

    per_page = 10
    total = len(tickets_grouped)
    start = (page - 1) * per_page
    end = start + per_page
    paginated_items = tickets_grouped[start:end]
    class Pagination:
        def __init__(self, items, page, per_page, total):
            self.items = items
            self.page = page
            self.per_page = per_page
            self.total = total
            self.pages = (total + per_page - 1) // per_page
            self.has_prev = page > 1
            self.has_next = page < self.pages
            self.prev_num = page - 1
            self.next_num = page + 1

    pagination = Pagination(paginated_items, page, per_page, total)

    count_by_nomor_ticket = dict(
        db.session.query(
            Ticket.nomor_ticket_id,
            func.count(Ticket.id)
        ).group_by(Ticket.nomor_ticket_id).all()
    )

    jumlah_tiket_aktif = db.session.query(NomorTicket)\
        .join(Ticket, Ticket.nomor_ticket_id == NomorTicket.id)\
        .filter(
            NomorTicket.status == 'close',
            Ticket.status_ticket == '4',
            Ticket.sla != 0,
            NomorTicket.id_qc == None
        )\
        .distinct()\
        .count()

    return render_template(
        'closed.html',
        user=current_user,
        tickets=pagination,
        count_by_nomor_ticket=count_by_nomor_ticket,
        jumlah_tiket_aktif=jumlah_tiket_aktif,
        tahapan_options=tahapan_options
    )

@app.route('/export-ticket-excel')
@login_required
def export_ticket_excel():
    if current_user.role != 'admin':
        return redirect(request.referrer)

    date_range = request.args.get('date', '') 

    try:
        start_date_str, end_date_str = date_range.split(' - ')
        start_date = datetime.strptime(start_date_str.strip(), "%Y-%m-%d")
        end_date = datetime.strptime(end_date_str.strip(), "%Y-%m-%d")
    except ValueError:
        flash('Format tanggal tidak valid. Gunakan format: YYYY-MM-DD - YYYY-MM-DD', 'danger')
        return redirect(url_for('pengaduan'))

    tickets = Ticket.query \
        .filter(Ticket.tanggal >= start_date, Ticket.tanggal <= end_date) \
        .order_by(Ticket.tanggal.desc()).all()

    if not tickets:
        flash('Tidak ada data ticket pada rentang tanggal tersebut.', 'warning')
        return redirect(url_for('pengaduan'))

    status_ticket_map = {
        '1': 'Aktif',
        '2': 'Perpanjangan',
        '3': 'Keberatan',
        '4': 'Tutup',
        '5': 'Reopen'
    }

    jenis_pengaduan_map = {
        '1': "Informasi Pengajuan",
        '2': "Permintaan Kode OTP",
        '3': "Informasi Tenor",
        '4': "Informasi Tagihan",
        '5': "Informasi Denda",
        '6': "Pembatalan Pinjaman",
        '7': "Informasi Pencairan Dana",
        '8': "Perilaku Petugas Penagihan",
        '9': "Informasi Pembayaran",
        '10': "Discount / Pemutihan"
    }

    data = []
    for t in tickets:
        status_label = status_ticket_map.get(str(t.status_ticket), t.status_ticket)
        jenis_label = jenis_pengaduan_map.get(str(t.jenis_pengaduan), t.jenis_pengaduan)

        file_links = ''
        if t.bukti_chat:
            filenames = [f.strip() for f in t.bukti_chat.split(',') if f.strip()]
            base_url = request.host_url.rstrip('/') + '/static/uploads'
            file_links = ', '.join([f"{base_url}/{filename}" for filename in filenames])
        
        data.append({
            "Channel": t.kanal_pengaduan,
            "Tanggal": t.tanggal.strftime('%Y-%m-%d') if t.tanggal else '',
            "No Ticket": t.nomor_ticket.nomor_ticket if t.nomor_ticket else '',
            "Order No": t.order_no,
            "Name": t.nama_nasabah,
            "Customer Phone Number": t.nomor_utama,
            "Email": t.email,
            "NIK": t.nik,
            "Detail Problem": t.detail_pengaduan,
            "Tipe Pengaduan": jenis_label,
            "Detail Pengaduan": t.detail_pengaduan,
            "Deskripsi Pengaduan": t.deskripsi_pengaduan,
            "Status Ticket": status_label,
            "DC": t.nama_dc,
            "OS": t.nama_os,
            "Bucket": t.nama_bucket,
            "Screenshoot Chat": file_links, 
        })

    df = pd.DataFrame(data)

    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Tickets')

    output.seek(0)
    return send_file(
        output,
        as_attachment=True,
        download_name=f"export_tickets_{start_date_str}_to_{end_date_str}.xlsx",
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

@app.route('/nomor-ticket/<int:nomor_ticket_id>')
@login_required
def list_ticket_by_nomor(nomor_ticket_id):
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    nomor_ticket = NomorTicket.query.get_or_404(nomor_ticket_id)

    tickets = Ticket.query.filter_by(nomor_ticket_id=nomor_ticket_id)\
        .order_by(Ticket.created_time.asc()).all()
    
    catatan_list = Catatan.query.filter_by(nomor_ticket_id=nomor_ticket_id)\
        .order_by(Catatan.tanggal.desc()).all()

    qc_users = User.query.filter_by(role='qc').all()

    jenis_pengaduan_map = {
        1: "Informasi Pengajuan",
        2: "Permintaan Kode OTP",
        3: "Informasi Tenor",
        4: "Informasi Tagihan",
        5: "Informasi Denda",
        6: "Pembatalan Pinjaman",
        7: "Informasi Pencairan Dana",
        8: "Perilaku Petugas Penagihan",
        9: "Informasi Pembayaran",
        10: "Discount / Pemutihan"
    }

    detail_pengaduan_map = {
        1: [
            "Hasil Pengajuan",
            "Pengajuan Ditolak",
            "Status Pengajuan sedang ditransfer",
            "Tidak bisa pengajuan ulang karena keterlambatan",
            "Verifikasi Bank gagal",
            "Verifikasi KTP gagal",
            "Cara pengajuan",
            "Perubahan Nomor Handphone",
            "Perubahan Nomor Rekening"
        ],
        2: [
            "OTP Limit",
            "Tidak terima SMS OTP"
        ],
        3: [
            "Informasi Pinjaman"
        ],
        4: [
            "Konsultasi detail pinjaman saat ini",
            "Konsultasi Perpanjangan",
            "Bukti Transfer"
        ],
        5: [
            "Denda Keterlambatan"
        ],
        6: [
            "Hapus Data (Penutupan Akun)",
            "Pembatalan Pinjaman"
        ],
        7: [
            "Pencairan dana berhasil",
            "Status pencairan dana gagal",
            "Status pengajuan pencairan dana ulang",
            "Tidak terima dana",
            "Operasi Gagal (tidak bisa verifikasi wajah dan KTP)"
        ],
        8: [
            "Keluhan Penagihan",
            "Keluhan Reminder",
            "Penipuan"
        ],
        9: [
            "Konfirmasi Pembayaran",
            "Pembayaran belum masuk",
            "Pembayaran bukan ke VA UATAS",
            "Refund (pembayaran double)",
            "Meminta VA (cicilan)",
            "Meminta VA (pelunasan)",
            "Meminta VA (perpanjangan)",
            "Tidak bisa ambil VA"
        ],
        10: [
            "Meminta keringanan pembayaran (cicilan)",
            "Meminta keringanan pembayaran (potongan denda)",
            "Tidak ada dana"
        ]
    }
    qc_users = User.query.filter_by(role='qc').all()

    return render_template(
        'list_ticket_by_nomor.html',
        nomor_ticket=nomor_ticket,
        tickets=tickets,
        user=current_user,
        jenis_pengaduan_map=jenis_pengaduan_map,
        detail_pengaduan_map=detail_pengaduan_map,
        qc_users=qc_users,
        catatan_list=catatan_list
    )

@app.route('/nomor-ticket-closed/<int:nomor_ticket_id>')
@login_required
def ticket_closed(nomor_ticket_id):
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    nomor_ticket = NomorTicket.query.get_or_404(nomor_ticket_id)

    tickets = Ticket.query.filter_by(nomor_ticket_id=nomor_ticket_id)\
        .order_by(Ticket.created_time.asc()).all()
    
    catatan_list = Catatan.query.filter_by(nomor_ticket_id=nomor_ticket_id)\
        .order_by(Catatan.tanggal.desc()).all()

    qc_users = User.query.filter_by(role='qc').all()

    jenis_pengaduan_map = {
        1: "Informasi Pengajuan",
        2: "Permintaan Kode OTP",
        3: "Informasi Tenor",
        4: "Informasi Tagihan",
        5: "Informasi Denda",
        6: "Pembatalan Pinjaman",
        7: "Informasi Pencairan Dana",
        8: "Perilaku Petugas Penagihan",
        9: "Informasi Pembayaran",
        10: "Discount / Pemutihan"
    }

    detail_pengaduan_map = {
        1: [
            "Hasil Pengajuan",
            "Pengajuan Ditolak",
            "Status Pengajuan sedang ditransfer",
            "Tidak bisa pengajuan ulang karena keterlambatan",
            "Verifikasi Bank gagal",
            "Verifikasi KTP gagal",
            "Cara pengajuan",
            "Perubahan Nomor Handphone",
            "Perubahan Nomor Rekening"
        ],
        2: [
            "OTP Limit",
            "Tidak terima SMS OTP"
        ],
        3: [
            "Informasi Pinjaman"
        ],
        4: [
            "Konsultasi detail pinjaman saat ini",
            "Konsultasi Perpanjangan",
            "Bukti Transfer"
        ],
        5: [
            "Denda Keterlambatan"
        ],
        6: [
            "Hapus Data (Penutupan Akun)",
            "Pembatalan Pinjaman"
        ],
        7: [
            "Pencairan dana berhasil",
            "Status pencairan dana gagal",
            "Status pengajuan pencairan dana ulang",
            "Tidak terima dana",
            "Operasi Gagal (tidak bisa verifikasi wajah dan KTP)"
        ],
        8: [
            "Keluhan Penagihan",
            "Keluhan Reminder",
            "Penipuan"
        ],
        9: [
            "Konfirmasi Pembayaran",
            "Pembayaran belum masuk",
            "Pembayaran bukan ke VA UATAS",
            "Refund (pembayaran double)",
            "Meminta VA (cicilan)",
            "Meminta VA (pelunasan)",
            "Meminta VA (perpanjangan)",
            "Tidak bisa ambil VA"
        ],
        10: [
            "Meminta keringanan pembayaran (cicilan)",
            "Meminta keringanan pembayaran (potongan denda)",
            "Tidak ada dana"
        ]
    }
    qc_users = User.query.filter_by(role='qc').all()

    return render_template(
        'ticket_closed.html',
        nomor_ticket=nomor_ticket,
        tickets=tickets,
        user=current_user,
        jenis_pengaduan_map=jenis_pengaduan_map,
        detail_pengaduan_map=detail_pengaduan_map,
        qc_users=qc_users,
        catatan_list=catatan_list
    )

@app.route('/nomor-ticket/<int:nomor_ticket_id>/update-status/<int:new_status>', methods=['POST'])
@login_required
def update_status_nomor_ticket(nomor_ticket_id, new_status):
    if current_user.role != 'staff':
        return redirect(request.referrer)

    nomor_ticket = NomorTicket.query.get_or_404(nomor_ticket_id)

    if new_status not in [2, 3, 5]:
        flash("Status tidak valid!", "danger")
        return redirect(url_for('list_ticket_by_nomor', nomor_ticket_id=nomor_ticket_id))

    Ticket.query.filter_by(nomor_ticket_id=nomor_ticket_id).update(
        {"status_ticket": str(new_status)}
    )
    db.session.commit()

    flash("Status semua tiket berhasil diperbarui.", "success")
    return redirect(url_for('list_ticket_by_nomor', nomor_ticket_id=nomor_ticket_id))

@app.route('/eskalasi-ticket-qc/<int:nomor_ticket_id>')
@login_required
def eskalasi_ticket_qc(nomor_ticket_id):
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    nomor_ticket = NomorTicket.query.get_or_404(nomor_ticket_id)

    tickets = Ticket.query.filter_by(nomor_ticket_id=nomor_ticket_id)\
        .order_by(Ticket.created_time.asc()).all()
    
    semua_kosong = all(ticket.deskripsi_qc in [None, ""] for ticket in tickets)

    if semua_kosong:
        status_feedback = "Belum ada Feedback"
        badge_feedback = "warning"
    else:
        status_feedback = "Ada Feedback"
        badge_feedback = "success"

    catatan_list = Catatan.query.filter_by(nomor_ticket_id=nomor_ticket_id)\
        .order_by(Catatan.tanggal.desc()).all()

    qc_users = User.query.filter_by(role='qc').all()

    jenis_pengaduan_map = {
        1: "Informasi Pengajuan",
        2: "Permintaan Kode OTP",
        3: "Informasi Tenor",
        4: "Informasi Tagihan",
        5: "Informasi Denda",
        6: "Pembatalan Pinjaman",
        7: "Informasi Pencairan Dana",
        8: "Perilaku Petugas Penagihan",
        9: "Informasi Pembayaran",
        10: "Discount / Pemutihan"
    }

    detail_pengaduan_map = {
        1: [
            "Hasil Pengajuan",
            "Pengajuan Ditolak",
            "Status Pengajuan sedang ditransfer",
            "Tidak bisa pengajuan ulang karena keterlambatan",
            "Verifikasi Bank gagal",
            "Verifikasi KTP gagal",
            "Cara pengajuan",
            "Perubahan Nomor Handphone",
            "Perubahan Nomor Rekening"
        ],
        2: [
            "OTP Limit",
            "Tidak terima SMS OTP"
        ],
        3: [
            "Informasi Pinjaman"
        ],
        4: [
            "Konsultasi detail pinjaman saat ini",
            "Konsultasi Perpanjangan",
            "Bukti Transfer"
        ],
        5: [
            "Denda Keterlambatan"
        ],
        6: [
            "Hapus Data (Penutupan Akun)",
            "Pembatalan Pinjaman"
        ],
        7: [
            "Pencairan dana berhasil",
            "Status pencairan dana gagal",
            "Status pengajuan pencairan dana ulang",
            "Tidak terima dana",
            "Operasi Gagal (tidak bisa verifikasi wajah dan KTP)"
        ],
        8: [
            "Keluhan Penagihan",
            "Keluhan Reminder",
            "Penipuan"
        ],
        9: [
            "Konfirmasi Pembayaran",
            "Pembayaran belum masuk",
            "Pembayaran bukan ke VA UATAS",
            "Refund (pembayaran double)",
            "Meminta VA (cicilan)",
            "Meminta VA (pelunasan)",
            "Meminta VA (perpanjangan)",
            "Tidak bisa ambil VA"
        ],
        10: [
            "Meminta keringanan pembayaran (cicilan)",
            "Meminta keringanan pembayaran (potongan denda)",
            "Tidak ada dana"
        ]
    }
    qc_users = User.query.filter_by(role='qc').all()

    return render_template(
        'hasil_eskalasi.html',
        nomor_ticket=nomor_ticket,
        tickets=tickets,
        user=current_user,
        jenis_pengaduan_map=jenis_pengaduan_map,
        detail_pengaduan_map=detail_pengaduan_map,
        qc_users=qc_users,
        catatan_list=catatan_list
    )

@app.route('/follow-up-pengaduan/<int:nomor_ticket_id>', methods=['POST'])
@login_required
def follow_up_pengaduan(nomor_ticket_id):
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    jenis_pengaduan = request.form.get("jenis_pengaduan")
    detail_pengaduan = request.form.get("detail_pengaduan")
    kronologis = request.form.get("kronologis")
    uploaded_files = request.files.getlist("bukti_chat")

    existing = request.form.getlist("existing_images")
    deleted = request.form.getlist("deleted_images")

    new_filenames = []
    for file in uploaded_files:
        if file and file.filename:
            filename = secure_filename(file.filename)
            save_path = os.path.join("static/uploads", filename)
            file.save(save_path)
            new_filenames.append(filename)

    final_files = [f for f in existing if f not in deleted] + new_filenames
    joined_filenames = ",".join(final_files)

    nomor_ticket = NomorTicket.query.get_or_404(nomor_ticket_id)
    tickets = Ticket.query.filter_by(nomor_ticket_id=nomor_ticket.id).all()

    for ticket in tickets:
        ticket.jenis_pengaduan = jenis_pengaduan
        ticket.detail_pengaduan = detail_pengaduan
        ticket.kronologis = kronologis
        ticket.bukti_chat = joined_filenames

    db.session.commit()
    flash("Data berhasil diupdate", "success")
    return redirect(url_for("list_ticket_by_nomor", nomor_ticket_id=nomor_ticket_id))

@app.route('/follow-up-pengaduan-reopen/<int:nomor_ticket_id>', methods=['POST'])
@login_required
def follow_up_pengaduan_reopen(nomor_ticket_id):
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    jenis_pengaduan = request.form.get("jenis_pengaduan")
    detail_pengaduan = request.form.get("detail_pengaduan")
    kronologis = request.form.get("kronologis")
    uploaded_files = request.files.getlist("bukti_chat")

    existing = request.form.getlist("existing_images")
    deleted = request.form.getlist("deleted_images")

    new_filenames = []
    for file in uploaded_files:
        if file and file.filename:
            filename = secure_filename(file.filename)
            save_path = os.path.join("static/uploads", filename)
            file.save(save_path)
            new_filenames.append(filename)

    final_files = [f for f in existing if f not in deleted] + new_filenames
    joined_filenames = ",".join(final_files)

    nomor_ticket = NomorTicket.query.get_or_404(nomor_ticket_id)
    tickets = Ticket.query.filter_by(nomor_ticket_id=nomor_ticket.id).all()

    for ticket in tickets:
        ticket.jenis_pengaduan = jenis_pengaduan
        ticket.detail_pengaduan = detail_pengaduan
        ticket.kronologis = kronologis
        ticket.bukti_chat = joined_filenames

    db.session.commit()
    flash("Data berhasil diupdate", "success")
    return redirect(url_for("list_reopen_ticket", nomor_ticket_id=nomor_ticket_id))

@app.route('/add-order/<int:ticket_id>', methods=['POST'])
@login_required
def add_order(ticket_id):
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    original_ticket = Ticket.query.get_or_404(ticket_id)

    order_no = request.form.get('order_no')
    nama_os = request.form.get('nama_os')
    nama_dc = request.form.get('nama_dc')
    nama_bucket = request.form.get('nama_bucket')
    deskripsi_pengaduan = request.form.get('deskripsi_pengaduan')
    tanggal_str = request.form.get('tanggal')

    if not deskripsi_pengaduan or not tanggal_str:
        flash('Deskripsi pengaduan dan tanggal wajib diisi.', 'danger')
        return redirect(url_for('list_ticket_by_nomor', nomor_ticket_id=original_ticket.nomor_ticket_id))

    print("Tanggal dari form:", tanggal_str)

    try:
        tanggal = datetime.strptime(tanggal_str, '%Y-%m-%d')
    except Exception as e:
        flash(f'Tanggal tidak valid. {str(e)}', 'danger')
        return redirect(url_for('list_ticket_by_nomor', nomor_ticket_id=original_ticket.nomor_ticket_id))

    new_ticket = Ticket(
        order_no=order_no,
        nama_os=nama_os,
        nama_dc=nama_dc,
        nama_bucket=nama_bucket,
        deskripsi_pengaduan=deskripsi_pengaduan,
        tanggal=tanggal,  

        kanal_pengaduan=original_ticket.kanal_pengaduan,
        kategori_pengaduan=original_ticket.kategori_pengaduan,
        jenis_pengaduan=original_ticket.jenis_pengaduan,
        detail_pengaduan=original_ticket.detail_pengaduan,
        nama_nasabah=original_ticket.nama_nasabah,
        email=original_ticket.email,
        nomor_utama=original_ticket.nomor_utama,
        nomor_kontak=original_ticket.nomor_kontak,
        nik=original_ticket.nik,

        input_by=current_user.id,
        sla=10,
        status_ticket='1',
        nomor_ticket_id=original_ticket.nomor_ticket_id
    )

    db.session.add(new_ticket)
    db.session.commit()

    flash('Order berhasil ditambahkan.', 'success')
    return redirect(url_for('list_ticket_by_nomor', nomor_ticket_id=original_ticket.nomor_ticket_id))

@app.route('/add-order-reopen/<int:ticket_id>', methods=['POST'])
@login_required
def add_order_reopen(ticket_id):
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    original_ticket = Ticket.query.get_or_404(ticket_id)

    order_no = request.form.get('order_no')
    nama_os = request.form.get('nama_os')
    nama_dc = request.form.get('nama_dc')
    nama_bucket = request.form.get('nama_bucket')
    deskripsi_pengaduan = request.form.get('deskripsi_pengaduan')
    tanggal_str = request.form.get('tanggal')

    if not deskripsi_pengaduan or not tanggal_str:
        flash('Deskripsi pengaduan dan tanggal wajib diisi.', 'danger')
        return redirect(url_for('list_ticket_by_nomor', nomor_ticket_id=original_ticket.nomor_ticket_id))

    print("Tanggal dari form:", tanggal_str)

    try:
        tanggal = datetime.strptime(tanggal_str, '%Y-%m-%d')
    except Exception as e:
        flash(f'Tanggal tidak valid. {str(e)}', 'danger')
        return redirect(url_for('list_ticket_by_nomor', nomor_ticket_id=original_ticket.nomor_ticket_id))

    new_ticket = Ticket(
        order_no=order_no,
        nama_os=nama_os,
        nama_dc=nama_dc,
        nama_bucket=nama_bucket,
        deskripsi_pengaduan=deskripsi_pengaduan,
        tanggal=tanggal,  

        kanal_pengaduan=original_ticket.kanal_pengaduan,
        kategori_pengaduan=original_ticket.kategori_pengaduan,
        jenis_pengaduan=original_ticket.jenis_pengaduan,
        detail_pengaduan=original_ticket.detail_pengaduan,
        nama_nasabah=original_ticket.nama_nasabah,
        email=original_ticket.email,
        nomor_utama=original_ticket.nomor_utama,
        nomor_kontak=original_ticket.nomor_kontak,
        nik=original_ticket.nik,

        input_by=current_user.id,
        sla=10,
        status_ticket='5',
        nomor_ticket_id=original_ticket.nomor_ticket_id
    )

    db.session.add(new_ticket)
    db.session.commit()

    flash('Order berhasil ditambahkan.', 'success')
    return redirect(url_for('list_reopen_ticket', nomor_ticket_id=original_ticket.nomor_ticket_id))

@app.route('/add-kontak/<int:ticket_id>', methods=['POST'])
@login_required
def add_kontak(ticket_id):
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    ticket = Ticket.query.get_or_404(ticket_id)

    nama_lengkap = request.form.get('nama_lengkap')
    nik = request.form.get('nik')
    phone = request.form.get('phone')
    phone_2 = request.form.get('phone_2')
    email = request.form.get('email')

    if not all([nama_lengkap, nik, phone]):
        flash('Field wajib diisi: Nama, NIK, dan No HP.', 'danger')
        return redirect(url_for('list_ticket_by_nomor', nomor_ticket_id=ticket.nomor_ticket_id))

    kontak = Kontak(
        nama_lengkap=nama_lengkap,
        nik=nik,
        phone=phone,
        phone_2=phone_2,
        email=email,
        id_ticket=ticket.id
    )

    db.session.add(kontak)
    db.session.commit()

    flash('Kontak berhasil ditambahkan.', 'success')
    return redirect(url_for('list_ticket_by_nomor', nomor_ticket_id=ticket.nomor_ticket_id))

def generate_nomor_ticket():
    today_str = datetime.now(JAKARTA_TZ).strftime("%d%m%y")

    prefix = f"AN{today_str}"

    last_ticket = (
        NomorTicket.query
        .filter(NomorTicket.nomor_ticket.like(f"{prefix}%"))
        .order_by(NomorTicket.nomor_ticket.desc())
        .first()
    )

    if last_ticket:
        try:
            last_number = int(last_ticket.nomor_ticket[len(prefix):]) 
        except ValueError:
            last_number = 0
        new_number = last_number + 1
    else:
        new_number = 1

    return f"{prefix}{new_number:02d}"

@app.route('/submit-ticket', methods=['POST'])
@login_required
def submit_ticket():
    if current_user.role != 'staff':
        return redirect(request.referrer)

    try:
        nomor_ticket_str = generate_nomor_ticket()

        nomor_ticket_obj = NomorTicket(nomor_ticket=nomor_ticket_str)
        db.session.add(nomor_ticket_obj)
        db.session.flush() 

        ticket = Ticket(
            kanal_pengaduan=request.form.get('country'),
            kategori_pengaduan=request.form.get('kategori'),
            jenis_pengaduan=request.form.get('jenis'),
            detail_pengaduan=request.form.get('detail_pengaduan'),
            tanggal=datetime.strptime(request.form.get('tanggal'), "%Y-%m-%d") if request.form.get('tanggal') else datetime.utcnow(),
            nama_nasabah=request.form.get('nama_nasabah'),
            email=request.form.get('email'),
            nomor_utama=request.form.get('nomor_utama'),
            nomor_kontak=request.form.get('nomor_kontak'),
            nik=request.form.get('nik'),
            nama_os=(request.form.get('nama_os') or "").replace(" ", ""),
            nama_dc=request.form.get('nama_dc'),
            nama_bucket=(request.form.get('nama_bucket') or "").replace(" ", ""),
            order_no=request.form.get('order_no'),
            deskripsi_pengaduan=request.form.get('deskripsi_pengaduan'),
            input_by=current_user.id,
            sla=10,
            status_ticket='1',
            nomor_ticket=nomor_ticket_obj,
            created_time=datetime.utcnow()
        )

        db.session.add(ticket)
        db.session.commit()

        flash(f'Ticket {nomor_ticket_str} berhasil ditambahkan!', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'Notice! Terjadi kesalahan: {e}', 'danger')

    return redirect(url_for('pengaduan'))

@app.route('/update-tahapan/<int:nomor_ticket_id>/<int:ticket_id>', methods=['POST'])
@login_required
def update_tahapan(nomor_ticket_id, ticket_id):
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    id_qc = request.form.get('id_qc')
    
    tiket = Ticket.query.get_or_404(ticket_id)

    tahapan = request.form.get('tahapan')
    nama_os = request.form.get('nama_os').strip() if request.form.get('nama_os') else None
    nama_bucket = request.form.get('nama_bucket').strip() if request.form.get('nama_bucket') else None
    nama_dc = request.form.get('nama_dc')
    nama_nasabah = request.form.get('nama_nasabah')
    nik = request.form.get('nik')
    nomor_utama = request.form.get('nomor_utama')
    nomor_kontak = request.form.get('nomor_kontak')
    email = request.form.get('email')
    deskripsi_pengaduan = request.form.get('deskripsi_pengaduan')
    order_no = request.form.get('order_no')

    tahapan_2 = None
    if tahapan == 'Follow Up':
        followup = request.form.get('tahapan_2_followup')
        tahapan_2 = followup if followup else None
    elif tahapan == 'Eskalasi QC':
        date = request.form.get('tahapan_2_date')
        desc = request.form.get('tahapan_2_desc')
        tahapan_2 = f"{date} - {desc}" if date and desc else None

    is_updating_tahapan = bool(tahapan or tahapan_2)

    if is_updating_tahapan:
        if tahapan: 
            tiket.tahapan = tahapan
        tiket.tahapan_2 = tahapan_2

        if tahapan == "Eskalasi ke QC" and id_qc:
            tiket.nomor_ticket.id_qc = int(id_qc)

        new_history = History(
            nomor_ticket=tiket.nomor_ticket.nomor_ticket,
            order_number=tiket.order_no,
            status_ticket=tiket.status_ticket,
            tahapan=tahapan if tahapan else tiket.tahapan, 
            create_by=current_user.id,
            nama_os=nama_os
        )
        db.session.add(new_history)

    tiket.nama_os = nama_os
    tiket.nama_bucket = nama_bucket
    tiket.nama_dc = nama_dc
    tiket.nama_nasabah = nama_nasabah
    tiket.nik = nik
    tiket.nomor_utama = nomor_utama
    tiket.nomor_kontak = nomor_kontak
    tiket.email = email
    tiket.deskripsi_pengaduan = deskripsi_pengaduan
    tiket.order_no = order_no

    db.session.commit()

    flash('Data berhasil diperbarui.', 'success')
    return redirect(url_for('list_ticket_by_nomor', nomor_ticket_id=nomor_ticket_id))

@app.route('/update-catatan/<int:ticket_id>', methods=['POST'])
@login_required
def update_catatan(ticket_id):
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    tiket = Ticket.query.get_or_404(ticket_id)
    
    catatan = request.form.get('catatan')
    
    if catatan:
        tiket.catatan = catatan
        tiket.tanggal_catatan = datetime.today().strftime('%Y-%m-%d')
        db.session.commit()
        flash('Catatan berhasil disimpan.', 'success')
    else:
        flash('Catatan tidak boleh kosong.', 'danger')

    return redirect(request.referrer or url_for('list_ticket_by_nomor', nomor_ticket_id=tiket.nomor_ticket_id))

@app.route('/mark-case-valid/<int:ticket_id>', methods=['POST'])
@login_required
def mark_case_valid(ticket_id):
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    tiket = Ticket.query.get_or_404(ticket_id)
    tiket.status_case = 'valid'
    db.session.commit()
    flash('Status case diubah menjadi VALID', 'success')
    return redirect(request.referrer or url_for('dashboard'))

@app.route('/update-tahapan-reopen/<int:nomor_ticket_id>/<int:ticket_id>', methods=['POST'])
@login_required
def update_tahapan_reopen(nomor_ticket_id, ticket_id):
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    tiket = Ticket.query.get_or_404(ticket_id)

    tahapan = request.form.get('tahapan')
    status_ticket = request.form.get('status_ticket')
    tahapan_2 = None

    if status_ticket == '3':
        date = request.form.get('tahapan_2_date')
        desc = request.form.get('tahapan_2_desc')
        tahapan_2 = f"{date} - {desc}" if date and desc else None
    elif status_ticket == '4':
        followup = request.form.get('tahapan_2_followup')
        tahapan_2 = followup if followup else None

    if not tahapan:
        flash('Tahapan wajib dipilih.', 'danger')
        return redirect(url_for('list_reopen_ticket', nomor_ticket_id=nomor_ticket_id))

    tiket.tahapan = tahapan
    tiket.status_ticket = status_ticket
    tiket.tahapan_2 = tahapan_2
    db.session.commit()

    new_history = History(
        nomor_ticket=tiket.nomor_ticket.nomor_ticket,
        order_number=tiket.order_no,
        status_ticket=status_ticket,
        tahapan=tahapan,
        create_by=current_user.id
    )
    db.session.add(new_history)
    db.session.commit()

    flash('Data berhasil diperbarui.', 'success')
    return redirect(url_for('list_reopen_ticket', nomor_ticket_id=nomor_ticket_id))

@app.route('/close-nomor-ticket/<int:nomor_ticket_id>', methods=['POST'])
@login_required
def close_nomor_ticket(nomor_ticket_id):
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    nomor_ticket = NomorTicket.query.get_or_404(nomor_ticket_id)
    
    nomor_ticket.status = 'close'
    nomor_ticket.closed_ticket = datetime.now(JAKARTA_TZ)

    tickets = Ticket.query.filter_by(nomor_ticket_id=nomor_ticket.id).all()
    for ticket in tickets:
        ticket.status_ticket = '4'

    db.session.commit()
    flash("Ticket berhasil ditutup.", "danger")
    return redirect(url_for('closed', nomor_ticket_id=nomor_ticket.id))

@app.route('/reopen-nomor-ticket/<int:nomor_ticket_id>', methods=['POST'])
@login_required
def reopen_nomor_ticket(nomor_ticket_id):
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    nomor_ticket = NomorTicket.query.get_or_404(nomor_ticket_id)
    
    nomor_ticket.status = 'reopen'

    tickets = Ticket.query.filter_by(nomor_ticket_id=nomor_ticket.id).all()
    for ticket in tickets:
        ticket.status_ticket = '5'

    db.session.commit()
    flash("Nomor ticket berhasil diubah menjadi Reopen.", "success")

    return redirect(url_for('reopen_ticket', nomor_ticket_id=nomor_ticket.id))

@app.route('/ticket-close')
@login_required
def close_ticket():
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    page = request.args.get('page', 1, type=int)
    jenis = request.args.get('jenis')
    status = request.args.get('status')
    tanggal = request.args.get('tanggal')
    q = request.args.get('q')

    nomor_ticket_query = NomorTicket.query.filter(NomorTicket.status == 'close')

    if q:
        matching_ids_from_nama = db.session.query(Ticket.nomor_ticket_id)\
            .filter(Ticket.nama_nasabah.ilike(f"%{q}%"))\
            .distinct().all()
        matching_ids_from_nama = [id[0] for id in matching_ids_from_nama]

        matching_ids_from_nomor = db.session.query(NomorTicket.id)\
            .filter(NomorTicket.nomor_ticket.ilike(f"%{q}%"))\
            .all()
        matching_ids_from_nomor = [id[0] for id in matching_ids_from_nomor]

        all_matching_ids = list(set(matching_ids_from_nama + matching_ids_from_nomor))

        nomor_ticket_query = nomor_ticket_query.filter(NomorTicket.id.in_(all_matching_ids))

    tickets_grouped = []

    for nt in nomor_ticket_query.all():
        query = Ticket.query.options(joinedload(Ticket.nomor_ticket))\
            .filter(Ticket.nomor_ticket_id == nt.id)

        if jenis:
            query = query.filter_by(jenis_pengaduan=jenis)
        if status:
            query = query.filter_by(status_ticket=status)
        if tanggal:
            try:
                tanggal_obj = datetime.strptime(tanggal, "%Y-%m-%d")
                query = query.filter(func.date(Ticket.tanggal) == tanggal_obj.date())
            except ValueError:
                pass

        first_ticket = query.order_by(Ticket.created_time.asc()).first()
        if first_ticket:
            tickets_grouped.append(first_ticket)

    tickets_grouped.sort(
        key=lambda x: x.created_time or datetime.min,
        reverse=True
    )

    per_page = 10
    total = len(tickets_grouped)
    start = (page - 1) * per_page
    end = start + per_page
    paginated_items = tickets_grouped[start:end]

    class Pagination:
        def __init__(self, items, page, per_page, total):
            self.items = items
            self.page = page
            self.per_page = per_page
            self.total = total
            self.pages = (total + per_page - 1) // per_page
            self.has_prev = page > 1
            self.has_next = page < self.pages
            self.prev_num = page - 1
            self.next_num = page + 1

    pagination = Pagination(paginated_items, page, per_page, total)

    count_by_nomor_ticket = dict(
        db.session.query(
            Ticket.nomor_ticket_id,
            func.count(Ticket.id)
        ).group_by(Ticket.nomor_ticket_id).all()
    )

    jumlah_tiket_close = NomorTicket.query.filter_by(status='close').count()

    return render_template(
        'ticket_close.html',
        user=current_user,
        tickets=pagination,
        count_by_nomor_ticket=count_by_nomor_ticket,
        jumlah_tiket_close=jumlah_tiket_close
    )

@app.route('/closed-ticket/<int:nomor_ticket_id>')
@login_required
def list_closed_ticket(nomor_ticket_id):
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    nomor_ticket = NomorTicket.query.get_or_404(nomor_ticket_id)

    tickets = Ticket.query.filter_by(nomor_ticket_id=nomor_ticket_id)\
        .order_by(Ticket.created_time.asc()).all()

    jenis_pengaduan_map = {
        1: "Informasi Pengajuan",
        2: "Permintaan Kode OTP",
        3: "Informasi Tenor",
        4: "Informasi Tagihan",
        5: "Informasi Denda",
        6: "Pembatalan Pinjaman",
        7: "Informasi Pencairan Dana",
        8: "Perilaku Petugas Penagihan",
        9: "Informasi Pembayaran",
        10: "Discount / Pemutihan"
    }

    detail_pengaduan_map = {
        1: [
            "Hasil Pengajuan",
            "Pengajuan Ditolak",
            "Status Pengajuan sedang ditransfer",
            "Tidak bisa pengajuan ulang karena keterlambatan",
            "Verifikasi Bank gagal",
            "Verifikasi KTP gagal",
            "Cara pengajuan",
            "Perubahan Nomor Handphone",
            "Perubahan Nomor Rekening"
        ],
        2: [
            "OTP Limit",
            "Tidak terima SMS OTP"
        ],
        3: [
            "Informasi Pinjaman"
        ],
        4: [
            "Konsultasi detail pinjaman saat ini",
            "Konsultasi Perpanjangan",
            "Bukti Transfer"
        ],
        5: [
            "Denda Keterlambatan"
        ],
        6: [
            "Hapus Data (Penutupan Akun)",
            "Pembatalan Pinjaman"
        ],
        7: [
            "Pencairan dana berhasil",
            "Status pencairan dana gagal",
            "Status pengajuan pencairan dana ulang",
            "Tidak terima dana",
            "Operasi Gagal (tidak bisa verifikasi wajah dan KTP)"
        ],
        8: [
            "Keluhan Penagihan",
            "Keluhan Reminder",
            "Penipuan"
        ],
        9: [
            "Konfirmasi Pembayaran",
            "Pembayaran belum masuk",
            "Pembayaran bukan ke VA UATAS",
            "Refund (pembayaran double)",
            "Meminta VA (cicilan)",
            "Meminta VA (pelunasan)",
            "Meminta VA (perpanjangan)",
            "Tidak bisa ambil VA"
        ],
        10: [
            "Meminta keringanan pembayaran (cicilan)",
            "Meminta keringanan pembayaran (potongan denda)",
            "Tidak ada dana"
        ]
    }

    return render_template(
        'list_closed_ticket.html',
        nomor_ticket=nomor_ticket,
        tickets=tickets,
        user=current_user,
        jenis_pengaduan_map=jenis_pengaduan_map,
        detail_pengaduan_map=detail_pengaduan_map
    )

@app.route('/reopen-ticket/<int:nomor_ticket_id>')
@login_required
def list_reopen_ticket(nomor_ticket_id):
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    nomor_ticket = NomorTicket.query.get_or_404(nomor_ticket_id)

    tickets = Ticket.query.filter_by(nomor_ticket_id=nomor_ticket_id)\
        .order_by(Ticket.created_time.asc()).all()

    jenis_pengaduan_map = {
        1: "Informasi Pengajuan",
        2: "Permintaan Kode OTP",
        3: "Informasi Tenor",
        4: "Informasi Tagihan",
        5: "Informasi Denda",
        6: "Pembatalan Pinjaman",
        7: "Informasi Pencairan Dana",
        8: "Perilaku Petugas Penagihan",
        9: "Informasi Pembayaran",
        10: "Discount / Pemutihan"
    }

    detail_pengaduan_map = {
        1: [
            "Hasil Pengajuan",
            "Pengajuan Ditolak",
            "Status Pengajuan sedang ditransfer",
            "Tidak bisa pengajuan ulang karena keterlambatan",
            "Verifikasi Bank gagal",
            "Verifikasi KTP gagal",
            "Cara pengajuan",
            "Perubahan Nomor Handphone",
            "Perubahan Nomor Rekening"
        ],
        2: [
            "OTP Limit",
            "Tidak terima SMS OTP"
        ],
        3: [
            "Informasi Pinjaman"
        ],
        4: [
            "Konsultasi detail pinjaman saat ini",
            "Konsultasi Perpanjangan",
            "Bukti Transfer"
        ],
        5: [
            "Denda Keterlambatan"
        ],
        6: [
            "Hapus Data (Penutupan Akun)",
            "Pembatalan Pinjaman"
        ],
        7: [
            "Pencairan dana berhasil",
            "Status pencairan dana gagal",
            "Status pengajuan pencairan dana ulang",
            "Tidak terima dana",
            "Operasi Gagal (tidak bisa verifikasi wajah dan KTP)"
        ],
        8: [
            "Keluhan Penagihan",
            "Keluhan Reminder",
            "Penipuan"
        ],
        9: [
            "Konfirmasi Pembayaran",
            "Pembayaran belum masuk",
            "Pembayaran bukan ke VA UATAS",
            "Refund (pembayaran double)",
            "Meminta VA (cicilan)",
            "Meminta VA (pelunasan)",
            "Meminta VA (perpanjangan)",
            "Tidak bisa ambil VA"
        ],
        10: [
            "Meminta keringanan pembayaran (cicilan)",
            "Meminta keringanan pembayaran (potongan denda)",
            "Tidak ada dana"
        ]
    }

    return render_template(
        'list_reopen_ticket.html',
        nomor_ticket=nomor_ticket,
        tickets=tickets,
        user=current_user,
        jenis_pengaduan_map=jenis_pengaduan_map,
        detail_pengaduan_map=detail_pengaduan_map
    )

@app.route('/reopen-ticket')
@login_required
def reopen_ticket():
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    page = request.args.get('page', 1, type=int)
    jenis = request.args.get('jenis')
    status = request.args.get('status')
    tanggal = request.args.get('tanggal')

    nomor_ticket_query = NomorTicket.query.filter(NomorTicket.status == 'reopen')

    tickets_grouped = []

    for nt in nomor_ticket_query.all():
        query = Ticket.query.options(joinedload(Ticket.nomor_ticket))\
            .filter(Ticket.nomor_ticket_id == nt.id)

        if jenis:
            query = query.filter_by(jenis_pengaduan=jenis)
        if status:
            query = query.filter_by(status_ticket=status)
        if tanggal:
            try:
                tanggal_obj = datetime.strptime(tanggal, "%Y-%m-%d")
                query = query.filter(func.date(Ticket.tanggal) == tanggal_obj.date())
            except ValueError:
                pass

        first_ticket = query.order_by(Ticket.created_time.asc()).first()
        if first_ticket:
            tickets_grouped.append(first_ticket)

    tickets_grouped.sort(
        key=lambda x: x.created_time or datetime.min,
        reverse=True
    )

    per_page = 10
    total = len(tickets_grouped)
    start = (page - 1) * per_page
    end = start + per_page
    paginated_items = tickets_grouped[start:end]

    class Pagination:
        def __init__(self, items, page, per_page, total):
            self.items = items
            self.page = page
            self.per_page = per_page
            self.total = total
            self.pages = (total + per_page - 1) // per_page
            self.has_prev = page > 1
            self.has_next = page < self.pages
            self.prev_num = page - 1
            self.next_num = page + 1

    pagination = Pagination(paginated_items, page, per_page, total)

    count_by_nomor_ticket = dict(
        db.session.query(
            Ticket.nomor_ticket_id,
            func.count(Ticket.id)
        ).group_by(Ticket.nomor_ticket_id).all()
    )

    jumlah_tiket_reopen = NomorTicket.query.filter_by(status='reopen').count()

    return render_template(
        'reopen_ticket.html',
        user=current_user,
        tickets=pagination,
        count_by_nomor_ticket=count_by_nomor_ticket,
        jumlah_tiket_reopen=jumlah_tiket_reopen
    )

@app.route('/download-template')
def download_template():
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    return send_from_directory(directory='static/files', path='template_cs.xlsx', as_attachment=True)

import re

def clean_alpha_only(val):
    cleaned = re.sub(r"[^a-zA-Z]", "", val) 
    return cleaned if cleaned else None

@app.route('/upload', methods=['POST'])
@login_required
def upload_excel():
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    def safe_val(val):
        return None if pd.isna(val) else str(val).strip()

    try:
        file = request.files.get('avatar')  
        if not file:
            flash("Tidak ada file yang diupload", 'danger')
            return redirect(request.referrer)

        df = pd.read_excel(file)

        expected_cols = ['kanal_pengaduan','tanggal', 'nama_nasabah', 'tipe_pengaduan',
                         'detail_pengaduan', 'order_no', 'os', 'dc', 'bucket']
        if not all(col in df.columns for col in expected_cols):
            flash("Kolom Excel tidak sesuai template.", 'danger')
            return redirect(request.referrer)

        jenis_pengaduan_map = {
            "Informasi Pengajuan": 1,
            "Permintaan Kode OTP": 2,
            "Informasi Tenor": 3,
            "Informasi Tagihan": 4,
            "Informasi Denda": 5,
            "Pembatalan Pinjaman": 6,
            "Informasi Pencairan Dana": 7,
            "Perilaku Petugas Penagihan": 8,
            "Informasi Pembayaran": 9,
            "Discount / Pemutihan": 10
        }

        existing_order_nos = {ticket.order_no for ticket in Ticket.query.with_entities(Ticket.order_no).all()}
        inserted_order_nos = set()
        inserted_count = 0

        for index, row in df.iterrows():
            order_no = safe_val(row['order_no'])

            if order_no and (order_no in existing_order_nos or order_no in inserted_order_nos):
                continue

            if order_no:
                inserted_order_nos.add(order_no)

            nomor_ticket = NomorTicket(nomor_ticket=generate_nomor_ticket())
            db.session.add(nomor_ticket)
            db.session.flush()

            tanggal_value = row['tanggal']
            if isinstance(tanggal_value, str):
                tanggal_value = datetime.strptime(tanggal_value, '%Y-%m-%d')
            elif pd.isna(tanggal_value):
                tanggal_value = datetime.utcnow()

            jenis_pengaduan_str = safe_val(row['tipe_pengaduan'])
            jenis_pengaduan_val = None
            if jenis_pengaduan_str:
                jenis_pengaduan_val = jenis_pengaduan_map.get(jenis_pengaduan_str)
                if not jenis_pengaduan_val:
                    raise ValueError(f"Jenis pengaduan tidak valid di baris {index + 2}: '{jenis_pengaduan_str}'")

            ticket = Ticket(
                kanal_pengaduan=safe_val(row['kanal_pengaduan']),
                nomor_ticket=nomor_ticket,
                tanggal=tanggal_value,
                nama_nasabah=safe_val(row['nama_nasabah']),
                jenis_pengaduan=jenis_pengaduan_val,
                detail_pengaduan=safe_val(row['detail_pengaduan']),
                order_no=order_no,
                nama_os=clean_alpha_only(safe_val(row['os']).replace(" ", "")) if safe_val(row['os']) and pd.notna(row['os']) else None,
                nama_dc=safe_val(row['dc']),
                nama_bucket=safe_val(row['bucket']).replace(" ", "") if safe_val(row['bucket']) and pd.notna(row['bucket']) else None,
                input_by=current_user.id,
                sla=10,
                status_ticket='1',
                created_time=datetime.utcnow(),
            )

            db.session.add(ticket)
            inserted_count += 1

        db.session.commit()
        flash(f"Berhasil mengimport data dari Excel. {inserted_count} data baru ditambahkan.", 'success')

    except Exception as e:
        db.session.rollback()
        flash(f"Gagal import: {e}", 'danger')

    return redirect(request.referrer)

@app.route("/case-valid")
@login_required
def case_valid():
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    page = request.args.get('page', 1, type=int)
    jenis = request.args.get('jenis')
    status = request.args.get('status')
    tanggal = request.args.get('tanggal')

    query = Ticket.query.options(joinedload(Ticket.nomor_ticket))\
        .filter(Ticket.status_case == 'valid')

    if jenis:
        query = query.filter(Ticket.jenis_pengaduan == jenis)
    if status:
        query = query.filter(Ticket.status_ticket == status)
    if tanggal:
        try:
            tanggal_obj = datetime.strptime(tanggal, "%Y-%m-%d")
            query = query.filter(func.date(Ticket.tanggal) == tanggal_obj.date())
        except ValueError:
            pass

    query = query.order_by(Ticket.created_time.desc())

    per_page = 10
    pagination = query.paginate(page=page, per_page=per_page, error_out=False)

    return render_template(
        "case_valid.html",
        user=current_user,
        tickets=pagination
    )

@app.route('/case-valid-qc')
@login_required
def case_valid_qc():
    if current_user.role != 'qc':
        flash('Akses ditolak: Anda bukan QC!')
        return redirect(url_for('index'))

    page = request.args.get('page', 1, type=int)
    jenis = request.args.get('jenis')
    status = request.args.get('status')
    tanggal = request.args.get('tanggal')
    q = request.args.get('q')

    nomor_ticket_query = NomorTicket.query.filter(
        NomorTicket.id_qc == current_user.id,
        NomorTicket.label_case == 'valid'
    )

    if q:
        matching_ids_from_nama = db.session.query(Ticket.nomor_ticket_id)\
            .filter(Ticket.nama_nasabah.ilike(f"%{q}%"))\
            .distinct().all()
        matching_ids_from_nama = [id[0] for id in matching_ids_from_nama]

        matching_ids_from_nomor = db.session.query(NomorTicket.id)\
            .filter(NomorTicket.nomor_ticket.ilike(f"%{q}%"))\
            .all()
        matching_ids_from_nomor = [id[0] for id in matching_ids_from_nomor]

        all_matching_ids = list(set(matching_ids_from_nama + matching_ids_from_nomor))
        nomor_ticket_query = nomor_ticket_query.filter(NomorTicket.id.in_(all_matching_ids))

    tickets_grouped = []

    for nt in nomor_ticket_query.all():
        query = Ticket.query.options(joinedload(Ticket.nomor_ticket))\
            .filter(Ticket.nomor_ticket_id == nt.id)

        if jenis:
            query = query.filter_by(jenis_pengaduan=jenis)
        if status:
            query = query.filter_by(status_ticket=status)
        if tanggal:
            try:
                tanggal_obj = datetime.strptime(tanggal, "%Y-%m-%d")
                query = query.filter(func.date(Ticket.tanggal) == tanggal_obj.date())
            except ValueError:
                pass

        first_ticket = query.order_by(Ticket.created_time.asc()).first()
        if first_ticket:
            tickets_grouped.append(first_ticket)

    tickets_grouped.sort(
        key=lambda x: x.created_time or datetime.min,
        reverse=True
    )

    per_page = 10
    total = len(tickets_grouped)
    start = (page - 1) * per_page
    end = start + per_page
    paginated_items = tickets_grouped[start:end]

    class Pagination:
        def __init__(self, items, page, per_page, total):
            self.items = items
            self.page = page
            self.per_page = per_page
            self.total = total
            self.pages = (total + per_page - 1) // per_page
            self.has_prev = page > 1
            self.has_next = page < self.pages
            self.prev_num = page - 1
            self.next_num = page + 1

    pagination = Pagination(paginated_items, page, per_page, total)

    count_by_nomor_ticket = dict(
        db.session.query(
            Ticket.nomor_ticket_id,
            func.count(Ticket.id)
        ).group_by(Ticket.nomor_ticket_id).all()
    )

    jumlah_tiket_valid = db.session.query(NomorTicket)\
        .join(Ticket, Ticket.nomor_ticket_id == NomorTicket.id)\
        .filter(
            NomorTicket.label_case == 'valid',
            NomorTicket.id_qc == current_user.id
        )\
        .distinct()\
        .count()

    return render_template(
        'case_valid_qc.html',
        user=current_user,
        tickets=pagination,
        count_by_nomor_ticket=count_by_nomor_ticket,
        jumlah_tiket_valid=jumlah_tiket_valid
    )

@app.route('/case-reopen-qc')
@login_required
def case_reopen_qc():
    if current_user.role != 'qc':
        flash('Akses ditolak: Anda bukan QC!')
        return redirect(url_for('index'))

    page = request.args.get('page', 1, type=int)
    jenis = request.args.get('jenis')
    status = request.args.get('status')
    tanggal = request.args.get('tanggal')
    q = request.args.get('q')

    nomor_ticket_query = NomorTicket.query.filter(
        NomorTicket.id_qc == current_user.id,
        NomorTicket.label_case == 'reopen'
    )

    if q:
        matching_ids_from_nama = db.session.query(Ticket.nomor_ticket_id)\
            .filter(Ticket.nama_nasabah.ilike(f"%{q}%"))\
            .distinct().all()
        matching_ids_from_nama = [id[0] for id in matching_ids_from_nama]

        matching_ids_from_nomor = db.session.query(NomorTicket.id)\
            .filter(NomorTicket.nomor_ticket.ilike(f"%{q}%"))\
            .all()
        matching_ids_from_nomor = [id[0] for id in matching_ids_from_nomor]

        all_matching_ids = list(set(matching_ids_from_nama + matching_ids_from_nomor))
        nomor_ticket_query = nomor_ticket_query.filter(NomorTicket.id.in_(all_matching_ids))

    tickets_grouped = []

    for nt in nomor_ticket_query.all():
        query = Ticket.query.options(joinedload(Ticket.nomor_ticket))\
            .filter(Ticket.nomor_ticket_id == nt.id)

        if jenis:
            query = query.filter_by(jenis_pengaduan=jenis)
        if status:
            query = query.filter_by(status_ticket=status)
        if tanggal:
            try:
                tanggal_obj = datetime.strptime(tanggal, "%Y-%m-%d")
                query = query.filter(func.date(Ticket.tanggal) == tanggal_obj.date())
            except ValueError:
                pass

        first_ticket = query.order_by(Ticket.created_time.asc()).first()
        if first_ticket:
            tickets_grouped.append(first_ticket)

    tickets_grouped.sort(
        key=lambda x: x.created_time or datetime.min,
        reverse=True
    )

    per_page = 10
    total = len(tickets_grouped)
    start = (page - 1) * per_page
    end = start + per_page
    paginated_items = tickets_grouped[start:end]

    class Pagination:
        def __init__(self, items, page, per_page, total):
            self.items = items
            self.page = page
            self.per_page = per_page
            self.total = total
            self.pages = (total + per_page - 1) // per_page
            self.has_prev = page > 1
            self.has_next = page < self.pages
            self.prev_num = page - 1
            self.next_num = page + 1

    pagination = Pagination(paginated_items, page, per_page, total)

    count_by_nomor_ticket = dict(
        db.session.query(
            Ticket.nomor_ticket_id,
            func.count(Ticket.id)
        ).group_by(Ticket.nomor_ticket_id).all()
    )

    jumlah_tiket_valid = db.session.query(NomorTicket)\
        .join(Ticket, Ticket.nomor_ticket_id == NomorTicket.id)\
        .filter(
            NomorTicket.label_case == 'reopen',
            NomorTicket.id_qc == current_user.id
        )\
        .distinct()\
        .count()

    return render_template(
        'case_reopen_qc.html',
        user=current_user,
        tickets=pagination,
        count_by_nomor_ticket=count_by_nomor_ticket,
        jumlah_tiket_valid=jumlah_tiket_valid
    )

@app.route('/case-not-valid-qc')
@login_required
def case_not_valid_qc():
    if current_user.role != 'qc':
        flash('Akses ditolak: Anda bukan QC!')
        return redirect(url_for('index'))

    page = request.args.get('page', 1, type=int)
    jenis = request.args.get('jenis')
    status = request.args.get('status')
    tanggal = request.args.get('tanggal')
    q = request.args.get('q')

    nomor_ticket_query = NomorTicket.query.filter(
        NomorTicket.id_qc == current_user.id,
        NomorTicket.label_case == 'tidak valid'
    )

    if q:
        matching_ids_from_nama = db.session.query(Ticket.nomor_ticket_id)\
            .filter(Ticket.nama_nasabah.ilike(f"%{q}%"))\
            .distinct().all()
        matching_ids_from_nama = [id[0] for id in matching_ids_from_nama]

        matching_ids_from_nomor = db.session.query(NomorTicket.id)\
            .filter(NomorTicket.nomor_ticket.ilike(f"%{q}%"))\
            .all()
        matching_ids_from_nomor = [id[0] for id in matching_ids_from_nomor]

        all_matching_ids = list(set(matching_ids_from_nama + matching_ids_from_nomor))
        nomor_ticket_query = nomor_ticket_query.filter(NomorTicket.id.in_(all_matching_ids))

    tickets_grouped = []

    for nt in nomor_ticket_query.all():
        query = Ticket.query.options(joinedload(Ticket.nomor_ticket))\
            .filter(Ticket.nomor_ticket_id == nt.id)

        if jenis:
            query = query.filter_by(jenis_pengaduan=jenis)
        if status:
            query = query.filter_by(status_ticket=status)
        if tanggal:
            try:
                tanggal_obj = datetime.strptime(tanggal, "%Y-%m-%d")
                query = query.filter(func.date(Ticket.tanggal) == tanggal_obj.date())
            except ValueError:
                pass

        first_ticket = query.order_by(Ticket.created_time.asc()).first()
        if first_ticket:
            tickets_grouped.append(first_ticket)

    tickets_grouped.sort(
        key=lambda x: x.created_time or datetime.min,
        reverse=True
    )

    per_page = 10
    total = len(tickets_grouped)
    start = (page - 1) * per_page
    end = start + per_page
    paginated_items = tickets_grouped[start:end]

    class Pagination:
        def __init__(self, items, page, per_page, total):
            self.items = items
            self.page = page
            self.per_page = per_page
            self.total = total
            self.pages = (total + per_page - 1) // per_page
            self.has_prev = page > 1
            self.has_next = page < self.pages
            self.prev_num = page - 1
            self.next_num = page + 1

    pagination = Pagination(paginated_items, page, per_page, total)

    count_by_nomor_ticket = dict(
        db.session.query(
            Ticket.nomor_ticket_id,
            func.count(Ticket.id)
        ).group_by(Ticket.nomor_ticket_id).all()
    )

    jumlah_tiket_tidak_valid = db.session.query(NomorTicket)\
        .join(Ticket, Ticket.nomor_ticket_id == NomorTicket.id)\
        .filter(
            NomorTicket.label_case == 'tidak valid',
            NomorTicket.id_qc == current_user.id
        )\
        .distinct()\
        .count()

    return render_template(
        'case_not_valid_qc.html',
        user=current_user,
        tickets=pagination,
        count_by_nomor_ticket=count_by_nomor_ticket,
        jumlah_tiket_tidak_valid=jumlah_tiket_tidak_valid
    )

@app.route('/upload-document/<int:ticket_id>', methods=['POST'])
@login_required
def upload_document(ticket_id):
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    ticket = Ticket.query.get_or_404(ticket_id)
    files = request.files.getlist('documents')
    uploaded_files = []

    for file in files:
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(UPLOAD_FOLDER, filename)

            if os.path.exists(filepath):
                name, ext = os.path.splitext(filename)
                filename = f"{name}_{datetime.utcnow().timestamp()}{ext}"
                filepath = os.path.join(UPLOAD_FOLDER, filename)

            file.save(filepath)
            uploaded_files.append(filename)

    if ticket.document:
        existing_files = ticket.document.split(',')
        all_files = existing_files + uploaded_files
    else:
        all_files = uploaded_files

    ticket.document = ','.join(all_files)
    db.session.commit()

    flash(f'{len(uploaded_files)} dokumen berhasil diupload.', 'success')
    return redirect(request.referrer or url_for('case_valid'))

@app.route('/hapus-dokumen/<int:ticket_id>', methods=['POST'])
@login_required
def hapus_dokumen(ticket_id):
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    filename = request.form.get('filename')
    ticket = Ticket.query.get_or_404(ticket_id)

    if not filename or filename not in (ticket.document or ''):
        flash('File tidak ditemukan atau tidak valid.', 'danger')
        return redirect(request.referrer)

    file_path = os.path.join('static/uploads', filename)
    if os.path.exists(file_path):
        os.remove(file_path)

    dokumen_list = ticket.document.split(',')
    dokumen_list.remove(filename)
    ticket.document = ','.join(dokumen_list)
    db.session.commit()

    flash(f'File {filename} berhasil dihapus.', 'success')
    return redirect(request.referrer)

@app.route('/sla')
@login_required
def sla():
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    page = request.args.get('page', 1, type=int)
    jenis = request.args.get('jenis')
    status = request.args.get('status')
    tanggal = request.args.get('tanggal')
    q = request.args.get('q')

    nomor_ticket_query = NomorTicket.query\
        .join(Ticket, Ticket.nomor_ticket_id == NomorTicket.id)\
        .filter(
            Ticket.sla == 0,
            NomorTicket.status == 'aktif'
        )

    if q:
        matching_ids_from_nama = db.session.query(Ticket.nomor_ticket_id)\
            .filter(Ticket.nama_nasabah.ilike(f"%{q}%"))\
            .distinct().all()
        matching_ids_from_nama = [id[0] for id in matching_ids_from_nama]

        matching_ids_from_nomor = db.session.query(NomorTicket.id)\
            .filter(NomorTicket.nomor_ticket.ilike(f"%{q}%"))\
            .all()
        matching_ids_from_nomor = [id[0] for id in matching_ids_from_nomor]

        all_matching_ids = list(set(matching_ids_from_nama + matching_ids_from_nomor))

        nomor_ticket_query = nomor_ticket_query.filter(NomorTicket.id.in_(all_matching_ids))

    tickets_grouped = []

    for nt in nomor_ticket_query.all():
        query = Ticket.query.options(joinedload(Ticket.nomor_ticket))\
            .filter(Ticket.nomor_ticket_id == nt.id, Ticket.sla == 0) 

        if jenis:
            query = query.filter_by(jenis_pengaduan=jenis)
        if status:
            query = query.filter_by(status_ticket=status)
        if tanggal:
            try:
                tanggal_obj = datetime.strptime(tanggal, "%Y-%m-%d")
                query = query.filter(func.date(Ticket.tanggal) == tanggal_obj.date())
            except ValueError:
                pass

        first_ticket = query.order_by(Ticket.created_time.asc()).first()
        if first_ticket:
            tickets_grouped.append(first_ticket)

    tickets_grouped.sort(
        key=lambda x: (
            0 if x.tahapan == "Eskalasi ke QC" else 1,
            -(x.created_time.timestamp() if x.created_time else 0)
        )
    )

    per_page = 10
    total = len(tickets_grouped)
    start = (page - 1) * per_page
    end = start + per_page
    paginated_items = tickets_grouped[start:end]

    class Pagination:
        def __init__(self, items, page, per_page, total):
            self.items = items
            self.page = page
            self.per_page = per_page
            self.total = total
            self.pages = (total + per_page - 1) // per_page
            self.has_prev = page > 1
            self.has_next = page < self.pages
            self.prev_num = page - 1
            self.next_num = page + 1

    pagination = Pagination(paginated_items, page, per_page, total)

    count_by_nomor_ticket = dict(
        db.session.query(
            Ticket.nomor_ticket_id,
            func.count(Ticket.id)
        ).filter(Ticket.sla == 0) 
        .group_by(Ticket.nomor_ticket_id).all()
    )

    jumlah_tiket_aktif = db.session.query(NomorTicket)\
        .join(Ticket, Ticket.nomor_ticket_id == NomorTicket.id)\
        .filter(NomorTicket.status == 'aktif', Ticket.sla == 0)\
        .distinct()\
        .count()

    return render_template(
        'sla.html',
        user=current_user,
        tickets=pagination,
        count_by_nomor_ticket=count_by_nomor_ticket,
        jumlah_tiket_aktif=jumlah_tiket_aktif
    )

@app.route('/add-detail-qc/<int:ticket_id>', methods=['POST'])
@login_required
def add_detail_qc(ticket_id):
    if current_user.role != 'qc':
        return redirect(request.referrer)

    tiket = Ticket.query.get_or_404(ticket_id)

    deskripsi_qc = request.form.get('deskripsi_qc')
    status_case = request.form.get('status_case')  

    uploaded_files = request.files.getlist('file_qc')
    existing_files = request.form.getlist('existing_images') 
    filenames = existing_files.copy()

    for file in uploaded_files:
        if file and file.filename:
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.root_path, 'static/uploads', filename)
            file.save(filepath)
            filenames.append(filename)

    tiket.deskripsi_qc = deskripsi_qc
    tiket.file_qc = ','.join(filenames) if filenames else None

    if tiket.nomor_ticket:
        tiket.nomor_ticket.label_case = status_case

        new_history = History(
            nomor_ticket=tiket.nomor_ticket.nomor_ticket,
            order_number=None,
            status_ticket=tiket.status_ticket,
            tahapan='Proses eskalasi QC',
            create_by=current_user.id,
            nama_os=None
        )
        db.session.add(new_history)

    db.session.commit()
    flash('Detail QC berhasil disimpan.', 'success')
    return redirect(request.referrer)

@app.route('/eskalasi-qc')
@login_required
def eskalasi_qc():
    if current_user.role != 'staff':
        return redirect(request.referrer)

    page = request.args.get('page', 1, type=int)
    jenis = request.args.get('jenis')
    status = request.args.get('status')
    tanggal = request.args.get('tanggal')
    q = request.args.get('q')

    nomor_ticket_query = NomorTicket.query.filter(
        or_(
            NomorTicket.label_case == None,
            NomorTicket.label_case == 'reopen'
        ),
        NomorTicket.id_qc.isnot(None), 
        or_(
            NomorTicket.status == None,
            and_(
                NomorTicket.status != 'close',
                NomorTicket.status != 'reopen'
            )
        )
    )

    if q:
        matching_ids_from_nama = db.session.query(Ticket.nomor_ticket_id)\
            .filter(Ticket.nama_nasabah.ilike(f"%{q}%"))\
            .distinct().all()
        matching_ids_from_nama = [id[0] for id in matching_ids_from_nama]

        matching_ids_from_nomor = db.session.query(NomorTicket.id)\
            .filter(NomorTicket.nomor_ticket.ilike(f"%{q}%"))\
            .all()
        matching_ids_from_nomor = [id[0] for id in matching_ids_from_nomor]

        all_matching_ids = list(set(matching_ids_from_nama + matching_ids_from_nomor))

        nomor_ticket_query = nomor_ticket_query.filter(NomorTicket.id.in_(all_matching_ids))

    tickets_grouped = []

    for nt in nomor_ticket_query.all():
        query = Ticket.query.options(joinedload(Ticket.nomor_ticket))\
            .filter(Ticket.nomor_ticket_id == nt.id)

        if jenis:
            query = query.filter_by(jenis_pengaduan=jenis)
        if status:
            query = query.filter_by(status_ticket=status)
        if tanggal:
            try:
                tanggal_obj = datetime.strptime(tanggal, "%Y-%m-%d")
                query = query.filter(func.date(Ticket.tanggal) == tanggal_obj.date())
            except ValueError:
                pass

        first_ticket = query.order_by(Ticket.created_time.asc()).first()
        if first_ticket:
            ada_feedback_qc = db.session.query(Ticket.id).filter(
                Ticket.nomor_ticket_id == nt.id,
                or_(
                    Ticket.deskripsi_qc.isnot(None),
                    Ticket.file_qc.isnot(None)
                )
            ).first() is not None

            first_ticket.feedback_qc_status = "Check Feedback QC" if ada_feedback_qc else "Belum ada Feedback QC"
            first_ticket.feedback_qc_badge = "success" if ada_feedback_qc else "warning"

            tickets_grouped.append(first_ticket)

    tickets_grouped.sort(
        key=lambda x: (
            0 if x.nomor_ticket.label_case == 'reopen' else 1,  
            -(x.nomor_ticket.change_date.timestamp() if x.nomor_ticket.change_date else 0)
        )
    )

    per_page = 10
    total = len(tickets_grouped)
    start = (page - 1) * per_page
    end = start + per_page
    paginated_items = tickets_grouped[start:end]

    class Pagination:
        def __init__(self, items, page, per_page, total):
            self.items = items
            self.page = page
            self.per_page = per_page
            self.total = total
            self.pages = (total + per_page - 1) // per_page
            self.has_prev = page > 1
            self.has_next = page < self.pages
            self.prev_num = page - 1
            self.next_num = page + 1

    pagination = Pagination(paginated_items, page, per_page, total)

    count_by_nomor_ticket = dict(
        db.session.query(
            Ticket.nomor_ticket_id,
            func.count(Ticket.id)
        ).group_by(Ticket.nomor_ticket_id).all()
    )

    jumlah_tiket_aktif = db.session.query(NomorTicket.id)\
        .filter(
            or_(
                NomorTicket.label_case == None,
                NomorTicket.label_case == 'reopen'
            ),
            NomorTicket.id_qc.isnot(None),
            or_(
                NomorTicket.status == None,
                and_(
                    NomorTicket.status != 'close',
                    NomorTicket.status != 'reopen'
                )
            )
        ).distinct().count()

    return render_template(
        'eskalasi.html',
        user=current_user,
        tickets=pagination,
        count_by_nomor_ticket=count_by_nomor_ticket,
        jumlah_tiket_aktif=jumlah_tiket_aktif
    )

@app.route('/eskalasi-qc-not-valid')
@login_required
def eskalasi_qc_not_valid():
    if current_user.role != 'staff':
        return redirect(request.referrer)

    page = request.args.get('page', 1, type=int)
    jenis = request.args.get('jenis')
    status = request.args.get('status')
    tanggal = request.args.get('tanggal')
    q = request.args.get('q')

    nomor_ticket_query = NomorTicket.query.filter(
        NomorTicket.label_case == "tidak valid",
        NomorTicket.id_qc.isnot(None),
        or_(
            NomorTicket.status == None,
            and_(
                NomorTicket.status != 'close',
                NomorTicket.status != 'reopen'
            )
        )
    )

    if q:
        matching_ids_from_nama = db.session.query(Ticket.nomor_ticket_id)\
            .filter(Ticket.nama_nasabah.ilike(f"%{q}%"))\
            .distinct().all()
        matching_ids_from_nama = [id[0] for id in matching_ids_from_nama]

        matching_ids_from_nomor = db.session.query(NomorTicket.id)\
            .filter(NomorTicket.nomor_ticket.ilike(f"%{q}%"))\
            .all()
        matching_ids_from_nomor = [id[0] for id in matching_ids_from_nomor]

        all_matching_ids = list(set(matching_ids_from_nama + matching_ids_from_nomor))
        nomor_ticket_query = nomor_ticket_query.filter(NomorTicket.id.in_(all_matching_ids))

    tickets_grouped = []

    for nt in nomor_ticket_query.all():
        query = Ticket.query.options(joinedload(Ticket.nomor_ticket))\
            .filter(Ticket.nomor_ticket_id == nt.id)

        if jenis:
            query = query.filter_by(jenis_pengaduan=jenis)
        if status:
            query = query.filter_by(status_ticket=status)
        if tanggal:
            try:
                tanggal_obj = datetime.strptime(tanggal, "%Y-%m-%d")
                query = query.filter(func.date(Ticket.tanggal) == tanggal_obj.date())
            except ValueError:
                pass

        first_ticket = query.order_by(Ticket.created_time.asc()).first()
        if first_ticket:
            ada_feedback_qc = db.session.query(Ticket.id).filter(
                Ticket.nomor_ticket_id == nt.id,
                or_(
                    Ticket.deskripsi_qc.isnot(None),
                    Ticket.file_qc.isnot(None)
                )
            ).first() is not None

            first_ticket.feedback_qc_status = "Check Feedback QC" if ada_feedback_qc else "Belum ada Feedback QC"
            first_ticket.feedback_qc_badge = "success" if ada_feedback_qc else "warning"

            tickets_grouped.append(first_ticket)

    tickets_grouped.sort(
        key=lambda x: (
            0 if x.nomor_ticket.label_case == 'reopen' else 1,  
            -(x.nomor_ticket.change_date.timestamp() if x.nomor_ticket.change_date else 0)
        )
    )

    per_page = 10
    total = len(tickets_grouped)
    start = (page - 1) * per_page
    end = start + per_page
    paginated_items = tickets_grouped[start:end]

    class Pagination:
        def __init__(self, items, page, per_page, total):
            self.items = items
            self.page = page
            self.per_page = per_page
            self.total = total
            self.pages = (total + per_page - 1) // per_page
            self.has_prev = page > 1
            self.has_next = page < self.pages
            self.prev_num = page - 1
            self.next_num = page + 1

    pagination = Pagination(paginated_items, page, per_page, total)

    count_by_nomor_ticket = dict(
        db.session.query(
            Ticket.nomor_ticket_id,
            func.count(Ticket.id)
        ).group_by(Ticket.nomor_ticket_id).all()
    )

    jumlah_tiket_aktif = db.session.query(NomorTicket.id)\
        .filter(
            NomorTicket.label_case == "tidak valid",
            NomorTicket.id_qc.isnot(None),
            or_(
                NomorTicket.status == None,
                and_(
                    NomorTicket.status != 'close',
                    NomorTicket.status != 'reopen'
                )
            )
        ).distinct().count()

    return render_template(
        'qc_not_valid.html',
        user=current_user,
        tickets=pagination,
        count_by_nomor_ticket=count_by_nomor_ticket,
        jumlah_tiket_aktif=jumlah_tiket_aktif
    )

@app.route('/eskalasi-qc-valid')
@login_required
def eskalasi_qc_valid():
    if current_user.role != 'staff':
        return redirect(request.referrer)

    page = request.args.get('page', 1, type=int)
    jenis = request.args.get('jenis')
    status = request.args.get('status')
    tanggal = request.args.get('tanggal')
    q = request.args.get('q')

    nomor_ticket_query = NomorTicket.query.filter(
        NomorTicket.label_case == "valid",
        NomorTicket.id_qc.isnot(None),
        or_(
            NomorTicket.status == None,
            and_(
                NomorTicket.status != 'close',
                NomorTicket.status != 'reopen'
            )
        )
    )

    if q:
        matching_ids_from_nama = db.session.query(Ticket.nomor_ticket_id)\
            .filter(Ticket.nama_nasabah.ilike(f"%{q}%"))\
            .distinct().all()
        matching_ids_from_nama = [id[0] for id in matching_ids_from_nama]

        matching_ids_from_nomor = db.session.query(NomorTicket.id)\
            .filter(NomorTicket.nomor_ticket.ilike(f"%{q}%"))\
            .all()
        matching_ids_from_nomor = [id[0] for id in matching_ids_from_nomor]

        all_matching_ids = list(set(matching_ids_from_nama + matching_ids_from_nomor))
        nomor_ticket_query = nomor_ticket_query.filter(NomorTicket.id.in_(all_matching_ids))

    tickets_grouped = []

    for nt in nomor_ticket_query.all():
        query = Ticket.query.options(joinedload(Ticket.nomor_ticket))\
            .filter(Ticket.nomor_ticket_id == nt.id)

        if jenis:
            query = query.filter_by(jenis_pengaduan=jenis)
        if status:
            query = query.filter_by(status_ticket=status)
        if tanggal:
            try:
                tanggal_obj = datetime.strptime(tanggal, "%Y-%m-%d")
                query = query.filter(func.date(Ticket.tanggal) == tanggal_obj.date())
            except ValueError:
                pass

        first_ticket = query.order_by(Ticket.created_time.asc()).first()
        if first_ticket:
            ada_feedback_qc = db.session.query(Ticket.id).filter(
                Ticket.nomor_ticket_id == nt.id,
                or_(
                    Ticket.deskripsi_qc.isnot(None),
                    Ticket.file_qc.isnot(None)
                )
            ).first() is not None

            first_ticket.feedback_qc_status = "Check Feedback QC" if ada_feedback_qc else "Belum ada Feedback QC"
            first_ticket.feedback_qc_badge = "success" if ada_feedback_qc else "warning"

            tickets_grouped.append(first_ticket)

    tickets_grouped.sort(
        key=lambda x: (
            0 if x.nomor_ticket.label_case == 'reopen' else 1,  
            -(x.nomor_ticket.change_date.timestamp() if x.nomor_ticket.change_date else 0)
        )
    )

    per_page = 10
    total = len(tickets_grouped)
    start = (page - 1) * per_page
    end = start + per_page
    paginated_items = tickets_grouped[start:end]

    class Pagination:
        def __init__(self, items, page, per_page, total):
            self.items = items
            self.page = page
            self.per_page = per_page
            self.total = total
            self.pages = (total + per_page - 1) // per_page
            self.has_prev = page > 1
            self.has_next = page < self.pages
            self.prev_num = page - 1
            self.next_num = page + 1

    pagination = Pagination(paginated_items, page, per_page, total)

    count_by_nomor_ticket = dict(
        db.session.query(
            Ticket.nomor_ticket_id,
            func.count(Ticket.id)
        ).group_by(Ticket.nomor_ticket_id).all()
    )

    jumlah_tiket_aktif = db.session.query(NomorTicket.id)\
        .filter(
            NomorTicket.label_case == "valid",
            NomorTicket.id_qc.isnot(None),
            or_(
                NomorTicket.status == None,
                and_(
                    NomorTicket.status != 'close',
                    NomorTicket.status != 'reopen'
                )
            )
        ).distinct().count()

    return render_template(
        'qc_valid.html',
        user=current_user,
        tickets=pagination,
        count_by_nomor_ticket=count_by_nomor_ticket,
        jumlah_tiket_aktif=jumlah_tiket_aktif
    )

@app.route('/eskalasi-ticket-qc-not-valid/<int:nomor_ticket_id>')
@login_required
def eskalasi_ticket_qc_not_valid(nomor_ticket_id):
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    nomor_ticket = NomorTicket.query.get_or_404(nomor_ticket_id)

    tickets = Ticket.query.filter_by(nomor_ticket_id=nomor_ticket_id)\
        .order_by(Ticket.created_time.asc()).all()
    
    semua_kosong = all(ticket.deskripsi_qc in [None, ""] for ticket in tickets)

    if semua_kosong:
        status_feedback = "Belum ada Feedback"
        badge_feedback = "warning"
    else:
        status_feedback = "Ada Feedback"
        badge_feedback = "success"

    catatan_list = Catatan.query.filter_by(nomor_ticket_id=nomor_ticket_id)\
        .order_by(Catatan.tanggal.desc()).all()

    qc_users = User.query.filter_by(role='qc').all()

    jenis_pengaduan_map = {
        1: "Informasi Pengajuan",
        2: "Permintaan Kode OTP",
        3: "Informasi Tenor",
        4: "Informasi Tagihan",
        5: "Informasi Denda",
        6: "Pembatalan Pinjaman",
        7: "Informasi Pencairan Dana",
        8: "Perilaku Petugas Penagihan",
        9: "Informasi Pembayaran",
        10: "Discount / Pemutihan"
    }

    detail_pengaduan_map = {
        1: [
            "Hasil Pengajuan",
            "Pengajuan Ditolak",
            "Status Pengajuan sedang ditransfer",
            "Tidak bisa pengajuan ulang karena keterlambatan",
            "Verifikasi Bank gagal",
            "Verifikasi KTP gagal",
            "Cara pengajuan",
            "Perubahan Nomor Handphone",
            "Perubahan Nomor Rekening"
        ],
        2: [
            "OTP Limit",
            "Tidak terima SMS OTP"
        ],
        3: [
            "Informasi Pinjaman"
        ],
        4: [
            "Konsultasi detail pinjaman saat ini",
            "Konsultasi Perpanjangan",
            "Bukti Transfer"
        ],
        5: [
            "Denda Keterlambatan"
        ],
        6: [
            "Hapus Data (Penutupan Akun)",
            "Pembatalan Pinjaman"
        ],
        7: [
            "Pencairan dana berhasil",
            "Status pencairan dana gagal",
            "Status pengajuan pencairan dana ulang",
            "Tidak terima dana",
            "Operasi Gagal (tidak bisa verifikasi wajah dan KTP)"
        ],
        8: [
            "Keluhan Penagihan",
            "Keluhan Reminder",
            "Penipuan"
        ],
        9: [
            "Konfirmasi Pembayaran",
            "Pembayaran belum masuk",
            "Pembayaran bukan ke VA UATAS",
            "Refund (pembayaran double)",
            "Meminta VA (cicilan)",
            "Meminta VA (pelunasan)",
            "Meminta VA (perpanjangan)",
            "Tidak bisa ambil VA"
        ],
        10: [
            "Meminta keringanan pembayaran (cicilan)",
            "Meminta keringanan pembayaran (potongan denda)",
            "Tidak ada dana"
        ]
    }

    return render_template(
        'eskalasi_ticket_qc_not_valid.html',
        nomor_ticket=nomor_ticket,
        tickets=tickets,
        user=current_user,
        jenis_pengaduan_map=jenis_pengaduan_map,
        detail_pengaduan_map=detail_pengaduan_map,
        qc_users=qc_users,
        catatan_list=catatan_list,
        status_feedback=status_feedback,
        badge_feedback=badge_feedback
    )

@app.route('/eskalasi-ticket-qc-valid/<int:nomor_ticket_id>')
@login_required
def eskalasi_ticket_qc_valid(nomor_ticket_id):
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    nomor_ticket = NomorTicket.query.get_or_404(nomor_ticket_id)

    tickets = Ticket.query.filter_by(nomor_ticket_id=nomor_ticket_id)\
        .order_by(Ticket.created_time.asc()).all()
    
    semua_kosong = all(ticket.deskripsi_qc in [None, ""] for ticket in tickets)

    if semua_kosong:
        status_feedback = "Belum ada Feedback"
        badge_feedback = "warning"
    else:
        status_feedback = "Ada Feedback"
        badge_feedback = "success"

    catatan_list = Catatan.query.filter_by(nomor_ticket_id=nomor_ticket_id)\
        .order_by(Catatan.tanggal.desc()).all()

    qc_users = User.query.filter_by(role='qc').all()

    jenis_pengaduan_map = {
        1: "Informasi Pengajuan",
        2: "Permintaan Kode OTP",
        3: "Informasi Tenor",
        4: "Informasi Tagihan",
        5: "Informasi Denda",
        6: "Pembatalan Pinjaman",
        7: "Informasi Pencairan Dana",
        8: "Perilaku Petugas Penagihan",
        9: "Informasi Pembayaran",
        10: "Discount / Pemutihan"
    }

    detail_pengaduan_map = {
        1: [
            "Hasil Pengajuan",
            "Pengajuan Ditolak",
            "Status Pengajuan sedang ditransfer",
            "Tidak bisa pengajuan ulang karena keterlambatan",
            "Verifikasi Bank gagal",
            "Verifikasi KTP gagal",
            "Cara pengajuan",
            "Perubahan Nomor Handphone",
            "Perubahan Nomor Rekening"
        ],
        2: [
            "OTP Limit",
            "Tidak terima SMS OTP"
        ],
        3: [
            "Informasi Pinjaman"
        ],
        4: [
            "Konsultasi detail pinjaman saat ini",
            "Konsultasi Perpanjangan",
            "Bukti Transfer"
        ],
        5: [
            "Denda Keterlambatan"
        ],
        6: [
            "Hapus Data (Penutupan Akun)",
            "Pembatalan Pinjaman"
        ],
        7: [
            "Pencairan dana berhasil",
            "Status pencairan dana gagal",
            "Status pengajuan pencairan dana ulang",
            "Tidak terima dana",
            "Operasi Gagal (tidak bisa verifikasi wajah dan KTP)"
        ],
        8: [
            "Keluhan Penagihan",
            "Keluhan Reminder",
            "Penipuan"
        ],
        9: [
            "Konfirmasi Pembayaran",
            "Pembayaran belum masuk",
            "Pembayaran bukan ke VA UATAS",
            "Refund (pembayaran double)",
            "Meminta VA (cicilan)",
            "Meminta VA (pelunasan)",
            "Meminta VA (perpanjangan)",
            "Tidak bisa ambil VA"
        ],
        10: [
            "Meminta keringanan pembayaran (cicilan)",
            "Meminta keringanan pembayaran (potongan denda)",
            "Tidak ada dana"
        ]
    }

    return render_template(
        'eskalasi_ticket_qc_valid.html',
        nomor_ticket=nomor_ticket,
        tickets=tickets,
        user=current_user,
        jenis_pengaduan_map=jenis_pengaduan_map,
        detail_pengaduan_map=detail_pengaduan_map,
        qc_users=qc_users,
        catatan_list=catatan_list,
        status_feedback=status_feedback,
        badge_feedback=badge_feedback
    )

@app.route('/ubah-label-valid/<int:nomor_ticket_id>', methods=['POST'])
@login_required
def ubah_label_valid(nomor_ticket_id):
    nomor_ticket = NomorTicket.query.get_or_404(nomor_ticket_id)
    nomor_ticket.label_case = 'reopen'
    db.session.commit()
    flash('Case di Re-open dan dioper kembali ke QC.', 'success')
    return redirect(request.referrer or url_for('eskalasi_qc_not_valid'))

@app.route('/add-catatan/<int:nomor_ticket_id>', methods=['POST'])
@login_required
def add_catatan(nomor_ticket_id):
    deskripsi = request.form.get('deskripsi')
    id_user = current_user.id
    tanggal = datetime.now(timezone('Asia/Jakarta'))

    if not deskripsi:
        flash('Deskripsi tidak boleh kosong.', 'danger')
        return redirect(request.referrer)

    catatan_baru = Catatan(
        nomor_ticket_id=nomor_ticket_id,
        deskripsi=deskripsi,
        user_id=id_user,
        tanggal=tanggal
    )

    db.session.add(catatan_baru)
    db.session.commit()
    flash('Catatan berhasil ditambahkan.', 'success')
    return redirect(request.referrer)

@app.route('/admin_statistik')
@login_required
def admin_statistik():
    if current_user.role != 'admin':
        flash('Akses ditolak: Anda bukan admin.')
        return redirect(url_for('login'))

    date_range = request.args.get('date_range')
    range2 = request.args.get('range2') 
    selected_os = request.args.getlist('os')
    selected_bucket = request.args.getlist('bucket')
    selected_kanal = request.args.getlist('kanal')
    selected_jenis_pengaduan = request.args.getlist('jenis_pengaduan')
    chart_by = request.args.get('chart_by')

    all_os = [r[0] for r in db.session.query(Ticket.nama_os).distinct().all() if r[0] is not None]
    all_bucket = [r[0] for r in db.session.query(Ticket.nama_bucket).distinct().all() if r[0] is not None]
    all_kanal = [r[0] for r in db.session.query(Ticket.kanal_pengaduan).distinct().all() if r[0] is not None]

    apply_os = len(selected_os) > 0
    apply_bucket = len(selected_bucket) > 0
    apply_kanal = len(selected_kanal) > 0
    jenis_pengaduan_filter = [int(j) for j in selected_jenis_pengaduan if j and j.strip().isdigit()]
    apply_jenis = len(jenis_pengaduan_filter) > 0

    if not date_range:
        start_date = db.session.query(func.min(Ticket.tanggal)).scalar()
        end_date = db.session.query(func.max(Ticket.tanggal)).scalar()
    else:
        try:
            start_str, end_str = date_range.split(' - ')
            start_date = datetime.strptime(start_str.strip(), '%Y-%m-%d')
            end_date = datetime.strptime(end_str.strip(), '%Y-%m-%d')
        except Exception:
            flash("Format tanggal tidak valid.", "danger")
            start_date = end_date = None

    range2_start = range2_end = None
    if range2:
        try:
            r2s, r2e = range2.split(' - ')
            range2_start = datetime.strptime(r2s.strip(), '%Y-%m-%d')
            range2_end = datetime.strptime(r2e.strip(), '%Y-%m-%d')
        except Exception:
            flash("Format range tanggal 2 tidak valid.", "warning")
            range2_start = range2_end = None

    def format_tanggal_indonesia(tanggal):
        if not tanggal:
            return ""
        bulan_dict = {
            1: 'Januari', 2: 'Februari', 3: 'Maret', 4: 'April',
            5: 'Mei', 6: 'Juni', 7: 'Juli', 8: 'Agustus',
            9: 'September', 10: 'Oktober', 11: 'November', 12: 'Desember'
        }
        return f"{tanggal.day} {bulan_dict[tanggal.month]} {tanggal.year}"

    def periode_str(sa, ea):
        if not sa or not ea:
            return "Semua"
        return f"{format_tanggal_indonesia(sa)} - {format_tanggal_indonesia(ea)}"

    jenis_pengaduan_labels = {
        1: "Informasi Pengajuan",
        2: "Permintaan Kode OTP",
        3: "Informasi Tenor",
        4: "Informasi Tagihan",
        5: "Informasi Denda",
        6: "Pembatalan Pinjaman",
        7: "Informasi Pencairan Dana",
        8: "Perilaku Petugas Penagihan",
        9: "Informasi Pembayaran",
        10: "Discount / Pemutihan"
    }

    total_nomor_ticket_all = db.session.query(func.count(distinct(NomorTicket.id))).scalar()

    if not (apply_os or apply_bucket or apply_jenis or apply_kanal or (start_date and end_date)):
        total_nomor_ticket = total_nomor_ticket_all
        total_open = db.session.query(NomorTicket).filter(
            or_(NomorTicket.status == 'aktif', NomorTicket.status == 'Reopen')
        ).count()
        total_close = db.session.query(NomorTicket).filter(NomorTicket.status == 'close').count()
    else:
        tq = db.session.query(NomorTicket.id).join(Ticket, Ticket.nomor_ticket_id == NomorTicket.id)
        if start_date and end_date:
            tq = tq.filter(Ticket.tanggal.between(start_date, end_date))
        if apply_os:
            tq = tq.filter(Ticket.nama_os.in_(selected_os))
        if apply_bucket:
            tq = tq.filter(Ticket.nama_bucket.in_(selected_bucket))
        if apply_kanal:
            tq = tq.filter(Ticket.kanal_pengaduan.in_(selected_kanal))
        if apply_jenis:
            tq = tq.filter(Ticket.jenis_pengaduan.in_(jenis_pengaduan_filter))

        total_nomor_ticket = db.session.query(func.count(distinct(tq.subquery().c.id))).scalar()

        nt_ids_subq = tq.subquery()
        total_open = db.session.query(func.count(NomorTicket.id)).filter(
            NomorTicket.id.in_(db.session.query(nt_ids_subq.c.id)),
            or_(NomorTicket.status == 'aktif', NomorTicket.status == 'Reopen')
        ).scalar()
        total_close = db.session.query(func.count(NomorTicket.id)).filter(
            NomorTicket.id.in_(db.session.query(nt_ids_subq.c.id)),
            NomorTicket.status == 'close'
        ).scalar()

    def agg_by_os(sa, ea):
        q = db.session.query(Ticket.nama_os, func.count(Ticket.id))
        if sa and ea:
            q = q.filter(Ticket.tanggal.between(sa, ea))
        if apply_os:
            q = q.filter(Ticket.nama_os.in_(selected_os))
        if apply_kanal:
            q = q.filter(Ticket.kanal_pengaduan.in_(selected_kanal))
        if apply_jenis:
            q = q.filter(Ticket.jenis_pengaduan.in_(jenis_pengaduan_filter))
        if apply_bucket:
            q = q.filter(Ticket.nama_bucket.in_(selected_bucket))
        rows = q.group_by(Ticket.nama_os).all()
        return { r[0]: r[1] for r in rows if r[0] }

    def agg_by_jenis(sa, ea):
        q = db.session.query(
            Ticket.jenis_pengaduan,
            func.count(distinct(Ticket.nomor_ticket_id))
        ).join(NomorTicket, Ticket.nomor_ticket_id == NomorTicket.id).filter(
            Ticket.jenis_pengaduan.in_(range(1, 11))
        )
        if sa and ea:
            q = q.filter(Ticket.tanggal.between(sa, ea))
        if apply_os:
            q = q.filter(Ticket.nama_os.in_(selected_os))
        if apply_bucket:
            q = q.filter(Ticket.nama_bucket.in_(selected_bucket))
        if apply_kanal:
            q = q.filter(Ticket.kanal_pengaduan.in_(selected_kanal))
        if apply_jenis:
            q = q.filter(Ticket.jenis_pengaduan.in_(jenis_pengaduan_filter))
        rows = q.group_by(Ticket.jenis_pengaduan).all()
        return {int(r[0]): r[1] for r in rows if r[0] is not None}

    from collections import defaultdict
    if apply_bucket:
        cq = db.session.query(
            Ticket.nama_os,
            Ticket.nama_bucket,
            func.count(Ticket.id)
        )
        if start_date and end_date:
            cq = cq.filter(Ticket.tanggal.between(start_date, end_date))
        if apply_os:
            cq = cq.filter(Ticket.nama_os.in_(selected_os))
        cq = cq.filter(Ticket.nama_bucket.in_(selected_bucket))
        if apply_kanal:
            cq = cq.filter(Ticket.kanal_pengaduan.in_(selected_kanal))
        if apply_jenis:
            cq = cq.filter(Ticket.jenis_pengaduan.in_(jenis_pengaduan_filter))

        chart_rows = cq.group_by(Ticket.nama_bucket, Ticket.nama_os).all()
        all_os_labels = sorted(set([r[0] for r in chart_rows if r[0]]))
        grouped = defaultdict(lambda: defaultdict(int))
        for os_name, bucket, cnt in chart_rows:
            if os_name:
                grouped[(bucket or "Tidak Diketahui")][os_name] = cnt

        chart_series = []
        for bucket_label in selected_bucket:
            bl = bucket_label or "Tidak Diketahui"
            series_data = [grouped[bl].get(os_label, 0) for os_label in all_os_labels]
            chart_series.append({"name": bl, "data": series_data})
        chart_labels = all_os_labels
        chart_title = "Jumlah Tiket per OS berdasarkan Bucket"
        if start_date and end_date:
            chart_title += f" ({format_tanggal_indonesia(start_date)} - {format_tanggal_indonesia(end_date)})"
        else:
            tmin = db.session.query(func.min(Ticket.tanggal)).scalar()
            tmax = db.session.query(func.max(Ticket.tanggal)).scalar()
            if tmin and tmax:
                chart_title += f" ({format_tanggal_indonesia(tmin)} - {format_tanggal_indonesia(tmax)})"
    else:
        aggA = agg_by_os(start_date, end_date)
        labels_set = set(aggA.keys())
        aggB = {}
        if range2_start and range2_end:
            aggB = agg_by_os(range2_start, range2_end)
            labels_set |= set(aggB.keys())

        labels_all = sorted([lbl for lbl in labels_set if lbl])

        dataA = [aggA.get(lbl, 0) for lbl in labels_all]
        chart_series = [{"name": f"Range 1 ({periode_str(start_date, end_date)})", "data": dataA}]
        if aggB:
            dataB = [aggB.get(lbl, 0) for lbl in labels_all]
            chart_series.append({"name": f"Range 2 ({periode_str(range2_start, range2_end)})", "data": dataB})

        chart_labels = labels_all
        chart_title = f"Jumlah Order per Tiket ({periode_str(start_date, end_date)})"
        if apply_kanal:
            chart_title += " - " + ", ".join(selected_kanal)
        if start_date and end_date:
            chart_title += f" ({format_tanggal_indonesia(start_date)} - {format_tanggal_indonesia(end_date)})"
        else:
            tmin = db.session.query(func.min(Ticket.tanggal)).scalar()
            tmax = db.session.query(func.max(Ticket.tanggal)).scalar()
            if tmin and tmax:
                chart_title += f" ({format_tanggal_indonesia(tmin)} - {format_tanggal_indonesia(tmax)})"

    aggA_jenis = agg_by_jenis(start_date, end_date)
    aggB_jenis = agg_by_jenis(range2_start, range2_end) if (range2_start and range2_end) else {}

    jenis_pengaduan_chart_labels = [k for k in range(1, 11) if (k in aggA_jenis) or (k in aggB_jenis)]
    dataA_jenis = [aggA_jenis.get(k, 0) for k in jenis_pengaduan_chart_labels]

    jenis_pengaduan_chart_series = [{
        "name": f"Range 1 ({periode_str(start_date, end_date)})",
        "data": dataA_jenis
    }]
    if aggB_jenis:
        dataB_jenis = [aggB_jenis.get(k, 0) for k in jenis_pengaduan_chart_labels]
        jenis_pengaduan_chart_series.append({
            "name": f"Range 2 ({periode_str(range2_start, range2_end)})",
            "data": dataB_jenis
        })

    jenis_pengaduan_chart_title = f"Jumlah Tiket per Jenis Pengaduan ({periode_str(start_date, end_date)})"
    if range2_start and range2_end:
        jenis_pengaduan_chart_title += f" vs ({periode_str(range2_start, range2_end)})"

    kq = db.session.query(
        Ticket.kanal_pengaduan,
        func.count(distinct(NomorTicket.id))
    ).join(NomorTicket, Ticket.nomor_ticket_id == NomorTicket.id).filter(
        Ticket.kanal_pengaduan.isnot(None)
    )
    if start_date and end_date:
        kq = kq.filter(Ticket.tanggal.between(start_date, end_date))
    if apply_os:
        kq = kq.filter(Ticket.nama_os.in_(selected_os))
    if apply_bucket:
        kq = kq.filter(Ticket.nama_bucket.in_(selected_bucket))
    if apply_jenis:
        kq = kq.filter(Ticket.jenis_pengaduan.in_(jenis_pengaduan_filter))
    if apply_kanal:
        kq = kq.filter(Ticket.kanal_pengaduan.in_(selected_kanal))

    kanal_pengaduan_chart_rows = kq.group_by(Ticket.kanal_pengaduan).all()
    kanal_pengaduan_chart_labels = [row[0] for row in kanal_pengaduan_chart_rows]
    kanal_pengaduan_chart_values = [row[1] for row in kanal_pengaduan_chart_rows]
    kanal_pengaduan_chart_series = [{"name": "Jumlah Ticket", "data": kanal_pengaduan_chart_values}]

    selected_range2 = None
    if range2_start and range2_end:
        selected_range2 = f"{range2_start.strftime('%Y-%m-%d')} - {range2_end.strftime('%Y-%m-%d')}"

    return render_template(
        'admin_statistik_1.html',
        user=current_user,
        total_nomor_ticket=total_nomor_ticket,
        total_nomor_ticket_all=total_nomor_ticket_all,
        total_open=total_open,
        total_close=total_close,
        selected_date_range=date_range or "Semua",
        selected_os=selected_os,
        selected_bucket=selected_bucket,
        selected_kanal=selected_kanal,
        selected_jenis_pengaduan=selected_jenis_pengaduan,
        all_os=all_os,
        all_bucket=all_bucket,
        all_kanal=all_kanal,
        chart_labels=chart_labels,
        chart_series=chart_series,
        chart_title=chart_title,
        jenis_pengaduan_chart_labels=jenis_pengaduan_chart_labels,
        jenis_pengaduan_chart_series=jenis_pengaduan_chart_series,
        jenis_pengaduan_chart_title=jenis_pengaduan_chart_title,
        ticket_chart_title="Jumlah Ticket per Status",
        tanggal_awal=format_tanggal_indonesia(start_date) if start_date else None,
        tanggal_akhir=format_tanggal_indonesia(end_date) if end_date else None,
        kanal_pengaduan_chart_labels=kanal_pengaduan_chart_labels,
        kanal_pengaduan_chart_values=kanal_pengaduan_chart_values,
        kanal_pengaduan_chart_series=kanal_pengaduan_chart_series,
        range2=range2,
        selected_range2=selected_range2
    )

@app.route('/list-ticket')
@login_required
def list_data():
    if current_user.role != 'admin':
        return redirect(request.referrer)

    page = request.args.get('page', 1, type=int)
    jenis = request.args.get('jenis')
    status = request.args.get('status')
    tanggal = request.args.get('tanggal')
    tanggal_penanganan = request.args.get('tanggal_penanganan')
    q = request.args.get('q')
    tahapan = request.args.get('tahapan')

    tahapan_options = [row[0] for row in db.session.query(Ticket.tahapan).distinct().all() if row[0]]

    nomor_ticket_query = NomorTicket.query\
        .join(Ticket, Ticket.nomor_ticket_id == NomorTicket.id)\
        .filter(
            NomorTicket.id_qc == None
        )\
        .distinct()

    if q:
        matching_ids_from_nama = db.session.query(Ticket.nomor_ticket_id)\
            .filter(Ticket.nama_nasabah.ilike(f"%{q}%"))\
            .distinct().all()
        matching_ids_from_nama = [id[0] for id in matching_ids_from_nama]

        matching_ids_from_nomor = db.session.query(NomorTicket.id)\
            .filter(NomorTicket.nomor_ticket.ilike(f"%{q}%"))\
            .all()
        matching_ids_from_nomor = [id[0] for id in matching_ids_from_nomor]

        all_matching_ids = list(set(matching_ids_from_nama + matching_ids_from_nomor))

        nomor_ticket_query = nomor_ticket_query.filter(NomorTicket.id.in_(all_matching_ids))

    tickets_grouped = []

    for nt in nomor_ticket_query.all():
        query = Ticket.query.options(joinedload(Ticket.nomor_ticket))\
            .filter(Ticket.nomor_ticket_id == nt.id, Ticket.sla != 0)

        if jenis:
            query = query.filter_by(kanal_pengaduan=jenis)
        if status:
            query = query.filter_by(status_ticket=status)
        if tanggal:
            try:
                tanggal_obj = datetime.strptime(tanggal, "%Y-%m-%d")
                query = query.filter(func.date(Ticket.tanggal) == tanggal_obj.date())
            except ValueError:
                pass
        if tanggal_penanganan:
            try:
                tanggal_penanganan_obj = datetime.strptime(tanggal_penanganan, "%Y-%m-%d")
                query = query.filter(func.date(Ticket.created_time) == tanggal_penanganan_obj.date())
            except ValueError:
                pass
        if tahapan:
            query = query.filter_by(tahapan=tahapan)

        first_ticket = query.order_by(Ticket.created_time.asc()).first()
        if first_ticket:
            tickets_grouped.append(first_ticket)

    tickets_grouped.sort(
        key=lambda x: x.created_time or datetime.min,
        reverse=True
    )

    per_page = 10
    total = len(tickets_grouped)
    start = (page - 1) * per_page
    end = start + per_page
    paginated_items = tickets_grouped[start:end]

    class Pagination:
        def __init__(self, items, page, per_page, total):
            self.items = items
            self.page = page
            self.per_page = per_page
            self.total = total
            self.pages = (total + per_page - 1) // per_page
            self.has_prev = page > 1
            self.has_next = page < self.pages
            self.prev_num = page - 1
            self.next_num = page + 1

    pagination = Pagination(paginated_items, page, per_page, total)

    count_by_nomor_ticket = dict(
        db.session.query(
            Ticket.nomor_ticket_id,
            func.count(Ticket.id)
        ).group_by(Ticket.nomor_ticket_id).all()
    )

    jumlah_tiket_aktif = db.session.query(NomorTicket)\
        .join(Ticket, Ticket.nomor_ticket_id == NomorTicket.id)\
        .filter(
            NomorTicket.id_qc == None
        )\
        .distinct()\
        .count()

    return render_template(
        'ticket_list.html',
        user=current_user,
        tickets=pagination,
        count_by_nomor_ticket=count_by_nomor_ticket,
        jumlah_tiket_aktif=jumlah_tiket_aktif,
        tahapan_options=tahapan_options
    )

@app.route('/list-ticket-info/<int:nomor_ticket_id>')
@login_required
def list_data_info(nomor_ticket_id):
    if current_user.role != 'admin':
        return redirect(request.referrer)
    
    nomor_ticket = NomorTicket.query.get_or_404(nomor_ticket_id)

    tickets = Ticket.query.filter_by(nomor_ticket_id=nomor_ticket_id)\
        .order_by(Ticket.created_time.asc()).all()
    
    catatan_list = Catatan.query.filter_by(nomor_ticket_id=nomor_ticket_id)\
        .order_by(Catatan.tanggal.desc()).all()

    qc_users = User.query.filter_by(role='qc').all()

    jenis_pengaduan_map = {
        1: "Informasi Pengajuan",
        2: "Permintaan Kode OTP",
        3: "Informasi Tenor",
        4: "Informasi Tagihan",
        5: "Informasi Denda",
        6: "Pembatalan Pinjaman",
        7: "Informasi Pencairan Dana",
        8: "Perilaku Petugas Penagihan",
        9: "Informasi Pembayaran",
        10: "Discount / Pemutihan"
    }

    detail_pengaduan_map = {
        1: [
            "Hasil Pengajuan",
            "Pengajuan Ditolak",
            "Status Pengajuan sedang ditransfer",
            "Tidak bisa pengajuan ulang karena keterlambatan",
            "Verifikasi Bank gagal",
            "Verifikasi KTP gagal",
            "Cara pengajuan",
            "Perubahan Nomor Handphone",
            "Perubahan Nomor Rekening"
        ],
        2: [
            "OTP Limit",
            "Tidak terima SMS OTP"
        ],
        3: [
            "Informasi Pinjaman"
        ],
        4: [
            "Konsultasi detail pinjaman saat ini",
            "Konsultasi Perpanjangan",
            "Bukti Transfer"
        ],
        5: [
            "Denda Keterlambatan"
        ],
        6: [
            "Hapus Data (Penutupan Akun)",
            "Pembatalan Pinjaman"
        ],
        7: [
            "Pencairan dana berhasil",
            "Status pencairan dana gagal",
            "Status pengajuan pencairan dana ulang",
            "Tidak terima dana",
            "Operasi Gagal (tidak bisa verifikasi wajah dan KTP)"
        ],
        8: [
            "Keluhan Penagihan",
            "Keluhan Reminder",
            "Penipuan"
        ],
        9: [
            "Konfirmasi Pembayaran",
            "Pembayaran belum masuk",
            "Pembayaran bukan ke VA UATAS",
            "Refund (pembayaran double)",
            "Meminta VA (cicilan)",
            "Meminta VA (pelunasan)",
            "Meminta VA (perpanjangan)",
            "Tidak bisa ambil VA"
        ],
        10: [
            "Meminta keringanan pembayaran (cicilan)",
            "Meminta keringanan pembayaran (potongan denda)",
            "Tidak ada dana"
        ]
    }
    qc_users = User.query.filter_by(role='qc').all()

    return render_template(
        'ticket_list_info.html',
        nomor_ticket=nomor_ticket,
        tickets=tickets,
        user=current_user,
        jenis_pengaduan_map=jenis_pengaduan_map,
        detail_pengaduan_map=detail_pengaduan_map,
        qc_users=qc_users,
        catatan_list=catatan_list
    )

@app.route('/kanal-email')
@login_required
def kanal_email():
    if current_user.role != 'admin':
        flash('Akses ditolak: Anda bukan Admin.')
        return redirect(url_for('admin_dashboard'))

    range1 = request.args.get('range1', '')
    range2 = request.args.get('range2', '')
    selected_os = request.args.getlist('os')
    selected_bucket = request.args.getlist('bucket')
    selected_jenis_pengaduan = request.args.getlist('jenis_pengaduan')
    chart_by = request.args.get('chart_by')

    all_os = [r[0] for r in db.session.query(Ticket.nama_os).distinct().all() if r[0] is not None]
    all_bucket = [r[0] for r in db.session.query(Ticket.nama_bucket).distinct().all() if r[0] is not None]

    apply_os = len(selected_os) > 0
    apply_bucket = len(selected_bucket) > 0
    jenis_pengaduan_filter = [int(j) for j in selected_jenis_pengaduan if j and j.strip().isdigit()]
    apply_jenis = len(jenis_pengaduan_filter) > 0

    KANAL_CONST = "EMAIL"

    def parse_range(date_range):
        try:
            start, end = date_range.split(' - ')
            return start.strip(), end.strip()
        except ValueError:
            return None, None

    range1_start, range1_end = parse_range(range1)
    range2_start, range2_end = parse_range(range2)

    if not range1_start or not range1_end:
        min_tanggal = db.session.query(func.min(Ticket.tanggal)).scalar()
        max_tanggal = db.session.query(func.max(Ticket.tanggal)).scalar()
        if min_tanggal and max_tanggal:
            range1_start = min_tanggal.strftime('%Y-%m-%d')
            range1_end = max_tanggal.strftime('%Y-%m-%d')

    start_date = datetime.strptime(range1_start, "%Y-%m-%d") if range1_start else None
    end_date = datetime.strptime(range1_end, "%Y-%m-%d") if range1_end else None
    start_date2 = datetime.strptime(range2_start, "%Y-%m-%d") if range2_start else None
    end_date2 = datetime.strptime(range2_end, "%Y-%m-%d") if range2_end else None

    def format_tanggal_indonesia(tanggal):
        if not tanggal:
            return ""
        bulan_dict = {
            1: 'Januari', 2: 'Februari', 3: 'Maret', 4: 'April',
            5: 'Mei', 6: 'Juni', 7: 'Juli', 8: 'Agustus',
            9: 'September', 10: 'Oktober', 11: 'November', 12: 'Desember'
        }
        return f"{tanggal.day} {bulan_dict[tanggal.month]} {tanggal.year}"

    def nomor_ticket_ids(start, end):
        q = (
            db.session.query(NomorTicket.id.label('id'))
            .join(Ticket, Ticket.nomor_ticket_id == NomorTicket.id)
            .filter(func.upper(Ticket.kanal_pengaduan) == KANAL_CONST)
        )
        if start and end:
            q = q.filter(Ticket.tanggal.between(start, end))
        if apply_os:
            q = q.filter(Ticket.nama_os.in_(selected_os))
        if apply_bucket:
            q = q.filter(Ticket.nama_bucket.in_(selected_bucket))
        if apply_jenis:
            q = q.filter(Ticket.jenis_pengaduan.in_(jenis_pengaduan_filter))
        return q.distinct()

    nt_subq1 = nomor_ticket_ids(start_date, end_date)

    if start_date2 and end_date2:
        nt_subq2 = nomor_ticket_ids(start_date2, end_date2)
        nt_union_subq = union(nt_subq1, nt_subq2).alias('nt_union')
    else:
        nt_union_subq = nt_subq1.subquery('nt_union')

    base_ids = db.session.query(nt_union_subq.c.id)

    total_nomor_ticket = db.session.query(func.count(nt_union_subq.c.id)).scalar()

    total_open = (
        db.session.query(func.count(NomorTicket.id))
        .filter(
            NomorTicket.id.in_(base_ids),
            NomorTicket.status.in_(["aktif", "Reopen"])
        )
        .scalar()
    )

    total_close = (
        db.session.query(func.count(NomorTicket.id))
        .filter(
            NomorTicket.id.in_(base_ids),
            NomorTicket.status == "close"
        )
        .scalar()
    )

    total_valid = (
        db.session.query(func.count(NomorTicket.id))
        .filter(
            NomorTicket.id.in_(base_ids),
            NomorTicket.label_case == "valid"
        )
        .scalar()
    )

    def get_chart_data(start_date, end_date):
        cq = db.session.query(Ticket.nama_os, func.count(Ticket.id)).filter(
            func.upper(Ticket.kanal_pengaduan) == KANAL_CONST
        )
        if start_date and end_date:
            cq = cq.filter(Ticket.tanggal.between(start_date, end_date))
        if apply_os:
            cq = cq.filter(Ticket.nama_os.in_(selected_os))
        if apply_jenis:
            cq = cq.filter(Ticket.jenis_pengaduan.in_(jenis_pengaduan_filter))
        if apply_bucket:
            cq = cq.filter(Ticket.nama_bucket.in_(selected_bucket))

        result = {}
        for os_name, cnt in cq.group_by(Ticket.nama_os).all():
            label = os_name or "Tidak Diketahui"
            result[label] = cnt
        return result

    def get_chart_data_bucket(start_date, end_date):
        cq = db.session.query(Ticket.nama_bucket, func.count(Ticket.id)).filter(
            func.upper(Ticket.kanal_pengaduan) == KANAL_CONST
        )
        if start_date and end_date:
            cq = cq.filter(Ticket.tanggal.between(start_date, end_date))
        if apply_os:
            cq = cq.filter(Ticket.nama_os.in_(selected_os))
        if apply_bucket:
            cq = cq.filter(Ticket.nama_bucket.in_(selected_bucket))
        if apply_jenis:
            cq = cq.filter(Ticket.jenis_pengaduan.in_(jenis_pengaduan_filter))

        result = {}
        for bucket_name, cnt in cq.group_by(Ticket.nama_bucket).all():
            label = bucket_name or "Tidak Diketahui"
            result[label] = cnt
        return result

    def get_jp_chart_data_filtered(start_date, end_date):
        jpq = (
            db.session.query(
                Ticket.jenis_pengaduan,
                func.count(func.distinct(Ticket.nomor_ticket_id))
            )
            .join(NomorTicket, NomorTicket.id == Ticket.nomor_ticket_id)
            .filter(func.upper(Ticket.kanal_pengaduan) == KANAL_CONST)
            .filter(Ticket.nomor_ticket_id.isnot(None))
        )

        if start_date and end_date:
            jpq = jpq.filter(Ticket.tanggal.between(start_date, end_date))
        if apply_os:
            jpq = jpq.filter(Ticket.nama_os.in_(selected_os))
        if apply_bucket:
            jpq = jpq.filter(Ticket.nama_bucket.in_(selected_bucket))
        if apply_jenis:
            jpq = jpq.filter(Ticket.jenis_pengaduan.in_(jenis_pengaduan_filter))

        data_dict = {}
        for jp, cnt in jpq.group_by(Ticket.jenis_pengaduan).all():
            if jp is None:
                continue
            try:
                key = str(int(jp))
            except (ValueError, TypeError):
                key = str(jp)
            data_dict[key] = int(cnt)

        return data_dict

    data_range1 = get_chart_data(start_date, end_date)
    data_range2 = get_chart_data(start_date2, end_date2) if (start_date2 and end_date2) else {}
    chart_labels = sorted(lbl for lbl in set(data_range1) | set(data_range2) if lbl != "Tidak Diketahui")
    chart_series = [{
        "name": f"Range 1 ({format_tanggal_indonesia(start_date)} - {format_tanggal_indonesia(end_date)})",
        "data": [data_range1.get(lbl, 0) for lbl in chart_labels]
    }]
    if start_date2 and end_date2:
        chart_series.append({
            "name": f"Range 2 ({format_tanggal_indonesia(start_date2)} - {format_tanggal_indonesia(end_date2)})",
            "data": [data_range2.get(lbl, 0) for lbl in chart_labels]
        })

    jp_data_1 = get_jp_chart_data_filtered(start_date, end_date) if (start_date and end_date) else {}
    jp_data_2 = get_jp_chart_data_filtered(start_date2, end_date2) if (start_date2 and end_date2) else {}
    all_jenis_ids = sorted(set(jp_data_1.keys()) | set(jp_data_2.keys()))
    jenis_pengaduan_labels = {
        1: "Informasi Pengajuan",
        2: "Permintaan Kode OTP",
        3: "Informasi Tenor",
        4: "Informasi Tagihan",
        5: "Informasi Denda",
        6: "Pembatalan Pinjaman",
        7: "Informasi Pencairan Dana",
        8: "Perilaku Petugas Penagihan",
        9: "Informasi Pembayaran",
        10: "Discount / Pemutihan"
    }
    jenis_pengaduan_chart_labels = [jenis_pengaduan_labels.get(jid, str(jid)) for jid in all_jenis_ids]
    jenis_pengaduan_chart_series = []
    if start_date and end_date:
        jenis_pengaduan_chart_series.append({
            "name": f"Range 1 ({format_tanggal_indonesia(start_date)} - {format_tanggal_indonesia(end_date)})",
            "data": [jp_data_1.get(jid, 0) for jid in all_jenis_ids]
        })
    if start_date2 and end_date2:
        jenis_pengaduan_chart_series.append({
            "name": f"Range 2 ({format_tanggal_indonesia(start_date2)} - {format_tanggal_indonesia(end_date2)})",
            "data": [jp_data_2.get(jid, 0) for jid in all_jenis_ids]
        })

    data_bucket_range1 = get_chart_data_bucket(start_date, end_date)
    data_bucket_range2 = get_chart_data_bucket(start_date2, end_date2) if (start_date2 and end_date2) else {}
    bucket_chart_labels = sorted(lbl for lbl in set(data_bucket_range1) | set(data_bucket_range2) if lbl != "Tidak Diketahui")
    bucket_chart_series = [{
        "name": f"Range 1 ({format_tanggal_indonesia(start_date)} - {format_tanggal_indonesia(end_date)})",
        "data": [data_bucket_range1.get(lbl, 0) for lbl in bucket_chart_labels]
    }]
    if start_date2 and end_date2:
        bucket_chart_series.append({
            "name": f"Range 2 ({format_tanggal_indonesia(start_date2)} - {format_tanggal_indonesia(end_date2)})",
            "data": [data_bucket_range2.get(lbl, 0) for lbl in bucket_chart_labels]
        })

    title_base = "Jumlah Tiket per OS"
    range1_txt = f"{format_tanggal_indonesia(start_date)} - {format_tanggal_indonesia(end_date)}" if (start_date and end_date) else None
    range2_txt = f"{format_tanggal_indonesia(start_date2)} - {format_tanggal_indonesia(end_date2)}" if (start_date2 and end_date2) else None

    if range1_txt and range2_txt:
        chart_title = f"{title_base} ({range1_txt} | {range2_txt})"
    elif range1_txt:
        chart_title = f"{title_base} ({range1_txt})"
    else:
        chart_title = title_base

    bucket_chart_title = chart_title.replace("Tiket per OS", "Order per Bucket")

    return render_template(
        'filtering_email.html',
        range1=range1,
        range2=range2,
        tanggal_awal2=format_tanggal_indonesia(start_date2) if start_date2 else None,
        tanggal_akhir2=format_tanggal_indonesia(end_date2) if end_date2 else None,
        user=current_user,
        total_nomor_ticket=total_nomor_ticket,
        total_open=total_open,
        total_close=total_close,
        total_valid=total_valid,
        selected_os=selected_os,
        selected_bucket=selected_bucket,
        all_os=all_os,
        all_bucket=all_bucket,
        chart_labels=chart_labels,
        chart_series=chart_series,
        chart_title=chart_title,
        bucket_chart_title=bucket_chart_title,
        jenis_pengaduan_chart_labels=jenis_pengaduan_chart_labels,
        jenis_pengaduan_chart_series=jenis_pengaduan_chart_series,
        ticket_chart_title="Jumlah Ticket per Status - Email",
        tanggal_awal=format_tanggal_indonesia(start_date) if start_date else None,
        tanggal_akhir=format_tanggal_indonesia(end_date) if end_date else None,
        bucket_chart_labels=bucket_chart_labels,
        bucket_chart_series=bucket_chart_series
    )

@app.route('/kanal-whatsapp')
@login_required
def kanal_whatsapp():
    if current_user.role != 'admin':
        flash('Akses ditolak: Anda bukan Admin.')
        return redirect(url_for('admin_dashboard'))

    range1 = request.args.get('range1', '')
    range2 = request.args.get('range2', '')
    selected_os = request.args.getlist('os')
    selected_bucket = request.args.getlist('bucket')
    selected_jenis_pengaduan = request.args.getlist('jenis_pengaduan')
    chart_by = request.args.get('chart_by')

    all_os = [r[0] for r in db.session.query(Ticket.nama_os).distinct().all() if r[0] is not None]
    all_bucket = [r[0] for r in db.session.query(Ticket.nama_bucket).distinct().all() if r[0] is not None]

    apply_os = len(selected_os) > 0
    apply_bucket = len(selected_bucket) > 0
    jenis_pengaduan_filter = [int(j) for j in selected_jenis_pengaduan if j and j.strip().isdigit()]
    apply_jenis = len(jenis_pengaduan_filter) > 0

    KANAL_CONST = "WHATSAPP"

    def parse_range(date_range):
        try:
            start, end = date_range.split(' - ')
            return start.strip(), end.strip()
        except ValueError:
            return None, None

    range1_start, range1_end = parse_range(range1)
    range2_start, range2_end = parse_range(range2)

    if not range1_start or not range1_end:
        min_tanggal = db.session.query(func.min(Ticket.tanggal)).scalar()
        max_tanggal = db.session.query(func.max(Ticket.tanggal)).scalar()
        if min_tanggal and max_tanggal:
            range1_start = min_tanggal.strftime('%Y-%m-%d')
            range1_end = max_tanggal.strftime('%Y-%m-%d')

    start_date = datetime.strptime(range1_start, "%Y-%m-%d") if range1_start else None
    end_date = datetime.strptime(range1_end, "%Y-%m-%d") if range1_end else None
    start_date2 = datetime.strptime(range2_start, "%Y-%m-%d") if range2_start else None
    end_date2 = datetime.strptime(range2_end, "%Y-%m-%d") if range2_end else None

    def format_tanggal_indonesia(tanggal):
        if not tanggal:
            return ""
        bulan_dict = {
            1: 'Januari', 2: 'Februari', 3: 'Maret', 4: 'April',
            5: 'Mei', 6: 'Juni', 7: 'Juli', 8: 'Agustus',
            9: 'September', 10: 'Oktober', 11: 'November', 12: 'Desember'
        }
        return f"{tanggal.day} {bulan_dict[tanggal.month]} {tanggal.year}"

    def nomor_ticket_ids(start, end):
        q = (
            db.session.query(NomorTicket.id.label('id'))
            .join(Ticket, Ticket.nomor_ticket_id == NomorTicket.id)
            .filter(func.upper(Ticket.kanal_pengaduan) == KANAL_CONST)
        )
        if start and end:
            q = q.filter(Ticket.tanggal.between(start, end))
        if apply_os:
            q = q.filter(Ticket.nama_os.in_(selected_os))
        if apply_bucket:
            q = q.filter(Ticket.nama_bucket.in_(selected_bucket))
        if apply_jenis:
            q = q.filter(Ticket.jenis_pengaduan.in_(jenis_pengaduan_filter))
        return q.distinct()

    nt_subq1 = nomor_ticket_ids(start_date, end_date)

    if start_date2 and end_date2:
        nt_subq2 = nomor_ticket_ids(start_date2, end_date2)
        nt_union_subq = union(nt_subq1, nt_subq2).alias('nt_union')
    else:
        nt_union_subq = nt_subq1.subquery('nt_union')

    base_ids = db.session.query(nt_union_subq.c.id)

    total_nomor_ticket = db.session.query(func.count(nt_union_subq.c.id)).scalar()

    total_open = (
        db.session.query(func.count(NomorTicket.id))
        .filter(
            NomorTicket.id.in_(base_ids),
            NomorTicket.status.in_(["aktif", "Reopen"])
        )
        .scalar()
    )

    total_close = (
        db.session.query(func.count(NomorTicket.id))
        .filter(
            NomorTicket.id.in_(base_ids),
            NomorTicket.status == "close"
        )
        .scalar()
    )

    total_valid = (
        db.session.query(func.count(NomorTicket.id))
        .filter(
            NomorTicket.id.in_(base_ids),
            NomorTicket.label_case == "valid"
        )
        .scalar()
    )

    def get_chart_data(start_date, end_date):
        cq = db.session.query(Ticket.nama_os, func.count(Ticket.id)).filter(
            func.upper(Ticket.kanal_pengaduan) == KANAL_CONST
        )
        if start_date and end_date:
            cq = cq.filter(Ticket.tanggal.between(start_date, end_date))
        if apply_os:
            cq = cq.filter(Ticket.nama_os.in_(selected_os))
        if apply_jenis:
            cq = cq.filter(Ticket.jenis_pengaduan.in_(jenis_pengaduan_filter))
        if apply_bucket:
            cq = cq.filter(Ticket.nama_bucket.in_(selected_bucket))

        result = {}
        for os_name, cnt in cq.group_by(Ticket.nama_os).all():
            label = os_name or "Tidak Diketahui"
            result[label] = cnt
        return result

    def get_chart_data_bucket(start_date, end_date):
        cq = db.session.query(Ticket.nama_bucket, func.count(Ticket.id)).filter(
            func.upper(Ticket.kanal_pengaduan) == KANAL_CONST
        )
        if start_date and end_date:
            cq = cq.filter(Ticket.tanggal.between(start_date, end_date))
        if apply_os:
            cq = cq.filter(Ticket.nama_os.in_(selected_os))
        if apply_bucket:
            cq = cq.filter(Ticket.nama_bucket.in_(selected_bucket))
        if apply_jenis:
            cq = cq.filter(Ticket.jenis_pengaduan.in_(jenis_pengaduan_filter))

        result = {}
        for bucket_name, cnt in cq.group_by(Ticket.nama_bucket).all():
            label = bucket_name or "Tidak Diketahui"
            result[label] = cnt
        return result

    def get_jp_chart_data_filtered(start_date, end_date):
        """
        Hasil key disimpan sebagai STRING (mis. "1","2",...) supaya aman untuk JSON & Apex.
        """
        jpq = (
            db.session.query(
                Ticket.jenis_pengaduan,
                func.count(func.distinct(Ticket.nomor_ticket_id))
            )
            .join(NomorTicket, NomorTicket.id == Ticket.nomor_ticket_id)
            .filter(func.upper(Ticket.kanal_pengaduan) == KANAL_CONST)
            .filter(Ticket.nomor_ticket_id.isnot(None))
        )

        if start_date and end_date:
            jpq = jpq.filter(Ticket.tanggal.between(start_date, end_date))
        if apply_os:
            jpq = jpq.filter(Ticket.nama_os.in_(selected_os))
        if apply_bucket:
            jpq = jpq.filter(Ticket.nama_bucket.in_(selected_bucket))
        if apply_jenis:
            jpq = jpq.filter(Ticket.jenis_pengaduan.in_(jenis_pengaduan_filter))

        data_dict = {}
        for jp, cnt in jpq.group_by(Ticket.jenis_pengaduan).all():
            if jp is None:
                continue
            try:
                key = str(int(jp))
            except (ValueError, TypeError):
                key = str(jp)
            data_dict[key] = int(cnt)

        return data_dict

    data_range1 = get_chart_data(start_date, end_date)
    data_range2 = get_chart_data(start_date2, end_date2) if (start_date2 and end_date2) else {}
    chart_labels = sorted(lbl for lbl in set(data_range1) | set(data_range2) if lbl != "Tidak Diketahui")
    chart_series = [{
        "name": f"Range 1 ({format_tanggal_indonesia(start_date)} - {format_tanggal_indonesia(end_date)})",
        "data": [data_range1.get(lbl, 0) for lbl in chart_labels]
    }]
    if start_date2 and end_date2:
        chart_series.append({
            "name": f"Range 2 ({format_tanggal_indonesia(start_date2)} - {format_tanggal_indonesia(end_date2)})",
            "data": [data_range2.get(lbl, 0) for lbl in chart_labels]
        })

    jp_data_1 = get_jp_chart_data_filtered(start_date, end_date) if (start_date and end_date) else {}
    jp_data_2 = get_jp_chart_data_filtered(start_date2, end_date2) if (start_date2 and end_date2) else {}

    all_jenis_ids = sorted(
        set(jp_data_1.keys()) | set(jp_data_2.keys()),
        key=lambda x: (0, int(x)) if str(x).isdigit() else (1, str(x))
    )

    jenis_pengaduan_labels = {
        1: "Informasi Pengajuan",
        2: "Permintaan Kode OTP",
        3: "Informasi Tenor",
        4: "Informasi Tagihan",
        5: "Informasi Denda",
        6: "Pembatalan Pinjaman",
        7: "Informasi Pencairan Dana",
        8: "Perilaku Petugas Penagihan",
        9: "Informasi Pembayaran",
        10: "Discount / Pemutihan"
    }

    jenis_pengaduan_chart_labels = []
    for jid in all_jenis_ids:
        label = None
        if str(jid).isdigit():
            label = jenis_pengaduan_labels.get(int(jid))
        jenis_pengaduan_chart_labels.append(label if label else str(jid))

    jenis_pengaduan_chart_series = []
    if start_date and end_date:
        jenis_pengaduan_chart_series.append({
            "name": f"Range 1 ({format_tanggal_indonesia(start_date)} - {format_tanggal_indonesia(end_date)})",
            "data": [jp_data_1.get(jid, 0) for jid in all_jenis_ids]
        })
    if start_date2 and end_date2:
        jenis_pengaduan_chart_series.append({
            "name": f"Range 2 ({format_tanggal_indonesia(start_date2)} - {format_tanggal_indonesia(end_date2)})",
            "data": [jp_data_2.get(jid, 0) for jid in all_jenis_ids]
        })

    data_bucket_range1 = get_chart_data_bucket(start_date, end_date)
    data_bucket_range2 = get_chart_data_bucket(start_date2, end_date2) if (start_date2 and end_date2) else {}
    bucket_chart_labels = sorted(lbl for lbl in set(data_bucket_range1) | set(data_bucket_range2) if lbl != "Tidak Diketahui")
    bucket_chart_series = [{
        "name": f"Range 1 ({format_tanggal_indonesia(start_date)} - {format_tanggal_indonesia(end_date)})",
        "data": [data_bucket_range1.get(lbl, 0) for lbl in bucket_chart_labels]
    }]
    if start_date2 and end_date2:
        bucket_chart_series.append({
            "name": f"Range 2 ({format_tanggal_indonesia(start_date2)} - {format_tanggal_indonesia(end_date2)})",
            "data": [data_bucket_range2.get(lbl, 0) for lbl in bucket_chart_labels]
        })

    title_base = "Jumlah Tiket per OS"
    range1_txt = f"{format_tanggal_indonesia(start_date)} - {format_tanggal_indonesia(end_date)}" if (start_date and end_date) else None
    range2_txt = f"{format_tanggal_indonesia(start_date2)} - {format_tanggal_indonesia(end_date2)}" if (start_date2 and end_date2) else None

    if range1_txt and range2_txt:
        chart_title = f"{title_base} ({range1_txt} | {range2_txt})"
    elif range1_txt:
        chart_title = f"{title_base} ({range1_txt})"
    else:
        chart_title = title_base

    bucket_chart_title = chart_title.replace("Tiket per OS", "Order per Bucket")

    return render_template(
        'filtering_wa.html',
        range1=range1,
        range2=range2,
        tanggal_awal2=format_tanggal_indonesia(start_date2) if start_date2 else None,
        tanggal_akhir2=format_tanggal_indonesia(end_date2) if end_date2 else None,
        user=current_user,
        total_nomor_ticket=total_nomor_ticket,
        total_open=total_open,
        total_close=total_close,
        total_valid=total_valid,
        selected_os=selected_os,
        selected_bucket=selected_bucket,
        all_os=all_os,
        all_bucket=all_bucket,
        chart_labels=chart_labels,
        chart_series=chart_series,
        chart_title=chart_title,
        bucket_chart_title=bucket_chart_title,
        jenis_pengaduan_chart_labels=jenis_pengaduan_chart_labels,
        jenis_pengaduan_chart_series=jenis_pengaduan_chart_series,
        ticket_chart_title="Jumlah Ticket per Status - WhatsApp",
        tanggal_awal=format_tanggal_indonesia(start_date) if start_date else None,
        tanggal_akhir=format_tanggal_indonesia(end_date) if end_date else None,
        bucket_chart_labels=bucket_chart_labels,
        bucket_chart_series=bucket_chart_series
    )

if __name__ == '__main__':
    if not is_running_from_reloader():
        if not scheduler.running:
            scheduler.start()
    with app.app_context():
        db.create_all()
    app.run(debug=True, port=5007, host='0.0.0.0')