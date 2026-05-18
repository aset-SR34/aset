import pandas as pd
from flask import send_file
from openpyxl import Workbook
import io
from datetime import datetime, timedelta

from flask import session, flash
from flask import Flask, render_template, request, redirect, send_file, url_for
from datetime import datetime   # ← TAMBAHKAN INI
import sqlite3
from datetime import datetime
tanggal_input = datetime.now()

app = Flask(__name__)
def is_admin():
    return session.get('role') == 'admin'
app.secret_key = "kunci-rahasia-super-aman"
app.permanent_session_lifetime = timedelta(minutes=5)

def get_db():
    conn = sqlite3.connect("aset.db", timeout=10)
    conn.row_factory = sqlite3.Row
    return conn

conn = get_db()

conn.execute("""CREATE TABLE IF NOT EXISTS ruang (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nama_ruang TEXT NOT NULL
)""")

conn.execute("""CREATE TABLE IF NOT EXISTS barang (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    kode_barang TEXT,
    uraian TEXT,
    tipe TEXT,
    merk TEXT,
    tahun_perolehan INTEGER,
    nilai INTEGER,
    ruang_id INTEGER
)""")

conn.execute("""CREATE TABLE IF NOT EXISTS pemeriksaan_fisik (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    barang_id INTEGER,
    jumlah_fisik INTEGER,
    selisih INTEGER,
    bb INTEGER DEFAULT 0,
    rr INTEGER DEFAULT 0,
    rb INTEGER DEFAULT 0,
    hl INTEGER DEFAULT 0,
    tanggal DATE DEFAULT CURRENT_DATE
)""")

conn.commit()
conn.close()

# membuat tabel jika belum ada
conn = get_db()
conn.execute("""
CREATE TABLE IF NOT EXISTS aset (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nama TEXT,
    jumlah INTEGER,
    lokasi TEXT
)
""")
conn.commit()
conn.close()

@app.before_request
def session_timeout():

    # hanya berlaku untuk admin
    if session.get('role') == 'admin':

        now = datetime.now()

        # cek waktu terakhir aktivitas
        last_activity = session.get('last_activity')

        if last_activity:
            last_activity = datetime.strptime(last_activity, "%Y-%m-%d %H:%M:%S")

            # hitung selisih waktu
            if now - last_activity > timedelta(minutes=5):
                session.clear()
                return redirect(url_for('login'))

        # update waktu aktivitas
        session['last_activity'] = now.strftime("%Y-%m-%d %H:%M:%S")

@app.route('/')
def index():
    conn = get_db()

    total_barang = conn.execute(
        "SELECT COUNT(*) FROM barang"
    ).fetchone()[0]

    total_ruang = conn.execute(
        "SELECT COUNT(*) FROM ruang"
    ).fetchone()[0]

    total_aset = conn.execute(
        "SELECT SUM(jumlah) FROM aset"
    ).fetchone()[0]

    grafik = conn.execute("""
        SELECT ruang.nama_ruang, SUM(aset.jumlah) as total
        FROM aset
        JOIN ruang ON aset.ruang_id = ruang.id
        GROUP BY ruang.nama_ruang
    """).fetchall()

    # 🔹 ambil kondisi barang (aman walau tabel kosong)
    kondisi = conn.execute("""
        SELECT
            COALESCE(SUM(bb),0) as bb,
            COALESCE(SUM(rr),0) as rr,
            COALESCE(SUM(rb),0) as rb,
            COALESCE(SUM(hl),0) as hl
        FROM pemeriksaan_fisik
    """).fetchone()

    conn.close()

    labels = [g["nama_ruang"] for g in grafik]
    values = [g["total"] for g in grafik]

    kondisi_labels = ["Baik", "Rusak Ringan", "Rusak Berat", "Hilang"]
    kondisi_values = [
        kondisi["bb"],
        kondisi["rr"],
        kondisi["rb"],
        kondisi["hl"]
    ]

    # 🔹 HITUNG STATUS (AMAN DARI ERROR)
    total_kondisi = sum(kondisi_values)

    if total_kondisi == 0:
        status = "Belum Ada Data"
        warna = "#95a5a6"
    else:
        baik_persen = kondisi["bb"] / total_kondisi * 100

        if kondisi["hl"] > 0 or kondisi["rb"] > kondisi["bb"]:
            status = "KRITIS"
            warna = "#e74c3c"
        elif baik_persen < 50:
            status = "PERLU PERBAIKAN"
            warna = "#e67e22"
        elif baik_persen < 75:
            status = "PERLU PERAWATAN"
            warna = "#f39c12"
        else:
            status = "AMAN"
            warna = "#2ecc71"

    return render_template(
        "dashboard.html",
        total_barang=total_barang,
        total_ruang=total_ruang,
        total_aset=total_aset,
        labels=labels,
        values=values,
        kondisi_labels=kondisi_labels,
        kondisi_values=kondisi_values,
        status=status,
        warna=warna
    )
    
@app.route('/login', methods=['GET','POST'])
def login():

    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']

        conn = get_db()

        user = conn.execute(
            "SELECT * FROM users WHERE username=? AND password=? AND role='admin'",
            (username, password)
        ).fetchone()

        conn.close()

        if user:
            session['role'] = 'admin'
            session['user_id'] = user['id']
            session['last_activity'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            return redirect('/')
        from flask import flash
        flash("Username atau password salah!", "error")
        return redirect(url_for('login'))

    return render_template('login.html')

@app.route('/logout')
def logout():
    flash("⏰ Session habis! Silakan login kembali.", "error")
    session.clear()
    return redirect('/')

@app.route('/barang', methods=['GET', 'POST'])
def barang():

    conn = get_db()

    # =====================================
    # TAMBAH BARANG
    # =====================================
    if request.method == 'POST':
        kode = request.form['kode_barang']
        nama = request.form['nama_barang']
        jumlah = request.form['jumlah_barang']
        tipe = request.form['tipe']
        merk = request.form['merk']
        tahun = request.form['tahun_perolehan']
        nilai = request.form['nilai']

        conn.execute("""
            INSERT INTO barang
            (kode_barang, nama_barang, jumlah_barang,
             tipe, merk, tahun_perolehan, nilai)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (kode, nama, jumlah, tipe, merk, tahun, nilai))

        conn.commit()

        return redirect(url_for('barang'))


    # =====================================
    # FILTER KODE BARANG
    # =====================================
    kode_filter = request.args.get('kode_filter', '')


    # =====================================
    # PAGINATION
    # =====================================
    page = request.args.get('page', 1, type=int)
    per_page = 10
    offset = (page - 1) * per_page


    query = "SELECT * FROM barang"
    params = []

    # ✅ jika filter aktif
    if kode_filter:
        query += " WHERE kode_barang LIKE ? "
        params.append(f"%{kode_filter}%")

    # =====================================
    # HITUNG TOTAL DATA
    # =====================================
    total = conn.execute(
        query.replace("*", "COUNT(*)"),
        params
    ).fetchone()[0]

    total_pages = (total + per_page - 1) // per_page


    # =====================================
    # AMBIL DATA PAGINATION
    # =====================================
    query += " ORDER BY id DESC LIMIT ? OFFSET ?"
    params.extend([per_page, offset])

    data = conn.execute(query, params).fetchall()

    conn.close()

    return render_template(
        "barang.html",
        data=data,
        page=page,
        total_pages=total_pages,
        kode_filter=kode_filter
    )

@app.route('/barang/tambah', methods=['POST'])
def tambah_barang():
    if not is_admin():
        flash("Akses ditolak!", "error")
        return redirect(url_for('barang'))
    conn = get_db()

    kode = request.form.get('kode_barang')
    nama = request.form.get('nama_barang')
    jumlah = int(request.form.get('jumlah_barang') or 0)
    tipe = request.form.get('tipe')
    merk = request.form.get('merk')
    tahun = request.form.get('tahun_perolehan')
    nilai = request.form.get('nilai')

    conn.execute("""
        INSERT INTO barang (
            kode_barang,
            nama_barang,
            jumlah_barang,
            tipe,
            merk,
            tahun_perolehan,
            nilai,
            tanggal_input
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
    """, (kode, nama, jumlah, tipe, merk, tahun, nilai))

    conn.commit()
    conn.close()

    return redirect(url_for('barang'))

@app.route('/barang/update/<int:id>', methods=['POST'])
def update_barang_data(id):
    if not is_admin():
        flash("Akses ditolak!", "error")
        return redirect(url_for('barang'))
    conn = get_db()

    kode = request.form.get('kode_barang')
    nama = request.form.get('nama_barang')
    jumlah = int(request.form.get('jumlah_barang') or 0)
    tipe = request.form.get('tipe')
    merk = request.form.get('merk')
    tahun = request.form.get('tahun_perolehan')
    nilai = request.form.get('nilai')

    conn.execute("""
        UPDATE barang
        SET
            kode_barang=?,
            nama_barang=?,
            jumlah_barang=?,
            tipe=?,
            merk=?,
            tahun_perolehan=?,
            nilai=?,
            tanggal_update=CURRENT_TIMESTAMP
        WHERE id=?
    """, (
        kode,
        nama,
        jumlah,
        tipe,
        merk,
        tahun,
        nilai,
        id
    ))

    conn.commit()
    conn.close()

    return redirect(url_for('barang'))

@app.route('/hapus/<int:id>')
def hapus(id):
    if not is_admin():
        flash("Akses ditolak!", "error")
        return redirect(url_for('barang'))
    conn = get_db()
    conn.execute("DELETE FROM barang WHERE id=?", (id,))
    conn.commit()
    conn.close()
    return redirect('/barang')

@app.route('/edit/<int:id>')
def edit_barang(id):
    if not is_admin():
        return "⛔ Akses ditolak!"
    # ambil data berdasarkan id
    return f"Edit data {id}"

@app.route('/update', methods=['POST'])
def update_barang():
    if not is_admin():
        flash("Akses ditolak!", "error")
        return redirect(url_for('barang'))
    id = request.form['id']
    nama = request.form['nama_barang']
    jumlah = request.form['jumlah_barang']
    tipe = request.form['tipe']
    merk = request.form['merk']
    tahun = request.form['tahun_perolehan']
    nilai = request.form['nilai']

    conn = sqlite3.connect('aset.db')
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE barang
        SET nama_barang=?,
            jumlah_barang=?,
            tipe=?,
            merk=?,
            tahun_perolehan=?,
            nilai=?
        WHERE id=?
    """, (nama, jumlah, tipe, merk, tahun, nilai, id))
    conn.commit()

    return redirect(url_for('barang'))

@app.route('/export')
def export_excel():
    if not is_admin():
        flash("Akses ditolak!", "error")
        return redirect(url_for('barang'))
    conn = get_db()
    data = conn.execute("SELECT * FROM aset").fetchall()
    conn.close()

    df = pd.DataFrame(data, columns=["ID", "Nama Aset", "Jumlah", "Lokasi"])

    file_path = "inventori_aset.xlsx"
    df.to_excel(file_path, index=False)

    return send_file(file_path, as_attachment=True)

@app.route('/periksa', methods=['GET', 'POST'])
def periksa():

    # 🔐 CEK ADMIN
    if request.method == 'POST' and not is_admin():
        flash("Akses ditolak!", "error")
        return redirect(url_for('periksa'))

    conn = get_db()

    # =============================
    # 💾 SIMPAN DATA
    # =============================
    if request.method == 'POST':

        barang_ids = request.form.getlist('barang_id[]')
        ruang_id = request.form.get('ruang_id')

        for barang_id in barang_ids:

            rr = int(request.form.get(f'rr_{barang_id}') or 0)
            rb = int(request.form.get(f'rb_{barang_id}') or 0)
            hl = int(request.form.get(f'hl_{barang_id}') or 0)
            jumlah_fisik = int(request.form.get(f'jumlah_{barang_id}') or 0)

            # 🔥 VALIDASI
            if (rr + rb + hl) > jumlah_fisik:
                flash("Jumlah kondisi tidak boleh melebihi jumlah fisik!", "error")
                return redirect(url_for('periksa', ruang_id=ruang_id))

            # 🔥 HITUNG BB
            bb = jumlah_fisik - (rr + rb + hl)

            # =========================
            # 🔥 AMBIL JUMLAH SISTEM
            # =========================
            jumlah_sistem = conn.execute("""
                SELECT COALESCE(SUM(jumlah),0)
                FROM aset
                WHERE barang_id=? AND ruang_id=?
            """, (barang_id, ruang_id)).fetchone()[0]

            selisih = jumlah_sistem - jumlah_fisik

            # =========================
            # 🔥 UPDATE ASET
            # =========================
            if selisih != 0:
                conn.execute("""
                    UPDATE aset
                    SET jumlah = jumlah - ?
                    WHERE barang_id=? AND ruang_id=?
                """, (selisih, barang_id, ruang_id))

            # =========================
            # 🔥 CEK DATA (TANPA TANGGAL)
            # =========================
            cek = conn.execute("""
                SELECT id FROM pemeriksaan_fisik WHERE barang_id=?
            """, (barang_id,)).fetchone()

            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            # =========================
            # 🔄 UPDATE / INSERT
            # =========================
            if cek:
                conn.execute("""
                    UPDATE pemeriksaan_fisik
                    SET jumlah_fisik=?, bb=?, rr=?, rb=?, hl=?, updated_at=?
                    WHERE barang_id=?
                """, (jumlah_fisik, bb, rr, rb, hl, now, barang_id))

            else:
                conn.execute("""
                    INSERT INTO pemeriksaan_fisik
                    (barang_id, jumlah_fisik, bb, rr, rb, hl, updated_at)
                    VALUES (?,?,?,?,?,?,?)
                """, (barang_id, jumlah_fisik, bb, rr, rb, hl, now))

        conn.commit()

        return redirect(url_for('periksa', ruang_id=ruang_id))

    # =============================
    # 📊 TAMPIL DATA
    # =============================
    ruang = conn.execute("SELECT * FROM ruang ORDER BY nama_ruang").fetchall()

    ruang_id = request.args.get('ruang_id')

    barang = []

    if ruang_id:
        barang = conn.execute("""
            SELECT
                barang.id as barang_id,
                barang.nama_barang,

                COALESCE(pf.jumlah_fisik, SUM(aset.jumlah)) as jumlah,

                COALESCE(pf.bb,0) as bb,
                COALESCE(pf.rr,0) as rr,
                COALESCE(pf.rb,0) as rb,
                COALESCE(pf.hl,0) as hl

            FROM aset
            JOIN barang ON aset.barang_id = barang.id

            LEFT JOIN pemeriksaan_fisik pf
                ON barang.id = pf.barang_id

            WHERE aset.ruang_id = ?

            GROUP BY barang.id
            ORDER BY barang.nama_barang
        """, (ruang_id,)).fetchall()

    conn.close()

    return render_template(
        "periksa.html",
        ruang=ruang,
        barang=barang,
        ruang_id=ruang_id
    )

@app.route('/pemeriksaan/cetak/<int:ruang_id>')
def cetak_pemeriksaan(ruang_id):
    if not is_admin():
        flash("Akses ditolak!", "error")
        return redirect(url_for('periksa'))

    conn = get_db()

    data = conn.execute("""
        SELECT 
            barang.nama_barang,
            aset.jumlah,
            COALESCE(pf.rr,0) as rr,
            COALESCE(pf.rb,0) as rb,
            COALESCE(pf.hl,0) as hl
        FROM aset
        JOIN barang ON aset.barang_id = barang.id
        LEFT JOIN pemeriksaan_fisik pf 
            ON barang.id = pf.barang_id
        WHERE aset.ruang_id = ?
        AND (pf.rr > 0 OR pf.rb > 0 OR pf.hl > 0)
    """, (ruang_id,)).fetchall()

    ruang = conn.execute(
        "SELECT nama_ruang FROM ruang WHERE id=?",
        (ruang_id,)
    ).fetchone()

    conn.close()

    return render_template(
        "cetak_pemeriksaan.html",
        data=data,
        ruang=ruang["nama_ruang"]
    )

@app.route('/barang_by_ruang/<int:id>')
def barang_by_ruang(id):
        conn = get_db()
        barang = conn.execute("""
            SELECT id, uraian 
            FROM barang 
            WHERE id = ?
        """, (id,)).fetchall()
        conn.close()

        return {"barang": [dict(b) for b in barang]}
    
@app.route('/pemeriksaan')
def pemeriksaan():
    conn = get_db()
    ruang = conn.execute("SELECT * FROM ruang").fetchall()
    conn.close()
    return render_template("pemeriksaan.html", ruang=ruang)

@app.route('/simpan_pemeriksaan', methods=['POST'])
def simpan_pemeriksaan():
    barang_id = request.form['barang_id']
    jumlah_fisik = request.form['jumlah_fisik']
    selisih = request.form['selisih']
    bb = request.form['bb']
    rr = request.form['rr']
    rb = request.form['rb']
    hl = request.form['hl']

    conn = get_db()
    conn.execute("""
        INSERT INTO pemeriksaan_fisik
        (barang_id, jumlah_fisik, selisih, bb, rr, rb, hl)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (barang_id, jumlah_fisik, selisih, bb, rr, rb, hl))

    conn.commit()
    conn.close()
    return redirect('/pemeriksaan')

@app.route('/laporan')
def laporan():
    conn = get_db()
    data = conn.execute("""
        SELECT barang.uraian, ruang.nama_ruang,
               jumlah_fisik, bb, rr, rb, hl, tanggal
        FROM pemeriksaan_fisik
        JOIN barang ON barang.id = pemeriksaan_fisik.barang_id
        JOIN ruang ON ruang.id = barang.ruang_id
        ORDER BY tanggal DESC
    """).fetchall()
    conn.close()
    return render_template("laporan.html", data=data)

@app.route('/grafik')
def grafik():
    conn = get_db()
    data = conn.execute("""
        SELECT
            SUM(bb) as bb,
            SUM(rr) as rr,
            SUM(rb) as rb,
            SUM(hl) as hl
        FROM pemeriksaan_fisik
    """).fetchone()
    conn.close()
    return render_template("grafik.html", data=data)

@app.route('/barang_ruang', methods=['GET', 'POST'])
def barang_ruang():
    if request.method == 'POST' and not is_admin():
        flash("Akses ditolak!", "error")
        return redirect(url_for('barang_ruang'))
    conn = get_db()

    # =====================================================
    # SIMPAN PENEMPATAN BARANG KE RUANGAN
    # =====================================================
    if request.method == 'POST':

        barang_id = request.form['barang_id']
        ruang_id = request.form['ruang_id']
        jumlah_input = int(request.form['jumlah'])

        # -----------------------------
        # ambil stok total barang
        # -----------------------------
        stok_total = conn.execute(
            "SELECT jumlah_barang FROM barang WHERE id=?",
            (barang_id,)
        ).fetchone()['jumlah_barang']

        # -----------------------------
        # hitung barang sudah ditempatkan
        # -----------------------------
        sudah_ditempatkan = conn.execute(
            """
            SELECT COALESCE(SUM(jumlah),0) AS total
            FROM aset
            WHERE barang_id=?
            """,
            (barang_id,)
        ).fetchone()['total']

        sisa_stok = stok_total - sudah_ditempatkan

        # -----------------------------
        # VALIDASI STOK
        # -----------------------------
        if jumlah_input > sisa_stok:
            flash(
                f"Stok tidak cukup! Sisa stok belum didistribusikan: {sisa_stok}",
                "error"
            )
            return redirect(url_for('barang_ruang'))

        # =====================================================
        # CEK DUPLIKASI BARANG DI RUANGAN
        # =====================================================
        cek = conn.execute("""
            SELECT id, jumlah
            FROM aset
            WHERE barang_id=? AND ruang_id=?
        """, (barang_id, ruang_id)).fetchone()

        if cek:
            # ✅ UPDATE jumlah jika sudah ada
            conn.execute("""
                UPDATE aset
                SET jumlah = jumlah + ?
                WHERE id=?
            """, (jumlah_input, cek["id"]))
        else:
            # ✅ INSERT jika belum ada
            conn.execute("""
                INSERT INTO aset (barang_id, ruang_id, jumlah)
                VALUES (?, ?, ?)
            """, (barang_id, ruang_id, jumlah_input))

        conn.commit()

        return redirect(url_for('barang_ruang'))

    # =====================================================
    # FILTER RUANGAN
    # =====================================================
    ruang_filter = request.args.get('ruang_id')

    query = """
        SELECT 
            ruang.nama_ruang,
            barang.nama_barang,
            SUM(aset.jumlah) AS jumlah
        FROM aset
        JOIN barang ON aset.barang_id = barang.id
        JOIN ruang ON aset.ruang_id = ruang.id
    """

    params = []

    if ruang_filter:
        query += " WHERE ruang.id=? "
        params.append(ruang_filter)

    query += """
        GROUP BY ruang.nama_ruang, barang.nama_barang
        ORDER BY ruang.nama_ruang
    """

    rows = conn.execute(query, params).fetchall()

    # =====================================================
    # DROPDOWN BARANG + SISA STOK
    # =====================================================
    barang = conn.execute("""
        SELECT 
            barang.id,
            barang.nama_barang,
            barang.jumlah_barang,
            COALESCE(SUM(aset.jumlah),0) AS terpakai,
            barang.jumlah_barang -
            COALESCE(SUM(aset.jumlah),0) AS sisa
        FROM barang
        LEFT JOIN aset 
        ON barang.id = aset.barang_id
        GROUP BY barang.id
    """).fetchall()

    # dropdown ruang
    ruang = conn.execute(
        "SELECT id, nama_ruang FROM ruang"
    ).fetchall()

    conn.close()

    # =====================================================
    # GROUPING DATA UNTUK TAMPILAN
    # =====================================================
    grouped = {}

    for r in rows:
        ruang_nama = r["nama_ruang"]

        if ruang_nama not in grouped:
            grouped[ruang_nama] = []

        grouped[ruang_nama].append({
            "nama_barang": r["nama_barang"],
            "jumlah": r["jumlah"]
        })

    # =====================================================
    # RENDER
    # =====================================================
    return render_template(
        'barang_ruang.html',
        barang=barang,
        ruang=ruang,
        grouped=grouped,
        ruang_filter=ruang_filter
    )
    
@app.route('/barang_ruang/cetak/<int:ruang_id>')
def cetak_per_ruang(ruang_id):
    if not is_admin():
        flash("Akses ditolak!", "error")
        return redirect(url_for('barang_ruang'))
    conn = get_db()

    ruang = conn.execute(
        "SELECT nama_ruang FROM ruang WHERE id=?",
        (ruang_id,)
    ).fetchone()

    data = conn.execute("""
        SELECT 
            barang.nama_barang,
            barang.merk,
            barang.kode_barang,
            barang.tahun_perolehan,
            aset.jumlah
        FROM aset
        JOIN barang ON aset.barang_id = barang.id
        WHERE aset.ruang_id=?
        ORDER BY barang.nama_barang
    """,(ruang_id,)).fetchall()

    conn.close()

    return render_template(
        "barang_ruang_print.html",
        ruang=ruang,
        data=data
    )
    
@app.route('/ruangan', methods=['GET', 'POST'])
def ruang():
    if (request.args.get("hapus") or request.args.get("edit_id")) and not is_admin():
        flash("Akses ditolak!", "error")
        return redirect(url_for('ruang'))

    conn = get_db()

    # ➜ tambah ruangan
    if request.method == 'POST':
        nama = request.form['nama_ruang'].strip()

        cek = conn.execute(
            "SELECT id FROM ruang WHERE nama_ruang=?",
            (nama,)
        ).fetchone()

        if not cek and nama:
            conn.execute(
                "INSERT INTO ruang (nama_ruang) VALUES (?)",
                (nama,)
            )
            conn.commit()

    # ➜ hapus ruangan
    if request.args.get("hapus"):
        conn.execute(
            "DELETE FROM ruang WHERE id=?",
            (request.args.get("hapus"),)
        )
        conn.commit()

    # ➜ edit ruangan
    if request.args.get("edit_id"):
        conn.execute(
            "UPDATE ruang SET nama_ruang=? WHERE id=?",
            (
                request.args.get("edit_nama"),
                request.args.get("edit_id")
            )
        )
        conn.commit()

    # ============================
    # 🔍 FILTER + PAGINATION
    # ============================
    keyword = request.args.get("cari", "")

    page = request.args.get('page', 1, type=int)
    per_page = 10
    offset = (page - 1) * per_page

    # total data
    total = conn.execute(
        "SELECT COUNT(*) FROM ruang WHERE nama_ruang LIKE ?",
        ('%' + keyword + '%',)
    ).fetchone()[0]

    total_pages = (total + per_page - 1) // per_page

    # ambil data sesuai page
    data = conn.execute(
        "SELECT * FROM ruang WHERE nama_ruang LIKE ? ORDER BY id DESC LIMIT ? OFFSET ?",
        ('%' + keyword + '%', per_page, offset)
    ).fetchall()

    conn.close()

    return render_template(
        "ruangan.html",
        data=data,
        keyword=keyword,
        page=page,
        total_pages=total_pages
    )
    
@app.route('/ruangan/update/<int:id>', methods=['POST'])
def update_ruangan(id):

    print("UPDATE ID:", id)  # DEBUG

    nama = request.form['nama_ruang']

    conn = get_db()
    conn.execute(
        "UPDATE ruang SET nama_ruang=? WHERE id=?",
        (nama, id)
    )
    conn.commit()
    conn.close()

    return redirect(url_for('ruang'))

@app.route('/ruangan/hapus/<int:id>', methods=['POST'])
def hapus_ruangan(id):

    if session.get('role') != 'admin':
        flash("Akses ditolak!", "error")
        return redirect(url_for('ruang'))

    conn = get_db()
    conn.execute("DELETE FROM ruang WHERE id=?", (id,))
    conn.commit()
    conn.close()

    flash("Data berhasil dihapus!", "success")
    return redirect(url_for('ruang'))
    
@app.route('/rekap/excel')
def export_rekap_excel():
    #if not is_admin():
    #   flash("Akses ditolak!", "error")
    #   return redirect(url_for('rekap_page'))
    conn = get_db()

    data = conn.execute("""
    SELECT
        barang.id,
        barang.nama_barang,

        barang.jumlah_barang AS jumlah_fisik,

        0 as bb,

        COALESCE(pf.rr,0) as rr,
        COALESCE(pf.rb,0) as rb,
        COALESCE(pf.hl,0) as hl,

        CASE 
            WHEN (COALESCE(pf.rr,0) + COALESCE(pf.rb,0) + COALESCE(pf.hl,0)) = 0
            THEN 0

            WHEN (
                barang.jumlah_barang -
                (COALESCE(pf.rr,0) +
                COALESCE(pf.rb,0) +
                COALESCE(pf.hl,0))
            ) < 0
            THEN 0

            ELSE (
                barang.jumlah_barang -
                (COALESCE(pf.rr,0) +
                COALESCE(pf.rb,0) +
                COALESCE(pf.hl,0))
            )
        END AS selisih,

        barang.tipe,
        barang.merk,
        barang.tahun_perolehan,
        barang.nilai

    FROM barang

    LEFT JOIN pemeriksaan_fisik pf
    ON barang.id = pf.barang_id
    AND pf.id = (
        SELECT MAX(id)
        FROM pemeriksaan_fisik
        WHERE barang_id = barang.id
    )

    ORDER BY barang.nama_barang
    """).fetchall()

    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "Rekap Barang"

    headers = [
        "Nama Barang", "Jumlah Fisik", "Selisih",
        "BB", "RR", "RB", "HL",
        "Tipe", "Merk", "Tahun", "Nilai"
    ]

    ws.append(headers)

    for d in data:
        ws.append([
            d["nama_barang"],
            d["jumlah_fisik"],
            d["selisih"],
            d["bb"],
            d["rr"],
            d["rb"],
            d["hl"],
            d["tipe"],
            d["merk"],
            d["tahun_perolehan"],
            d["nilai"]
        ])

    file_stream = io.BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)

    return send_file(
        file_stream,
        download_name="rekap_barang.xlsx",
        as_attachment=True
    )
    
@app.route('/rekap/cetak')
def rekap_cetak():
    #if not is_admin():
    #    flash("Akses ditolak!", "error")
    #    return redirect(url_for('rekap_page'))
    conn = get_db()

    data = conn.execute("""
        SELECT
            barang.id,
            barang.nama_barang,
            barang.jumlah_barang,

            COALESCE(pf.bb,0) as bb,
            COALESCE(pf.rr,0) as rr,
            COALESCE(pf.rb,0) as rb,
            COALESCE(pf.hl,0) as hl,

            barang.jumlah_barang AS jumlah_fisik,
            barang.tipe,
            barang.merk,
            barang.tahun_perolehan,
            barang.nilai

        FROM barang
        LEFT JOIN pemeriksaan_fisik pf
            ON barang.id = pf.barang_id
        GROUP BY barang.id
        ORDER BY barang.nama_barang
    """).fetchall()

    conn.close()
    return render_template("rekap_print.html", data=data)
    
@app.route('/rekap')
def rekap_page():
    conn = get_db()

    page = request.args.get('page', 1, type=int)
    per_page = 10
    offset = (page - 1) * per_page

    data = conn.execute("""
    SELECT
        barang.id,
        barang.kode_barang,
        barang.nama_barang,
        barang.jumlah_barang,
        COALESCE(pf.bb,0) as bb,
        COALESCE(pf.rr,0) as rr,
        COALESCE(pf.rb,0) as rb,
        COALESCE(pf.hl,0) as hl,
        barang.jumlah_barang AS jumlah_fisik,
        barang.tipe,
        barang.merk,
        barang.tahun_perolehan,
        barang.nilai
    FROM barang
    LEFT JOIN pemeriksaan_fisik pf
        ON barang.id = pf.barang_id
    GROUP BY barang.id
    ORDER BY barang.nama_barang
    LIMIT ? OFFSET ?
    """, (per_page, offset)).fetchall()

    total = conn.execute("SELECT COUNT(*) FROM barang").fetchone()[0]
    total_pages = (total + per_page - 1) // per_page

    conn.close()

    return render_template(
        "rekap.html",
        data=data,
        page=page,
        total_pages=total_pages
    )
    
if __name__ == '__main__':
    app.run(debug=True)
    #app.run(host='0.0.0.0', port=5001)