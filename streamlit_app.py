"""
Aplikasi web lokal (Streamlit) untuk mengisi otomatis dokumen dari
template.docx + data pemohon (upload Excel dan/atau ketik manual).

Cara jalankan:
    pip install streamlit docxtpl openpyxl pandas
    streamlit run app.py

Aplikasi akan otomatis terbuka di browser (http://localhost:8501).
Semua proses tetap 100% lokal di komputer ini — tidak ada data yang
dikirim ke internet.
"""

import io
import json
import os
import zipfile
from datetime import datetime

import openpyxl
import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate
from shapely.geometry import Point, shape as shapely_shape
from shapely.ops import transform as shapely_transform

st.set_page_config(
    page_title="Isi Dokumen Otomatis",
    page_icon="📄",
    layout="centered",
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


# ---------------------------------------------------------------------- #
# Helpers
# ---------------------------------------------------------------------- #
def cari_file_default(nama_file):
    path = os.path.join(BASE_DIR, nama_file)
    return path if os.path.exists(path) else None


def ambil_field_dari_template(template_bytes):
    """Baca semua placeholder {{ ... }} di template.docx secara otomatis."""
    doc = DocxTemplate(io.BytesIO(template_bytes))
    variabel = doc.get_undeclared_template_variables()
    return sorted(variabel)


def baca_data_excel(file_bytes, kolom_target):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    headers = [str(h).strip() if h is not None else "" for h in rows[0]]
    data_rows = [r for r in rows[1:] if r and r[0] is not None and str(r[0]).strip() != ""]

    daftar = []
    for row in data_rows:
        item = {h: ("" if v is None else str(v)) for h, v in zip(headers, row)}
        # Samakan dengan kolom target (field template) — kolom yang tidak
        # dikenal template diabaikan, kolom template yang tidak ada di
        # Excel diisi kosong.
        selaras = {kol: item.get(kol, "") for kol in kolom_target}
        daftar.append(selaras)
    return daftar


def buat_dokumen(template_bytes, daftar_pemohon, kolom_nama="nama_pemohon"):
    hasil = []
    for i, context in enumerate(daftar_pemohon, start=1):
        doc = DocxTemplate(io.BytesIO(template_bytes))
        doc.render({k: ("" if v is None else str(v)) for k, v in context.items()})

        sumber_nama = context.get(kolom_nama) or f"pemohon_{i}"
        nama_file = "".join(
            ch for ch in str(sumber_nama) if ch.isalnum() or ch in (" ", "_", "-")
        ).strip().replace(" ", "_") or f"pemohon_{i}"

        buffer = io.BytesIO()
        doc.save(buffer)
        hasil.append((f"{nama_file}.docx", buffer.getvalue()))
    return hasil


def parse_koordinat(teks):
    """Parse 'lat, lon' desimal jadi (lat, lon). Kembalikan None kalau gagal."""
    try:
        bagian = str(teks).replace(";", ",").split(",")
        if len(bagian) != 2:
            return None
        lat = float(bagian[0].strip())
        lon = float(bagian[1].strip())
        return lat, lon
    except Exception:
        return None


def muat_rtrw_geojson(file_bytes):
    data = json.loads(file_bytes)
    records = []
    for feat in data.get("features", []):
        geom = shapely_shape(feat["geometry"])
        records.append({"geom": geom, "attrs": feat.get("properties", {}) or {}})
    field_tersedia = sorted(records[0]["attrs"].keys()) if records else []
    return records, field_tersedia


def muat_rtrw_shapefile_zip(zip_bytes):
    import shapefile  # pyshp

    zf = zipfile.ZipFile(io.BytesIO(zip_bytes))
    peta_ekstensi = {os.path.splitext(n)[1].lower(): n for n in zf.namelist()}

    for wajib in (".shp", ".dbf"):
        if wajib not in peta_ekstensi:
            raise ValueError(f"File {wajib} tidak ditemukan di dalam .zip. Pastikan .shp, .dbf, .shx, dan .prj ada semua.")

    shp_bytes = zf.read(peta_ekstensi[".shp"])
    dbf_bytes = zf.read(peta_ekstensi[".dbf"])
    shx_bytes = zf.read(peta_ekstensi[".shx"]) if ".shx" in peta_ekstensi else None
    prj_teks = zf.read(peta_ekstensi[".prj"]).decode("utf-8", errors="ignore") if ".prj" in peta_ekstensi else None

    sf = shapefile.Reader(
        shp=io.BytesIO(shp_bytes),
        dbf=io.BytesIO(dbf_bytes),
        shx=io.BytesIO(shx_bytes) if shx_bytes else None,
    )
    nama_field = [f[0] for f in sf.fields[1:]]  # field pertama selalu DeletionFlag, dilewati

    records = []
    for sr in sf.shapeRecords():
        geom = shapely_shape(sr.shape.__geo_interface__)
        attrs = dict(zip(nama_field, sr.record))
        records.append({"geom": geom, "attrs": attrs})

    # Reproject ke WGS84 (lat/lon) kalau shapefile memakai sistem koordinat lain (mis. UTM)
    if prj_teks:
        try:
            from pyproj import CRS, Transformer

            crs_asal = CRS.from_wkt(prj_teks)
            crs_wgs84 = CRS.from_epsg(4326)
            if crs_asal != crs_wgs84:
                transformer = Transformer.from_crs(crs_asal, crs_wgs84, always_xy=True)
                for rec in records:
                    rec["geom"] = shapely_transform(transformer.transform, rec["geom"])
        except Exception:
            pass  # Kalau gagal reproject, tetap lanjut pakai koordinat asli

    return records, sorted(nama_field)


def cari_zona_rtrw(lat, lon, records, kolom_zona):
    titik = Point(lon, lat)  # shapely: urutan (x=lon, y=lat)
    for rec in records:
        geom = rec["geom"]
        if geom.contains(titik) or geom.intersects(titik):
            return rec["attrs"].get(kolom_zona, "")
    return None


def buat_zip(daftar_file):
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for nama, data in daftar_file:
            zf.writestr(nama, data)
    buffer.seek(0)
    return buffer


# ---------------------------------------------------------------------- #
# 1. Template
# ---------------------------------------------------------------------- #
st.title("📄 Isi Dokumen Otomatis")
st.caption("Mengisi template Word secara otomatis dari data pemohon — berjalan 100% lokal di komputer ini.")

st.divider()
st.subheader("1. Template Word")

default_template = cari_file_default("template.docx")
if default_template:
    st.success("Ditemukan: template.docx")
    pakai_default_template = st.checkbox("Pakai file ini", value=True, key="pakai_template")
else:
    pakai_default_template = False

template_upload = None
if not pakai_default_template:
    template_upload = st.file_uploader("Upload template.docx", type=["docx"])

template_bytes = None
if pakai_default_template:
    with open(default_template, "rb") as f:
        template_bytes = f.read()
elif template_upload is not None:
    template_bytes = template_upload.read()

if not template_bytes:
    st.info("Upload atau pilih template.docx dulu untuk melanjutkan.")
    st.stop()

try:
    kolom_field = ambil_field_dari_template(template_bytes)
except Exception as exc:
    st.error(f"Gagal membaca placeholder dari template: {exc}")
    st.stop()

if not kolom_field:
    st.warning("Template ini tidak punya placeholder {{ ... }} yang terdeteksi.")
    st.stop()

# Kolom yang dipakai sebagai nama file hasil — utamakan nama_pemohon kalau ada
kolom_nama_file = "nama_pemohon" if "nama_pemohon" in kolom_field else kolom_field[0]

with st.expander(f"Field terdeteksi di template ({len(kolom_field)})"):
    st.write(", ".join(kolom_field))

# ---------------------------------------------------------------------- #
# 2. Data pemohon — state gabungan
# ---------------------------------------------------------------------- #
st.divider()
st.subheader("2. Data Pemohon")

# Reset tabel kalau field template berubah (mis. ganti template)
if st.session_state.get("_kolom_field") != kolom_field:
    st.session_state["_kolom_field"] = kolom_field
    st.session_state["data_pemohon"] = pd.DataFrame(columns=kolom_field)

tab_excel, tab_manual = st.tabs(["📥 Upload Excel", "✍️ Ketik Manual"])

with tab_excel:
    file_excel = st.file_uploader("Upload data_pemohon.xlsx", type=["xlsx"], key="excel_uploader")
    default_data = cari_file_default("data_pemohon.xlsx")
    sumber_bytes = None
    if file_excel is not None:
        sumber_bytes = file_excel.read()
    elif default_data:
        st.caption("Atau pakai file default yang ditemukan di folder:")
        if st.button("Pakai data_pemohon.xlsx dari folder"):
            with open(default_data, "rb") as f:
                sumber_bytes = f.read()

    if sumber_bytes:
        try:
            baris_baru = baca_data_excel(sumber_bytes, kolom_field)
            st.write(f"Ditemukan **{len(baris_baru)} baris** di file Excel.")
            if st.button(f"➕ Tambahkan {len(baris_baru)} baris ke daftar", type="primary"):
                df_baru = pd.DataFrame(baris_baru, columns=kolom_field)
                st.session_state["data_pemohon"] = pd.concat(
                    [st.session_state["data_pemohon"], df_baru], ignore_index=True
                )
                st.rerun()
        except Exception as exc:
            st.error(f"Gagal membaca file Excel: {exc}")

with tab_manual:
    st.caption("Isi satu pemohon, lalu klik tambah. Bisa diulang untuk pemohon berikutnya.")
    with st.form("form_manual", clear_on_submit=True):
        nilai_input = {}
        for kol in kolom_field:
            nilai_input[kol] = st.text_input(kol, key=f"input_{kol}")
        submit_manual = st.form_submit_button("➕ Tambahkan ke Daftar", type="primary")

    if submit_manual:
        df_baru = pd.DataFrame([nilai_input], columns=kolom_field)
        st.session_state["data_pemohon"] = pd.concat(
            [st.session_state["data_pemohon"], df_baru], ignore_index=True
        )
        st.rerun()

# ---------------------------------------------------------------------- #
# 3. Daftar gabungan — bisa diedit / dihapus langsung
# ---------------------------------------------------------------------- #
st.divider()
st.subheader("3. Daftar Pemohon")
st.caption("Bisa diedit langsung di tabel. Hapus baris lewat ikon 🗑 di ujung kanan baris (arahkan kursor ke baris).")

df_terkini = st.session_state["data_pemohon"]
df_hasil_edit = st.data_editor(
    df_terkini,
    num_rows="dynamic",
    use_container_width=True,
    key="editor_pemohon",
)
st.session_state["data_pemohon"] = df_hasil_edit

jumlah_valid = df_hasil_edit.dropna(how="all").shape[0]
st.write(f"Total: **{jumlah_valid} pemohon** siap diproses.")

# ---------------------------------------------------------------------- #
# 4. Cek kesesuaian RTRW (opsional)
# ---------------------------------------------------------------------- #
st.divider()
st.subheader("4. Cek Kesesuaian RTRW (Opsional)")
st.caption(
    "Upload data pola ruang RTRW untuk mengecek zona di titik koordinat setiap pemohon "
    "(dari kolom 'koordinat_lokasi', format desimal: `1.234567, 124.567890`)."
)

rtrw_file = st.file_uploader(
    "Upload pola ruang RTRW — GeoJSON (.geojson/.json) atau Shapefile (.zip berisi .shp + .dbf + .shx + .prj)",
    type=["geojson", "json", "zip"],
    key="rtrw_uploader",
)

if rtrw_file is not None:
    try:
        if rtrw_file.name.lower().endswith(".zip"):
            rtrw_records, field_tersedia = muat_rtrw_shapefile_zip(rtrw_file.read())
        else:
            rtrw_records, field_tersedia = muat_rtrw_geojson(rtrw_file.read())

        st.success(f"Berhasil memuat {len(rtrw_records)} poligon zona pola ruang.")

        if not field_tersedia:
            st.warning("Tidak ada kolom atribut yang terdeteksi di file ini.")
        else:
            kolom_zona = st.selectbox(
                "Kolom mana yang berisi nama zona/pola ruang?",
                field_tersedia,
                help="Contoh nama kolom yang umum dipakai: NAMOBJ, POLA_RUANG, KETERANGAN",
            )

            if st.button("🔍 Cek Kesesuaian untuk Semua Pemohon", use_container_width=True):
                df_cek = df_hasil_edit.dropna(how="all").copy()
                hasil_zona = []
                for _, baris in df_cek.iterrows():
                    koordinat = baris.get("koordinat_lokasi", "")
                    hasil_parse = parse_koordinat(koordinat)
                    if not hasil_parse:
                        hasil_zona.append("⚠️ Format koordinat tidak terbaca")
                        continue
                    lat, lon = hasil_parse
                    zona = cari_zona_rtrw(lat, lon, rtrw_records, kolom_zona)
                    hasil_zona.append(zona if zona not in (None, "") else "❌ Di luar seluruh zona / tidak ditemukan")
                df_cek["zona_rtrw"] = hasil_zona
                st.session_state["hasil_cek_rtrw"] = df_cek

    except Exception as exc:
        st.error(f"Gagal membaca file RTRW: {exc}")

if "hasil_cek_rtrw" in st.session_state:
    kolom_tampil = [k for k in [kolom_nama_file, "koordinat_lokasi", "zona_rtrw"] if k in st.session_state["hasil_cek_rtrw"].columns]
    st.write("**Hasil pengecekan zona per pemohon:**")
    st.dataframe(st.session_state["hasil_cek_rtrw"][kolom_tampil], use_container_width=True)
    st.caption(
        "⚠️ Hasil ini murni perhitungan geometris (titik vs poligon) dari data yang kamu upload. "
        "Selalu verifikasi manual untuk keputusan resmi, terutama untuk titik yang dekat batas zona "
        "atau kalau data RTRW yang diupload belum yang terbaru."
    )

# ---------------------------------------------------------------------- #
# 5. Buat dokumen
# ---------------------------------------------------------------------- #
st.divider()
st.subheader("5. Buat Dokumen")

if st.button("🚀 Buat Semua Dokumen", type="primary", use_container_width=True):
    daftar_final = df_hasil_edit.dropna(how="all").fillna("").to_dict("records")
    if not daftar_final:
        st.warning("Belum ada data pemohon. Tambahkan lewat Upload Excel atau Ketik Manual di atas.")
    else:
        progress = st.progress(0, text="Memproses...")
        try:
            hasil = []
            for i, context in enumerate(daftar_final, start=1):
                satu = buat_dokumen(template_bytes, [context], kolom_nama_file)
                hasil.extend(satu)
                progress.progress(i / len(daftar_final), text=f"Memproses {i}/{len(daftar_final)}...")

            progress.empty()
            st.success(f"Selesai! {len(hasil)} dokumen berhasil dibuat.")

            zip_buffer = buat_zip(hasil)
            nama_zip = f"hasil_dokumen_{datetime.now().strftime('%Y%m%d_%H%M')}.zip"

            st.download_button(
                label="⬇️ Download Semua (.zip)",
                data=zip_buffer,
                file_name=nama_zip,
                mime="application/zip",
                use_container_width=True,
            )

            with st.expander("Download satuan per pemohon"):
                for nama, data in hasil:
                    st.download_button(
                        label=f"📄 {nama}",
                        data=data,
                        file_name=nama,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"dl_{nama}",
                    )
        except Exception as exc:
            progress.empty()
            st.error(f"Terjadi kesalahan saat membuat dokumen: {exc}")