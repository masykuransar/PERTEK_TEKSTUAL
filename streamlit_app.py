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
import os
import zipfile
from datetime import datetime

import openpyxl
import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate

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
# 4. Buat dokumen
# ---------------------------------------------------------------------- #
st.divider()
st.subheader("4. Buat Dokumen")

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
