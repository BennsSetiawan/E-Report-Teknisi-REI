import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os
import requests
import json

# --- KONFIGURASI GOOGLE SHEETS ---
URL_SHEETS = "https://script.google.com/macros/s/AKfycbwRGAikWG-9qIVSZ7ZFVDpVGm9nTn4DAptb4sHFfc7x1LzY7oJCKGpiH190HLFfRdOX/exec"

def save_to_sheets(data_laporan):
    try:
        response = requests.post(URL_SHEETS, json=data_laporan)
        if response.status_code == 200:
            return True
        else:
            st.error(f"Gagal kirim ke Sheets. Status: {response.status_code}")
            return False
    except Exception as e:
        st.error(f"Error Koneksi Google Sheets: {e}")
        return False

# --- 1. INITIALIZING SESSION STATE ---
if 'db_checklist' not in st.session_state:
    st.session_state.db_checklist = {
        "Zybio": {
            "default": ["Cek Kebocoran Reagen", "Pembersihan Probe", "Pemeriksaan Waste Line", "Kalibrasi Background"],
            "Z3": ["Cek Reagent,Diluent & Lyse", "Rendam Chamber RBC & WBC", "Cleaning Selang Yang Kotor", "Cleaning Jarum", "Cleaning Part-Part Dari Debu", "Cleaning Body", "Cek Mekanikal (Self Check)", "Cek Background", "Quality Control (QC)"],
            "Z31": ["Cek Reagent,Diluent & Lyse", "Rendam Chamber RBC & WBC", "Cleaning Selang Yang Kotor", "Cleaning Jarum", "Cleaning Part-Part Dari Debu", "Cleaning Body", "Cek Mekanikal (Self Check)", "Cek Background", "Quality Control (QC)"],
            "Z3CRP": ["Cek Reagent,Diluent & Lyse", "Rendam Chamber RBC & WBC", "Cleaning Selang Yang Kotor", "Cleaning Jarum", "Cleaning Part-Part Dari Debu", "Cleaning Body", "Cek Mekanikal (Self Check)", "Cek Background", "Quality Control (QC)"],
            "Z5": ["Cek Reagent,Diluent & Lyse", "Rendam Chamber RBC & WBC", "Cleaning Selang Yang Kotor", "Cleaning Jarum", "Cleaning Part-Part Dari Debu", "Cleaning Body", "Cek Mekanikal (Self Check)", "Cek Background", "Quality Control (QC)"],
            "Z50": ["Cek Reagent,Diluent & Lyse", "Rendam Chamber RBC & WBC", "Cleaning Selang Yang Kotor", "Cleaning Jarum", "Cleaning Part-Part Dari Debu", "Cleaning Body", "Cek Mekanikal (Self Check)", "Cek Background", "Quality Control (QC)"],
        },
        "Porlak": {
            "default": ["Cek Tanggal Open Reagent", "Cek Tegangan", "Cek Suhu Ruangan", "Cek Lingkungan Udara", "Cek Hasil Quality Control (QC)", "Tes Background"],
            "PJH-H360": ["Cek Reagent,Diluent & Lyse", "Rendam Chamber RBC & WBC", "Cleaning Selang Yang Kotor", "Cleaning Jarum", "Cleaning Part-Part Dari Debu", "Cleaning Body", "Cek Mekanikal (Self Check)", "Cek Background", "Quality Control (QC)"],
            "PJH-B60": ["Cleaning Cuvet", "Cleaning Selang Aspirate", "Cleaning Fan", "Cleaning Part-Part Dari Debu", "Cek Mekanikal", "Cek Background", "Quality Control (Optional)"],
            "PJH-U300":["Cleaning Liquid Box", "Cleaning Waste Box", "Cleaning Belt", "Cleaning Body", "Cek Mekanikal", "Quality Control (Optional)"],
            "PJH-B101": ["Cleaning Cuvet", "Cleaning Selang Aspirate", "Cleaning Fan", "Cleaning Part-Part Dari Debu", "Cek Mekanikal", "Cek Background", "Quality Control (Optional)"],
            "PJH-B200": ["Cek Kondisi Air", "Cek Kondisi Filter Air", "CLeaning Tray Reagent", "Cleaning Cuvet", "Cleaning Jarum Sample", "Cleaning Jarum Wash", "Cleaning Mixer", "Power On Test/Startup Check", "Reset", "Air Purge (2-3 Kali)","Detergent Wash", "Wash Cell All", "Cek Mekanikal (Pastikan Air Mengalir Sempurna/Baik)", "Cell Blank (Range On 20.000-60.000)", "Calibration Tes", "Quality Control Tes (QC)", "Precision Tes"]
        },
        "Bioway": {
            "default": ["Cek Tanggal Open Reagent", "Cek Tegangan", "Cek Suhu Ruangan", "Cek Lingkungan Udara", "Cek Hasil Quality Control (QC)", "Tes Background"],
            "BW-300": ["Cleaning Liquid Box", "Cleaning Waste Box", "Cleaning Belt", "Cleaning Body", "Cek Mekanikal", "Quality Control (Optional)"],
        },
        "Biontech": { 
            "default": ["Cek Tanggal Open Reagent", "Cek Tegangan", "Cek Suhu Ruangan", "Cek Lingkungan Udara", "Cek Hasil Quality Control (QC)", "Tes Background"],
            "DX-60": ["Cleaning Cuvet", "Cleaning Selang Aspirate", "Cleaning Fan", "Cleaning Part-Part Dari Debu", "Cek Mekanikal", "Cek Background", "Quality Control (Optional)"],
            "DX-101": ["Cleaning Cuvet", "Cleaning Selang Aspirate", "Cleaning Fan", "Cleaning Part-Part Dari Debu", "Cek Mekanikal", "Cek Background", "Quality Control (Optional)"],
            "DX-200": ["Cek Kondisi Air", "Cek Kondisi Filter Air", "CLeaning Tray Reagent", "Cleaning Cuvet", "Cleaning Jarum Sample", "Cleaning Jarum Wash", "Cleaning Mixer", "Power On Test/Startup Check", "Reset", "Air Purge (2-3 Kali)","Detergent Wash", "Wash Cell All", "Cek Mekanikal (Pastikan Air Mengalir Sempurna/Baik)", "Cell Blank (Range On 20.000-60.000)", "Calibration Tes", "Quality Control Tes (QC)", "Precision Tes"]
        },
        "Heto": {
            "default": ["Cek Tanggal Open Reagent", "Cek Tegangan", "Cek Suhu Ruangan", "Cek Lingkungan Udara", "Cek Hasil Quality Control (QC)", "Tes Background"],
            "HT-H360": ["Cek Reagent,Diluent & Lyse", "Rendam Chamber RBC & WBC", "Cleaning Selang Yang Kotor", "Cleaning Jarum", "Cleaning Part-Part Dari Debu", "Cleaning Body", "Cek Mekanikal (Self Check)", "Cek Background", "Quality Control (QC)"],
            "HT-B60": ["Cleaning Cuvet", "Cleaning Selang Aspirate", "Cleaning Fan", "Cleaning Part-Part Dari Debu", "Cek Mekanikal", "Cek Background", "Quality Control (Optional)"],
            "HT-U300":["Cleaning Liquid Box", "Cleaning Waste Box", "Cleaning Belt", "Cleaning Body", "Cek Mekanikal", "Quality Control (Optional)"],
            "HT-B101": ["Cleaning Cuvet", "Cleaning Selang Aspirate", "Cleaning Fan", "Cleaning Part-Part Dari Debu", "Cek Mekanikal", "Cek Background", "Quality Control (Optional)"],
            "HT-B200": ["Cek Kondisi Air", "Cek Kondisi Filter Air", "CLeaning Tray Reagent", "Cleaning Cuvet", "Cleaning Jarum Sample", "Cleaning Jarum Wash", "Cleaning Mixer", "Power On Test/Startup Check", "Reset", "Air Purge (2-3 Kali)","Detergent Wash", "Wash Cell All", "Cek Mekanikal (Pastikan Air Mengalir Sempurna/Baik)", "Cell Blank (Range On 20.000-60.000)", "Calibration Tes", "Quality Control Tes (QC)", "Precision Tes"]
        },
        "Bioelab": {
            "default": ["Cek Mixer", "Pembersihan Kuvet", "Kalibrasi Cairan"],
            "ES-380": ["Cek Kondisi Air", "Cek Kondisi Filter Air", "CLeaning Tray Reagent", "Cleaning Cuvet", "Cleaning Jarum Sample", "Cleaning Jarum Wash", "Cleaning Mixer", "Power On Test/Startup Check", "Reset", "Air Purge (2-3 Kali)","Detergent Wash", "Wash Cell All", "Cek Mekanikal (Pastikan Air Mengalir Sempurna/Baik)", "Cell Blank (Range On 20.000-60.000)", "Calibration Tes", "Quality Control Tes (QC)", "Precision Tes"]
        },
        "Lamuno": {"default": ["Update Software", "Cek Konektivitas", "Pembersihan Scanner"]}
    }

if 'db_alat' not in st.session_state:
    st.session_state.db_alat = {
        "Zybio": ["Z5", "Z3", "Z50", "Z31", "Z3CRP"],
        "Porlak": ["PJH-B60", "PJH-H360", "PJH-B200", "PJH-B101", "PJH-U300"],
        "Bioway": ["BW-300"],
        "Biontech": ["DX-200", "DX-101", "DX-60"],
        "Heto": ["HT-H360", "HT-B200", "HT-B60", "HT-B101", "HT-U300"],
        "Bioelab": ["ES-380", "ES-480"],
        "Lamuno": ["Lamuno X", "Lamuno Pro"]
    }

# --- 2. KONFIGURASI HALAMAN ---
st.set_page_config(page_title="Service Report PJH Pro", layout="wide")

# --- 3. HEADER ---
st.markdown("<h1 style='text-align: center;'>PT. RAJAERBA INDOCHEM</h1>", unsafe_allow_html=True)
st.divider()

# --- 4. SIDEBAR ---
st.sidebar.header("👤 Data Teknisi")
list_nama = ["Muhammad Edy Surya", "Muhammad Abdul Majid", "Ari Nugroho", "Leonardus Naldo Poju"]
nama_teknisi = st.sidebar.selectbox("Nama Teknisi", list_nama, key="sb_name")
nik_map = {"Muhammad Edy Surya": "TEKNISI", "Muhammad Abdul Majid": "TEKNISI", "Ari Nugroho": "TEKNISI", "Leonardus Naldo Poju": "TEKNISI"}
id_teknisi = st.sidebar.text_input("Jabatan", nik_map[nama_teknisi], key="sb_id")
nama_pt = st.sidebar.text_input("Nama Perusahaan", "PT. RAJAERBA INDOCHEM", key="sb_pt")
lokasi_pt = st.sidebar.selectbox("Lokasi Cabang", ["Tangerang", "Semarang", "Jakarta", "Medan"], key="sb_loc")

# --- 5. IDENTIFIKASI FORM ---
st.subheader("🔍 Identifikasi Alat & Customer")
c1, c2 = st.columns(2)
with c1:
    merek = st.selectbox("Merek Alat", list(st.session_state.db_alat.keys()), key="main_mrk")
    tipe = st.selectbox("Tipe Alat", st.session_state.db_alat[merek], key="main_tp")
    sn_alat = st.text_input("Nomor Seri (S/N)", key="main_sn")
with c2:
    nama_customer = st.text_input("Nama Customer / Instansi", key="main_cust")
    alamat_customer = st.text_input("Alamat Customer", key="main_addr")
    kontak_customer = st.text_input("Kontak Customer", key="main_contact")

# --- 6. CHECKLIST ---
st.subheader("✅ Checklist Tindakan")
steps = st.session_state.db_checklist.get(merek, {}).get(tipe, st.session_state.db_checklist.get(merek, {}).get("default", []))
hasil_checklist = []
for i, step in enumerate(steps):
    ct, cs = st.columns([3, 1])
    with ct: st.write(f"**{i+1}.** {step}")
    with cs:
        res = st.radio("Hasil", ["Pass (✔)", "Fail (✘)", "Perlu Ganti (!)"], key=f"step_{i}", horizontal=True, label_visibility="collapsed")
        hasil_checklist.append({"Langkah": step, "Hasil": res})

# --- 7. TANDA TANGAN ---
st.subheader("📝 Validasi & TTD")
catatan = st.text_area("Catatan Tambahan", key="main_note")
v1, v2 = st.columns(2)
with v1:
    st.info("TTD Teknisi"); ttd_tek = st.file_uploader("Upload TTD", type=["png", "jpg", "jpeg"], key="up_tek")
    if ttd_tek: st.image(ttd_tek, width=150)
    st.write(f"({nama_teknisi})")
with v2:
    st.info("TTD Customer"); ttd_cus = st.file_uploader("Upload TTD ", type=["png", "jpg", "jpeg"], key="up_cus")
    if ttd_cus: st.image(ttd_cus, width=150)
    st.write(f"({nama_customer if nama_customer else 'Customer'})")

# --- 8. FUNGSI GENERATE ---
def generate_docx(checklist, tech, nik, cust, sn, note, pt, branch, img_tek, img_cus):
    doc = Document()
    doc.add_heading('MAINTENANCE REPORT', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
    p = doc.add_paragraph()
    p.add_run(f"Perusahaan: {pt}\n").bold = True
    p.add_run(f"Cabang: {branch}\nTeknisi: {tech} ({nik})\nCustomer: {cust}\nAlamat: {alamat_customer}\nKontak: {kontak_customer}\nS/N Alat: {sn}\nMerek Alat: {merek}\nTipe Alat: {tipe}\nTanggal: {datetime.now().strftime('%d/%m/%Y')}")
    
    table = doc.add_table(rows=1, cols=2); table.style = 'Table Grid'
    hdr = table.rows[0].cells; hdr[0].text = 'Pekerjaan'; hdr[1].text = 'Hasil'
    for item in checklist:
        row = table.add_row().cells
        row[0].text = item['Langkah']; row[1].text = item['Hasil']
    
    doc.add_paragraph(f"\nCatatan: {note}")
    
    ttd_table = doc.add_table(rows=3, cols=2); ttd_table.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cells_header = ttd_table.rows[0].cells; cells_header[0].text = "Teknisi,"; cells_header[1].text = "Customer,"
    for cell in cells_header: cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    cells_img = ttd_table.rows[1].cells
    if img_tek: cells_img[0].paragraphs[0].add_run().add_picture(img_tek, width=Inches(1.2))
    if img_cus: cells_img[1].paragraphs[0].add_run().add_picture(img_cus, width=Inches(1.2))
    for cell in cells_img: cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    cells_footer = ttd_table.rows[2].cells; cells_footer[0].text = f"( {tech} )"; cells_footer[1].text = f"( {cust if cust else '........'} )"
    for cell in cells_footer: cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    target = BytesIO(); doc.save(target)
    return target.getvalue()

def generate_excel(checklist, tech, nik, cust, sn, pt, branch):
    df_detail = pd.DataFrame({"Kategori": ["Teknisi", "Jabatan", "Perusahaan", "Cabang", "Customer", "S/N Alat"], "Informasi": [tech, nik, pt, branch, cust, sn]})
    df_checklist = pd.DataFrame(checklist)
    target = BytesIO()
    with pd.ExcelWriter(target, engine='xlsxwriter') as writer:
        df_detail.to_excel(writer, index=False, sheet_name='Data_Teknisi')
        df_checklist.to_excel(writer, index=False, sheet_name='Checklist_Report')
    return target.getvalue()

# --- 10. TOMBOL SUBMIT & EKSPOR ---
st.divider()
st.subheader("💾 Finalisasi & Kirim Laporan")

ex1, ex2 = st.columns(2)
with ex1:
    st.download_button(label="📥 Unduh Excel (.xlsx)", data=generate_excel(hasil_checklist, nama_teknisi, id_teknisi, nama_customer, sn_alat, nama_pt, lokasi_pt), file_name=f"E-Report_{sn_alat}.xlsx", key="btn_excel")
with ex2:
    word_content = generate_docx(hasil_checklist, nama_teknisi, id_teknisi, nama_customer, sn_alat, catatan, nama_pt, lokasi_pt, ttd_tek, ttd_cus)
    st.download_button(label="📥 Unduh Word (.docx)", data=word_content, file_name=f"E-Report_{sn_alat}.docx", key="btn_word")

st.write("---")
if st.button("🚀 SUBMIT LAPORAN ", use_container_width=True):
    if not nama_customer or not sn_alat:
        st.warning("Mohon isi Form Di Atas Sebelum Submit.")
    else:
        with st.spinner("Mengirim data ke Google Sheet..."):
            data_sheet = {
                "teknisi": nama_teknisi,
                "customer": nama_customer,
                "sn_alat": sn_alat,
                "merek": merek,
                "tipe": tipe,
                "catatan": catatan
            }
            sheet_ok = save_to_sheets(data_sheet)
            if sheet_ok:
                st.success(f"✅ Berhasil! Data sudah masuk ke Google Sheet.")
                st.balloons()

# --- 11. EDIT DATABASE ---
with st.expander("⚙️ Edit Checklist"):
    new_steps = st.text_area("Ubah Langkah", ", ".join(steps), key="edit_area")
    if st.button("Update", key="btn_up"):
        st.session_state.db_checklist[merek][tipe] = [x.strip() for x in new_steps.split(",")]
        st.success("Tersimpan!"); st.rerun()