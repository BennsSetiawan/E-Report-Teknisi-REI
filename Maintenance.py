import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import requests
import urllib.parse

# --- KONFIGURASI GOOGLE SHEETS ---
URL_SHEETS = "https://script.google.com/macros/s/AKfycbzpj7xT43iJOxL-f9CksWBuUYswKHirEZ_i7mX3t0AWEdwy2R6OzOVRTsdftMnaTPPt/exec"

def save_to_sheets(data_laporan):
    try:
        response = requests.post(URL_SHEETS, json=data_laporan)
        return response.status_code == 200
    except: return False

# --- 1. DATABASE CHECKLIST (DATA LU 100% AMAN) ---
if 'db_checklist' not in st.session_state:
    st.session_state.db_checklist = {
        "AeHealth": {
            "AERC-3": {
                "Bulanan": {
                    "Mati": ["Periksa kondisi Reagent Diluent & Lyse", "Rendam chamber RBC & WBC", "Cleaning Flow Cell", "Cleaning Selang Sampling / Tubing", "Cleaning Jarum Aspirate", "Cleaning Fan Pendingin", "Bersihkan Debu Pada Komponen Internal", "Cleaning Body Alat"],
                    "Hidup": ["Prime Diluent & Lyse", "Prime Seluruh Reagent", "Jalankan Self Check", "Test Background 2–3 kali", "Test Printer", "QC (Quality Control)"]
                },
                "Service": ["Periksa kondisi reagen Diluent", "Periksa kondisi Reagent Lyse", "Rendam chamber RBC", "Rendam chamber WBC", "Cleaning Flow Cell", "Cleaning Selang Sampling / Tubing", "Cleaning Jarum Aspirate", "Cleaning Fan Pendingin", "Bersihkan Debu Pada Komponen Internal", "Cleaning Body Alat", "Prime Diluent", "Prime Lyse", "Prime Seluruh Reagen", "Jalankan Self Check", "Test Background 2–3 kali", "Test Printer", "QC (Quality Control)"]
            }
        },
        "Zybio": {
            "Z5": {
                "Bulanan": {
                    "Mati": ["Periksa kondisi Reagent Diluent & Lyse", "Rendam chamber RBC & WBC", "Cleaning Flow Cell", "Cleaning Selang Sampling / Tubing", "Cleaning Jarum Aspirate", "Cleaning Fan Pendingin", "Bersihkan Debu Pada Komponen Internal", "Cleaning Body Alat"],
                    "Hidup": ["Prime Diluent & Lyse", "Prime Seluruh Reagent", "Jalankan Self Check", "Test Background 3 kali", "Test Printer", "QC (Quality Control)"]
                }
            },
            "Z3 / Z31 / Z3 CRP": {
                "Bulanan": {
                    "Mati": ["Periksa kondisi Reagent Diluent & Lyse", "Rendam chamber RBC & WBC", "Cleaning Flow Cell", "Cleaning Selang Sampling / Tubing", "Cleaning Jarum Aspirate", "Cleaning Fan Pendingin", "Bersihkan Debu Pada Komponen Internal", "Cleaning Body Alat"],
                    "Hidup": ["Prime Diluent & Lyse", "Prime Seluruh Reagent", "Jalankan Self Check", "Test Background 2–3 kali", "Test Printer", "QC (Quality Control)"]
                }
            }
        },
        "Porlak": {
            "PJH-B200": {
                "Bulanan": {
                    "Mati": ["Periksa Kondisi Air Secara Visual", "Periksa Kualitas Air Menggunakan TDS meter", "Cleaning Filter Air", "Cleaning Cuvet", "Cleaning Reagent tray", "Cleaning Jarum Sample", "Cleaning Jarum Washer", "Cleaning Mixer / Pengaduk", "Cleaning Body Alat"],
                    "Hidup": ["Reset System", "Air Purge 3 kali", "Detergent Wash", "Wash Cell All", "Cek Cell Blank (Range 20.000–60.000)", "QC (Quality Control)"]
                }
            },
            "PJH-B60 / PJHB-101": {
                "Bulanan": {
                    "Mati": ["Cleaning Cuvet", "Cleaning Selang Aspirasi", "Cleaning Inkubator", "Cleaning Fan Pendingin", "Bersihkan Debu Pada Komponen Internal", "Cleaning Body Alat"],
                    "Hidup": ["Washing / Cuci 10 Detik", "Cek Background", "Periksa Suhu Inkubator (37°C)", "Periksa Kondisi Lampu (Range 20.000–60.000)", "Test Printer", "QC (Quality Control)"]
                }
            },
            "PJH-U300": {
                "Bulanan": {
                    "Hidup": ["Periksa Liquid Box", "Periksa Waste Box", "Cleaning Belt Conveyor", "Cleaning Body Alat", "Bersihkan Debu pada Komponen Internal", "QC (Quality Control)"]
                
                }
            },
            "PJH-H360": {
                "Bulanan": {
                    "Mati": ["Periksa kondisi Reagent Diluent & Lyse", "Rendam Chamber RBC & WBC", "Cleaning Flow Cell", "Cleaning Selang Sampling / Tubing", "Cleaning Jarum Aspirate", "Cleaning Fan Pendingin", "Bersihkan Debu Pada Komponen Internal", "Cleaning Body Alat"],
                    "Hidup": ["Prime Diluent", "Prime Lyse", "Prime Seluruh Reagent", "Jalankan Self Check", "Test Background 3 Kali", "Test Printer", "QC (Quality Control)"]
                }
            },
            "PJH-H610": {
                "Bulanan": {
                    "Mati": ["Periksa kondisi Reagent Diluent & Lyse", "Rendam chamber RBC & WBC", "Cleaning Flow Cell", "Cleaning Selang Sampling / Tubing", "Cleaning Jarum Aspirate", "Cleaning Fan Pendingin", "Bersihkan Debu Pada Komponen Internal", "Cleaning Body Alat"],
                    "Hidup": ["Prime Diluent", "Prime Lyse", "Prime Seluruh Reagent", "Jalankan Self Check", "Test Background 3 Kali", "Test Printer", "QC (Quality Control)"]
                }
            }
        },
        "Caretium": {
            "XI-931": {
                "Bulanan": {
                    "Mati": ["Cleaning Jarum Sample", "Cleaning Selang", "Cleaning Fan Pendingin", "Cleaning Body Alat", "Check Kondisi Peristaltic Tube"],
                    "Hidup": ["Prime Reagen", "Test Calibration Slope", "Test Printer", "QC (Quality Control)"]
                }
            }
        },
        "Lamuno": {
            "Lamuno X / Lamuno Pro": {
                "Bulanan": {
                    "Hidup": ["Periksa QC Internal, Pastikan Tidak Expired", "Periksa kondisi Reagent", "Pastikan Nomor Lot Reagent Sesuai Dengan Data Pada Alat", "Test Printer", "Cleaning Body Alat", "QC (Quality Control)"]
                }
            }
        },
        "Biontech": {
            "ES-380": {
                "Bulanan": {
                    "Mati": ["Periksa Kondisi Air Secara Visual", "Perikas Kualitas Air Menggunakan TDS Meter", "Cleaning Filter Air", "Cleaning Cuvet", "Cleaning Reagent Tray", "Cleaning Jarum Sample", "Cleaning Jarum Washer", "Cleaning Mixer / Pengaduk", "Cleaning Body Alat"],
                    "Hidup": ["Reset System", "Air Purge 3 kali", "Detergent Wash", "Wash Cell All", "Cek Cell Blank (Range 20.000–60.000)", "QC (Quality Control)"]
                }
            },
            "ES-101": {
                "Bulanan": {
                    "Mati": ["Cleaning Cuvet", "Cleaning Selang Aspirasi", "Cleaning Inkubator", "Cleaning Fan Pendingin", "Bersihkan Debu Pada Komponen Internal", "Cleaning Body Alat"],
                    "Hidup": ["Washing / Cuci 10 Detik", "Cek Background", "Periksa Suhu Inkubator (37°C)", "Periksa Kondisi Lampu (Range 20.000–60.000)", "Test Printer", "QC (Quality Control)"]
                }
            }
        },
        "Singseng": {
            "RD6S": {
                "Bulanan": {
                    "Mati": ["Periksa kondisi reagen Diluent & Lyse", "Rendam chamber RBC & WBC", "Cleaning Flow Cell", "Cleaning Selang Sampling / Tubing", "Cleaning Jarum Aspirate", "Cleaning Fan Pendingin", "Bersihkan Debu Pada Komponen Internal", "Cleaning Body Alat"],
                    "Hidup": ["Prime Diluent", "Prime Lyse", "Prime Seluruh Reagent", "Jalankan Self Check", "Test Background 3 Kali", "Test Printer", "QC (Quality Control)"]
                }
            },
            "RD3S": {
                "Bulanan": {
                    "Mati": ["Periksa kondisi reagen Diluent & Lyse", "Rendam chamber RBC & WBC", "Cleaning Flow Cell", "Cleaning Selang Sampling / Tubing", "Cleaning Jarum Aspirate", "Cleaning Fan Pendingin", "Bersihkan Debu Pada Komponen Internal", "Cleaning Body Alat"],
                    "Hidup": ["Prime Diluent", "Prime Lyse", "Prime Seluruh Reagent", "Jalankan Self Check", "Test Background 3 Kali", "Test Printer", "QC (Quality Control)"]
                }
            }
        },
        "Bioway": {
            "BW-300": {
                "Bulanan": {
                    "Hidup": ["Periksa Liquid Box", "Periksa Waste Box", "Cleaning Belt Conveyor", "Cleaning Body Alat", "Bersihkan Debu pada Komponen Internal", "QC (Quality Control)"]
                }
            }
        },
        "Heto": {
            "HT-H360": {
                "Bulanan": {
                    "Mati": ["Periksa kondisi Reagent Diluent & Lyse", "Rendam chamber RBC & WBC", "Cleaning Flow Cell", "Cleaning Selang Sampling / Tubing", "Cleaning Jarum Aspirate", "Cleaning Fan Pendingin", "Bersihkan Debu Pada Komponen Internal", "Cleaning Body Alat"],
                    "Hidup": ["Prime Diluent", "Prime Lyse", "Prime Seluruh Reagent", "Jalankan Self Check", "Test Background 3 Kali", "Test Printer", "QC (Quality Control)"]
                }
            },
            "HT-B200": {
                "Bulanan": {
                    "Mati": ["Periksa Kondisi Air Secara Visual", "Perikas Kualitas Air Menggunakan TDS Meter", "Cleaning Filter Air", "Cleaning Cuvet", "Cleaning Reagent Tray", "Cleaning Jarum Sample", "Cleaning Jarum Washer", "Cleaning Mixer / Pengaduk", "Cleaning Body Alat"],
                    "Hidup": ["Reset System", "Air Purge 3 kali", "Detergent Wash", "Wash Cell All", "Cek Cell Blank (Range 20.000–60.000)", "QC (Quality Control)"]
                }
            },
            "HT-U300": {
                "Bulanan": {
                    "Hidup": ["Periksa Liquid Box", "Periksa Waste Box", "Cleaning Belt Conveyor", "Cleaning Body Alat", "Bersihkan Debu pada Komponen Internal", "QC (Quality Control)"]
                }
            },
            "HT-B60 / HT-B101": {
                "Bulanan": {
                    "Mati": ["Cleaning Cuvet", "Cleaning Selang Aspirasi", "Cleaning Inkubator", "Cleaning Fan Pendingin", "Bersihkan Debu Pada Komponen Internal", "Cleaning Body Alat"],
                    "Hidup": ["Washing / Cuci 10 Detik", "Cek Background", "Periksa Suhu Inkubator (37°C)", "Periksa Kondisi Lampu (Range 20.000–60.000)", "Test Printer", "QC (Quality Control)"]
                }
        }
    }
}

if 'db_alat' not in st.session_state:
    st.session_state.db_alat = {
        "AeHealth": ["AERC-3"],
        "Biontech": ["ES-380, ES-101"],
        "Bioway":   ["BW-300"],
        "Caretium": ["XI-931"],
        "Heto": ["HT-H360", "HT-200", "HT-U300", "HT-B60 / HT-B101"],
        "Lamuno": ["Lamuno X / Lamuno Pro"],
        "Porlak": ["PJH-B200", "PJH-H360", "PJH-H610", "PJH-B60 / PJHB-101", "PJH-U300"],
        "Singseng": ["RD6S", "RD3S"],
        "Zybio": ["Z5", "Z3 / Z31 / Z3-CRP"]
    }

if 'list_teknisi' not in st.session_state:
    st.session_state.list_teknisi = ["Muhammad Edy Surya", "Muhammad Abdul Majid", "Ari Nugroho", "Leonardus Naldo Poju"]

# --- 2. KONFIGURASI HALAMAN ---
st.set_page_config(page_title="REI E-Report Teknisi", layout="wide", page_icon="⚙️")

st.markdown("""
<link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">
<style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;600;800&display=swap');
    html, body, [class*="css"] { font-family: 'Poppins', sans-serif; background-color: #f8fafc; }
    [data-testid="stSidebar"] { background-color: #1a1e23 !important; }
    [data-testid="stSidebar"] p, [data-testid="stSidebar"] label, [data-testid="stSidebar"] h3 { color: #ffffff !important; }
    div[data-testid="stVerticalBlockBorderWrapper"] {
        background-color: white; border-radius: 12px; padding: 30px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.08); border: 1px solid #e2e8f0;
    }
    .header-container { display: flex; align-items: center; justify-content: center; padding: 25px; background: linear-gradient(to right, #f76300, #ff8447); border-radius: 15px; margin-bottom: 25px; border: 1px solid #e2e8f0; }
    .rei-logo-main { font-size: 80px; font-weight: 800; color: #00477f; display: flex; align-items: center; letter-spacing: -4px; }
    .text-title-wrapper { border-left: 5px solid #00477f; padding-left: 20px; margin-left: 20px; line-height: 1.1; }
    .text-title-wrapper .e-report { font-size: 35px; font-weight: 700; color: #00477f; margin: 0; }
    .text-title-wrapper .teknisi { font-size: 50px; font-weight: 800; color: #000000; margin: 0; letter-spacing: 2px; }
    div.stButton > button:first-child {
        background-color: #00477f !important; color: white !important;
        font-weight: 800 !important; height: 60px !important; width: 100% !important; border-radius: 10px !important; font-size: 20px !important;
    }
    .whatsapp-btn {
        background-color: #25d366; color: white; padding: 15px; border-radius: 10px;
        text-decoration: none; font-weight: bold; display: block; text-align: center; margin-top: 10px;
    }
</style>
""", unsafe_allow_html=True)

# --- 3. FUNGSI GENERATE DOCX ---
def generate_docx(checklist, tech, nik, cust, sn, note, merek, tipe, periode, pt, branch, foto_list, ttd_t, ttd_c):
    doc = Document()
    
    # --- HALAMAN 1: LAPORAN UTAMA ---
    doc.add_heading('MAINTENANCE REPORT', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
    p = doc.add_paragraph()
    p.add_run(f"Perusahaan: {pt}\nCabang: {branch}\nTeknisi: {tech}\nCustomer: {cust}\nS/N: {sn}\nAlat: {merek} {tipe}\nPeriode: {periode}\nTanggal: {datetime.now().strftime('%d/%m/%Y')}")
    
    # Tabel Checklist
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    hdr = table.rows[0].cells
    hdr[0].text = 'Pekerjaan'
    hdr[1].text = 'Hasil'
    for item in checklist:
        row = table.add_row().cells
        row[0].text = item['Langkah']
        row[1].text = item['Hasil']
    
    doc.add_paragraph(f"\nCatatan: {note}")

    # --- TANDA TANGAN ---
    doc.add_paragraph("\n")
    ttd_table = doc.add_table(rows=3, cols=2)
    ttd_table.autofit = True
    
    ttd_table.rows[0].cells[0].text = "Technician,"
    ttd_table.rows[0].cells[1].text = "Customer Approval,"
    
    if ttd_t:
        run_tek = ttd_table.rows[1].cells[0].paragraphs[0].add_run()
        run_tek.add_picture(ttd_t, width=Inches(1.3))
    if ttd_c:
        run_cus = ttd_table.rows[1].cells[1].paragraphs[0].add_run()
        run_cus.add_picture(ttd_c, width=Inches(1.3))
    
    ttd_table.rows[2].cells[0].text = f"( {tech} )"
    ttd_table.rows[2].cells[1].text = f"( {cust if cust else '.......'} )"

    # --- DOKUMENTASI FOTO ---
    if foto_list:
        doc.add_page_break()
        doc.add_heading('DOKUMENTASI FOTO', level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        table_foto = doc.add_table(rows=0, cols=2)
        row_cells = None
        for i, foto in enumerate(foto_list):
            if i % 2 == 0:
                row_cells = table_foto.add_row().cells
            cell = row_cells[i % 2]
            p_foto = cell.paragraphs[0]
            run_foto = p_foto.add_run()
            run_foto.add_picture(foto, width=Inches(2.8))
            p_foto.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
    target = BytesIO()
    doc.save(target)
    return target.getvalue()

# --- 4. NAVIGATION & HEADER ---
with st.sidebar:
    st.markdown("### MAIN MENU")
    menu = st.radio("NAVIGASI", ["Maintenance", "Service", "Pengaturan"], label_visibility="collapsed")
    st.markdown("---")
    nama_teknisi = st.selectbox("Nama Teknisi", st.session_state.list_teknisi)
    id_teknisi = st.text_input("Jabatan", "TEKNISI")
    nama_pt = st.text_input("Perusahaan", "PT. RAJAERBA INDOCHEM")
    lokasi_pt = st.selectbox("Cabang", ["Tangerang", "Semarang", "Jakarta", "Medan"])

st.markdown(f"""<div class="header-container"><div class="rei-logo-main">REI</div><div class="text-title-wrapper"><p class="e-report">E-REPORT</p><p class="teknisi">TEKNISI</p></div></div>""", unsafe_allow_html=True)

# --- 5. MAIN FORM LOGIC ---
if menu in ["Maintenance", "Service"]:
    st.subheader(f"📝 Form {menu}")
    
    if menu == "Maintenance":
        periode = st.selectbox("Periode Maintenance:", ["Bulanan", "3 Bulan", "6 Bulan", "12 Bulan"])
    else:
        periode = "Service"
    
    with st.container(border=True):
        st.markdown("##### Unit & Customer")
        c1, c2 = st.columns(2)
        with c1:
            merek = st.selectbox("Merek Alat", list(st.session_state.db_alat.keys()))
            tipe = st.selectbox("Tipe Alat", st.session_state.db_alat[merek])
            sn_alat = st.text_input("S/N Alat")
        with c2:
            nama_customer = st.text_input("Nama Customer")
            alamat_customer = st.text_input("Alamat")
            nomor_telepon = st.text_input("Nomor Telepon")

    # Ambil Data Checklist
    data_alat = st.session_state.db_checklist.get(merek, {}).get(tipe, {}).get(periode if periode in st.session_state.db_checklist.get(merek, {}).get(tipe, {}) else "Bulanan", {})
    hasil_checklist = []

    if menu == "Maintenance":
        if "Mati" in data_alat:
            st.markdown("##### 🔴 KONDISI ALAT MATI")
            with st.container(border=True):
                for i, step in enumerate(data_alat["Mati"]):
                    col1, col2 = st.columns([3, 2])
                    col1.write(f"**{i+1}.** {step}")
                    res = col2.radio("Status", ["Cleaned", "Not Cleaned"], key=f"m_{i}", horizontal=True, label_visibility="collapsed")
                    hasil_checklist.append({"Langkah": f"(Mati) {step}", "Hasil": res})
                    if i < len(data_alat["Mati"]) - 1: st.divider()
        if "Hidup" in data_alat:
            st.markdown("##### 🟢 KONDISI ALAT HIDUP")
            with st.container(border=True):
                for i, step in enumerate(data_alat["Hidup"]):
                    col1, col2 = st.columns([3, 2])
                    col1.write(f"**{i+1}.** {step}")
                    res = col2.radio("Status", ["Pass", "Fail", "⚠️ Warning"], key=f"h_{i}", horizontal=True, label_visibility="collapsed")
                    hasil_checklist.append({"Langkah": f"(Hidup) {step}", "Hasil": res})
                    if i < len(data_alat["Hidup"]) - 1: st.divider()
    else:
        st.markdown("##### 🛠️ DETAIL TINDAKAN SERVICE")
        data_service = st.session_state.db_checklist.get(merek, {}).get(tipe, {}).get("Service")
        if not data_service:
            template = st.session_state.db_checklist.get(merek, {}).get(tipe, {}).get("Bulanan", {})
            data_service = template.get("Mati", []) + template.get("Hidup", []) if isinstance(template, dict) else template

        with st.container(border=True):
            if data_service:
                for i, step in enumerate(data_service):
                    col1, col2 = st.columns([3, 2])
                    col1.write(f"**{i+1}.** {step}")
                    res = col2.radio("Hasil", ["Selesai", "Pending", "Ganti Part"], key=f"svc_{i}", horizontal=True, label_visibility="collapsed")
                    hasil_checklist.append({"Langkah": f"(Service) {step}", "Hasil": res})
                    if i < len(data_service) - 1: st.divider()

    st.markdown("##### DOKUMENTASI & VALIDASI")
    with st.container(border=True):
        catatan = st.text_area("Catatan Tambahan")
        foto_alat = st.file_uploader("Upload Foto Lapangan", accept_multiple_files=True, type=['jpg','png'])
        v1, v2 = st.columns(2)
        with v1:
            ttd_tek = st.file_uploader("TTD Teknisi", type=['png','jpg'], key="tek_ttd")
        with v2:
            ttd_cus = st.file_uploader("TTD Customer", type=['png','jpg'], key="cus_ttd")

    st.markdown("---")
    
    # Inisialisasi status di session_state biar gak ilang pas didownload
    if 'proses_selesai' not in st.session_state:
        st.session_state.proses_selesai = False
        st.session_state.data_excel = None
        st.session_state.data_word = None
        st.session_state.nama_file = ""
    
    if st.button("➣ SUBMIT & GENERATE REPORT"):
        if not sn_alat or not nama_customer:
            st.error("⚠️ S/N Alat dan Nama Customer wajib diisi!")
        else:
            # --- 1. SIAPKAN DATA LENGKAP UNTUK GOOGLE SHEETS ---
            # Gabungkan checklist jadi satu teks panjang buat kolom Detail
            checklist_text = "\n".join([f"{c['Langkah']}: {c['Hasil']}" for c in hasil_checklist])
            
            data_laporan = {
                "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "Teknisi": nama_teknisi,
                "Customer": nama_customer,
                "Alat": f"{merek} {tipe}",
                "SN": sn_alat,
                "Periode": periode,
                "Catatan": catatan,
                "Detail_Checklist": checklist_text
            }
            
            with st.status("Memproses laporan...", expanded=True) as status:
                st.write("Mengirim data lengkap ke database...")
                # Panggil fungsi save_to_sheets dengan data_laporan yang sudah lengkap
                sheets_ok = save_to_sheets(data_laporan)
                
                if sheets_ok:
                    st.write("✅ Data berhasil masuk ke Google Sheets!")
                else:
                    st.write("❌ Gagal mengirim ke Sheets, tapi file tetap dibuat...")
                
                # --- 2. LOGIKA EXCEL RAPI (SESUAI FILE REFERENSI) ---
                f_tek = nama_teknisi.replace(" ", "_")
                f_cus = nama_customer.replace(" ", "_")
                f_sn = sn_alat.replace("/", "-")
                
            #Excel 
        if menu == "Maintenance":
            if periode == "Bulanan": p_title = "MONTHLY MAINTENANCE"
            elif periode == "3 Bulan": p_title = "QUARTERLY MAINTENANCE (3 MONTHS)"
            elif periode == "6 Bulan": p_title = "SEMI-ANNUAL MAINTENANCE (6 MONTHS)"
            elif periode == "12 Bulan": p_title = "ANNUAL MAINTENANCE (12 MONTHS)"
            else: p_title = "PERIODIC MAINTENANCE"
        else:
            p_title = "SERVICE & REPAIR"

        clean_cus = nama_customer.replace(" ", "_") if nama_customer else "Customer"
        nama_file_excel =clean_cus = nama_customer.replace(" ", "_") if nama_customer else "Customer"
        tgl_file = datetime.now().strftime('%d%m%y') # Format tanggal: 150526
        
        # Nama file baru: Judul-NamaCustomer-Tanggal.xlsx
        nama_file_excel = f"{p_title.replace(' ', '_')}-{clean_cus}-{tgl_file}.xlsx"
        nama_file_word = f"{p_title.replace(' ', '_')}-{clean_cus}-{tgl_file}.docx"
        
        # 2. Proses Pembuatan File
        output_excel = BytesIO() # Gunakan nama output_excel biar sinkron sama tombol download lu
        with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
            workbook  = writer.book
            worksheet = workbook.add_worksheet('Report')
            
            # FORMATTING
            navy_blue = '#1c3d73'
            border_col = '#D1D5DB'
            title_fmt = workbook.add_format({'bold': True, 'font_size': 20, 'font_color': navy_blue, 'align': 'center', 'valign': 'vcenter'})
            type_fmt = workbook.add_format({'bold': True, 'font_size': 13, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#E2E8F0', 'font_color': '#1E293B', 'border': 1, 'border_color': border_col})
            label_fmt = workbook.add_format({'bold': True, 'bg_color': '#F3F4F6', 'border': 1, 'border_color': border_col, 'font_size': 10})
            data_fmt = workbook.add_format({'border': 1, 'border_color': border_col, 'text_wrap': True, 'font_size': 10, 'valign': 'top'})
            header_table = workbook.add_format({'bold': True, 'fg_color': navy_blue, 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
            name_label_fmt = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'top', 'top': 1})
            
            # LAYOUT
            worksheet.set_column('A:A', 6); worksheet.set_column('B:B', 50); worksheet.set_column('C:C', 18); worksheet.set_column('E:E', 20); worksheet.set_column('F:F', 35)
            
            # HEADER
            worksheet.merge_range('A1:F2', 'PT. RAJAERBA INDOCHEM', title_fmt)
            worksheet.merge_range('A3:F3', f"{p_title} REPORT", type_fmt)
            
            # --- IDENTITAS (Kanan) ---
            info_pairs = [
                ("No. Laporan", f"REI/{menu[:3].upper()}/{datetime.now().strftime('%Y%m%d')}/{sn_alat[:30] if sn_alat else '000'}/{nama_teknisi.split()[0][:10].upper() if nama_teknisi else 'XXX'}"),
                ("Tanggal", datetime.now().strftime('%d %B %Y')),
                ("Teknisi", nama_teknisi),
                ("Customer", nama_customer),
                ("Model Alat", f"{merek} {tipe}"),
                ("S/N Number", sn_alat)
            ]
             
            curr_row = 4
            for label, val in info_pairs:
                worksheet.write(curr_row, 4, label, label_fmt)
                worksheet.write(curr_row, 5, val, data_fmt)
                curr_row += 1
            
            # TABEL CHECKLIST (KIRI)
            worksheet.write(4, 0, 'NO', header_table)
            worksheet.write(4, 1, 'DESKRIPSI PEKERJAAN', header_table)
            worksheet.write(4, 2, 'STATUS', header_table)
            
            row_idx = 5
            for i, item in enumerate(hasil_checklist):
                bg = '#FFFFFF' if i % 2 == 0 else '#F9FAFB'
                row_f = workbook.add_format({'border': 1, 'border_color': border_col, 'bg_color': bg, 'text_wrap': True})
                worksheet.write(row_idx, 0, i+1, row_f)
                worksheet.write(row_idx, 1, item['Langkah'], row_f)
                worksheet.write(row_idx, 2, item['Hasil'], row_f)
                row_idx += 1

            # --- TANDA TANGAN ---
            sig_text_fmt = workbook.add_format({'italic': True, 'align': 'center', 'font_size': 10})
            sig_row = len(hasil_checklist) + 6
            worksheet.set_row(sig_row + 1, 75)
            
            worksheet.merge_range(sig_row, 0, sig_row, 1, "Customer Approval,", workbook.add_format({'italic': True, 'align': 'center'}))
            worksheet.write(sig_row, 2, "Technician,", workbook.add_format({'italic': True, 'align': 'center'}))
            
            if ttd_cus:
                worksheet.insert_image(sig_row + 1, 0, "ttd_cus.png", {'image_data': ttd_cus, 'x_scale': 0.45, 'y_scale': 0.45, 'x_offset': 60})
            else:
                worksheet.merge_range(sig_row + 1, 0, sig_row + 1, 1, "(Tanda Tangan Belum Terlampir)", sig_text_fmt)
            
            if ttd_tek:
                worksheet.insert_image(sig_row + 1, 2, "ttd_tek.png", {'image_data': ttd_tek, 'x_scale': 0.45, 'y_scale': 0.45, 'x_offset': 10})
            else:
                worksheet.write(sig_row + 1, 2, "(Tanda Tangan Belum Terlampir)", sig_text_fmt)

            worksheet.merge_range(sig_row + 2, 0, sig_row + 2, 1, f"( {nama_customer if nama_customer else '........'} )", name_label_fmt)
            worksheet.write(sig_row + 2, 2, f"( {nama_teknisi if nama_teknisi else '........'} )", name_label_fmt)

        list_foto_input = [f for f in foto_alat] if foto_alat else []
        word_data = generate_docx(hasil_checklist, nama_teknisi, id_teknisi, nama_customer, sn_alat, catatan, merek, tipe, periode, nama_pt, lokasi_pt, list_foto_input, ttd_tek, ttd_cus)
                
        # 4. Simpan ke Session State
        st.session_state.data_excel = output_excel.getvalue()
        st.session_state.data_word = word_data
        st.session_state.nama_file = f"{p_title.replace(' ', '_')}-{nama_customer.replace(' ', '_')}"
        st.session_state.proses_selesai = True
                
        status.update(label="Laporan Siap!", state="complete", expanded=False)
        st.balloons()
        
        # --- AREA DOWNLOAD (TAMPIL TERUS SELAMA PROSES_SELESAI = TRUE) ---
    if st.session_state.proses_selesai:
        st.markdown("### 📥 DOWNLOAD AREA")
        st.info("Data sudah tersimpan di Database. Silakan unduh file di bawah:")
        
        c_dl1, c_dl2 = st.columns(2)
        with c_dl1:
            st.download_button(label="📥 Download Excel", data=st.session_state.data_excel, file_name=f"{st.session_state.nama_file}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        with c_dl2:
            st.download_button(label="📄 Download Word", data=st.session_state.data_word, file_name=f"{st.session_state.nama_file}.docx", use_container_width=True)
        
        # Tombol Reset buat laporan baru
        if st.button("🔄 Buat Laporan Baru"):
            st.session_state.proses_selesai = False
            st.rerun()
            
            no_wa_cs = "6281545939640"
            pesan_wa = f"*LAPORAN REI*\nTeknisi: {nama_teknisi}\nCustomer: {nama_customer}\nSN: {sn_alat}\nStatus: Laporan Terkirim."
            encoded_msg = urllib.parse.quote(pesan_wa)
            st.markdown(f'<a href="https://wa.me/{no_wa_cs}?text={encoded_msg}" target="_blank" class="whatsapp-btn">📲 Share ke WhatsApp CS</a>', unsafe_allow_html=True)