import streamlit as st
import pandas as pd
import numpy as np
import datetime
from PIL import Image
import io

# --- PAGE CONFIGURATION (MOBILE OPTIMIZED) ---
st.set_page_config(
    page_title="App Presensi & LMS Les Bahasa Inggris (Mobile V4)",
    page_icon="ðŸ“±",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- CUSTOM MOBILE-FRIENDLY CSS ---
st.markdown("""
    <style>
    .main-title {
        font-size: 22px;
        font-weight: 800;
        color: #1E3A8A;
        text-align: center;
        margin-bottom: 5px;
    }
    .role-badge {
        background-color: #DBEAFE;
        color: #1E40AF;
        padding: 6px 12px;
        border-radius: 20px;
        font-size: 13px;
        font-weight: bold;
        text-align: center;
        margin-bottom: 15px;
    }
    .mobile-card {
        background-color: #FFFFFF;
        border: 1px solid #E5E7EB;
        padding: 12px;
        border-radius: 12px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        margin-bottom: 10px;
    }
    .stButton>button {
        width: 100%;
        border-radius: 10px;
        height: 45px;
        font-weight: bold;
    }
    .download-btn-card {
        background-color: #F0FDF4;
        border: 1px solid #BBF7D0;
        padding: 15px;
        border-radius: 12px;
        margin-top: 10px;
        margin-bottom: 15px;
    }
    </style>
""", unsafe_allow_html=True)

# --- HELPER FUNCTION: EXCEL EXPORTER ---
def generate_excel_recap(df_presensi, df_jurnal, df_nilai, df_siswa):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_presensi.to_excel(writer, sheet_name='Rekap Presensi', index=False)
        df_jurnal.to_excel(writer, sheet_name='Jurnal Mengajar', index=False)
        df_nilai.to_excel(writer, sheet_name='Daftar Nilai', index=False)
        df_siswa.to_excel(writer, sheet_name='Data Siswa', index=False)
    return output.getvalue()

# --- INITIALIZE SESSION STATE ---
if 'data_siswa' not in st.session_state:
    st.session_state['data_siswa'] = pd.DataFrame([
        {"ID": "S01", "Nama": "Andi Pratama", "Kelas": "General Speaking - A", "Tipe": "Les Online", "Status": "Aktif"},
        {"ID": "S02", "Nama": "Siti Rahmawati", "Kelas": "IELTS Prep - B", "Tipe": "Tatap Muka", "Status": "Aktif"},
        {"ID": "S03", "Nama": "Budi Santoso", "Kelas": "Private Conversation", "Tipe": "Privat Rumah", "Status": "Aktif"},
        {"ID": "S04", "Nama": "Citra Dewi", "Kelas": "General Speaking - A", "Tipe": "Les Online", "Status": "Aktif"},
        {"ID": "S05", "Nama": "Doni Wijaya", "Kelas": "IELTS Prep - B", "Tipe": "Tatap Muka", "Status": "Aktif"},
    ])

if 'data_presensi' not in st.session_state:
    st.session_state['data_presensi'] = [
        {"Tanggal": datetime.date(2026, 10, 1), "ID Siswa": "S01", "Nama": "Andi Pratama", "Kelas": "General Speaking - A", "Modus": "Les Online", "Status Kehadiran": "Hadir", "Lokasi/GPS": "Zoom Meeting (Online)", "Foto Bukti": None, "Catatan": "Koneksi lancar"},
        {"Tanggal": datetime.date(2026, 10, 1), "ID Siswa": "S02", "Nama": "Siti Rahmawati", "Kelas": "IELTS Prep - B", "Modus": "Tatap Muka", "Status Kehadiran": "Hadir", "Lokasi/GPS": "Gedung Les R.01", "Foto Bukti": None, "Catatan": "Datang tepat waktu"},
        {"Tanggal": datetime.date(2026, 10, 2), "ID Siswa": "S03", "Nama": "Budi Santoso", "Kelas": "Private Conversation", "Modus": "Privat Rumah", "Status Kehadiran": "Hadir", "Lokasi/GPS": "-6.2088, 106.8456 (GPS Verified)", "Foto Bukti": None, "Catatan": "Check-in rumah siswa"},
    ]

if 'jurnal_mengajar' not in st.session_state:
    st.session_state['jurnal_mengajar'] = pd.DataFrame([
        {"Pertemuan Ke": 1, "Tanggal": datetime.date(2026, 10, 1), "Kelas": "General Speaking - A", "Modus": "Les Online", "Pengajar": "Miss Sarah", "Topik/Materi": "Introducing Yourself & Small Talk", "Ringkasan Pembelajaran": "Siswa berlatih roleplay perkenalan."},
        {"Pertemuan Ke": 1, "Tanggal": datetime.date(2026, 10, 1), "Kelas": "IELTS Prep - B", "Modus": "Tatap Muka", "Pengajar": "Mr. John", "Topik/Materi": "IELTS Speaking Part 1 Practice", "Ringkasan Pembelajaran": "Latihan topik hobi & pekerjaan."},
        {"Pertemuan Ke": 1, "Tanggal": datetime.date(2026, 10, 2), "Kelas": "Private Conversation", "Modus": "Privat Rumah", "Pengajar": "Miss Sarah", "Topik/Materi": "Business Vocabulary & Negotiation", "Ringkasan Pembelajaran": "Latihan negosiasi bahasa Inggris."},
    ])

if 'materi_pelajaran' not in st.session_state:
    st.session_state['materi_pelajaran'] = pd.DataFrame([
        {"Kode Modul": "MOD-01", "Judul": "Mastering English Pronunciation", "Kelas": "General Speaking - A", "Format": "PDF & Audio MP3", "Link/Akses": "https://drive.google.com/materi-01"},
        {"Kode Modul": "MOD-02", "Judul": "IELTS Speaking Cue Cards 2026", "Kelas": "IELTS Prep - B", "Format": "E-Book PDF", "Link/Akses": "https://drive.google.com/materi-02"},
        {"Kode Modul": "MOD-03", "Judul": "Daily Conversation & Idioms", "Kelas": "Private Conversation", "Format": "Audio & Video", "Link/Akses": "https://drive.google.com/materi-03"},
    ])

if 'daftar_nilai' not in st.session_state:
    st.session_state['daftar_nilai'] = pd.DataFrame([
        {"ID Siswa": "S01", "Nama": "Andi Pratama", "Kelas": "General Speaking - A", "Speaking Test": 85, "Listening Quiz": 90, "Grammar & Vocab": 80, "Tugas Mandiri": 88, "Nilai Akhir": 85.75},
        {"ID Siswa": "S02", "Nama": "Siti Rahmawati", "Kelas": "IELTS Prep - B", "Speaking Test": 78, "Listening Quiz": 82, "Grammar & Vocab": 85, "Tugas Mandiri": 80, "Nilai Akhir": 81.25},
        {"ID Siswa": "S03", "Nama": "Budi Santoso", "Kelas": "Private Conversation", "Speaking Test": 92, "Listening Quiz": 88, "Grammar & Vocab": 90, "Tugas Mandiri": 95, "Nilai Akhir": 91.25},
    ])

if 'daftar_penugasan' not in st.session_state:
    st.session_state['daftar_penugasan'] = pd.DataFrame([
        {"ID Tugas": "TUG-01", "Judul Tugas": "Audio Recording: Self Introduction", "Kelas": "General Speaking - A", "Tenggat Waktu": datetime.date(2026, 10, 15), "Status": "Terbit", "Instruksi": "Kirim rekaman suara (MP3/M4A) via HP"},
        {"ID Tugas": "TUG-02", "Judul Tugas": "IELTS Speaking Part 2 Mock Record", "Kelas": "IELTS Prep - B", "Tenggat Waktu": datetime.date(2026, 10, 18), "Status": "Terbit", "Instruksi": "Topik: Describe a memorable journey"},
    ])

# --- TOP MOBILE HEADER & ROLE SWITCHER ---
st.markdown("<div class='main-title'>ðŸ“± Presensi & LMS English Class</div>", unsafe_allow_html=True)

role = st.selectbox(
    "ðŸ“² Pilih Mode Peran Pengguna (Tampilan HP):",
    ["ðŸ‘‘ Mode Admin / Pemilik Les", "ðŸ‘¨â€ðŸ« Mode Guru / Tutor", "ðŸ‘¤ Mode Siswa / Murid", "ðŸ“ Form Pendaftaran Mandiri (Public)"],
    index=0
)

st.markdown(f"<div class='role-badge'>Aktif: {role}</div>", unsafe_allow_html=True)


# ==============================================================================
# MODE 1: ADMIN / PEMILIK LES (Dashboard Rekap, Export Excel, Foto Bukti, Master Data)
# ==============================================================================
if role == "ðŸ‘‘ Mode Admin / Pemilik Les":
    tab_a1, tab_a2, tab_a3, tab_a4 = st.tabs([
        "ðŸ“Š Dashboard & Monitoring",
        "ðŸ“¥ Ekspor & Download Laporan",
        "ðŸ“· Rekap Presensi & Foto GPS",
        "âš™ï¸ Kelola Master Data"
    ])

    with tab_a1:
        st.subheader("ðŸ“Š Ringkasan Operasional Kursus")
        m1, m2, m3 = st.columns(3)
        with m1:
            st.metric("Total Siswa", len(st.session_state['data_siswa']))
        with m2:
            st.metric("Record Absensi", len(st.session_state['data_presensi']))
        with m3:
            st.metric("Pertemuan Selesai", len(st.session_state['jurnal_mengajar']))

        st.markdown("---")
        st.write("ðŸ“ˆ **Distribusi Modus Pembelajaran Siswa:**")
        st.bar_chart(st.session_state['data_siswa']['Tipe'].value_counts())

    with tab_a2:
        st.subheader("ðŸ“¥ Ekspor Laporan Otomatis ke HP")
        st.caption("Unduh laporan lengkap dalam format **Excel (.xlsx)** atau **CSV** langsung ke galeri/file manager HP Anda.")

        # Prepare export DataFrame (removing raw photo bytes for clean tabular output)
        presensi_export_data = []
        for p in st.session_state['data_presensi']:
            p_copy = p.copy()
            p_copy['Ada Foto Bukti'] = "Ya" if p['Foto Bukti'] is not None else "Tidak"
            p_copy.pop('Foto Bukti', None)
            presensi_export_data.append(p_copy)
        
        df_p_export = pd.DataFrame(presensi_export_data)

        # Date Range & Filter
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            st_date = st.date_input("Dari Tanggal", datetime.date(2026, 10, 1))
        with col_f2:
            end_date = st.date_input("Sampai Tanggal", datetime.date.today())

        # Filtered Export Data
        if not df_p_export.empty and 'Tanggal' in df_p_export.columns:
            df_p_filtered = df_p_export[(df_p_export['Tanggal'] >= st_date) & (df_p_export['Tanggal'] <= end_date)]
        else:
            df_p_filtered = df_p_export

        st.markdown("<div class='download-btn-card'>", unsafe_allow_html=True)
        st.write("ðŸ“Š **1. Laporan Terpadu (Excel - Multi Sheet)**")
        st.caption("Mencakup Sheet: Rekap Presensi, Jurnal Mengajar, Daftar Nilai, & Master Data Siswa.")
        
        excel_bytes = generate_excel_recap(
            df_p_filtered,
            st.session_state['jurnal_mengajar'],
            st.session_state['daftar_nilai'],
            st.session_state['data_siswa']
        )
        
        st.download_button(
            label="ðŸ“¥ Download Laporan Lengkap (.xlsx Excel)",
            data=excel_bytes,
            file_name=f"Rekap_Lengkap_Les_Bahasa_Inggris_{datetime.date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div class='download-btn-card'>", unsafe_allow_html=True)
        st.write("ðŸ“„ **2. Ekspor Khusus Presensi Kehadiran (.csv)**")
        csv_presensi = df_p_filtered.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="ðŸ“„ Download Rekap Absensi (.csv)",
            data=csv_presensi,
            file_name=f"Rekap_Presensi_{datetime.date.today()}.csv",
            mime="text/csv",
            use_container_width=True
        )
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div class='download-btn-card'>", unsafe_allow_html=True)
        st.write("ðŸ’¯ **3. Ekspor Laporan Nilai Siswa (.csv)**")
        csv_nilai = st.session_state['daftar_nilai'].to_csv(index=False).encode('utf-8')
        st.download_button(
            label="ðŸ’¯ Download Laporan Nilai (.csv)",
            data=csv_nilai,
            file_name=f"Laporan_Nilai_Siswa_{datetime.date.today()}.csv",
            mime="text/csv",
            use_container_width=True
        )
        st.markdown("</div>", unsafe_allow_html=True)

    with tab_a3:
        st.subheader("ðŸ“· Galeri Rekap Presensi & Bukti Selfie / GPS")
        for idx, item in enumerate(reversed(st.session_state['data_presensi'])):
            with st.expander(f"ðŸ“ {item['Tanggal']} | {item['Nama']} - {item['Modus']} ({item['Status Kehadiran']})"):
                st.write(f"**Kelas:** {item['Kelas']}")
                st.write(f"**Lokasi / Titik GPS:** {item['Lokasi/GPS']}")
                st.write(f"**Catatan:** {item['Catatan']}")
                
                if item['Foto Bukti'] is not None:
                    try:
                        image = Image.open(io.BytesIO(item['Foto Bukti']))
                        st.image(image, caption=f"Foto Selfie / Verifikasi Lokasi ({item['Nama']})", width=300)
                    except Exception as e:
                        st.caption("ðŸ“· Foto tersimpan dalam format file.")
                else:
                    st.caption("â„¹ï¸ Tidak ada lampiran foto (Presensi Manual/Online).")

    with tab_a4:
        st.subheader("âš™ï¸ Kelola Data Siswa & Kursus")
        st.dataframe(st.session_state['data_siswa'], use_container_width=True)

        with st.expander("âž• Tambah Siswa Baru"):
            with st.form("form_tambah_siswa_admin"):
                s_id = f"S0{len(st.session_state['data_siswa'])+1}"
                s_n = st.text_input("Nama Siswa Baru")
                s_k = st.selectbox("Kelas", ["General Speaking - A", "IELTS Prep - B", "Private Conversation", "Kids English"])
                s_t = st.selectbox("Tipe Pembelajaran", ["Privat Rumah", "Tatap Muka", "Les Online"])

                btn_add_siswa = st.form_submit_button("âž• Daftarkan Siswa")
                if btn_add_siswa:
                    new_s_data = {"ID": s_id, "Nama": s_n, "Kelas": s_k, "Tipe": s_t, "Status": "Aktif"}
                    st.session_state['data_siswa'] = pd.concat([st.session_state['data_siswa'], pd.DataFrame([new_s_data])], ignore_index=True)
                    
                    # Add to score table too
                    new_n_data = {"ID Siswa": s_id, "Nama": s_n, "Kelas": s_k, "Speaking Test": 0, "Listening Quiz": 0, "Grammar & Vocab": 0, "Tugas Mandiri": 0, "Nilai Akhir": 0.0}
                    st.session_state['daftar_nilai'] = pd.concat([st.session_state['daftar_nilai'], pd.DataFrame([new_n_data])], ignore_index=True)
                    st.success(f"âœ… Siswa {s_n} berhasil didaftarkan!")


# ==============================================================================
# MODE 2: GURU / TUTOR (Absensi Kamera Selfie, GPS, Jurnal, Penilaian, Download Rekap)
# ==============================================================================
elif role == "ðŸ‘¨â€ðŸ« Mode Guru / Tutor":
    tab_g1, tab_g2, tab_g3, tab_g4, tab_g5 = st.tabs([
        "ðŸ“¸ Absensi & Selfie",
        "ðŸ“– Jurnal Mengajar",
        "ðŸ’¯ Input Nilai",
        "ðŸ“Œ Terbit Tugas",
        "ðŸ“¥ Download Rekap"
    ])

    with tab_g1:
        st.subheader("ðŸ“¸ Input Presensi Siswa (Kamera HP)")
        st.caption("Gunakan kamera HP untuk verifikasi foto selfie / lokasi saat kelas privat atau tatap muka.")

        with st.form("form_presensi_kamera", clear_on_submit=True):
            s_nama = st.selectbox("Pilih Siswa", options=st.session_state['data_siswa']['Nama'].tolist())
            selected_student = st.session_state['data_siswa'][st.session_state['data_siswa']['Nama'] == s_nama].iloc[0]
            
            col1, col2 = st.columns(2)
            with col1:
                tgl_sesi = st.date_input("Tanggal", datetime.date.today())
                modus = st.selectbox("Modus Pembelajaran", ["Privat Rumah", "Tatap Muka", "Les Online"], index=["Privat Rumah", "Tatap Muka", "Les Online"].index(selected_student['Tipe']))
            with col2:
                status_kehadiran = st.selectbox("Status Kehadiran", ["Hadir", "Izin", "Sakit", "Alpha"])
                lokasi_gps = st.text_input("Koordinat GPS / Alamat / Platform", value="-6.2088, 106.8456 (Verified GPS HP)" if modus=="Privat Rumah" else ("Ruang Kelas 1" if modus=="Tatap Muka" else "Zoom Meeting"))

            st.write("ðŸ“· **Ambil Foto Selfie / Bukti Lokasi via HP:**")
            img_file_buffer = st.camera_input("Buka Kamera HP untuk Ambil Foto Bukti")
            upload_backup = st.file_uploader("Atau Unggah Foto dari Galeri HP", type=['jpg', 'png', 'jpeg'])

            catatan_sesi = st.text_area("Catatan Evaluasi Sesi", placeholder="Contoh: Pronunciation and vocabulary extended well.")

            btn_simpan_presensi = st.form_submit_button("ðŸš€ Simpan Presensi & Bukti Foto")

            if btn_simpan_presensi:
                foto_final = None
                if img_file_buffer is not None:
                    foto_final = img_file_buffer.getvalue()
                elif upload_backup is not None:
                    foto_final = upload_backup.getvalue()

                new_record = {
                    "Tanggal": tgl_sesi,
                    "ID Siswa": selected_student['ID'],
                    "Nama": s_nama,
                    "Kelas": selected_student['Kelas'],
                    "Modus": modus,
                    "Status Kehadiran": status_kehadiran,
                    "Lokasi/GPS": lokasi_gps,
                    "Foto Bukti": foto_final,
                    "Catatan": catatan_sesi
                }
                st.session_state['data_presensi'].append(new_record)
                st.success(f"âœ… Presensi {s_nama} berhasil dicatat dengan foto bukti!")

    with tab_g2:
        st.subheader("ðŸ“– Isi Jurnal Mengajar Harian")
        with st.form("form_jurnal_guru"):
            p_ke = st.number_input("Pertemuan Ke-", min_value=1, value=len(st.session_state['jurnal_mengajar']) + 1)
            tgl_j = st.date_input("Tanggal Pertemuan", datetime.date.today())
            kls_j = st.selectbox("Kelas Target", st.session_state['data_siswa']['Kelas'].unique().tolist())
            mds_j = st.selectbox("Modus Kelas", ["Privat Rumah", "Tatap Muka", "Les Online"])
            pengajar_n = st.text_input("Nama Pengajar / Tutor", value="Miss Sarah")
            topik_m = st.text_input("Topik Bahasa Inggris Hari Ini", placeholder="Contoh: Phrasal Verbs & Speaking Practice")
            ringkasan_m = st.text_area("Ringkasan Materi & Evaluasi Siswa")

            btn_jurnal_save = st.form_submit_button("ðŸ’¾ Simpan Jurnal Mengajar")
            if btn_jurnal_save:
                new_j = {
                    "Pertemuan Ke": p_ke,
                    "Tanggal": tgl_j,
                    "Kelas": kls_j,
                    "Modus": mds_j,
                    "Pengajar": pengajar_n,
                    "Topik/Materi": topik_m,
                    "Ringkasan Pembelajaran": ringkasan_m
                }
                st.session_state['jurnal_mengajar'] = pd.concat([st.session_state['jurnal_mengajar'], pd.DataFrame([new_j])], ignore_index=True)
                st.success("âœ… Jurnal mengajar berhasil disimpan!")

    with tab_g3:
        st.subheader("ðŸ’¯ Penilaian Siswa")
        df_n = st.session_state['daftar_nilai']
        s_dinilai = st.selectbox("Pilih Nama Siswa", df_n['Nama'].tolist())
        
        with st.form("form_nilai_guru"):
            spk = st.slider("Speaking Test", 0, 100, 85)
            lis = st.slider("Listening Quiz", 0, 100, 80)
            grm = st.slider("Grammar & Vocab", 0, 100, 85)
            tgs = st.slider("Tugas Mandiri", 0, 100, 90)

            btn_save_nilai = st.form_submit_button("ðŸ’¾ Simpan Nilai")
            if btn_save_nilai:
                n_akhir = (spk * 0.35) + (lis * 0.25) + (grm * 0.20) + (tgs * 0.20)
                idx = st.session_state['daftar_nilai'][st.session_state['daftar_nilai']['Nama'] == s_dinilai].index[0]
                st.session_state['daftar_nilai'].at[idx, 'Speaking Test'] = spk
                st.session_state['daftar_nilai'].at[idx, 'Listening Quiz'] = lis
                st.session_state['daftar_nilai'].at[idx, 'Grammar & Vocab'] = grm
                st.session_state['daftar_nilai'].at[idx, 'Tugas Mandiri'] = tgs
                st.session_state['daftar_nilai'].at[idx, 'Nilai Akhir'] = n_akhir
                st.success(f"âœ… Nilai {s_dinilai} berhasil diupdate: {n_akhir:.2f}")

    with tab_g4:
        st.subheader("ðŸ“Œ Terbitkan Tugas Baru")
        with st.form("form_tugas_guru"):
            t_judul = st.text_input("Judul Tugas", placeholder="Contoh: Record Speaking Audio 2 Menit")
            t_kelas = st.selectbox("Untuk Kelas", st.session_state['data_siswa']['Kelas'].unique().tolist())
            t_dl = st.date_input("Deadline", datetime.date.today() + datetime.timedelta(days=7))
            t_inst = st.text_area("Instruksi Tugas Mandiri")

            btn_tugas_pub = st.form_submit_button("ðŸš€ Diterbitkan ke Siswa")
            if btn_tugas_pub:
                new_t =
