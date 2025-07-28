from docx import Document

# Membuat dokumen Word
doc = Document()
doc.add_heading('Program Pecahkan Angka – Versi Unplugged', 0)

# Tujuan Pembelajaran
doc.add_heading('🎯 Tujuan Pembelajaran', level=1)
doc.add_paragraph(
    "- Menyebutkan dan memecah bilangan cacah hingga 100.000 berdasarkan nilai tempatnya.\n"
    "- Menghubungkan konsep nilai tempat dengan bentuk panjang bilangan.\n"
    "- Melatih logika berpikir sistematis seperti dalam proses pemrograman."
)

# Alat dan Bahan
doc.add_heading('🛠️ Alat dan Bahan', level=1)
doc.add_paragraph(
    "- Kartu angka besar (misal: 5 digit dan 6 digit, contoh: 42.317, 89.062)\n"
    "- Kartu nilai tempat: Puluhan Ribu, Ribuan, Ratusan, Puluhan, Satuan\n"
    "- Kartu bilangan (angka dari 0–9)\n"
    "- Worksheet / lembar kerja siswa\n"
    "- Papan tulis atau chart “nilai tempat”\n"
    "- Spidol dan kertas\n"
    "- Amplop untuk setiap kelompok (berisi angka acak dan kartu nilai tempat)"
)

# Alokasi Waktu
doc.add_heading('⏰ Alokasi Waktu', level=1)
table = doc.add_table(rows=1, cols=2)
hdr_cells = table.rows[0].cells
hdr_cells[0].text = 'Kegiatan'
hdr_cells[1].text = 'Waktu'

rows = [
    ('Pendahuluan (orientasi)', '5 menit'),
    ('Demonstrasi guru', '5 menit'),
    ('Aktivitas kelompok', '20 menit'),
    ('Refleksi dan diskusi kelas', '5 menit'),
    ('Penutup (review & soal tantangan)', '5 menit')
]

for kegiatan, waktu in rows:
    row_cells = table.add_row().cells
    row_cells[0].text = kegiatan
    row_cells[1].text = waktu

# Langkah-langkah Kegiatan
doc.add_heading('🧩 Langkah-Langkah Kegiatan', level=1)
doc.add_paragraph('1. Pendahuluan – 5 menit\n'
    '   - Guru mengingatkan kembali konsep nilai tempat pada bilangan cacah besar.\n'
    '   - Tanya jawab ringan dengan siswa.\n\n'
    '2. Demonstrasi Guru – 5 menit\n'
    '   - Guru menuliskan contoh angka besar di papan (misalnya 72.645).\n'
    '   - Siswa bersama-sama memecah angka tersebut berdasarkan nilai tempat.\n\n'
    '3. Aktivitas Kelompok – 20 menit\n'
    '   - Siswa dibagi menjadi kelompok kecil (3–5 orang).\n'
    '   - Setiap kelompok mendapat amplop berisi 3 angka dan kartu nilai tempat.\n'
    '   - Tugas mereka adalah memecah angka ke bentuk panjang dan menuliskannya di worksheet.\n\n'
    '4. Refleksi dan Diskusi – 5 menit\n'
    '   - Beberapa kelompok mempresentasikan hasil mereka.\n\n'
    '5. Penutup – 5 menit\n'
    '   - Guru memberi soal tantangan untuk dipecahkan bersama.'
)

# Contoh Worksheet
doc.add_heading('📝 Contoh Worksheet Sederhana', level=1)
table = doc.add_table(rows=1, cols=7)
hdr_cells = table.rows[0].cells
hdr_cells[0].text = 'Angka'
hdr_cells[1].text = 'Puluhan Ribu'
hdr_cells[2].text = 'Ribuan'
hdr_cells[3].text = 'Ratusan'
hdr_cells[4].text = 'Puluhan'
hdr_cells[5].text = 'Satuan'
hdr_cells[6].text = 'Bentuk Panjang'

data_rows = [
    ['58.109', '50.000', '8.000', '100', '0', '9', '50.000 + 8.000 + 100 + 0 + 9'],
    ['91.840', '90.000', '1.000', '800', '40', '0', '90.000 + 1.000 + 800 + 40 + 0']
]

for row_data in data_rows:
    row_cells = table.add_row().cells
    for i, value in enumerate(row_data):
        row_cells[i].text = value

# Simpan dokumen
doc.save("Program_Pecahkan_Angka_Unplugged.docx")
