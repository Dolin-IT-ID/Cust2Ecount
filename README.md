# Excel Column Mapper with AI 📊

Program konversi Excel menggunakan Streamlit dan Ollama AI untuk membantu pengguna mengidentifikasi dan memetakan kolom-kolom yang relevan antara file sumber dan file target konversi.

## 🚀 Fitur Utama

- **AI-Powered Column Mapping**: Menggunakan Ollama AI untuk analisis semantik kolom
- **Fuzzy String Matching**: Pencarian kolom berdasarkan kesamaan string
- **Interactive UI**: Interface yang user-friendly dengan Streamlit
- **Multi-format Support**: Mendukung Excel (.xlsx, .xls) dan CSV
- **Manual Adjustment**: Kemampuan untuk menyesuaikan mapping secara manual
- **Preview & Download**: Preview hasil konversi dan download file hasil

## 📋 Prerequisites

1. **Python 3.8+**
2. **Ollama** - Install dan jalankan Ollama dengan model yang diinginkan
   ```bash
   # Install Ollama (https://ollama.ai/)
   # Pull model yang diinginkan
   ollama pull llama3.2
   ```

## 🛠️ Installation

1. Clone atau download repository ini
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## 🎯 Cara Penggunaan

### 1. Persiapan File
- Pastikan file `Template_Ecount.xlsx` ada di direktori yang sama
- Siapkan file sumber yang akan dikonversi

### 2. Menjalankan Aplikasi
```bash
streamlit run app.py
```

### 3. Langkah-langkah di Aplikasi

1. **Load Template File**
   - Klik tombol "Load Template File" untuk memuat template
   - Template akan menampilkan struktur kolom target

2. **Upload Source File**
   - Upload file Excel atau CSV yang akan dikonversi
   - Aplikasi akan menampilkan struktur kolom sumber

3. **AI Column Mapping**
   - Pilih model Ollama yang diinginkan di sidebar
   - Klik "Get AI Suggestions" untuk mendapatkan saran mapping dari AI
   - AI akan menganalisis kolom berdasarkan:
     - Kesamaan semantik
     - Kesamaan string
     - Terminologi bisnis umum
     - Dukungan multi-bahasa (English, Chinese, Indonesian)

4. **Fuzzy Match Analysis** (Opsional)
   - Klik "Fuzzy Match Analysis" untuk analisis kesamaan string
   - Atur threshold di sidebar (default: 70%)

5. **Manual Adjustments**
   - Review hasil AI mapping
   - Lakukan penyesuaian manual jika diperlukan
   - Setiap kolom template dapat dipetakan ke kolom sumber yang sesuai

6. **Generate & Download**
   - Klik "Generate Converted File" untuk membuat file hasil
   - Preview data hasil konversi
   - Download file Excel yang sudah dikonversi

## 📊 Contoh Penggunaan

**File Target (Template_Ecount.xlsx):**
- Kolom: Name, Email, Phone, Address, Website, etc.

**File Sumber (SCRAP_DATA.xlsx):**
- Kolom: 公司名称 (Nama Perusahaan), 邮箱 (Email), 电话 (Telepon), etc.

**Hasil AI Mapping:**
- "Email" → "邮箱 (Email)"
- "Website" → "网址 (Website)"
- "Phone" → "电话 (Telepon)"

## ⚙️ Konfigurasi

### Model Ollama
Aplikasi mendukung berbagai model Ollama:
- llama3.2 (recommended)
- llama3.1
- llama2
- mistral
- codellama

### Fuzzy Match Threshold
- Range: 50-100%
- Default: 70%
- Semakin tinggi threshold, semakin ketat kriteria kesamaan

## 🔧 Troubleshooting

### Error: "Could not connect to Ollama"
**Solusi:**
1. Pastikan Ollama sudah terinstall dan berjalan
2. Jalankan: `ollama serve`
3. Pastikan model sudah di-pull: `ollama pull llama3.2`

### Error: "Template file not found"
**Solusi:**
1. Pastikan file `Template_Ecount.xlsx` ada di direktori yang sama dengan `app.py`
2. Periksa nama file dan ekstensi

### Error: "Unsupported file format"
**Solusi:**
1. Gunakan file dengan format .xlsx, .xls, atau .csv
2. Pastikan file tidak corrupt

## 📁 Struktur File

```
Cust2Ecount/
├── app.py                 # Aplikasi utama Streamlit
├── requirements.txt       # Dependencies Python
├── README.md             # Dokumentasi
├── analyze_files.py      # Script analisis file (untuk development)
├── Template_Ecount.xlsx  # File template target
└── SCRAP_DATA.xlsx      # File contoh sumber
```

## 🤝 Contributing

Kontribusi sangat diterima! Silakan:
1. Fork repository
2. Buat feature branch
3. Commit perubahan
4. Push ke branch
5. Buat Pull Request

## 📝 License

Project ini menggunakan MIT License.

## 📞 Support

Jika mengalami masalah atau memiliki pertanyaan, silakan buat issue di repository ini.

---

**Dibuat dengan ❤️ menggunakan Streamlit dan Ollama AI**