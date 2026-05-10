# poinmeet — Catatan Sesi & Konsep Tools
Dibuat: 7 Mei 2026 | Terakhir diupdate: 10 Mei 2026
Lokasi: `D:\proj qc tools\MOM-TOOLS\`

---

## KONSEP TOOLS

Aplikasi desktop (berjalan di browser lokal) untuk generate Minutes of Meeting (MOM) dari transcript rapat secara otomatis menggunakan AI, lalu export ke file Word (.docx) dengan format template DBC.

### Alur Kerja
1. User paste info meeting dari Google Calendar (nama, tanggal, jam)
2. User paste transcript dari Fathom (copy Summary) atau notetaker lain
3. User paste list attendees dari Google Calendar (Guests section)
4. Klik Generate → AI (Groq) proses transcript → hasilkan poin-poin MOM
5. User edit hasil (nama, tanggal, waktu, lokasi, attendees, issues, next meeting)
6. Export ke Word (.docx) format template DBC

---

## STACK TEKNOLOGI

- **Backend**: Python Flask + python-docx + Groq AI
- **Frontend**: HTML/JS single file (index.html), buka di browser
- **AI Model**: llama-3.3-70b-versatile via Groq API
- **Tidak butuh**: Docker, Node.js, database
- **Python deps**: flask, flask-cors, python-docx, groq>=0.11.0, httpx>=0.27.0, python-dotenv

---

## STRUKTUR FILE

```
MOM-TOOLS/
├── install.bat          ← buat venv, install deps, jalankan start.bat
├── start.bat            ← jalankan backend, buka browser
├── README.md
├── logo_dbc.png         ← logo untuk header Word
├── backend/
│   ├── app.py           ← Flask API + serve frontend di /app
│   ├── requirements.txt
│   ├── .env             ← GROQ_API_KEY=...
│   └── venv/
└── frontend/
    └── index.html       ← UI lengkap single file
```

---

## KONFIGURASI

- **Groq API Key**: disimpan di `backend/.env` → `GROQ_API_KEY=gsk_...`
- **Port backend**: 5000
- **URL app**: http://localhost:5000/app
- **Frontend**: di-serve langsung oleh Flask dari folder `frontend/`

---

## FITUR UTAMA

### Input (Step 1)
- Paste info meeting dari Google Calendar → auto-parse nama, tanggal Indonesia, jam 24h
  - Format input: `[Book] Nama Meeting\nThursday, May 7⋅3:30 – 4:30pm`
  - Output: nama meeting, `Kamis, 7 Mei 2026`, `15:30 - 16:30`
- Paste transcript (Fathom Summary atau format speaker)
- Paste attendees dari Google Calendar Guests section
- Input manual tanggal & waktu sebagai fallback
- Koneksi status AI ditampilkan otomatis saat halaman load (tidak ada input API key di UI)

### Edit (Step 2)
- Edit semua field meeting info
- Tanggal: date picker → auto-format ke Indonesia (Jumat, 07 Mei 2026)
- Waktu: dua input HH:MM (start – end)
- Lokasi: default "Online", bisa diedit
- Attendees: tag list, bisa tambah/hapus, auto-detect dari transcript + calendar
- Issues & Actions: tabel editable inline (textarea per baris)
  - Deadline: date picker + text fallback (TBA/Note)
  - Status: dropdown (Open/On Progress/Done/Monitoring/Note/Noted/Progress)
  - PIC: text input
- Next Meeting: tanggal, waktu, lokasi, chairperson, agenda
- Tab step di atas bisa diklik langsung untuk navigasi

### Export (Step 3)
- Format filename: `yyyymmdd_nama-meeting.docx`
  - Tanggal dari transcript → fallback ke input manual
- Template Word: DBC format

---

## TEMPLATE WORD (DBC FORMAT)

### Page Setup
- Ukuran: Letter (21.59 x 27.94 cm), Portrait
- Margin: Top/Bottom/Right 2.54cm, Left 1.25cm
- Font utama: Calibri 9pt

### Header (tiap halaman)
- Tabel 2 kolom: logo_dbc.png (2.97×2.97cm) | "Minute of Meeting" Arial 12pt bold
- Border: #C0C0C0

### Footer (tiap halaman)
- "Page X of Y" — field otomatis PAGE & NUMPAGES, Arial 8pt, rata kanan

### Struktur Dokumen
1. **Meeting Information** — header #4C4C4C putih, label #E6E6E6
   - Fields: Group/Meeting/Project Name, Meeting Date, Meeting Location, Chairperson, Attendees, With Apologies, Doc Circulation, Prepared By, Additional Info
2. **Topic & Attachment** — label #E6E6E6
3. **Issues and Actions** — header bold italic #E6E6E6, 4 kolom
   - Issues and Actions (lebar ~10cm), Deadline (center ~3cm), Status (center ~3cm), PIC (center ~2cm)
   - Header repeat tiap halaman
   - Warna status: Open/On Progress/Progress = `#FFFF00`, Done = `#00FF00`, Monitoring/Note/Noted = `#00FFFF`
4. **Tentative Agenda for the Next Meeting** — header #000000 putih, label #D9D9D9
   - Fields: Proposed Date, Proposed Location, Proposed Time, Chairperson, Proposed Agenda, Additional Info

---

## LOGIKA PENTING

### Parse Info Meeting dari Calendar
- Regex: `(\w+),\s+(\w+)\s+(\d{1,2})(?:\s*·\s*(\d{1,2}:\d{2})\s*[–-]\s*(\d{1,2}:\d{2})(am|pm)?)?`
- Konversi jam: suffix `pm` diterapkan ke KEDUA waktu (start & end)
- Hari/bulan dikonversi ke Indonesia

### Deteksi Attendees
- Dari transcript: regex speaker `[00:01] Name:` atau `Name:`
- Dari calendar paste: regex email → ambil local part → title case
- Keduanya digabung, bisa tambah manual
- Jika keduanya kosong → user isi manual di Step 2 (ada hint)

### AI Generate MOM
- Model: llama-3.3-70b-versatile (Groq)
- Instruksi: deteksi bahasa transcript, jika Inggris translate ke Indonesia (kata per kata + konteks, istilah tanpa padanan tetap asli)
- Output: JSON dengan meeting_name, meeting_date, items[], dll
- API key dari `backend/.env`, tidak pernah expose ke frontend

### Prioritas Data
- Tanggal: calendar paste > manual input Step 1 > AI result
- Waktu: calendar paste > manual input Step 1 > AI result
- Nama meeting: calendar paste > AI result
- Lokasi: AI result, default "Online" jika kosong

---

## INSTALL.BAT PATTERN
Mengikuti pola AKURASIT-TOOLS (simpel):
```bat
cd /d "%~dp0backend"
if exist venv rmdir /s /q venv
python -m venv venv
call venv\Scripts\activate
pip install -r requirements.txt
cd /d "%~dp0"
call start.bat
```

## START.BAT PATTERN
```bat
cd /d "%~dp0backend"
:: kill port 5000
start /B python app.py
:: tunggu health check http://localhost:5000/api/health
start http://localhost:5000/app
pause → kill backend saat ditutup
```

---

## API ENDPOINTS

| Method | Endpoint | Fungsi |
|--------|----------|--------|
| GET | /app | Serve frontend HTML |
| GET | /api/health | Cek status + groq_configured |
| POST | /api/parse-meeting-info | Parse nama/tanggal/jam dari teks Calendar |
| POST | /api/extract-attendees | Detect attendees dari transcript + calendar text |
| POST | /api/generate | Generate MOM via Groq AI (X-API-Key header opsional) |
| POST | /api/export | Build & download .docx |

---

## MASALAH & SOLUSI YANG PERNAH TERJADI

1. **groq + httpx conflict** — `TypeError: proxies`
   - Fix: `pip install "httpx>=0.27.0" "groq>=0.11.0"`

2. **install.bat false failure** — pip notice bikin exit code non-zero
   - Fix: verifikasi pakai `python -c "import flask, groq, docx, dotenv"` bukan exit code pip

3. **Permission denied venv** — venv lama tidak bisa ditimpa
   - Fix: `rmdir /s /q venv` sebelum `python -m venv venv`

4. **extract_attendees not defined** — terhapus saat refactor
   - Fix: pastikan fungsi didefinisikan sebelum `api_generate`

5. **Waktu tidak terkonversi ke 24h** — suffix pm tidak ke start time
   - Fix: `is_pm = 'pm' in suffix` lalu terapkan ke kedua waktu (start & end)

6. **call activate.bat tidak reliable** — python setelah activate masih pakai sistem
   - Fix: gunakan path eksplisit `venv\Scripts\python.exe` untuk semua perintah

---

## SUMBER DATA PER FIELD

| Field | Sumber |
|-------|--------|
| Nama Meeting | Paste Google Calendar → AI → manual |
| Tanggal | Paste Google Calendar → manual input → AI |
| Waktu | Paste Google Calendar → manual input → AI |
| Lokasi | AI → default "Online" |
| Attendees | Transcript (regex speaker) + Calendar paste (email) → manual |
| Issues & Actions | AI dari transcript (translate jika Inggris) |
| Chairperson, Prepared By | AI dari transcript |
| Next Meeting | AI dari transcript |


---

## UPDATE SESI 8 MEI 2026

### Fitur Baru
- **Lampiran gambar**: upload jpg/png/gif/webp (banyak), preview thumbnail + caption, ditempatkan di halaman lampiran Word
- **Auto-compress**: jika docx >15MB, gambar dikompres bertahap (quality 70→50→35→20, max-width 1200→800→600px) — pola dari BRDSIT-TOOLS menggunakan Pillow
- **Rate limit notification**: jika kuota Groq habis, tampilkan pesan + info waktu tunggu dari provider
- **Clear all items**: tombol hapus semua baris Issues & Actions (dengan konfirmasi)
- **Generate ulang**: tombol kembali ke Step 1 dari section Issues & Actions
- **Header Word**: tinggi row 1.44cm exact, logo menyesuaikan tinggi, semua row tabel lain min 0.48cm
- **CLEAN_FOR_SHARE.bat**: bersihkan venv/cache/uploads/log/output untuk share, `.env` tetap disertakan

### Dependencies Baru
- `Pillow>=10.0.0` — untuk compress gambar

### File Baru
- `CLEAN_FOR_SHARE.bat` — script bersih-bersih sebelum share
- `backend/uploads/` — folder penyimpanan gambar yang diupload

### API Endpoints Baru
| Method | Endpoint | Fungsi |
|--------|----------|--------|
| POST | /api/upload-image | Upload gambar, return {id, name} |
| DELETE | /api/delete-image/<uid> | Hapus gambar |
| GET | /api/image/<uid> | Serve gambar untuk preview |

### Catatan Masalah Baru
- **Syntax error unmatched '}'**: terjadi saat strReplace menyisipkan kode di tengah dict Python
  - Fix: pastikan dict STATUS_COLORS utuh setelah insert `_apply_min_row_height`


---

## UPDATE SESI 9 MEI 2026

### Perubahan backend/app.py

1. **Fix dead code** — blok `extract_attendees` duplikat yang tidak terjangkau setelah `return` di `api_parse_meeting_info` dihapus

2. **Perluas blacklist `extract_attendees`** — tambah kata umum agar tidak ikut jadi attendees:
   `recording, solution, proposed, summary, agenda, topic, discussion, meeting, update, follow, review, next, view, item, issue, status, deadline, attachment, information, additional, prepared, chairperson, circulation, apologies, location, date, time`

3. **Hapus blok ketiga `extract_from_calendar`** — regex plain "Name:" yang menyebabkan "View Recording", "Proposed Solution" dll ikut terdeteksi sebagai attendees. Sekarang hanya ambil dari: (1) `Name <email>` dan (2) email local part → title case

4. **Fix logika tahun di `parse_meeting_info`** — sebelumnya pakai "jika tanggal sudah lewat → tahun depan" yang salah. Sekarang: jika bulan di kalender > bulan berjalan → pakai tahun sebelumnya (karena MOM = meeting yang sudah terjadi)

5. **Next Meeting field — replace TBA/kosong dengan "-"** di `build_docx` via helper `_nm(val)`

6. **`serve_app` tambah `Cache-Control: no-store`** — browser tidak cache `index.html` lagi

### Perubahan frontend/index.html

1. **Deadline kolom Issues & Actions**:
   - Hapus text input TBA, sisakan date picker saja
   - Default value = tanggal meeting (`eMeetingDatePicker`)
   - `collectData` disesuaikan (hanya 1 input, index td bergeser karena ada kolom `#`)

2. **Nomor urut Issues & Actions**:
   - Kolom `#` di kiri tabel, auto-fill nomor urut
   - `renumberItems()` dipanggil saat tambah/hapus baris
   - Saat export: issue text jadi `"1. Lorem ipsum"` (nomor digabung)
   - Index `tds` di `collectData` bergeser: deadline=tds[2], status=tds[3], pic=tds[4]

3. **Next Meeting — field tanggal jadi date picker**:
   - `eNextDate` → `type="date"` + `onchange="formatNextDate()"`
   - `eNextDateFormatted` hidden input menyimpan format Indonesia
   - `formatNextDate()` fungsi baru mirip `formatTanggal()`
   - `_nm()` helper: kosong/TBA → "-"
   - `populateEditForm` parse tanggal next meeting ke ISO untuk date picker
   - `collectData` ambil dari `eNextDateFormatted`, fallback hitung dari date picker

4. **Prepared By mandatory**:
   - Default value "TBA" (dari HTML `value="TBA"` dan `populateEditForm`)
   - Label ada tanda `*` merah
   - `exportDocx` validasi: jika kosong/TBA → highlight border merah + error alert, block export

5. **Chairperson sync dengan Prepared By**:
   - `eChairperson` punya `oninput="this.dataset.edited='1'"` untuk track manual edit
   - `ePreparedBy` punya `oninput` yang update chairperson jika belum diedit manual
   - `populateEditForm`: reset `dataset.edited`, jika chairperson kosong → ikut preparedBy

6. **Tombol Generate disabled saat health check**:
   - `window.onload` disable `btnGenerate` dulu
   - `checkHealth` enable tombol hanya jika `groq_configured: true`
   - Jika API key tidak ada / backend mati → tombol tetap disabled

7. **Fix bug duplikat `const alertEl`** di `exportDocx` yang menyebabkan JS error saat load

8. **`autoResize` pakai `setTimeout`** — agar browser sempat render sebelum hitung `scrollHeight`

9. **Meeting Location default "Online"** — `populateEditForm` set `d.meeting_location || 'Online'`

### Perubahan start.bat

- Setelah backend siap, terminal menunggu input user
- Ketik `restart` + Enter → kill backend, start ulang, tunggu health check, buka tab browser baru
- Enter kosong → shutdown backend
- `start.bat` manual tetap bisa dijalankan dari luar seperti biasa

### Catatan Masalah Sesi Ini

- **Browser cache agresif** — meski `no-store` sudah ditambah, tab lama tetap cache. Solusi: buka tab baru dengan URL langsung
- **JS error duplikat `const alertEl`** — menyebabkan `window.onload` tidak jalan, status stuck "⏳ Mengecek koneksi ke AI..."
- **`autoResize` tidak akurat** — dipanggil sebelum browser render, fix dengan `setTimeout(..., 0)`



---

## UPDATE SESI 10 MEI 2026

### Perubahan frontend/index.html

1. **Field Lokasi default "Online"** (fix menyeluruh):
   - HTML `eMeetingLocation`: `value="TBA"` → `value="Online"`, placeholder diubah ke `"Online"`
   - `populateEditForm`: fallback diperluas — selain kosong/null, nilai `"TBA"` (case-insensitive) juga diganti `"Online"`
   - Prioritas lokasi: AI result (jika bukan TBA/kosong) → `"Online"`

2. **Field Prepared By — UX placeholder**:
   - HTML: hapus `value="TBA"`, ganti ke `placeholder="TBA"`
   - `populateEditForm`: jika AI return `"TBA"` atau kosong → set value `""` (placeholder tampil)
   - Saat user mulai ketik → placeholder hilang; saat dihapus → placeholder `TBA` muncul kembali
   - Validasi export tetap berlaku: kosong = dianggap TBA → block export + highlight merah
