# poinmeet — AI Meeting Minutes Generator

Generate Minutes of Meeting (MOM) dari transcript rapat secara otomatis, lalu export ke file Word (.docx) format template DBC.

---

## Cara Install & Jalankan

### Install (sekali saja)
Klik dua kali `install.bat`
- Otomatis buat virtual environment Python & install semua dependencies
- Langsung buka aplikasi setelah selesai

### Jalankan (setiap mau pakai)
Klik dua kali `start.bat`

### Bersihkan sebelum share
Klik dua kali `CLEAN_FOR_SHARE.bat`
- Hapus venv, cache, uploads, log, file output
- API key di `backend/.env` **tetap disertakan** agar penerima langsung bisa pakai

---

## Cara Pakai

### Step 1 — Input

#### Info Meeting (dari Google Calendar)
1. Buka Google Calendar → klik event meeting
2. Copy nama event + baris tanggal/waktu:
   ```
   [Book] Preview Internal SIT BRD: 26-022 Mengoptimalkan kinerja tim
   Thursday, May 7⋅3:30 – 4:30pm
   ```
3. Paste ke kolom **"Paste info meeting dari Google Calendar"**
4. Nama meeting, tanggal (Indonesia), dan jam (24h) otomatis terisi

#### Transcript (dari Fathom)
1. Buka Fathom → buka hasil recording meeting
2. Klik **"Copy Summary"**
3. Paste ke kolom **Transcript**

#### Attendees (dari Google Calendar)
1. Buka Google Calendar → klik event → scroll ke bagian **Guests**
2. Copy semua nama/email peserta:
   ```
   Budi Santoso <budi@company.com>
   Sari Dewi <sari@company.com>
   ```
3. Paste ke kolom **"Paste deskripsi undangan Google Calendar"**
4. Nama otomatis terdeteksi — bisa diedit di Step 2

> Jika transcript dan Calendar dikosongkan, attendees bisa diisi manual di Step 2.

---

### Step 2 — Edit MOM

- **Info Meeting**: nama, tanggal (date picker → format Indonesia), waktu (HH:MM – HH:MM), lokasi (default: Online), chairperson, prepared by
- **Attendees**: tag list, tambah/hapus nama
- **Issues & Actions**: tabel editable inline
  - Deadline: date picker + field teks untuk TBA/Note
  - Status: dropdown dengan warna otomatis di Word
  - PIC: text input
  - Tombol: **+ Tambah Item** | **🗑 Hapus Semua** | **↩ Generate Ulang**
- **Next Meeting**: tanggal, waktu, lokasi, chairperson, agenda
- **Lampiran**: upload gambar (jpg/png/gif/webp, bisa banyak) + caption per gambar

---

### Step 3 — Export

Klik **Export ke Word (.docx)**

- Nama file: `yyyymmdd_nama-meeting.docx`
- Jika total file >15MB → gambar otomatis dikompres
- Tab step di atas bisa diklik langsung untuk navigasi antar step

---

## Format File Word (Template DBC)

| Section | Keterangan |
|---------|-----------|
| Header | Logo DBC + "Minute of Meeting", tinggi 1.44cm, tiap halaman |
| Footer | Page X of Y, Arial 8pt, kanan bawah, tiap halaman |
| Meeting Information | Header #4C4C4C, label #E6E6E6 |
| Topic & Attachment | Label #E6E6E6 |
| Issues and Actions | Header bold italic #E6E6E6, header repeat tiap halaman |
| Tentative Agenda | Header hitam #000000, label #D9D9D9 |
| Lampiran | Gambar full-width + caption italic biru #1F497D |

**Warna status:**
- Open / On Progress / Progress → Kuning `#FFFF00`
- Done → Hijau `#00FF00`
- Monitoring / Note / Noted → Cyan `#00FFFF`

---

## Troubleshooting

| Masalah | Solusi |
|---------|--------|
| Status AI merah saat buka app | Jalankan `start.bat` terlebih dahulu |
| Browser tidak terbuka otomatis | Buka manual: http://localhost:5000/app |
| Error "Kuota AI habis" | Tunggu sesuai info waktu yang ditampilkan, lalu coba lagi |
| install.bat error Python not found | Install Python 3.10+ dari https://python.org |
| Generate gagal / JSON error | Cek koneksi internet |

---

## Stack Teknologi
- Backend: Python Flask + python-docx + Pillow + Groq AI (llama-3.3-70b-versatile)
- Frontend: HTML/JS single file, buka di browser lokal
- Tidak butuh Docker, Node.js, atau database
