# Moodle Question Studio

Aplikasi Flask untuk membaca berkas **DOCX, Moodle XML, dan CSV**, menyunting soal hasil konversi dan mengekspor ke **Moodle XML, GIFT, Aiken, atau CSV**.

## Deploy di Vercel

1. Import repository `eltonmalumbot/convertsoal` dari dashboard Vercel.
2. Root Directory: `./`. Framework: pilih **Other** bila Flask tidak terdeteksi otomatis. Jangan menambahkan build command/output directory.
3. Vercel menggunakan `app.py` sebagai Python/Flask entrypoint dan `requirements.txt` untuk dependensi.
4. Deploy lalu kunjungi `/api/health` (harus mengembalikan `status: ok`), dan tes upload DOCX/XML dari halaman utama.

Python versi 3.12 disetel oleh `.python-version`; konfigurasi fungsi di `vercel.json`.

## Menjalankan lokal

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

Aplikasi akan tampil di <http://127.0.0.1:8080>.

## Batasan dan keamanan

- Maksimal upload **3,8 MB** per file dan respons API kira-kira **4 MB**. Untuk dokumen bergambar besar, pecah file agar tetap berada di bawah batas Vercel Functions.
- Dokumen diproses di memory per-request, tidak sengaja disimpan sebagai arsip permanen. Namun file tetap dikirim ke server aplikasi.
- Tidak ada login/authentication; **jangan publikasikan soal ujian rahasia** di domain terbuka tanpa menambahkan proteksi akses.
- Deteksi struktur Word berbasis heuristik: opsi, kunci, tabel, dan gambar wajib diperiksa manual sebelum digunakan untuk asesmen.
- Soal tanpa kunci jawaban ditandai dan akan ditolak saat ekspor ke format ujian. CSV dapat dipakai menyimpan draf.
- CSV adalah format internal Moodle Question Studio, **bukan CSV impor bawaan Moodle**.
- Berkas `.doc` lama, PDF, tipe soal plugin eksternal, dan transkripsi gambar tidak didukung pada versi ini.
- Mode Cloze pada XML dipertahankan sebagai sintaks dalam field HTML; tidak mengedit Cloze secara visual.

## Endpoint

| Metode | Path | Fungsi |
| --- | --- | --- |
| GET | `/` | Halaman konverter |
| GET | `/api/health` | Pemeriksaan layanan |
| POST | `/api/parse` | Multipart form-data `file`: DOCX/XML/CSV |
| POST | `/api/export` | JSON `{questions,format,category}` |

## QA

Periksa file Word contoh dengan penomoran otomatis, kunci jawaban majemuk, tabel Benar/Salah, dan gambar. Sebelum menggunakan pada produksi, lakukan pengujian import hasil XML langsung di **Moodle 4.4.3 staging** dan pastikan soal/gambar/kunci/skor sesuai.

Aplikasi independen, tidak berafiliasi resmi dengan Moodle.
