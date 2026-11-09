# QA / Release note

- Source berasal dari Moodle Question Studio MVP yang dipersiapkan untuk Flask di Vercel.
- Unit test dan HTTP smoke test harus dijalankan sebelum menyatakan production-ready.
- Manual test: upload DOCX dan XML, periksa jumlah soal, koreksi kunci jawaban, ekspor Moodle XML, impor kembali di Moodle staging.
- Pastikan gambar embedded XML, tabel Cloze, dan skor pilihan majemuk sesuai harapan.
- Batas serverless untuk request/response 4 MB harus dipantau pada dokumen bergambar.
- Aplikasi publik belum memiliki autentikasi. Jangan masukkan materi rahasia.
