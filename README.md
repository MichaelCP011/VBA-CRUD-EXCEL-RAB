<img width="1920" height="1020" alt="image" src="https://github.com/user-attachments/assets/716d6ef9-db0b-4202-a08a-2c2dbbe3d91d" /># Aplikasi Pencatatan Anggaran Sederhana Berbasis Excel

Pernahkah Anda merasa kesulitan mengelola daftar anggaran atau inventaris barang? Proses mencatat, mengubah, dan menghapus data secara manual di Excel seringkali memakan waktu dan rentan terhadap kesalahan. Bagi banyak pelaku usaha, mahasiswa, atau siapa pun yang bukan seorang programmer, membuat sebuah sistem otomatis terasa seperti hal yang mustahil.

Proyek ini hadir sebagai solusi. Kami telah membangun sebuah sistem **CRUD (Create, Read, Update, Delete)** yang fungsional, sepenuhnya di dalam lingkungan Microsoft Excel yang sudah akrab bagi Anda. Lupakan kerumitan coding; aplikasi ini memungkinkan Anda untuk:

- **Menambah data** melalui formulir yang mudah digunakan.
- **Mencari dan melihat** data yang sudah ada dengan cepat.
- **Memperbarui data** jika ada perubahan.
- **Menghapus data** yang sudah tidak relevan.
- **Mencetak laporan** anggaran ke dalam format PDF profesional dengan sekali klik.

Tujuannya sederhana: memberikan kemudahan bagi siapa pun untuk memiliki sistem pencatatan data yang semi-otomatis, rapi, dan efisien tanpa perlu menulis satu baris kode pun.

---

# Panduan Penggunaan

Berikut adalah langkah-langkah untuk menggunakan aplikasi ini, mulai dari membuka file hingga mencetak laporan.

## 1. Membuka File dan Aktivasi Macro
Aplikasi ini dibangun menggunakan VBA (Visual Basic for Applications), yaitu bahasa program di dalam Excel. Agar semua tombol dan fungsi otomatis berjalan, Anda wajib mengaktifkan macro.

1. Buka file `.xlsm` yang telah Anda unduh.
2. Di bagian atas, akan muncul bar kuning dengan notifikasi keamanan **"SECURITY WARNING Macros have been disabled."**
3. Klik tombol **"Enable Content"**.
4. Setelah itu, semua fitur aplikasi siap untuk digunakan.

## 2. Menambah Data Barang Baru
Gunakan formulir di sheet **`INPUT BARANG`** untuk memasukkan data.

1. Isi kolom **Nama Barang**, **Deskripsi**, **Kuantitas**, dan **Harga Satuan**.
2. Klik tombol **ADD**.
3. Data Anda akan otomatis tersimpan di sheet **`RAB`** dengan ID unik yang baru.
<img width="1920" height="1020" alt="image" src="https://github.com/user-attachments/assets/61f2d724-f1fa-48c7-9860-1159a7f83db5" />

## 3. Mencari dan Mengubah Data
Jika Anda ingin memperbaiki data yang sudah ada:

1. Ketik **ID Barang** yang ingin diubah pada kolom yang tersedia (contoh: `RAB002`).
2. Klik tombol **SEARCH**. Data yang sesuai akan otomatis muncul di formulir.
3. Ubah informasi yang diperlukan (misalnya, ubah kuantitas atau harganya).
4. Setelah selesai, klik tombol **UPDATE**. Perubahan Anda akan langsung tersimpan.
<img width="1920" height="1020" alt="image" src="https://github.com/user-attachments/assets/fc7075ca-4004-4b32-88e6-664a2a00c5e9" />
<img width="1920" height="1020" alt="image" src="https://github.com/user-attachments/assets/0bf4f1b3-7bf7-4bca-8b86-9c8fe3657bab" />

## 4. Menghapus Data
Untuk menghapus data yang tidak lagi diperlukan:

1. Cari data yang ingin dihapus menggunakan langkah-langkah di atas (ketik ID, lalu klik **SEARCH**).
2. Setelah data muncul di formulir, klik tombol **DELETE**.
3. Sebuah kotak konfirmasi akan muncul. Klik **Yes** untuk menghapus data secara permanen.
<img width="1920" height="1020" alt="image" src="https://github.com/user-attachments/assets/fc7075ca-4004-4b32-88e6-664a2a00c5e9" />
<img width="1920" height="1020" alt="image" src="https://github.com/user-attachments/assets/3cd695dc-6020-4e0b-a58f-9c4449c35d6d" />

## 5. Melihat Daftar Barang (Database)
Semua data yang Anda masukkan tersimpan rapi dalam bentuk tabel di sheet bernama **`RAB`**. Anda bisa membuka sheet ini kapan saja untuk melihat keseluruhan daftar barang dan total anggaran secara langsung.
<img width="1920" height="1020" alt="image" src="https://github.com/user-attachments/assets/a1d81fc2-544c-410c-ada5-80d730075400" />

## 6. Mencetak Laporan ke PDF
Aplikasi ini dapat membuat laporan PDF profesional secara otomatis.

1. **Menyesuaikan Template (Opsional)**: Jika Anda ingin mengubah tampilan laporan (misalnya mengubah judul atau logo), buka sheet **`FORMAT LAPORAN`**. Anda bisa mengedit tata letak, teks, atau desain di sheet ini sesuai keinginan.
2. **Mencetak Laporan**: Kembali ke sheet **`INPUT BARANG`**, lalu klik tombol **Cetak ke PDF**.
3. Sebuah jendela "Save As" akan muncul. Pilih lokasi penyimpanan, beri nama file Anda, dan klik **Save**.
4. File PDF akan dibuat dan otomatis terbuka.
<img width="1920" height="1020" alt="image" src="https://github.com/user-attachments/assets/99999d34-51a6-478f-a6d2-a03d251f0fd9" />
<img width="1920" height="1020" alt="image" src="https://github.com/user-attachments/assets/533ba20b-7b21-4a3c-9097-0ab1f751302d" />

> **Catatan untuk Pengguna Mahir:** Jika Anda ingin melihat atau memodifikasi kode di balik aplikasi ini, Anda dapat menekan `Alt + F11` untuk membuka VBA Editor.

---

# Credits & Kloning Repositori

## Credits
Proyek ini dibuat dan dikembangkan oleh Michael Christia Putra.
Follow my IG :>
@mcputra011

## Kloning Repositori
Untuk mendapatkan salinan proyek ini di komputer lokal Anda, ikuti langkah-langkah berikut:

1. **Clone repositori ini**
   ```bash
   git clone https://github.com/MichaelCP011/VBA-CRUD-EXCEL-RAB.git
