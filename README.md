# Aplikasi Pencatatan Anggaran Sederhana Berbasis Excel

Pernahkah Anda merasa kesulitan mengelola daftar anggaran atau inventaris barang? Proses mencatat, mengubah, dan menghapus data secara manual di Excel seringkali memakan waktu dan rentan terhadap kesalahan. Bagi banyak pelaku usaha, mahasiswa, atau siapa pun yang bukan seorang programmer, membuat sebuah sistem otomatis terasa seperti hal yang mustahil.

Proyek ini hadir sebagai solusi. Kami telah membangun sebuah sistem **CRUD (Create, Read, Update, Delete)** yang fungsional, sepenuhnya di dalam lingkungan Microsoft Excel yang sudah akrab bagi Anda. Lupakan kerumitan coding; aplikasi ini memungkinkan Anda untuk:

- **Menambah data** melalui formulir yang mudah digunakan.
- **Mencari dan melihat** data yang sudah ada dengan cepat.
- **Memperbarui data** jika ada perubahan.
- **Menghapus data** yang sudah tidak relevan.
- **Mencetak laporan** anggaran ke dalam format PDF profesional dengan sekali klik.

Tujuannya sederhana: memberikan kemudahan bagi siapa pun untuk memiliki sistem pencatatan data yang semi-otomatis, rapi, dan efisien tanpa perlu menulis satu baris kode pun.

---

## Panduan Penggunaan

Berikut adalah langkah-langkah untuk menggunakan aplikasi ini, mulai dari membuka file hingga mencetak laporan.

### 1. Membuka File dan Aktivasi Macro
Aplikasi ini dibangun menggunakan VBA (Visual Basic for Applications), yaitu bahasa program di dalam Excel. Agar semua tombol dan fungsi otomatis berjalan, Anda wajib mengaktifkan macro.

1. Buka file `.xlsm` yang telah Anda unduh.
2. Di bagian atas, akan muncul bar kuning dengan notifikasi keamanan **"SECURITY WARNING Macros have been disabled."**
3. Klik tombol **"Enable Content"**.
   ![Enable Content](https://i.imgur.com/e3J2p1w.png) 4. Setelah itu, semua fitur aplikasi siap untuk digunakan.

### 2. Menambah Data Barang Baru
Gunakan formulir di sheet **`INPUT BARANG`** untuk memasukkan data.

1. Isi kolom **Nama Barang**, **Deskripsi**, **Kuantitas**, dan **Harga Satuan**.
2. Klik tombol **ADD**.
3. Data Anda akan otomatis tersimpan di sheet **`RAB`** dengan ID unik yang baru.

### 3. Mencari dan Mengubah Data
Jika Anda ingin memperbaiki data yang sudah ada:

1. Ketik **ID Barang** yang ingin diubah pada kolom yang tersedia (contoh: `RAB002`).
2. Klik tombol **SEARCH**. Data yang sesuai akan otomatis muncul di formulir.
3. Ubah informasi yang diperlukan (misalnya, ubah kuantitas atau harganya).
4. Setelah selesai, klik tombol **UPDATE**. Perubahan Anda akan langsung tersimpan.

### 4. Menghapus Data
Untuk menghapus data yang tidak lagi diperlukan:

1. Cari data yang ingin dihapus menggunakan langkah-langkah di atas (ketik ID, lalu klik **SEARCH**).
2. Setelah data muncul di formulir, klik tombol **DELETE**.
3. Sebuah kotak konfirmasi akan muncul. Klik **Yes** untuk menghapus data secara permanen.

### 5. Melihat Daftar Barang (Database)
Semua data yang Anda masukkan tersimpan rapi dalam bentuk tabel di sheet bernama **`RAB`**. Anda bisa membuka sheet ini kapan saja untuk melihat keseluruhan daftar barang dan total anggaran secara langsung.

### 6. Mencetak Laporan ke PDF
Aplikasi ini dapat membuat laporan PDF profesional secara otomatis.

1. **Menyesuaikan Template (Opsional)**: Jika Anda ingin mengubah tampilan laporan (misalnya mengubah judul atau logo), buka sheet **`FORMAT LAPORAN`**. Anda bisa mengedit tata letak, teks, atau desain di sheet ini sesuai keinginan.
2. **Mencetak Laporan**: Kembali ke sheet **`INPUT BARANG`**, lalu klik tombol **Cetak ke PDF**.
3. Sebuah jendela "Save As" akan muncul. Pilih lokasi penyimpanan, beri nama file Anda, dan klik **Save**.
4. File PDF akan dibuat dan otomatis terbuka.

> **Catatan untuk Pengguna Mahir:** Jika Anda ingin melihat atau memodifikasi kode di balik aplikasi ini, Anda dapat menekan `Alt + F11` untuk membuka VBA Editor.

---

## Credits & Kloning Repositori

### Credits
Proyek ini dibuat dan dikembangkan oleh [Nama Anda].

### Kloning Repositori
Untuk mendapatkan salinan proyek ini di komputer lokal Anda, ikuti langkah-langkah berikut:

1. **Clone repositori ini**
   ```bash
   git clone [https://github.com/](https://github.com/)[Username-Anda]/[Nama-Repositori-Anda].git
