# Bank DKI PDF to Excel

Script Python ini digunakan untuk mengekstrak data mutasi rekening **Bank DKI** dari file PDF e-statement, membersihkan data transaksi, memisahkan debit dan kredit, lalu mengekspornya ke file Excel.

Script ini dibuat khusus berdasarkan format e-statement Bank DKI, di mana:

- **halaman 1** berisi **ringkasan / summary**
- **halaman berikutnya** berisi **detail transaksi**

---

## Fitur

- Membaca file PDF mutasi rekening Bank DKI
- Mengambil **summary** dari halaman pertama
- Mengambil **detail transaksi** dari seluruh halaman PDF
- Menggabungkan tabel dari beberapa halaman
- Menggabungkan baris transaksi yang terpotong ke beberapa baris
- Memisahkan nilai **Debit** dan **Kredit**
- Mengekspor hasil ke file Excel
- Menyesuaikan lebar kolom Excel otomatis

---

## Struktur Output

File Excel hasil export akan memiliki 2 sheet:

### 1. Sheet `Transaksi`

Berisi data transaksi dengan kolom:

- `Tanggal`
- `Jam`
- `Rincian`
- `Debit`
- `Kredit`
- `Saldo`

### 2. Sheet `Summary`

Berisi data ringkasan rekening dengan kolom:

- `Rekening`
- `Saldo_Awal`
- `Transaksi_Masuk`
- `Transaksi_Keluar`
- `Saldo_Akhir`

---

## Struktur Folder

Disarankan menggunakan struktur folder seperti berikut:

```bash
project_folder/
│
├── pdf_file/
│   └── Estatement-Maret-2026-Jakom-1776346789668941.pdf
│
├── excel_file/
│   └── Estatement-Maret-2026-Jakom-1776346789668941.xlsx
│
├── script.py
└── README.md
