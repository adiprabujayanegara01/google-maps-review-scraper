# Google Maps Review Scraper

Python-based web scraper untuk mengambil data ulasan dari Google Maps menggunakan **Selenium** dan **Pandas**.

Proyek ini dapat digunakan untuk mengumpulkan dan mengolah data ulasan Google Maps dari satu atau beberapa lokasi secara otomatis, kemudian menyimpannya dalam format **CSV** dan **Excel (.xlsx)**.

## ✨ Features

* Scraping ulasan Google Maps secara otomatis.
* Mendukung satu atau beberapa URL Google Maps.
* Mengambil informasi reviewer dan ulasan.
* Mengambil rating setiap ulasan.
* Mengambil tanggal ulasan.
* Menghasilkan perkiraan tanggal berdasarkan informasi relatif seperti "2 bulan lalu".
* Mengambil balasan dari pemilik tempat.
* Mengambil informasi Local Guide.
* Mengambil jumlah ulasan reviewer.
* Mengambil jumlah foto dan informasi keberadaan foto.
* Mengambil jumlah suka pada ulasan.
* Mengurutkan ulasan berdasarkan:

  * Relevan
  * Terbaru
  * Rating tertinggi
  * Rating terendah
* Mendukung mode **headless**.
* Dapat memblokir gambar untuk membantu mempercepat proses scraping.
* Menghasilkan file CSV.
* Menghasilkan file Excel.
* Mendukung penggabungan hasil dari beberapa cabang/lokasi.

---

## 📁 Project Structure

```text
google-maps-review-scraper/
│
├── scraping_data_ulasan_google_ulasan.py
│
├── output/
│   ├── C28_Social_Space_Malioboro_20260929_154254.csv
│   └── C28_Social_Space_Malioboro_20260929_154254.xlsx
│
└── README.md
```

### Penjelasan

| File/Folder                             | Keterangan                        |
| --------------------------------------- | --------------------------------- |
| `scraping_data_ulasan_google_ulasan.py` | Script utama untuk scraping       |
| `output/`                               | Folder penyimpanan hasil scraping |
| `.csv`                                  | Hasil scraping dalam format CSV   |
| `.xlsx`                                 | Hasil scraping dalam format Excel |
| `README.md`                             | Dokumentasi proyek                |

---

# ⚙️ Requirements

Pastikan komputer telah memiliki:

* Python 3.9 atau lebih baru
* Google Chrome
* Internet connection

Library Python yang digunakan:

* Selenium
* Pandas
* OpenPyXL

---

# 🚀 Installation

Clone repository:

```bash
git clone https://github.com/USERNAME/google-maps-review-scraper.git
```

Masuk ke folder repository:

```bash
cd google-maps-review-scraper
```

Install dependencies:

```bash
pip install selenium pandas openpyxl
```

### Optional

Jika Selenium Manager tidak dapat menemukan ChromeDriver secara otomatis:

```bash
pip install webdriver-manager
```

---

# ▶️ Usage

## 1. Menggunakan URL langsung

Jalankan:

```bash
python scraping_data_ulasan_google_ulasan.py "URL_GOOGLE_MAPS"
```

Contoh:

```bash
python scraping_data_ulasan_google_ulasan.py "https://www.google.com/maps/place/..."
```

---

## 2. Membatasi jumlah ulasan

Gunakan parameter `--max-reviews`.

```bash
python scraping_data_ulasan_google_ulasan.py "URL_GOOGLE_MAPS" --max-reviews 200
```

Contoh tersebut akan membatasi proses scraping hingga sekitar 200 ulasan.

---

## 3. Mengurutkan ulasan berdasarkan yang terbaru

```bash
python scraping_data_ulasan_google_ulasan.py "URL_GOOGLE_MAPS" --sort newest
```

Pilihan sorting:

```text
relevant
newest
highest
lowest
```

---

## 4. Menggunakan beberapa URL

Beberapa URL dapat diberikan sekaligus:

```bash
python scraping_data_ulasan_google_ulasan.py "URL_1" "URL_2" "URL_3"
```

Script akan memproses setiap lokasi secara berurutan.

Apabila terdapat lebih dari satu lokasi, hasil seluruh lokasi juga dapat digabungkan ke dalam satu dataset.

---

## 5. Menggunakan file URL

Buat file:

```text
urls.txt
```

Isi satu URL Google Maps pada setiap baris:

```text
https://www.google.com/maps/place/LOKASI_1
https://www.google.com/maps/place/LOKASI_2
https://www.google.com/maps/place/LOKASI_3
```

Kemudian jalankan:

```bash
python scraping_data_ulasan_google_ulasan.py --urls-file urls.txt
```

Baris yang diawali dengan `#` akan diabaikan.

---

# 🖥️ Headless Mode

Jika ingin menjalankan browser tanpa menampilkan jendela Chrome:

```bash
python scraping_data_ulasan_google_ulasan.py "URL_GOOGLE_MAPS" --headless
```

Mode ini berguna apabila scraper dijalankan pada server atau environment tanpa GUI.

---

# 🖼️ Block Images

Untuk mengurangi pemuatan gambar selama proses scraping:

```bash
python scraping_data_ulasan_google_ulasan.py "URL_GOOGLE_MAPS" --block-images
```

Parameter ini dapat dikombinasikan dengan `--headless`:

```bash
python scraping_data_ulasan_google_ulasan.py "URL_GOOGLE_MAPS" --headless --block-images
```

---

# 📊 Output

Secara default, hasil scraping disimpan di folder:

```text
output/
```

Format yang dihasilkan:

```text
.csv
.xlsx
```

Contoh:

```text
output/
├── C28_Social_Space_Malioboro_20260929_154254.csv
└── C28_Social_Space_Malioboro_20260929_154254.xlsx
```

Jika scraping dilakukan terhadap beberapa lokasi, script juga dapat menghasilkan:

```text
gabungan_semua_cabang.csv
gabungan_semua_cabang.xlsx
```

---

# 🧾 Data yang Dikumpulkan

Dataset hasil scraping dapat memiliki beberapa field berikut:

| Field                    | Keterangan                         |
| ------------------------ | ---------------------------------- |
| `no`                     | Nomor data                         |
| `nama_tempat`            | Nama lokasi                        |
| `review_id`              | ID ulasan                          |
| `nama_reviewer`          | Nama reviewer                      |
| `local_guide`            | Status Local Guide                 |
| `jumlah_ulasan_reviewer` | Jumlah ulasan reviewer             |
| `rating`                 | Rating yang diberikan              |
| `tanggal`                | Informasi tanggal dari Google Maps |
| `perkiraan_tanggal`      | Perkiraan tanggal dalam format ISO |
| `ulasan`                 | Isi ulasan                         |
| `balasan_pemilik`        | Balasan dari pemilik tempat        |
| `jumlah_foto`            | Jumlah foto                        |
| `foto`                   | Informasi keberadaan foto          |
| `jumlah_suka`            | Jumlah suka                        |
| `waktu_scrape`           | Waktu proses scraping              |

---

# 📌 Example

Contoh data yang dihasilkan:

```text
nama_tempat:
C28 Social Space - Malioboro

rating:
5.0

nama_reviewer:
Melia Ariani

ulasan:
Tempatnya nyaman kerenn banget, staff nya jugaa ramah bangettt...

local_guide:
True

jumlah_foto:
4
```

---

# 🔧 Command Options

| Parameter        | Fungsi                                                  |
| ---------------- | ------------------------------------------------------- |
| `--urls-file`    | Membaca daftar URL dari file `.txt`                     |
| `--max-reviews`  | Membatasi jumlah ulasan                                 |
| `--sort`         | Mengatur urutan ulasan                                  |
| `--stall-limit`  | Menentukan batas percobaan ketika tidak ada ulasan baru |
| `--max-minutes`  | Membatasi waktu proses scrolling                        |
| `--output-dir`   | Menentukan folder output                                |
| `--headless`     | Menjalankan Chrome tanpa GUI                            |
| `--block-images` | Memblokir gambar                                        |
| `--no-xlsx`      | Menyimpan CSV saja                                      |

Contoh penggunaan beberapa parameter sekaligus:

```bash
python scraping_data_ulasan_google_ulasan.py \
  "URL_GOOGLE_MAPS" \
  --max-reviews 500 \
  --sort newest \
  --headless \
  --block-images
```

---

# ⚠️ Notes

Script ini bergantung pada struktur halaman Google Maps. Perubahan pada struktur HTML, selector, mekanisme pemuatan halaman, atau fitur Google Maps dapat menyebabkan scraper perlu diperbarui.

Gunakan scraper secara bertanggung jawab dan pastikan penggunaannya sesuai dengan ketentuan layanan, robots/policies yang berlaku, serta peraturan yang relevan.

Data hasil scraping juga perlu dikelola dengan memperhatikan aspek privasi dan penggunaan data.
