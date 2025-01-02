### Konten README:

````markdown
# Automasi Excel - Untuk Dini

Program ini dirancang untuk bantuin Dini lebih santai dan janji pelq cyum. Dengan program ini, kamu cuma perlu jalanin beberapa langkah sederhana untuk menghasilkan laporan otomatis. 😊

---

## Cara Menggunakan Program

1. **Jalankan program pertama untuk memproses data CSV:**
   ```bash
   python fillNew.py
   ```
````

Setelah program selesai, akan muncul pesan seperti ini:

`Semua file BA Siap yaaa Dini Sayaang!!!`
`Sekarang kamu jalanin 👉 python tambahinFooter.py`

2. **Tambahkan footer pada file Excel:**
   Jalankan perintah berikut setelah langkah pertama selesai:

   ```bash
   python tambahinFooter.py
   ```

3. \*\*Laporan Selesai`
   Kalau semua laporan udah selesai kamu bisa ntar muncul pesan

   `Semua footer udah ada di excell.`
   `Sekarang kamu udah bisa leha-lehaa Dini sayaaaaang.`

   Folder Laporan bisa kamu liat di folder "Output_With_Footer"

---

## Prasyarat (Prerequisites)

### 1. Membuat Environment Virtual

Agar program berjalan dengan baik, pastikan kamu menggunakan virtual environment Python. Berikut langkah-langkahnya:

1. **Buat environment virtual:**

   ```bash
   python -m venv env
   ```

2. **Aktifkan environment virtual:**

   - **Windows (PowerShell):**

     ```bash
     .\\env\\Scripts\\Activate.ps1
     ```

   - **Windows (Command Prompt):**

     ```bash
     env\\Scripts\\activate
     ```

   - **Linux/MacOS:**
     ```bash
     source env/bin/activate
     ```

3. **Install semua dependencies:**
   Jalankan perintah berikut untuk menginstal semua library yang diperlukan:
   ```bash
   pip install -r requirements.txt
   ```

---

## Struktur Folder

Pastikan struktur foldernya sesuai seperti ini:

        ```
        project-folder/
        │
        ├── env/  # Environment virtual Python
        │   ├── Perwakilan JP Des/  # Folder dengan file CSV
        │   ├── BA_Perwakilan_Template.xlsx
        │   ├── Footer_Template.xlsx
        │   ├── gambarHeader.png
        │   └── requirements.txt
        │
        ├── Output/  # Hasil akhir laporan Excel
        │   ├── fillNew.py
        │   ├── tambahinFooter.py
        │   └── Readme.md
        ```

Semoga membantu dan selalu sukses ya Dini......
Kalau ada masalah, jangan ragu untuk kasih tahu aku. Ane bantuin sebisanya

# automation-BA-Perwakilan-BLT-POS-Mbo-2024
