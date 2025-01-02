### Konten README:

````markdown
# Automasi Excel - Untuk Dini Sayang 💖

Program ini dirancang untuk membantu Dini Sayang agar lebih santai dan mendapatkan waktu ekstra untuk peluk cium aku. Dengan program ini, kamu hanya perlu menjalankan beberapa langkah sederhana untuk menghasilkan laporan otomatis. 😊

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

Semoga membantu dan selalu sukses ya, Dini Sayang! Kalau ada masalah, jangan ragu untuk kasih tahu aku. Aku siap bantu. 💕
# automation-BA-Perwakilan-BLT-POS-Mbo-2024
