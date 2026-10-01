# 📋 Generator SPO UNIMOF

> **Generator Standar Prosedur Operasional berbasis web yang modern, responsif, dan mudah dipelihara.**

Generator SPO UNIMOF merupakan aplikasi berbasis web untuk membantu membuat dokumen **Standar Prosedur Operasional (SPO)** dalam format **Microsoft Word `.docx`**.

Project dirancang dengan struktur terpisah antara **HTML, CSS, JavaScript, asset, dan template Word**, sehingga lebih mudah dikembangkan, diperbaiki, dan dipelihara.

---

## ✨ Fitur Utama

| Fitur | Keterangan |
|---|---|
| 📝 Form SPO | Membuat dokumen SPO melalui form |
| 🏢 Unit | Input unit/instalasi terkait |
| 📄 Judul SPO | Input judul Standar Prosedur Operasional |
| 📖 Pengertian | Menambahkan pengertian prosedur |
| 🎯 Tujuan | Mendukung penomoran otomatis |
| 📜 Kebijakan | Mendukung penomoran otomatis |
| 🔢 Prosedur | Input langkah-langkah prosedur |
| 🔀 Bagan Alir | Generator bagan alir otomatis |
| 🖼️ Upload Bagan | Menggunakan gambar bagan alir |
| 📋 Bagan Teks | Menambahkan bagan alir dalam bentuk teks |
| 👁️ Preview | Melihat preview bagan sebelum dibuat |
| 💾 Draft | Menyimpan draft menggunakan `localStorage` |
| 🔄 Reset | Menghapus seluruh isian form |
| 📥 DOCX | Menghasilkan dokumen Microsoft Word |
| 📱 Responsive | Mendukung desktop, tablet, dan mobile |
| 🎨 UI Modern | Tema visual menyesuaikan logo UNIMOF |

---

## 🖥️ Tampilan

Aplikasi menggunakan desain yang:

- 🎨 Modern
- 📱 Responsive
- 🖱️ User friendly
- ⚡ Ringan
- 🧩 Modular
- 🛠️ Mudah dikembangkan

Tema visual menggunakan kombinasi warna yang disesuaikan dengan identitas visual **UNIMOF**.

---

## 📁 Struktur Project

```text
spo-generator-project/
│
├── 📄 index.html
├── 📄 README.md
│
├── 📂 css/
│   └── 🎨 style.css
│
├── 📂 js/
│   └── ⚙️ app.js
│
└── 📂 assets/
    ├── 🖼️ logo-unimof.png
    └── 📄 template-spo.docx
