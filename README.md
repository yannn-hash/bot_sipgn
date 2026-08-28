# SIPGN Hybrid Automation v4

Model:
1. Bot membuka Google Chrome asli dengan profile khusus.
2. User melakukan Cloudflare/SSO verification dan login secara manual.
3. Bot tidak menyelesaikan, melewati, atau mengakali Cloudflare.
4. Setelah login, Playwright terhubung ke Chrome yang sama melalui CDP.
5. Bot membuka form distribusi dan mengisi pekerjaan repetitif.
6. Prototype berhenti sebelum Submit.

## Instalasi & Panduan Pemindahan ke Komputer Baru

Jika Anda ingin menjalankan bot ini di komputer Windows yang baru, ikuti panduan berikut.

### 1. Persiapan File
Saat memindahkan/meng-copy folder bot ini dari satu komputer ke komputer lain, **JANGAN** ikut menyalin folder berikut untuk menghemat ruang dan menghindari error (konfigurasi tiap komputer berbeda):
- `.venv`
- `__pycache__`
- `chrome-profile` (jika Anda ingin login ulang agar aman)
- `web_app/frontend/node_modules`

### 2. Prasyarat Sistem
Pastikan komputer baru sudah terinstall aplikasi berikut:
- **Google Chrome** (Browser utama)
- **Python** (Centang *"Add Python to PATH"* saat instalasi)
- **Node.js** (Gunakan versi LTS)

### 3. Setup Backend (Python)
Buka terminal (CMD/PowerShell) di dalam folder utama project ini, lalu jalankan:

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python -m playwright install
```

### 4. Setup Frontend (Web UI)
Buka terminal baru atau gunakan terminal yang sama, masuk ke dalam folder frontend, lalu jalankan npm install:

```powershell
cd web_app\frontend
npm install
```

> **Catatan**: Jangan lupa untuk meng-copy file `.env.example` menjadi `.env` jika belum ada sebelum bot dijalankan.

## Login manual

```powershell
python bot.py --login
```

Browser Google Chrome akan terbuka. Selesaikan Cloudflare dan SSO secara manual.
Jangan menutup Chrome selama proses login.

## Menjalankan Bot (Web UI) - Cara Termudah

Setelah Anda berhasil melakukan login manual, Anda bisa langsung menjalankan bot beserta tampilan web (antarmuka) nya.

1. Buka folder utama bot ini.
2. Klik ganda (*double-click*) pada file **`Mulai_SIPGN.bat`**.
3. File ini akan otomatis menyalakan server backend (API) dan frontend (Web UI).
4. Browser akan terbuka secara otomatis yang menampilkan halaman kontrol bot (atau akses secara manual di `http://localhost:5170`).
5. **Penting:** Biarkan dua jendela CMD (hitam) tetap terbuka selama Anda menggunakan aplikasi web.

## Menjalankan Automation (via CLI/Terminal)

Jika Anda ingin menjalankan bot secara langsung dari terminal (tanpa Web UI), jalankan perintah berikut setelah melakukan login:

```powershell
python bot.py --excel template_bot_distribusi_sipgn_v2.xlsx --row 2
```

Untuk seluruh data dalam file excel:

```powershell
python bot.py --excel template_bot_distribusi_sipgn_v2.xlsx
```

## Catatan

- Profile khusus disimpan di `chrome-profile/`.
- Jangan menggunakan profile Chrome utama.
- Jangan membagikan folder `chrome-profile` karena dapat berisi session/cookie.
- Screenshot error/success disimpan di `artifacts/`.
- Bot tidak mencoba bypass Cloudflare.
- Submit otomatis sengaja dinonaktifkan pada prototype.
