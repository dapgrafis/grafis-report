DASHBOARD TIKET GRAPHIC DESIGN (GitHub Pages)

Isi paket:
  index.html                     halaman dashboard
  data.json                      data awal (akan diperbarui otomatis)
  config.json                    ID folder Google Drive dan bulan mulai
  package.json                   pustaka pembaca Excel
  scripts/sync.js                skrip pembaca folder Google Drive
  .github/workflows/sync.yml     jadwal otomatis (tiap ~10 menit) dan deploy

Langkah singkat:
  1. Buat repository baru di GitHub, lalu unggah SEMUA file ini (termasuk folder .github dan scripts).
  2. Settings > Pages > Build and deployment > Source: pilih "GitHub Actions".
  3. Tab Actions > "Sinkron data Google Drive dan deploy" > Run workflow.
  4. Setelah hijau, buka alamat Pages Anda (Settings > Pages).

Folder Google Drive harus dibagikan sebagai "Siapa saja yang memiliki link" (hanya lihat).
Opsional: tambahkan secret GOOGLE_API_KEY (Settings > Secrets and variables > Actions) agar
pembacaan memakai API resmi Google dan lebih stabil.
