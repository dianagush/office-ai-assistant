# Memasang add-in secara permanen lewat katalog tepercaya ("Shared Folder") di Office, supaya tombol
# AI Assistant ada walaupun Word/Excel/PowerPoint dibuka dengan cara biasa.
#
# Skrip ini mengubah pengaturan keamanan Office (Trust Center) dan membuat satu folder berbagi
# HANYA-BACA berisi manifest add-in, jadi sengaja dijalankan sendiri oleh pemilik komputer:
#   Pasang Permanen.cmd            -> memasang
#   Pasang Permanen.cmd -Periksa   -> hanya menampilkan apa yang akan dilakukan, tanpa mengubah apa pun
#   Pasang Permanen.cmd -Lepas     -> membatalkan semuanya
param([switch]$Periksa, [switch]$Lepas)
$ErrorActionPreference = "Stop"

$root = Split-Path $PSScriptRoot -Parent
$manifest = Join-Path $root "manifest.prod.xml"
$folder = Join-Path $env:LOCALAPPDATA "Office AI Assistant\Katalog"
$shareName = "OfficeAIAddin"
$url = "\\$env:COMPUTERNAME\$shareName"
$guid = "{7c1f3a52-6d9e-4b08-9a3c-2e5f8d1b4a67}"
$key = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$guid"
$me = [Security.Principal.WindowsIdentity]::GetCurrent().Name

# Hanya pembuatan/penghapusan folder berbagi yang butuh administrator; Windows akan meminta persetujuan (UAC).
function Elevated($command) {
  Start-Process powershell -Verb RunAs -Wait -ArgumentList "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", $command
}

function Visible { Test-Path (Join-Path $url "manifest.prod.xml") }

try {
  if ($Lepas) {
    Write-Host "Yang akan dilepas:"
    Write-Host "  - katalog tepercaya di Office ($key)"
    Write-Host "  - folder berbagi $url"
    Write-Host "  - folder $folder"
    if ($Periksa) { Write-Host "`n(-Periksa: tidak ada yang diubah)"; return }
    if (Test-Path $key) { Remove-Item $key -Recurse }
    if (Visible) { Elevated "Remove-SmbShare -Name '$shareName' -Force" }
    if (Test-Path $folder) {
      Remove-Item (Join-Path $folder "manifest.prod.xml") -Force -ErrorAction SilentlyContinue
      Remove-Item $folder -ErrorAction SilentlyContinue
    }
    Write-Host "`nSelesai. Buka ulang Word/Excel/PowerPoint; tombol AI Assistant hilang dari tab Home."
    return
  }

  if (-not (Test-Path $manifest)) { throw "manifest.prod.xml tidak ditemukan di $root. Buat dulu: npm run manifest -- https://NAMA.github.io/REPO" }

  Write-Host "Yang akan dilakukan:"
  Write-Host "  1. Menyalin manifest.prod.xml ke $folder"
  Write-Host "  2. Membuat folder berbagi HANYA-BACA $url (hanya untuk akun $me; Windows meminta persetujuan administrator)"
  Write-Host "  3. Mendaftarkannya sebagai katalog tepercaya di Office (Trust Center) untuk akun ini"
  if ($Periksa) { Write-Host "`n(-Periksa: tidak ada yang diubah)"; return }

  New-Item -ItemType Directory -Force $folder | Out-Null
  Copy-Item $manifest (Join-Path $folder "manifest.prod.xml") -Force

  if (-not (Visible)) {
    try { Elevated "New-SmbShare -Name '$shareName' -Path '$folder' -ReadAccess '$me' | Out-Null" }
    catch { throw "Persetujuan administrator dibatalkan, jadi folder berbagi tidak dibuat. Tidak ada pengaturan Office yang diubah." }
  }
  if (-not (Visible)) {
    throw "Folder berbagi sudah dibuat, tetapi belum bisa dibuka dari komputer ini ($url). Tidak ada pengaturan Office yang diubah."
  }

  New-Item $key -Force | Out-Null
  New-ItemProperty $key -Name Id -Value $guid -PropertyType String -Force | Out-Null
  New-ItemProperty $key -Name Url -Value $url -PropertyType String -Force | Out-Null
  New-ItemProperty $key -Name Flags -Value 3 -PropertyType DWord -Force | Out-Null

  Write-Host "`nSelesai. Satu langkah terakhir, cukup sekali per aplikasi:"
  Write-Host "  1. Tutup Word/Excel/PowerPoint sampai benar-benar keluar, lalu buka seperti biasa."
  Write-Host "  2. Insert > Add-ins > My Add-ins > tab SHARED FOLDER > pilih AI Assistant > Add."
  Write-Host "Sesudah itu tombol AI Assistant ada di tab Home setiap kali aplikasinya dibuka."
  Write-Host "`nKalau tab SHARED FOLDER tidak ada: File > Options > Trust Center > Trust Center Settings >"
  Write-Host "Trusted Add-in Catalogs. Isi Catalog Url dengan  $url , klik Add catalog, centang Show in Menu,"
  Write-Host "OK, lalu buka ulang aplikasinya."
}
catch {
  Write-Host "`nGAGAL: $($_.Exception.Message)" -ForegroundColor Red
  exit 1
}
