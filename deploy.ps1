<#
  Fenotipik_Indeks_App - TEK KOMUTLA YAYIN

  Bu bir Streamlit projesi; canliya cikis GitHub uzerinden olur
  (Streamlit Community Cloud main dalini izler). Yani "deploy" = commit + push.

  Kullanim:
    .\deploy.ps1 "Ne degisti"
    .\deploy.ps1 -Deneme "..."   hicbir sey yapmaz, ne yapacagini yazar
#>
param(
  [Parameter(Position=0)][string]$Mesaj,
  [switch]$Deneme
)

$ErrorActionPreference = "Stop"
$PROJE = "Fenotipik_Indeks_App"
Set-Location $PSScriptRoot

function Basla($m) { Write-Host "`n=== $m ===" -ForegroundColor Cyan }
function Dur($m)   { Write-Host "HATA: $m" -ForegroundColor Red; exit 1 }

if (-not $Mesaj) {
  $Mesaj = Read-Host "Commit mesaji (ne degisti?)"
  if (-not $Mesaj) { Dur "Mesaj bos birakilamaz." }
}

Write-Host "PROJE : $PROJE"
Write-Host "MESAJ : $Mesaj"

Basla "Sozdizimi denetimi (py_compile)"
$py = if (Test-Path ".\venv\Scripts\python.exe") { ".\venv\Scripts\python.exe" } else { "python" }
$kaynaklar = Get-ChildItem -Filter *.py -Recurse |
             Where-Object { $_.FullName -notmatch '\\venv\\|\\\.venv\\|\\__pycache__\\' } |
             Select-Object -ExpandProperty FullName
if ($kaynaklar) {
  & $py -m py_compile @kaynaklar
  if ($LASTEXITCODE -ne 0) { Dur "Python sozdizimi hatasi - hicbir sey yayinlanmadi." }
  Write-Host "$($kaynaklar.Count) dosya temiz." -ForegroundColor Green
}

Basla "git"
$degisiklik = git status --porcelain
if (-not $degisiklik) {
  Write-Host "Degisiklik yok - commit atlaniyor." -ForegroundColor DarkGray
} elseif ($Deneme) {
  Write-Host "[DENEME] su dosyalar commit edilecekti:" -ForegroundColor Yellow
  $degisiklik
} else {
  git add -A
  git commit -m $Mesaj
  if ($LASTEXITCODE -ne 0) { Dur "git commit basarisiz." }
  git push
  if ($LASTEXITCODE -ne 0) { Dur "git push basarisiz - commit YERELDE duruyor." }
}

Write-Host "`nTAMAM." -ForegroundColor Green