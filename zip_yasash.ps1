# tarqatish\AdolatAvtomat.zip ni yasaydi: dastur, qo'llanma va sudlar jadvali.
# Shaxsiy mijoz hujjatlari, jurnallar va hisob ma'lumotlari arxivga kiritilmaydi.
$ErrorActionPreference = "Stop"
$loyiha = $PSScriptRoot
$tarqatish = Join-Path $loyiha "tarqatish"
New-Item -ItemType Directory -Path $tarqatish -Force | Out-Null
# AdolatAvtomat (da'vo) va AdolatAvtomat2 (buyruq) - bir xil tartibda
foreach ($nom in @("AdolatAvtomat", "AdolatAvtomat2")) {
    $vaqtincha = Join-Path $env:TEMP "${nom}_zip"
    $ichki = Join-Path $vaqtincha $nom
    if (Test-Path -LiteralPath $vaqtincha) { Remove-Item -LiteralPath $vaqtincha -Recurse -Force }
    New-Item -ItemType Directory -Path $ichki | Out-Null
    # Faqat yig'ilgan dastur fayllari; ish va login ma'lumotlari alohida papkada qoladi.
    foreach ($x in @("$nom.exe", "_internal")) {
        $manba = Join-Path $loyiha "$nom\$x"
        if (-not (Test-Path -LiteralPath $manba)) { throw "Dastur fayli topilmadi: $manba" }
        Copy-Item -LiteralPath $manba -Destination $ichki -Recurse
    }
    foreach ($x in @("QOLLANMA.txt", "Fuqarolik_sudlari_yurisdiksiya.xlsx")) {
        $manba = Join-Path $loyiha $x
        if (Test-Path -LiteralPath $manba) { Copy-Item -LiteralPath $manba -Destination $ichki }
    }
    foreach ($x in @("mijozlar", "umumiy_hujjatlar")) {
        New-Item -ItemType Directory -Path (Join-Path $ichki $x) | Out-Null
    }
    # Yangi foydalanuvchi login va parolni o'z sozlamalarida kiritadi.
    [System.IO.File]::WriteAllText((Join-Path $ichki "sozlamalar.ini"), "[ADOLAT]`r`nlogin =`r`nparol =`r`n", [System.Text.UTF8Encoding]::new($false))
    Compress-Archive -Path $ichki -DestinationPath (Join-Path $tarqatish "$nom.zip") -Force
    Remove-Item -LiteralPath $vaqtincha -Recurse -Force
}
# ALOHIDA dastur: KvitansiyaYaratish - faqat dastur (mijozlar/jurnal/sozlamalar - o'zida yaratiladi)
$kv = Join-Path $loyiha "KvitansiyaYaratish"
if (Test-Path -LiteralPath $kv) {
    $kvv = Join-Path $env:TEMP "KvitansiyaYaratish_zip"
    if (Test-Path -LiteralPath $kvv) { Remove-Item -LiteralPath $kvv -Recurse -Force }
    New-Item -ItemType Directory -Path $kvv | Out-Null
    $kvd = Join-Path $kvv "KvitansiyaYaratish"
    New-Item -ItemType Directory -Path $kvd -Force | Out-Null
    Get-ChildItem -LiteralPath $kv | Where-Object { $_.Name -in @("KvitansiyaYaratish.exe", "_internal", "Fuqarolik_sudlari_yurisdiksiya.xlsx") } | ForEach-Object { Copy-Item -LiteralPath $_.FullName -Destination $kvd -Recurse -Force }
    Compress-Archive -LiteralPath $kvd -DestinationPath (Join-Path $tarqatish "KvitansiyaYaratish.zip") -Force
    Remove-Item -LiteralPath $kvv -Recurse -Force
}
Write-Output "OK: tarqatish\AdolatAvtomat.zip va KvitansiyaYaratish.zip tayyor"
