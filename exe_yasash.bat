@echo off
rem adolat_avtomat.py dan AdolatAvtomat\AdolatAvtomat.exe (oynali, papka ko'rinishida) yasaydi va tarqatish\AdolatAvtomat.zip ni yangilaydi
cd /d "%~dp0"
if "%BUILD_VENV%"=="" set BUILD_VENV=%LOCALAPPDATA%\AdolatAvtomat_build
if not exist "%BUILD_VENV%\Scripts\python.exe" (
  echo Birinchi marta: yig'ish muhiti tayyorlanmoqda...
  python -m venv "%BUILD_VENV%" || goto xato
)
"%BUILD_VENV%\Scripts\python.exe" -m pip install -q selenium pymupdf openpyxl python-docx opencv-python numpy pytesseract rapidocr-onnxruntime==1.2.3 pyinstaller || goto xato
"%BUILD_VENV%\Scripts\python.exe" -m PyInstaller --noconfirm --onedir --windowed --noupx --version-file "%~dp0versiya_info.txt" --name AdolatAvtomat --collect-all selenium --collect-all pytesseract --collect-all rapidocr_onnxruntime --workpath "%BUILD_VENV%\work" --specpath "%BUILD_VENV%\work" --distpath "%~dp0." adolat_avtomat.py || goto xato
rem 2-dastur: AdolatAvtomat2 (BUYRUQ tartibida) - AdolatAvtomat2\ dagi mijozlar/sozlamalar/natijalar saqlanib qoladi
"%BUILD_VENV%\Scripts\python.exe" -m PyInstaller --noconfirm --onedir --windowed --noupx --version-file "%~dp0versiya_info.txt" --name AdolatAvtomat2 --collect-all selenium --collect-all pytesseract --collect-all rapidocr_onnxruntime --workpath "%BUILD_VENV%\work2" --specpath "%BUILD_VENV%\work2" --distpath "%BUILD_VENV%\dist2" adolat_avtomat2.py || goto xato
if exist "%~dp0AdolatAvtomat2\_internal" rmdir /s /q "%~dp0AdolatAvtomat2\_internal"
xcopy /e /i /y /q "%BUILD_VENV%\dist2\AdolatAvtomat2" "%~dp0AdolatAvtomat2" >nul || goto xato
if not exist "%~dp0AdolatAvtomat2\Fuqarolik_sudlari_yurisdiksiya.xlsx" copy /y "%~dp0Fuqarolik_sudlari_yurisdiksiya.xlsx" "%~dp0AdolatAvtomat2\" >nul
if not exist "%~dp0AdolatAvtomat2\QOLLANMA.txt" copy /y "%~dp0QOLLANMA.txt" "%~dp0AdolatAvtomat2\" >nul
"%BUILD_VENV%\Scripts\python.exe" -m PyInstaller --noconfirm --onedir --windowed --noupx --name KvitansiyaYaratish --collect-all selenium --exclude-module cv2 --exclude-module numpy --exclude-module rapidocr_onnxruntime --exclude-module onnxruntime --workpath "%BUILD_VENV%\work_kv" --specpath "%BUILD_VENV%\work_kv" --distpath "%BUILD_VENV%\dist_kv" kvitansiya_yaratish.py || goto xato
rem KvitansiyaYaratish\ dagi mijozlar/sozlamalar/jurnal saqlanib qolsin - faqat dastur fayllari almashtiriladi
if exist "%~dp0KvitansiyaYaratish\_internal" rmdir /s /q "%~dp0KvitansiyaYaratish\_internal"
xcopy /e /i /y /q "%BUILD_VENV%\dist_kv\KvitansiyaYaratish" "%~dp0KvitansiyaYaratish" >nul || goto xato
if not exist "%~dp0KvitansiyaYaratish\Fuqarolik_sudlari_yurisdiksiya.xlsx" copy /y "%~dp0Fuqarolik_sudlari_yurisdiksiya.xlsx" "%~dp0KvitansiyaYaratish\" >nul
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0zip_yasash.ps1" || goto xato
echo.
echo OK: AdolatAvtomat.exe, KvitansiyaYaratish.exe va tarqatish\*.zip yangilandi
pause
exit /b 0
:xato
echo.
echo XATO! Yuqoridagi xabarni o'qing.
pause
exit /b 1
