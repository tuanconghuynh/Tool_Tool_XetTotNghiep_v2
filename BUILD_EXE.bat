@echo off
cd /d "%~dp0"
py -m PyInstaller --noconfirm --clean --onefile --windowed --collect-all customtkinter --name XetTotNghiep_v2 xet_tot_nghiep.py
if errorlevel 1 goto failed
copy /Y cau_hinh.json dist\cau_hinh.json
echo Thanh cong. File nam trong thu muc dist.
:failed
pause
