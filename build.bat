@echo off
chcp 65001 > nul
cd /d "%~dp0"
py -3 -m venv .venv
if errorlevel 1 (
  echo Python 3가 설치되어 있지 않거나 PATH에 등록되지 않았습니다.
  pause
  exit /b 1
)
call .venv\Scripts\activate
python -m pip install --upgrade pip
pip install pandas openpyxl xlrd pyinstaller
pyinstaller --noconfirm --clean --onefile --windowed --name "고객사별_금액대조" amount_checker.py
if errorlevel 1 (
  echo EXE 생성에 실패했습니다.
  pause
  exit /b 1
)
echo.
echo 생성 완료: dist\고객사별_금액대조.exe
pause
