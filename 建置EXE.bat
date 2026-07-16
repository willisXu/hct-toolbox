@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo  HCT 工具箱 - EXE 一鍵建置
echo  (只需在這台電腦執行一次,約 5-10 分鐘)
echo ============================================

where python >nul 2>nul
if errorlevel 1 (
  echo [錯誤] 找不到 Python,請先到 python.org 安裝 3.10 以上版本,
  echo        安裝時勾選 "Add python.exe to PATH"
  pause & exit /b 1
)

echo [1/4] 建立建置用虛擬環境...
if not exist build_venv ( python -m venv build_venv )
call build_venv\Scripts\activate.bat

echo [2/4] 安裝套件(streamlit/openpyxl/xlrd/pyinstaller)...
pip install -q --upgrade pip
pip install -q -r requirements.txt pyinstaller
if errorlevel 1 ( echo [錯誤] 套件安裝失敗,請檢查網路 & pause & exit /b 1 )

echo [3/4] 打包 EXE(PyInstaller,輸出到 release\)...
pyinstaller --noconfirm --clean --onedir --name "HCT工具箱" ^
  --distpath release --workpath build ^
  --collect-all streamlit ^
  --copy-metadata streamlit ^
  --collect-all pandas ^
  --collect-all openpyxl ^
  --collect-all xlrd ^
  --collect-all pyarrow ^
  --hidden-import streamlit.runtime.scriptrunner.magic_funcs ^
  run_app.py
if errorlevel 1 ( echo [錯誤] 打包失敗,請把上面訊息截圖回報 & pause & exit /b 1 )

echo [4/4] 複製程式檔與設定到輸出資料夾(保持可編輯)...
set OUT=release\HCT工具箱
copy /y app.py "%OUT%\" >nul
xcopy /y /e /i core "%OUT%\core" >nul
xcopy /y /e /i mappings "%OUT%\mappings" >nul
copy /y "使用說明.txt" "%OUT%\" >nul

echo.
echo ============================================
echo  完成!成品資料夾:release\HCT工具箱\
echo  把整個資料夾複製給使用者,雙擊裡面的「HCT工具箱.exe」即可使用。
echo  日後修改 app.py / core / mappings\HCT範本.xlsx 不需重新打包。
echo ============================================
pause
