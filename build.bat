@echo off
python -m pip install -U pip pyinstaller pyinstaller-hooks-contrib -r requirements.txt
pyinstaller --onefile --clean --name wb-ads-report ^
  --collect-all pandas --collect-all numpy --collect-all openpyxl --collect-all XlsxWriter ^
  build_wb_ads_report.py
echo Done. See dist\wb-ads-report.exe