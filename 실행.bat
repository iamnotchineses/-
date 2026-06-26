@echo off
cd /d "%~dp0"

set "APP="
for %%f in ("브랜드매출*.py") do set "APP=%%f"
if not defined APP for %%f in ("app_auto.py") do set "APP=%%f"
if not defined APP for %%f in ("*.py") do set "APP=%%f"

if not defined APP (
  echo [오류] 이 폴더에 실행할 .py 파일이 없습니다.
  echo 대시보드 .py 와 이 .bat 을 같은 폴더에 두세요.
  pause
  exit /b
)

echo ===========================================
echo  실행: %APP%
echo  주소: http://localhost:8503
echo  사내공유: 같은 와이파이에서 http://이PC의IP:8503
echo ===========================================
echo.

py -m streamlit run "%APP%" --server.address=0.0.0.0 --server.port=8503

pause
