@echo off
chcp 65001 > nul
cd /d %~dp0
if not exist .venv (
  echo [설치] 가상환경 생성 및 패키지 설치 중...
  python -m venv .venv
  call .venv\Scripts\activate.bat
  python -m pip install --upgrade pip
  pip install -r requirements.txt
) else (
  call .venv\Scripts\activate.bat
)
streamlit run app.py
pause
