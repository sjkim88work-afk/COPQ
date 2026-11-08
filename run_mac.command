#!/bin/bash
# 실패비용(COPQ) 관리 프로그램 — macOS 실행 파일 (Finder에서 더블클릭)
cd "$(dirname "$0")" || exit 1

if ! command -v python3 >/dev/null 2>&1; then
  echo "Python 3이 설치되어 있지 않습니다."
  echo "https://www.python.org/downloads/macos/ 에서 설치한 뒤 다시 실행하세요."
  read -r -p "엔터를 누르면 창을 닫습니다..."
  exit 1
fi

if [ ! -d ".venv" ]; then
  echo "[설치] 가상환경 생성 및 패키지 설치 중... (처음 한 번, 1~2분)"
  python3 -m venv .venv || { echo "가상환경 생성 실패"; read -r; exit 1; }
  source .venv/bin/activate
  python -m pip install --upgrade pip -q
  pip install -r requirements.txt -q || { echo "패키지 설치 실패 (인터넷 연결 확인)"; read -r; exit 1; }
else
  source .venv/bin/activate
fi

echo "브라우저에서 http://localhost:8501 이 열립니다. 종료하려면 이 창에서 Control+C"
streamlit run app.py --browser.gatherUsageStats false
