#!/bin/bash
# ============================================================
# 인스턴스 리소스 분석 대시보드 설치 스크립트
# GitHub: https://github.com/data-jy/instance-resource-dashboard
# 포트: 5001
# ============================================================
set -e

REPO_URL="https://github.com/data-jy/instance-resource-dashboard.git"
APP_DIR="$HOME/resource-analyze"
PORT=5001
XLSX_JS_URL="https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js"

echo "========================================"
echo " 인스턴스 리소스 분석 대시보드 설치"
echo "========================================"

# Python 버전 확인
python3 --version || { echo "[오류] python3 가 없습니다."; exit 1; }

# git clone
if [ -d "$APP_DIR" ]; then
  echo "[INFO] 디렉토리가 이미 존재합니다. git pull 로 업데이트합니다."
  cd "$APP_DIR" && git pull
else
  echo "[INFO] 저장소 클론 중..."
  git clone "$REPO_URL" "$APP_DIR"
  cd "$APP_DIR"
fi

# pip 패키지 설치
echo "[INFO] 패키지 설치 중..."
pip3 install pandas openpyxl

# JS 라이브러리 로컬 다운로드 (외부 CDN 의존성 제거 → 폐쇄망 호환)
echo "[INFO] xlsx.bundle.js 다운로드 중..."
mkdir -p "$APP_DIR/static"
if [ ! -f "$APP_DIR/static/xlsx.bundle.js" ]; then
  curl -fsSL "$XLSX_JS_URL" -o "$APP_DIR/static/xlsx.bundle.js" \
    && echo "[INFO] xlsx.bundle.js 다운로드 완료 ($(wc -c < "$APP_DIR/static/xlsx.bundle.js") bytes)" \
    || echo "[경고] xlsx.bundle.js 다운로드 실패. 인터넷 연결을 확인하세요."
else
  echo "[INFO] xlsx.bundle.js 이미 존재함 — 건너뜀"
fi

# 기존 프로세스 종료
echo "[INFO] 기존 프로세스 종료..."
pkill -f "server.py --port $PORT" 2>/dev/null || true
sleep 1

# 앱 실행
echo "[INFO] 앱 시작 (포트 $PORT)..."
nohup python3 "$APP_DIR/server.py" --port "$PORT" > "$APP_DIR/server.log" 2>&1 &
APP_PID=$!
sleep 2

if ps -p $APP_PID > /dev/null 2>&1; then
  echo ""
  echo "========================================"
  echo " 설치 및 실행 완료!"
  echo " PID: $APP_PID"
  echo " URL: http://$(hostname -I | awk '{print $1}'):$PORT"
  echo " 로그: $APP_DIR/server.log"
  echo "========================================"
else
  echo "[오류] 앱 실행 실패. 로그를 확인하세요:"
  cat "$APP_DIR/server.log"
  exit 1
fi

# systemd 서비스 등록 여부 묻기
read -p "[선택] systemd 서비스로 등록하시겠습니까? (재부팅 후 자동 시작) [y/N]: " ans
if [[ "$ans" =~ ^[Yy]$ ]]; then
  USER_NAME=$(whoami)
  sudo tee /etc/systemd/system/resource-analyze.service > /dev/null << EOF
[Unit]
Description=Instance Resource Analysis Dashboard
After=network.target

[Service]
Type=simple
User=$USER_NAME
WorkingDirectory=$APP_DIR
ExecStart=/usr/bin/python3 server.py --port $PORT
Restart=always
RestartSec=5
Environment=PYTHONUNBUFFERED=1

[Install]
WantedBy=multi-user.target
EOF
  sudo systemctl daemon-reload
  sudo systemctl enable resource-analyze
  echo "[INFO] systemd 서비스 등록 완료 (resource-analyze)"
fi

echo "완료."
