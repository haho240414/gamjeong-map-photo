#!/bin/bash
cd "$(dirname "$0")"

if [ ! -d ".venv" ]; then
  echo "최초 실행: 가상환경을 만들고 필요한 패키지를 설치합니다..."
  python3 -m venv .venv
  ./.venv/bin/python -m pip install -q --upgrade pip
  ./.venv/bin/python -m pip install -q -r requirements.txt
fi

if [ ! -f ".env" ]; then
  cp .env.example .env
fi

./.venv/bin/python webapp_server.py --host 127.0.0.1 --port 8877 --open
