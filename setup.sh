#!/bin/bash
# lab_chore setup.sh — 최초 1회 실행: bash setup.sh
set -e
echo "🔬 CSNL lab_chore 설치 시작..."

PYTHON=""
for py in "$HOME/.pyenv/shims/python3" "/opt/homebrew/bin/python3" "/usr/local/bin/python3" "python3"; do
    if "$py" --version &>/dev/null 2>&1; then PYTHON="$py"; break; fi
done

if [ -z "$PYTHON" ]; then
    echo "❌ Python 3를 찾을 수 없습니다. brew install python 을 먼저 실행하세요."; exit 1
fi
echo "✅ Python: $($PYTHON --version)"

echo "📦 패키지 설치 중..."
"$PYTHON" -m pip install -r "$(dirname "$0")/requirements.txt" --quiet \
    --break-system-packages 2>/dev/null \
    || "$PYTHON" -m pip install -r "$(dirname "$0")/requirements.txt" --quiet
echo "✅ 패키지 설치 완료"

APP="$(dirname "$0")/실험참여자비GUI.app"
chmod +x "$APP/Contents/MacOS/run"
xattr -dr com.apple.quarantine "$APP" 2>/dev/null || true
echo "✅ .app 권한 설정 완료"

echo ""
echo "═══════════════════════════════════════════════════════"
echo "🎉 설치 완료!"
echo ""
echo "  1. '실험참여자비GUI.app' 을 Excel 템플릿 파일들과"
echo "     같은 폴더에 복사하세요."
echo "  2. .app 더블클릭 → 브라우저에서 GUI 실행"
echo ""
echo "  업로드 양식 자동화 (템플릿 폴더에서 실행):"
echo "  LAB_CHORE_DIR=. python3 <path>/upload_updater.py --all"
echo "═══════════════════════════════════════════════════════"
