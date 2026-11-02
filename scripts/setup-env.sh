#!/usr/bin/env bash
# 클라우드 세션이 시작될 때마다 자동 실행: watch 스킬에 필요한 프로그램과 설정을 준비한다.
# 내 PC(로컬)에서 열었을 때는 아무것도 하지 않는다. 실패해도 세션은 정상 시작된다.

[ "${CLAUDE_CODE_REMOTE:-}" = "true" ] || exit 0

log() { echo "[setup-env] $*"; }

# 1) 영상 다운로드 도구 — 유튜브가 자주 바뀌므로 매번 최신으로 맞춘다
pip install --break-system-packages -q -U yt-dlp >/dev/null 2>&1 \
  || pip install -q -U yt-dlp >/dev/null 2>&1 \
  || log "yt-dlp 설치 실패"

# 2) 영상 처리 도구
if ! command -v ffmpeg >/dev/null 2>&1; then
  SUDO=""; [ "$(id -u)" != "0" ] && command -v sudo >/dev/null 2>&1 && SUDO="sudo"
  ($SUDO apt-get update -qq && $SUDO apt-get install -y -qq ffmpeg) >/dev/null 2>&1 \
    || log "ffmpeg 설치 실패"
fi

# 3) watch 기본 설정: 로컬 엔진 / balanced / 음성 인식 fallback 없음
CONF_DIR="$HOME/.config/watch"
if [ ! -f "$CONF_DIR/.env" ]; then
  mkdir -p "$CONF_DIR"
  cat > "$CONF_DIR/.env" <<'EOF'
GEMINI_API_KEY=
GROQ_API_KEY=
OPENAI_API_KEY=
WATCH_ENGINE="local"
WATCH_DETAIL="balanced"
WATCH_WHISPER_BACKEND="none"
SETUP_COMPLETE="true"
EOF
  chmod 600 "$CONF_DIR/.env"
fi

log "준비 완료 (yt-dlp: $(yt-dlp --version 2>/dev/null || echo 없음), ffmpeg: $(command -v ffmpeg >/dev/null && echo 있음 || echo 없음))"
exit 0
