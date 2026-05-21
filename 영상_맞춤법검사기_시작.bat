@echo off
chcp 65001 > nul
cd /d "%~dp0"

REM ── VS Code 자동 연결 차단 ──────────────────────
set VSCODE_PID=
set VSCODE_IPC_HOOK=
set VSCODE_IPC_HOOK_CLI=
set ELECTRON_RUN_AS_NODE=
set TERM_PROGRAM=

REM ── 필수 패키지 확인 ────────────────────────────
pip show pillow > nul 2>&1
if errorlevel 1 (
    echo Pillow 설치 중...
    pip install pillow -q
)

REM ── Streamlit 실행 ──────────────────────────────
python -m streamlit run app_video.py
if errorlevel 1 (
    echo.
    echo [오류] 실행 실패. 필수 패키지를 설치합니다...
    pip install streamlit openai moviepy opencv-python numpy pandas openpyxl pillow
    python -m streamlit run app_video.py
)
pause
