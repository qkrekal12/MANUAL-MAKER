"""Local launcher that prepares a private venv, then runs run_streamlit.py."""
from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path


APP_DIR = Path(__file__).resolve().parent
ENTRY_FILE = APP_DIR / "run_streamlit.py"
REQUIREMENTS = APP_DIR / "requirements.txt"
VENV_DIR = APP_DIR / ".venv"
PYTHON = VENV_DIR / ("Scripts/python.exe" if os.name == "nt" else "bin/python")
PORT = os.getenv("STREAMLIT_SERVER_PORT", "8767")


def run(command: list[str]) -> None:
    subprocess.run(command, cwd=APP_DIR, check=True)


def main() -> int:
    if sys.version_info < (3, 10):
        print("Python 3.10 이상이 필요합니다.")
        input("종료하려면 Enter를 누르세요.")
        return 1
    if not ENTRY_FILE.is_file() or not REQUIREMENTS.is_file():
        print("run_streamlit.py 또는 requirements.txt를 찾을 수 없습니다.")
        input("종료하려면 Enter를 누르세요.")
        return 1
    try:
        if not PYTHON.is_file():
            print("앱 전용 가상환경을 준비합니다...")
            run([sys.executable, "-m", "venv", str(VENV_DIR)])
        streamlit = VENV_DIR / ("Scripts/streamlit.exe" if os.name == "nt" else "bin/streamlit")
        if not streamlit.is_file():
            print("필요한 패키지를 설치합니다. 최초 실행 시 인터넷 연결이 필요합니다...")
            run([str(PYTHON), "-m", "pip", "install", "-r", str(REQUIREMENTS)])
        command = [str(PYTHON), "-m", "streamlit", "run", str(ENTRY_FILE), "--server.address", "127.0.0.1", "--server.port", PORT, "--browser.gatherUsageStats", "false"]
        print(f"매뉴얼 스튜디오를 시작합니다: http://127.0.0.1:{PORT}")
        run(command)
        return 0
    except subprocess.CalledProcessError as error:
        print(f"실행 단계에서 오류가 발생했습니다 (종료 코드 {error.returncode}).")
    except OSError as error:
        print(f"실행 파일을 시작하지 못했습니다: {error}")
    input("오류 내용을 확인한 뒤 Enter를 눌러 종료하세요.")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
