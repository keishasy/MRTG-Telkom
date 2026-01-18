import subprocess, sys, time, webbrowser
from pathlib import Path

def main():
    base_dir = Path(sys.argv[0]).resolve().parent
    app_path = base_dir / "app.py"

    cmd = [
        sys.executable, "-m", "streamlit", "run", str(app_path),
        "--server.headless=true",
        "--browser.gatherUsageStats=false",
        "--server.port=8501"
    ]

    p = subprocess.Popen(cmd, cwd=str(base_dir))
    time.sleep(2)
    webbrowser.open("http://localhost:8501")
    p.wait()

if __name__ == "__main__":
    main()
