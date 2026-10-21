import os, sys
from streamlit.web import cli as stcli

def resource_path(rel):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)

if __name__ == "__main__":
    sys.argv = [
        "streamlit", "run", resource_path("app.py"),
        "--global.developmentMode=false",
        "--server.port=8501",
    ]
    sys.exit(stcli.main())
