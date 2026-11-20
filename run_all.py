# -*- coding: utf-8 -*-
"""One command: extract GitHub data -> SQLite -> sync Excel dashboard.   python run_all.py"""
import subprocess
import sys

for script in ("pipeline.py", "update_excel.py"):
    print(f"\n=== {script} ===")
    if subprocess.run([sys.executable, script]).returncode != 0:
        sys.exit(f"{script} failed. Stopped.")
print("\nDone. Open GitHub_Trending_Dashboard.xlsx")
