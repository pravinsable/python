# C:\pravin\src\github\pravinsable\python\dexcom.py
import os
import sys
from pydexcom import Dexcom

# Ensure output is UTF-8 for trend arrows
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')

try:
    dexcom = Dexcom(
        username=os.getenv("DEXCOM_USER"), 
        password=os.getenv("DEXCOM_PASS"), 
        region=os.getenv("DEXCOM_REGION", "us")
    )
    bg = dexcom.get_current_glucose_reading()
    if bg:
        # 1.1.5: trend_arrow provides the Unicode icon (↑, ↓, →)
        print(f"{bg.value} {bg.trend_arrow}")
    else:
        print("No Data")
except Exception:
    print("Error")

    
