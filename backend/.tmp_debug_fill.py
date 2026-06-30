import sys
import time
import traceback
from pathlib import Path

sys.path.insert(0, r'.')
import main

uploads = Path('uploads')
excel = max(uploads.glob('*.xlsm'), key=lambda p: p.stat().st_mtime)
template = max(uploads.glob('*.hwp'), key=lambda p: p.stat().st_mtime)
out = Path('outputs/test_web_fill_2.hwp')

print('start', time.time(), flush=True)
print('excel', excel.name, flush=True)
print('template', template.name, flush=True)
try:
    result = main.fill_hwp_with_extracted_data(excel, template, out)
    print('after call', time.time(), result, out.exists(), out.stat().st_size if out.exists() else 0, flush=True)
except Exception:
    print('error', flush=True)
    print(traceback.format_exc(), flush=True)
