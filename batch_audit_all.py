"""批量运行16科目AI审计，保存底稿到demo_outputs"""
import httpx, json, os, time, shutil, sys

API = 'http://127.0.0.1:8800'
OUT = 'D:/audit-agent/demo_outputs'
os.makedirs(OUT, exist_ok=True)

ALL = [
    ('C','货币资金'),('D','应收票据'),('E','其他应收款'),('F','预付账款'),
    ('G','预收账款'),('H','短期借款'),('I','应收账款'),('J','应付账款'),
    ('K','应付票据'),('L','应付职工薪酬'),('M','应交税费'),('N','其他应付款'),
    ('O','存货'),('P','固定资产'),('Q','无形资产'),('R','长期借款'),
]

for sid, sname in ALL:
    total_start = time.time()
    print(f'\n{"="*50}\n[{sid}] {sname}\n{"="*50}')

    r = httpx.get(f'{API}/api/demo-files?subject={sid}', timeout=10)
    files = r.json().get('files', [])
    if not files:
        print(f'  SKIP: no data files')
        continue
    print(f'  Files: {len(files)}')

    params = {
        'subject': sid,
        'entity_name': '重庆蛮先进智能制造有限公司',
        'period': '2024-12-31',
        'auditor': 'AI审计智能体',
        'firm_name': 'XX会计师事务所',
        'files': ','.join(files)
    }

    cevent = ''
    xlsx_saved = False
    log_saved = False

    try:
        resp = httpx.get(f'{API}/api/audit-stream', params=params, timeout=900)
        for line in resp.iter_lines():
            if not line: continue
            try:
                s = line.decode() if isinstance(line, bytes) else line
                if s.startswith('event: '): cevent = s[7:].strip(); continue
                if s.startswith('data: '):
                    d = json.loads(s[6:])
                    if cevent == 'progress' and d.get('pct', 0) % 25 < 5:
                        elapsed = time.time() - total_start
                        print(f"  [{d['pct']}%] {d.get('step','')[:40]} ({elapsed:.0f}s)")
                    if cevent == 'result':
                        src_xlsx = d.get('output_file')
                        src_log = d.get('log_file')
                        if src_xlsx and os.path.exists(src_xlsx):
                            dst = os.path.join(OUT, f'{sid}_{sname}_审计底稿.xlsx')
                            shutil.copy(src_xlsx, dst); xlsx_saved = True
                            print(f'  [OK] Excel: {dst}')
                        if src_log and os.path.exists(src_log):
                            dst = os.path.join(OUT, f'{sid}_{sname}_审计日志.docx')
                            shutil.copy(src_log, dst); log_saved = True
                            print(f'  [OK] Log: {dst}')
                    if cevent == 'error':
                        print(f'  [ERR] {d.get("msg","")[:120]}')
            except: pass
        elapsed = time.time() - total_start
        status = 'OK' if (xlsx_saved or log_saved) else 'NO OUTPUT'
        print(f'  [{status}] {elapsed:.0f}s')
    except Exception as e:
        print(f'  [FAIL] {time.time()-total_start:.0f}s: {e}')

print(f'\n{"="*50}')
print('ALL DONE!')
for f in sorted(os.listdir(OUT)):
    sz = os.path.getsize(os.path.join(OUT, f))
    print(f'  {f} ({sz:,} bytes)')
