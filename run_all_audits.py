"""
逐一运行6科目审计并保存生成的底稿Excel
输出到: D:/audit-agent/demo_outputs/
"""
import httpx, json, os, time, re
from datetime import datetime

API = "http://127.0.0.1:8800"
OUT = "D:/audit-agent/demo_outputs"
os.makedirs(OUT, exist_ok=True)

SUBJECTS = [
    {"id": "C", "name": "货币资金", "entity": "重庆蛮先进智能制造有限公司", "period": "2024-12-31"},
    {"id": "D", "name": "应收票据", "entity": "重庆蛮先进智能制造有限公司", "period": "2024-12-31"},
    {"id": "E", "name": "其他应收款", "entity": "重庆蛮先进智能制造有限公司", "period": "2024-12-31"},
    {"id": "F", "name": "预付账款", "entity": "重庆蛮先进智能制造有限公司", "period": "2024-12-31"},
    {"id": "G", "name": "预收账款", "entity": "重庆蛮先进智能制造有限公司", "period": "2024-12-31"},
    {"id": "H", "name": "短期借款", "entity": "重庆蛮先进智能制造有限公司", "period": "2024-12-31"},
]

for subj in SUBJECTS:
    sid = subj["id"]
    print(f"\n{'='*60}")
    print(f"开始审计: {sid} - {subj['name']}")
    print(f"{'='*60}")

    # 获取该科目的demo文件
    r = httpx.get(f"{API}/api/demo-files?subject={sid}")
    files = r.json().get("files", [])
    if not files:
        print(f"  无数据文件，跳过")
        continue

    print(f"  数据文件: {len(files)} 个")
    for f in files:
        print(f"    - {os.path.basename(f)}")

    # 调用SSE审计
    params = {
        "subject": sid,
        "entity_name": subj["entity"],
        "period": subj["period"],
        "auditor": "AI审计智能体",
        "reviewer": "",
        "firm_name": "XX会计师事务所",
        "files": ",".join(files)
    }

    start = time.time()
    current_event = ""
    try:
        resp = httpx.get(f"{API}/api/audit-stream", params=params, timeout=900)
        reader = resp.iter_lines()
        result_data = None
        last_pct = 0

        for line in reader:
            if not line: continue
            try:
                line_str = line.decode() if isinstance(line, bytes) else line
                if line_str.startswith("event: "):
                    current_event = line_str[7:].strip()
                    continue
                if line_str.startswith("data: "):
                    data = json.loads(line_str[6:])
                    if current_event == "progress":
                        pct = data.get("pct", 0)
                        if pct >= last_pct + 5:
                            last_pct = pct
                            elapsed = time.time() - start
                            print(f"  [{pct}%] {data.get('step','')} ({elapsed:.0f}s)")
                    elif current_event == "result":
                        result_data = data
                        print(f"  >>> 审计完成! 底稿: {data.get('output_file')}")
                    elif current_event == "error":
                        print(f"  !!! 错误: {data.get('msg')}")
            except:
                pass

        elapsed = time.time() - start
        if result_data:
            # 复制底稿到demo_outputs
            src = result_data.get("output_file")
            if src and os.path.exists(src):
                dst_name = f"{sid}_{subj['name']}_审计底稿_{subj['period'].replace('-','')}.xlsx"
                dst = os.path.join(OUT, dst_name)
                import shutil
                shutil.copy(src, dst)
                print(f"  底稿已保存: {dst}")
                print(f"  总耗时: {elapsed:.0f}s")
            print(f"  {sid} [OK]")
        else:
            print(f"  {sid} [FAILED] - 无结果数据")
    except Exception as e:
        elapsed = time.time() - start
        print(f"  {sid} [ERROR] {elapsed:.0f}s: {e}")

print("\n全部审计完成!")
print(f"输出目录: {OUT}/")
for f in sorted(os.listdir(OUT)):
    print(f"  {f}")
