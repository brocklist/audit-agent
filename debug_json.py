import re, json

raw = open('D:/audit-agent/output/llm_raw_20260509_125931.txt', encoding='utf-8').read()

m = re.search(r'```json\s*([\s\S]*?)\s*```', raw)
if m:
    j = m.group(1)
    pos = 58247
    start = max(0, pos - 80)
    end = min(len(j), pos + 80)
    context = j[start:end]
    print('Context around char 58247:')
    print(repr(context))

    before = j[:pos]
    line_num = before.count('\n') + 1
    print(f'\nError at line {line_num}')
    lines = j.split('\n')
    for i in range(max(0, line_num-4), min(len(lines), line_num+3)):
        marker = '>>>' if i == line_num-1 else '   '
        print(f'{marker} L{i+1}: {lines[i][:200]}')
