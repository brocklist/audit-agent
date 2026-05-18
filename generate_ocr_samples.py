"""生成 3 张演示 OCR 凭证图片（完税凭证 / 纳税申报表 / 增值税发票）"""
import os
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

OUT = Path(__file__).parent / "samples" / "ocr"
OUT.mkdir(parents=True, exist_ok=True)

# 查找中文字体
def find_font(sizes=[36, 28, 22, 18, 14]):
    """返回 {size: ImageFont} 字典"""
    candidates = [
        "C:/Windows/Fonts/simsun.ttc",
        "C:/Windows/Fonts/msyh.ttc",
        "C:/Windows/Fonts/simhei.ttf",
        "C:/Windows/Fonts/simkai.ttf",
        "C:/Windows/Fonts/mingliu.ttc",
    ]
    font_path = None
    for p in candidates:
        if os.path.exists(p):
            font_path = p
            break
    if not font_path:
        raise RuntimeError("未找到中文字体，请安装宋体/微软雅黑")
    return {s: ImageFont.truetype(font_path, s) for s in sizes}

F = find_font()

def draw_table(draw, x, y, w, cols, rows, col_widths=None):
    """在 draw 上绘制表格，返回右下角 y 坐标"""
    if col_widths is None:
        col_widths = [w // len(cols)] * len(cols)
    # 表头
    cx = x
    for ci, (col_name, cw) in enumerate(zip(cols, col_widths)):
        draw.rectangle([cx, y, cx + cw, y + 32], outline="black", width=1)
        draw.text((cx + 8, y + 6), col_name, fill="black", font=F[14])
        cx += cw
    y += 32
    # 数据行
    for row in rows:
        cx = x
        row_h = 28
        for ci, (val, cw) in enumerate(zip(row, col_widths)):
            draw.rectangle([cx, y, cx + cw, y + row_h], outline="black", width=1)
            draw.text((cx + 8, y + 4), str(val), fill="black", font=F[14])
            cx += cw
        y += row_h
    return y

# ================================================================
# 1. 完税凭证 — 增值税
# ================================================================
img = Image.new("RGB", (860, 600), "white")
d = ImageDraw.Draw(img)

# 标题
d.text((280, 24), "中 华 人 民 共 和 国", fill="black", font=F[22])
d.text((280, 56), "税 收 完 税 证 明", fill="black", font=F[36])

# 信息区
info_lines = [
    ("凭证号：","SW20250115001234"),
    ("纳税人识别号：","91500107MA60XXXXXXX"),
    ("纳税人名称：","重庆蛮先进智能制造有限公司"),
    ("税种：","增值税"),
    ("税款所属期：","2024年10月1日 — 2024年12月31日"),
    ("缴款金额（大写）：","陆拾柒万陆仟元整"),
    ("缴款金额（小写）：","¥ 676,000.00"),
    ("缴款日期：","2025年01月15日"),
    ("征收机关：","国家税务总局重庆市渝北区税务局"),
]
y = 100
for label, value in info_lines:
    d.text((120, y), label, fill="black", font=F[18])
    d.text((340, y), value, fill="black", font=F[18])
    y += 38

# 表格: 税种明细
y += 10
d.text((120, y), "税款明细：", fill="black", font=F[18])
y += 36
cols = ["税种", "税率", "计税金额", "应纳税额", "已预缴", "应补退"]
cws = [130, 80, 140, 140, 140, 140]
rows = [
    ["增值税", "13%", "5,200,000.00", "676,000.00", "0.00", "676,000.00"],
]
draw_table(d, 120, y, 770, cols, rows, cws)

# 印章
d.ellipse([620, 400, 800, 540], outline="red", width=3)
d.text((645, 450), "国家税务总局", fill="red", font=F[18])
d.text((645, 478), "重庆市渝北区", fill="red", font=F[18])
d.text((653, 506), "税务局", fill="red", font=F[18])

img.save(f"{OUT}/完税凭证_增值税.png")
print(f"[OK] 完税凭证_增值税.png")

# ================================================================
# 2. 纳税申报表 — 企业所得税季度预缴
# ================================================================
img = Image.new("RGB", (880, 600), "white")
d = ImageDraw.Draw(img)

d.text((240, 20), "中华人民共和国企业所得税", fill="black", font=F[18])
d.text((260, 46), "月（季）度预缴纳税申报表（A类）", fill="black", font=F[28])

# 基本信息行
meta = [
    ("纳税人识别号", "91500107MA60XXXXXXX"),
    ("纳税人名称", "重庆蛮先进智能制造有限公司"),
    ("申报期间", "2024年10月1日 — 2024年12月31日"),
    ("填报日期", "2025年01月15日"),
]
y = 84
for label, val in meta:
    d.text((80, y), f"{label}：{val}", fill="black", font=F[18])
    y += 32

# 主表
y += 8
cols = ["行次", "项    目", "本期金额", "累计金额"]
cws = [50, 380, 200, 200]
rows = [
    ["1", "一、营业收入", "5,200,000.00", "15,600,000.00"],
    ["2", "  减：营业成本", "3,120,000.00", "9,360,000.00"],
    ["3", "  减：税金及附加", "67,600.00", "202,800.00"],
    ["4", "  减：期间费用", "1,560,000.00", "4,680,000.00"],
    ["5", "二、利润总额（1-2-3-4）", "452,400.00", "1,357,200.00"],
    ["6", "三、应纳税所得额", "12,000,000.00", "36,000,000.00"],
    ["", "  （含纳税调整增加额）", "", ""],
    ["7", "四、适用税率", "25%", "25%"],
    ["8", "五、应纳所得税额（6×7）", "3,000,000.00", "9,000,000.00"],
    ["9", "  减：已预缴所得税额", "2,500,000.00", "7,500,000.00"],
    ["10", "六、本期应补（退）所得税额", "500,000.00", "1,500,000.00"],
]
y = draw_table(d, 80, y, 830, cols, rows, cws)

y += 16
d.text((80, y), "声明：此纳税申报表是根据《中华人民共和国企业所得税法》及相关规定填报的，是真实、可靠、完整的。", fill="black", font=F[14])

img.save(f"{OUT}/纳税申报表_企业所得税.png")
print(f"[OK] 纳税申报表_企业所得税.png")

# ================================================================
# 3. 增值税发票 — 进项发票
# ================================================================
img = Image.new("RGB", (880, 480), "white")
d = ImageDraw.Draw(img)

d.text((320, 14), "增 值 税 专 用 发 票", fill="#cc0000", font=F[28])
d.text((360, 50), "发     票     联", fill="#cc0000", font=F[18])

# 发票信息区
left_items = [
    ("发票代码：", "4401234567"),
    ("发票号码：", "0987654321"),
    ("开票日期：", "2024年12月20日"),
]
y = 80
for label, val in left_items:
    d.text((40, y), label + val, fill="black", font=F[18])
    y += 30

# 购买方 / 销售方
parties = [
    ("购", "名称：重庆蛮先进智能制造有限公司"),
    ("买", "纳税人识别号：91500107MA60XXXXXXX"),
    ("方", "地址：重庆市渝北区龙兴镇智能制造产业园A区"),
    ("销", "名称：重庆江南新材料科技有限公司"),
    ("售", "纳税人识别号：91500108MA60YYYYYYY"),
    ("方", "地址：重庆市南岸区茶园新区B栋"),
]
y = 80
for role, text in parties:
    d.text((500, y), f"{role}：{text}", fill="black", font=F[14])
    y += 26

# 货物明细表
y = 210
cols = ["货物/服务名称", "规格型号", "数量", "单价", "金额", "税率", "税额"]
cws = [200, 100, 80, 100, 140, 70, 140]
rows = [
    ["高精度伺服电机", "SV-2000", "50台", "17,000.00", "850,000.00", "13%", "110,500.00"],
]
draw_table(d, 40, y, 830, cols, rows, cws)

# 合计
y = 270
d.text((40, y), "价税合计（大写）：玖拾陆万零伍佰元整", fill="black", font=F[18])
d.text((600, y), "（小写）：¥ 960,500.00", fill="black", font=F[18])

# 销售方盖章
d.ellipse([680, 340, 830, 450], outline="red", width=2)
d.text((695, 380), "重庆江南新材", fill="red", font=F[14])
d.text((698, 400), "料科技有限公", fill="red", font=F[14])
d.text((710, 420), "司发票专用章", fill="red", font=F[14])

img.save(f"{OUT}/增值税发票_进项.png")
print(f"[OK] 增值税发票_进项.png")

print(f"\n全部完成：{len(list(OUT.glob('*.png')))} 张凭证图片 → {OUT}")
