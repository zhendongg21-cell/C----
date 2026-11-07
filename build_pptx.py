# -*- coding: utf-8 -*-
"""生成《基于 ESP32-S3 的智能健康手环》答辩/评审汇报 PPTX（19 页，16:9，深色主题）。"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

OUT = r"C:\Users\gaogo\Desktop\智能手环\ESP32-S3智能健康手环_答辩汇报.pptx"
TOTAL = 19

def I(v):
    return Inches(v)

# ---- 配色：沿用文档第 3 章界面规范 ----
BG    = RGBColor(0x0B, 0x12, 0x20)   # 背景 #0B1220
CARD  = RGBColor(0x15, 0x22, 0x38)   # 卡片 #152238
CARD2 = RGBColor(0x1B, 0x2B, 0x45)
ROW_A = RGBColor(0x11, 0x1C, 0x2F)
LINE  = RGBColor(0x2A, 0x3B, 0x57)
TXT   = RGBColor(0xF8, 0xFA, 0xFC)   # 主文字
SUB   = RGBColor(0x94, 0xA3, 0xB8)   # 次要文字
DIM   = RGBColor(0x5B, 0x6B, 0x85)
TEAL  = RGBColor(0x2D, 0xD4, 0xBF)   # 运动
RED   = RGBColor(0xFF, 0x5A, 0x6E)   # 心率
AMBER = RGBColor(0xFB, 0xBF, 0x24)   # 体温/久坐
BLUE  = RGBColor(0x60, 0xA5, 0xFA)
FONT  = "Microsoft YaHei"
LINEHEX = "2A3B57"


def set_run(r, size=13, color=TXT, bold=False):
    f = r.font
    f.size = Pt(size)
    f.bold = bold
    f.color.rgb = color
    f.name = FONT
    rPr = r._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set("typeface", FONT)


def tb(slide, x, y, w, h, text, size=13, color=TXT, bold=False,
       align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spacing=1.12,
       space_after=2, wrap=True):
    sh = slide.shapes.add_textbox(I(x), I(y), I(w), I(h))
    tf = sh.text_frame
    tf.word_wrap = wrap
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    for i, line in enumerate(str(text).split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(space_after)
        p.line_spacing = spacing
        r = p.add_run()
        r.text = line
        set_run(r, size, color, bold)
    return sh


def rect(slide, x, y, w, h, fill, line=None, rounded=False, lw=0.75):
    typ = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    sh = slide.shapes.add_shape(typ, I(x), I(y), I(w), I(h))
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(lw)
    if rounded:
        try:
            sh.adjustments[0] = 0.08
        except Exception:
            pass
    return sh


def bullets(slide, x, y, w, h, items, size=12.5, gap=7, color=SUB,
            dot=TEAL, spacing=1.16):
    sh = slide.shapes.add_textbox(I(x), I(y), I(w), I(h))
    tf = sh.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, it in enumerate(items):
        if isinstance(it, tuple):
            head, body = it
        else:
            head, body = "", it
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(gap)
        p.line_spacing = spacing
        r0 = p.add_run()
        r0.text = "•  "
        set_run(r0, size, dot, True)
        if head:
            rh = p.add_run()
            rh.text = head
            set_run(rh, size, TXT, True)
        r1 = p.add_run()
        r1.text = body
        set_run(r1, size, color)
    return sh


def set_cell_border(cell, hexcolor=LINEHEX, width=6350):
    tcPr = cell._tc.get_or_add_tcPr()
    idx = 0
    for edge in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        ln = tcPr.find(qn(edge))
        if ln is None:
            ln = tcPr.makeelement(qn(edge), {})
            tcPr.insert(idx, ln)
        idx += 1
        for ch in list(ln):
            ln.remove(ch)
        ln.set("w", str(width))
        sf = ln.makeelement(qn("a:solidFill"), {})
        ln.append(sf)
        sc = ln.makeelement(qn("a:srgbClr"), {"val": hexcolor})
        sf.append(sc)


def table(slide, x, y, w, h, data, col_w=None, fs=10.5, hs=10.5,
          head_color=TEAL, align_left_cols=(0,)):
    rows, cols = len(data), len(data[0])
    gf = slide.shapes.add_table(rows, cols, I(x), I(y), I(w), I(h))
    tbl = gf.table
    tbl.first_row = False
    tbl.first_col = False
    tbl.horz_banding = False
    tbl.vert_banding = False
    if col_w:
        tot = float(sum(col_w))
        for j, cw in enumerate(col_w):
            tbl.columns[j].width = Emu(int(I(w) * cw / tot))
    for i, row in enumerate(data):
        for j, val in enumerate(row):
            c = tbl.cell(i, j)
            c.text = str(val)
            c.fill.solid()
            c.fill.fore_color.rgb = CARD2 if i == 0 else (CARD if i % 2 else ROW_A)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            c.margin_left = I(0.07)
            c.margin_right = I(0.07)
            c.margin_top = I(0.025)
            c.margin_bottom = I(0.025)
            for p in c.text_frame.paragraphs:
                p.alignment = PP_ALIGN.LEFT if j in align_left_cols else PP_ALIGN.CENTER
                p.space_after = Pt(0)
                p.line_spacing = 1.0
                for r in p.runs:
                    set_run(r, hs if i == 0 else fs,
                            head_color if i == 0 else TXT, i == 0 or j == 0)
            set_cell_border(c)
    return gf


def new_slide(prs, num, title, subtitle=None, accent=TEAL):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = BG
    rect(s, 0.55, 0.40, 0.075, 0.46, accent)
    tb(s, 0.78, 0.33, 10.6, 0.58, title, 25, TXT, True)
    if subtitle:
        tb(s, 0.80, 0.92, 11.4, 0.30, subtitle, 11.5, SUB)
    tb(s, 11.50, 0.46, 1.30, 0.30, "%02d / %d" % (num, TOTAL),
       10, DIM, align=PP_ALIGN.RIGHT)
    rect(s, 0.55, 1.30, 12.23, 0.014, LINE)
    tb(s, 0.55, 7.10, 10.0, 0.24,
       "基于 ESP32-S3 的智能健康手环 · 项目开发文档说明书（C++版）", 8.5, DIM)
    tb(s, 10.90, 7.10, 1.88, 0.24, "答辩 / 评审汇报", 8.5, DIM,
       align=PP_ALIGN.RIGHT)
    return s


def card_title(slide, x, y, w, title, accent=TEAL, size=14):
    rect(slide, x, y + 0.035, 0.05, 0.25, accent)
    tb(slide, x + 0.15, y, w - 0.2, 0.32, title, size, TXT, True)


def chip(slide, x, y, w, h, text, color=TEAL, size=10.5, fill=CARD2):
    rect(slide, x, y, w, h, fill, rounded=True)
    tb(slide, x, y, w, h, text, size, color, True,
       align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def layer(slide, x, y, w, h, name, desc, color):
    rect(slide, x, y, w, h, CARD, rounded=True)
    rect(slide, x, y, 0.055, h, color)
    tb(slide, x + 0.18, y + 0.05, 1.35, h - 0.1, name, 12, color, True,
       anchor=MSO_ANCHOR.MIDDLE)
    tb(slide, x + 1.62, y + 0.05, w - 1.80, h - 0.1, desc, 10.2, SUB,
       spacing=1.05, anchor=MSO_ANCHOR.MIDDLE)


def hbar(slide, x, y, w, label, val, maxv, color, lw=1.95, unit="学时"):
    tb(slide, x, y + 0.02, lw, 0.28, label, 10.5, TXT, True)
    bx = x + lw + 0.10
    bw = w - lw - 0.66
    rect(slide, bx, y + 0.065, bw, 0.19, ROW_A)
    fw = bw * val / float(maxv)
    if fw > 0.02:
        rect(slide, bx, y + 0.065, fw, 0.19, color)
    tb(slide, bx + bw + 0.10, y + 0.02, 0.58, 0.28,
       "%d%s" % (val, unit), 10.5, color, True)


prs = Presentation()
prs.slide_width = I(13.333)
prs.slide_height = I(7.5)

# ============================================================ 01 封面
s = prs.slides.add_slide(prs.slide_layouts[6])
s.background.fill.solid()
s.background.fill.fore_color.rgb = BG
rect(s, 0.0, 0.0, 0.16, 7.5, TEAL)
rect(s, 0.16, 0.0, 0.06, 7.5, RGBColor(0x12, 0x3A, 0x47))
tb(s, 1.05, 1.05, 10.8, 0.32, "课 程 设 计 · 答 辩 汇 报", 13, TEAL, True)
tb(s, 1.02, 1.62, 11.2, 1.25, "基于 ESP32-S3 的智能健康手环", 40, TXT, True)
tb(s, 1.05, 2.92, 11.0, 0.46,
   "项目开发文档说明书（C++版）", 20, SUB)
rect(s, 1.05, 3.62, 2.2, 0.055, TEAL)
tb(s, 1.05, 3.92, 11.0, 0.90,
   "含 DeepSeek AI 辅助开发过程记录\n运动与久坐提醒为核心的软硬结合综合实训",
   14, SUB, spacing=1.45)

info = [
    ("核心平台", "ESP32-S3 + FreeRTOS + LVGL + C++17"),
    ("教学安排", "80 学时 · 10 天集中实训（可裁剪 64）"),
    ("组织形式", "3～4 人 / 组 · 五类角色分工"),
    ("选做扩展", "PC 管理端 Qt 5 / Wi-Fi 同步 / 心率血氧；新增：大模型健康问答 · RAG 检索 · 手环操控鼠标"),
]
for i, (k, v) in enumerate(info):
    yy = 5.02 + i * 0.34
    tb(s, 1.05, yy, 1.35, 0.28, k, 11.5, TEAL, True)
    tb(s, 2.45, yy, 8.6, 0.28, v, 11.5, SUB)

rect(s, 1.05, 6.60, 11.20, 0.014, LINE)
for i, t in enumerate(["提醒能触发", "数据能显示", "断电不丢设置", "连续运行不死机"]):
    chip(s, 1.05 + i * 2.85, 6.78, 2.62, 0.38, t,
         color=AMBER if i >= 2 else TEAL, size=10.5, fill=CARD)
tb(s, 10.65, 1.10, 2.1, 0.30, "2026 年 10 月", 11.5, SUB,
   align=PP_ALIGN.RIGHT)

# ============================================================ 02 汇报议程
s = new_slide(prs, 2, "汇报议程", "Agenda · 从问题定义到教学落地与答辩验收")
agenda = [
    ("01", "项目定位与需求", "背景 / 核心闭环 / 功能范围 / 约束与取舍", TEAL),
    ("02", "总体架构与硬件", "终端四层 + 服务端四层 / 选型 / 引脚 / 结构", BLUE),
    ("03", "软件设计与核心算法", "FreeRTOS 任务 / 计步 / 久坐状态机 / 界面", TEAL),
    ("04", "实现方案与可靠性", "六个设计模式 / 本地存储 / 同步 / 容错", AMBER),
    ("05", "测试与课程设计", "测试方案 / 量化指标 / 80 学时 / 10 天计划", BLUE),
    ("06", "考核评价与总结", "评分构成 / 设计原则自评 / 难易度 / 展望", TEAL),
]
for i, (n, t, d, c) in enumerate(agenda):
    col, row = i % 3, i // 3
    x, y = 0.55 + col * 4.15, 1.62 + row * 2.62
    rect(s, x, y, 3.92, 2.34, CARD, rounded=True)
    rect(s, x, y, 0.055, 2.34, c)
    tb(s, x + 0.28, y + 0.24, 1.3, 0.62, n, 34, c, True)
    tb(s, x + 0.30, y + 1.00, 3.35, 0.40, t, 16, TXT, True)
    tb(s, x + 0.30, y + 1.48, 3.35, 0.70, d, 11, SUB, spacing=1.25)

# ============================================================ 03 背景与定位
s = new_slide(prs, 3, "项目背景与课程定位", "为什么做：让嵌入式教学里的每一行代码都有直接反馈")
rect(s, 0.55, 1.58, 5.95, 4.05, CARD, rounded=True)
card_title(s, 0.80, 1.80, 5.4, "背景与教学痛点")
bullets(s, 0.82, 2.28, 5.35, 3.10, [
    ("健康诉求：", "久坐行为危害明确，可穿戴设备需要“察觉—提醒—反馈”的闭环能力。"),
    ("教学痛点：", "传统实验反馈链路长、结果不可见，学生难建立“代码—现象”的直觉。"),
    ("项目主张：", "屏幕亮起、步数跳动、马达震动，替代串口里的 success。"),
    ("能力目标：", "能跑 → 好用 → 可靠，三台阶递进，而非只交一份能编译的工程。"),
    ("软硬结合：", "从寄存器自检、算法调参到装配验收，覆盖完整工程链条。"),
], size=12, gap=11)

rect(s, 6.78, 1.58, 6.00, 4.05, CARD, rounded=True)
card_title(s, 7.03, 1.80, 5.4, "课程定位（表 1-1 / 表 5-1）", accent=BLUE)
facts = [
    ("总学时", "80 学时（可按 64 学时裁剪）"),
    ("实施周期", "10 天集中实训"),
    ("分组规模", "3～4 人 / 组，设 5 类角色"),
    ("先修基础", "C++ / 单片机 / 数字电路基础"),
    ("技术主线", "ESP32-S3 · FreeRTOS · LVGL · C++17"),
    ("选做平台", "PC 管理端 Qt 5 + SQLite"),
]
for i, (k, v) in enumerate(facts):
    yy = 2.26 + i * 0.52
    tb(s, 7.05, yy, 1.35, 0.30, k, 11, BLUE, True)
    tb(s, 8.42, yy, 4.10, 0.30, v, 11, SUB)
    if i < len(facts) - 1:
        rect(s, 7.05, yy + 0.39, 5.45, 0.01, LINE)

rect(s, 0.55, 5.82, 12.23, 1.02, CARD2, rounded=True)
card_title(s, 0.80, 5.98, 11.7, "预期学习成果", accent=AMBER, size=12.5)
tb(s, 0.82, 6.38, 11.7, 0.34,
   "搭建四层工程骨架  ·  编写传感器驱动与 ID 自检  ·  调通计步与久坐闭环  ·  "
   "完成 LVGL 六页交互  ·  执行联调测试与故障演练  ·  产出文档并答辩",
   11.5, TXT)

# ============================================================ 04 核心闭环
s = new_slide(prs, 4, "项目概述：核心闭环与范围", "一条主线：察觉 → 提醒 → 反馈（P0 必做）")
steps = [
    ("察 觉", "IMU 50 Hz 采样 + SVM 合加速度\n佩戴判定 / 清醒时段 / 连续无有效活动", TEAL),
    ("提 醒", "连续 45 min 达阈值 → 马达 2 s × 3\n+ 全屏弹窗（已起身 / 稍后 10 min）", AMBER),
    ("反 馈", "点“已起身”或 5 min 内 ≥ 30 步 → 清零\n稍后 10 min 再提醒，最多 3 次", RED),
]
for i, (t, d, c) in enumerate(steps):
    x = 0.55 + i * 4.15
    rect(s, x, 1.62, 3.92, 1.92, CARD, rounded=True)
    rect(s, x, 1.62, 3.92, 0.06, c)
    tb(s, x + 0.26, 1.86, 3.4, 0.44, t, 19, c, True)
    tb(s, x + 0.26, 2.40, 3.42, 0.95, d, 11.3, SUB, spacing=1.30)
    if i < 2:
        tb(s, x + 3.94, 2.30, 0.22, 0.40, "›", 22, DIM, True,
           align=PP_ALIGN.CENTER)
rect(s, 0.55, 3.72, 12.23, 0.52, CARD2, rounded=True)
tb(s, 0.80, 3.72, 11.8, 0.52,
   "阈值 30 / 45 / 60 min 设置页可调，写入 NVS 掉电保存；仅“佩戴中 + 非免打扰 + 连续无有效活动”三条件同时满足才累加计时。",
   11.5, TXT, anchor=MSO_ANCHOR.MIDDLE)

rect(s, 0.55, 4.48, 5.95, 1.50, CARD, rounded=True)
card_title(s, 0.80, 4.66, 5.4, "必做范围（P0）")
bullets(s, 0.82, 5.08, 5.35, 0.85, [
    "运动监测：计步与活动量统计",
    "久坐提醒：状态机 + 震动 + 交互",
    "表盘与提醒交互：可触可达可见",
], size=11.5, gap=3)
rect(s, 6.78, 4.48, 6.00, 1.50, CARD, rounded=True)
card_title(s, 7.03, 4.62, 5.4, "选做范围（P1 / P2 · 含新增）", accent=BLUE)
bullets(s, 7.05, 4.98, 5.40, 1.00, [
    "生理 / 联网：心率血氧 / 跌倒 / Wi-Fi / PC 端",
    "AI 选做：PC 大模型问答 + 基于归档数据 RAG（附来源）",
    "HID 选做：手环操控鼠标（默认关 / 二次确认 / 可断开）",
], size=10.0, gap=3, dot=BLUE)

rect(s, 0.55, 6.18, 12.23, 0.62, CARD2, rounded=True)
tb(s, 0.78, 6.18, 1.75, 0.62, "四 条 底 线", 12, AMBER, True,
   anchor=MSO_ANCHOR.MIDDLE)
for i, t in enumerate(["提醒能触发", "数据能显示", "断电不丢设置", "连续运行不死机"]):
    tb(s, 2.70 + i * 2.50, 6.18, 2.40, 0.62, "✓  " + t, 11.5, TXT,
       anchor=MSO_ANCHOR.MIDDLE)

# ============================================================ 05 功能介绍
s = new_slide(prs, 5, "功能介绍", "五组功能 · 六个页面 · F-01～F-16 与实训阶段映射（表 2-1 / 2-4），新增三项选做见 FR-19～21")
groups = [("运动感知", TEAL), ("生理监测", RED), ("交互显示", BLUE),
          ("提醒与电源", AMBER), ("PC 管理（选做）", DIM)]
for i, (g, c) in enumerate(groups):
    chip(s, 0.55 + i * 2.48, 1.56, 2.32, 0.42, g, color=c, size=11)

card_title(s, 0.55, 2.16, 7.0, "终端页面与交互（表 2-1）")
table(s, 0.55, 2.54, 7.55, 3.90, [
    ["页面", "主内容", "关键交互"],
    ["表盘页", "时间 / 步数 / 电量 / 佩戴状态", "短按翻页，长按 2 s 熄屏"],
    ["运动页", "今日步数 / 活动时长 / 目标环", "实时刷新，1 s 节流"],
    ["心率血氧页", "HR / SpO₂ / PI 与信号质量", "开始测量，低灌注提示"],
    ["体温页", "TMP117 滑动平均值", "单位与趋势显示"],
    ["久坐页", "倒计时 / 今日提醒次数", "阈值 30 / 45 / 60 min 可调"],
    ["设置页", "阈值 / 免打扰 / 恢复出厂", "写 NVS，掉电保存"],
    ["提醒弹窗", "全屏 + 马达 2 s × 3", "“已起身” / “稍后 10 min”"],
], col_w=[1.35, 3.45, 3.05], fs=10.2, hs=10.5)

rect(s, 8.36, 2.16, 4.42, 4.28, CARD, rounded=True)
card_title(s, 8.60, 2.34, 4.0, "功能清单映射（表 2-4）", accent=BLUE)
tb(s, 8.62, 2.78, 3.95, 0.30, "F-01 ～ F-16 · 按优先级分期", 11, SUB)
for i, (p, name, desc, c) in enumerate([
    ("P0", "必做 · 核心闭环", "计步 / 久坐状态机 / 表盘 / 提醒交互 / 设置 NVS", TEAL),
    ("P1", "按能力 · 进阶", "心率血氧 / 体温 / 跌倒 / 六页完整界面 / 存储", BLUE),
    ("P2", "加分 · 扩展", "Wi-Fi 同步 / PC 端 / 新增：大模型问答 + RAG / 手环鼠标（FR-19～21）", AMBER),
]):
    yy = 3.20 + i * 1.02
    rect(s, 8.62, yy, 3.92, 0.90, CARD2, rounded=True)
    tb(s, 8.80, yy + 0.10, 0.60, 0.28, p, 13, c, True)
    tb(s, 9.42, yy + 0.08, 2.95, 0.28, name, 11.5, TXT, True)
    tb(s, 8.80, yy + 0.44, 3.55, 0.40, desc, 10.2, SUB, spacing=1.12)

# ============================================================ 06 需求与约束
s = new_slide(prs, 6, "需求分析与约束", "功能需求 FR-01～FR-21（新增 FR-19～21 选做）· 非功能 NFR-01～12 · 硬件形态取舍")
rect(s, 0.55, 1.58, 5.95, 3.30, CARD, rounded=True)
card_title(s, 0.80, 1.78, 5.4, "功能性需求（分组归纳）")
bullets(s, 0.82, 2.22, 5.35, 2.50, [
    ("采集：", "IMU 50 Hz、PPG FIFO、体温 1 s，三传感器 ID 自检。"),
    ("算法：", "计步与活动量、久坐状态机、心率血氧与跌倒判定。"),
    ("交互：", "六页流转 + 全屏提醒弹窗 + 长短按与触摸。"),
    ("数据：", "分钟聚合写盘、NVS 设置、断点续传与 Wi-Fi 同步。"),
    ("安全：", "看门狗、I2C NACK 恢复、免打扰时段与佩戴判定。"),
    ("新增选做：", "FR-19 大模型问答 / FR-20 RAG 检索 / FR-21 手环鼠标（P2 加分）。"),
], size=11.0, gap=7)

rect(s, 6.78, 1.58, 6.00, 3.30, CARD, rounded=True)
card_title(s, 7.03, 1.78, 5.4, "非功能需求与量化约束", accent=BLUE)
bullets(s, 7.05, 2.22, 5.40, 2.50, [
    ("精度：", "100 步误计率 < 3%、漏计率 < 5%；45 min 误差 ≤ 10 s。"),
    ("稳定性：", "连续 8 h 无死机；断电不丢设置；断网 2 h 补传无重复。"),
    ("可用性：", "触摸目标 ≥ 48 px、按钮 ≥ 52 px，90 s 内完成核心操作。"),
    ("生理指标：", "心率 ≤ ±5 bpm，SpO₂ ≤ ±3%（选做标定后）。"),
    ("可维护性：", "四层解耦、越层检查三条、驱动可替换。"),
], size=11.5, gap=9, dot=BLUE)

rect(s, 0.55, 5.06, 12.23, 1.78, CARD2, rounded=True)
card_title(s, 0.80, 5.22, 11.7, "硬件形态取舍：为教学可视化让路", accent=AMBER, size=13)
取舍 = [
    ("4.3″ 800×480 屏", "表体 110×70×12 mm，≤ 85 g，24 mm 表带"),
    ("USB-C 供电", "不配电池、不做低功耗 / 休眠"),
    ("CPU 固定 240 MHz", "电量固定显示 100%，熄屏只关背光"),
]
for i, (a, b) in enumerate(取舍):
    x = 0.82 + i * 3.98
    tb(s, x, 5.68, 3.70, 0.30, a, 12, AMBER, True)
    tb(s, x, 6.04, 3.70, 0.60, b, 10.5, SUB, spacing=1.18)
tb(s, 0.82, 6.62, 11.7, 0.26,
   "取舍理由：把“可观测、可调试、可复现”放在功耗与形态之前，降低实训门槛；同时用局部刷新缓解 SPI 带宽压力。",
   10.5, DIM)

# ============================================================ 07 总体架构
s = new_slide(prs, 7, "总体架构", "终端四层 + 服务端四层（选做）· 依赖单向向下，越层可静态检查")
tb(s, 0.55, 1.52, 5.9, 0.30, "手环终端（ESP32-S3 / FreeRTOS）", 12.5, TEAL, True)
term = [
    ("应用层", "AppController · ScreenManager · 六个 Screen / 提醒弹窗", TEAL),
    ("服务层", "SensorHub · StepService · SedentaryService · VitalService\nFallService · WearService · RecordStore · SyncClient", BLUE),
    ("驱动层", "Mpu6050 / Max30102 / Tmp117 / Ft6236 / ST7262 寄存器封装", AMBER),
    ("HAL 层", "I2CBus（互斥锁保证多任务原子性）· SpiBus · Gpio · Ledc · Nvs", RED),
]
for i, (n, d, c) in enumerate(term):
    layer(s, 0.55, 1.90 + i * 0.90, 5.95, 0.78, n, d, c)

tb(s, 6.78, 1.52, 5.9, 0.30, "PC 服务端（选做 / Qt 5 + SQLite）", 12.5, BLUE, True)
serv = [
    ("接入层", "SyncServer TCP 监听 · ClientSession 拆包鉴权 · Dispatcher", BLUE),
    ("业务层", "UserService · RecordService · AlertService", TEAL),
    ("存储层", "SQLite WAL 模式，(device_id, ts) 唯一键去重", AMBER),
    ("展示层", "Qt Widgets + QSqlTableModel + Qt Charts 曲线", RED),
]
for i, (n, d, c) in enumerate(serv):
    layer(s, 6.78, 1.90 + i * 0.90, 6.00, 0.78, n, d, c)

rect(s, 0.55, 5.62, 7.55, 1.30, CARD, rounded=True)
card_title(s, 0.80, 5.78, 7.1, "越层检查三条（可脚本化）", accent=AMBER, size=12.5)
bullets(s, 0.82, 6.14, 7.00, 0.70, [
    "ui/ 目录不得出现 i2c_master_ 开头 API；driver/ 不得出现 lv_ 开头 API；service/ 不直接操作 GPIO。",
], size=10.8, gap=2)
rect(s, 8.36, 5.62, 4.42, 1.30, CARD, rounded=True)
card_title(s, 8.60, 5.78, 4.0, "部署拓扑", accent=BLUE, size=12.5)
tb(s, 8.62, 6.12, 3.95, 0.72,
   "教师机固定 192.168.10.10:9527\n2.4 GHz Wi-Fi · deviceId HB-2026-G03-01\n手环主动连接，密码不明文存 NVS",
   10.5, SUB, spacing=1.22)

# ============================================================ 08 硬件设计
s = new_slide(prs, 8, "硬件设计", "表 3-6～3-13 / 附录 A：选型、引脚、供电与装调验收")
card_title(s, 0.55, 1.56, 7.4, "核心器件与接口（表 3-7 / 表 A-1）")
table(s, 0.55, 1.94, 7.55, 3.34, [
    ["器件", "型号 / 接口", "地址 / 引脚", "用途"],
    ["主控", "ESP32-S3 N16R8", "双核 240 MHz", "16 MB Flash / 8 MB PSRAM"],
    ["IMU", "MPU6050 · I2C", "0x68", "三轴加速度，计步"],
    ["PPG", "MAX30102 · I2C", "0x57", "心率 / 血氧（选做）"],
    ["温度", "TMP117 · I2C", "0x48", "体表温度（选做）"],
    ["触摸", "FT6236 · I2C", "0x38", "电容触摸"],
    ["显示", "4.3″ IPS 800×480 · SPI", "ST7262 / 80 MHz DMA", "LVGL 六页"],
    ["执行", "按键 GPIO4 / 马达 GPIO5", "4 kHz PWM", "翻页 / 熄屏 / 震动"],
], col_w=[1.05, 2.55, 1.85, 2.45], fs=10.0, hs=10.3)

rect(s, 8.36, 1.56, 4.42, 3.72, CARD, rounded=True)
card_title(s, 8.60, 1.74, 4.0, "引脚总表（表 3-11）", accent=BLUE)
tb(s, 8.62, 2.18, 3.95, 2.95,
   "I2C    SDA GPIO8   SCL GPIO9   400 kHz\n"
   "SPI    SCK GPIO12  MOSI GPIO11\n"
   "       CS GPIO10   DC  GPIO14\n"
   "       RST GPIO15  BL  GPIO16\n"
   "触摸   INT GPIO17  RST GPIO18\n"
   "按键   KEY GPIO4（10 kΩ 上拉 + 100 nF）\n"
   "马达   PWM GPIO5（2N7002 + 1N4148）\n"
   "USB    D− GPIO19   D+ GPIO20\n\n"
   "原则：总线集中、模拟分离、启动引脚留空\n"
   "避开 GPIO0/3/45/46、33–37、19/20、43/44",
   10.6, SUB, spacing=1.35)

rect(s, 0.55, 5.48, 12.23, 1.44, CARD2, rounded=True)
card_title(s, 0.80, 5.62, 11.7, "供电、结构与装调验收", accent=AMBER, size=12.5)
items = [
    ("供电", "USB-C VBUS 加 1 A 保险丝；先确认 3V3 与 GND"),
    ("电流预算", "按表 3-10 逐模块核算，屏背光与马达为峰值大户"),
    ("结构", "3D 打印外壳 + 24 mm 硅胶表带，≤ 85 g"),
    ("排障顺序", "3.3 V 与 GND → 总线波形 → 最后怀疑代码"),
]
for i, (a, b) in enumerate(items):
    x = 0.82 + i * 3.02
    tb(s, x, 6.04, 2.85, 0.28, a, 11.5, AMBER, True)
    tb(s, x, 6.36, 2.85, 0.50, b, 10.2, SUB, spacing=1.15)

# ============================================================ 09 软件设计
s = new_slide(prs, 9, "软件设计：FreeRTOS 任务划分", "表 3-15 · 7 个任务，双核分工，采集与算法分离")
card_title(s, 0.55, 1.56, 8.0, "任务表（优先级 / 核 / 栈 / 周期 / 职责）")
table(s, 0.55, 1.94, 8.35, 4.10, [
    ["任务", "优先", "核", "栈", "周期", "职责"],
    ["imu_task", "5", "1", "6 KB", "20 ms", "读 MPU6050，发布 ImuFrame"],
    ["ppg_task", "4", "1", "6 KB", "10 ms", "测量时读 FIFO，发布 PpgSample"],
    ["temp_task", "2", "1", "4 KB", "1 s", "读 TMP117 并滑动平均"],
    ["ui_task", "3", "0", "8 KB", "16–33 ms", "LVGL 事务与页面刷新"],
    ["record_task", "1", "0", "4 KB", "60 s", "分钟聚合写 LittleFS"],
    ["sync_task", "2", "0", "6 KB", "5 min", "Wi-Fi 同步与游标推进"],
    ["button_task", "3", "0", "3 KB", "10 ms", "消抖与长短按识别"],
], col_w=[1.55, 0.75, 0.60, 0.85, 1.15, 3.45], fs=10.3, hs=10.3)

rect(s, 9.16, 1.56, 3.62, 4.48, CARD, rounded=True)
card_title(s, 9.40, 1.74, 3.2, "并发三不要", accent=RED, size=13)
bullets(s, 9.42, 2.22, 3.10, 2.10, [
    "不在中断里调 LVGL 或动态分配内存",
    "不持 I2C 锁等待队列",
    "不在 UI 任务做阻塞式文件写入",
], size=11.2, gap=12, dot=RED)
card_title(s, 9.40, 4.10, 3.2, "数据流原则", accent=TEAL, size=13)
bullets(s, 9.42, 4.52, 3.10, 1.35, [
    "队列削峰，采集与算法分离",
    "跨任务传值不传指针",
    "UI 1 s 节流，每分钟聚合写盘",
], size=11.2, gap=9)

rect(s, 0.55, 6.20, 12.23, 0.62, CARD2, rounded=True)
tb(s, 0.80, 6.20, 11.8, 0.62,
   "驱动自检 ID：MPU6050 WHO_AM_I = 0x68 · MAX30102 Part ID = 0x15 · TMP117 Device ID = 0x0117（上电先自检再入网）",
   11.5, TXT, anchor=MSO_ANCHOR.MIDDLE)

# ============================================================ 10 计步算法
s = new_slide(prs, 10, "核心算法 ①：计步", "从 50 Hz 原始加速度到一步：动态阈值 + 双重约束 + 震动屏蔽")
pipe = [
    ("1", "采样", "IMU 50 Hz"), ("2", "预处理", "零偏校准 + 一阶低通"),
    ("3", "合成", "SVM 合加速度，消除方向影响"), ("4", "滑窗", "5 s 窗口去均值"),
    ("5", "阈值", "μ + 0.8σ 动态阈值"), ("6", "候选峰", "SVM 上穿记峰"),
    ("7", "判定", "峰谷差 ≥ 0.6 g 且间隔 250–2000 ms"),
]
for i, (n, t, d) in enumerate(pipe):
    x = 0.55 + i * 1.77
    rect(s, x, 1.62, 1.62, 1.55, CARD, rounded=True)
    chip(s, x + 0.12, 1.76, 0.36, 0.32, n, color=TEAL, size=10, fill=CARD2)
    tb(s, x + 0.12, 2.16, 1.40, 0.30, t, 12, TXT, True)
    tb(s, x + 0.12, 2.50, 1.40, 0.62, d, 9.6, SUB, spacing=1.12)
    if i < 6:
        tb(s, x + 1.63, 2.18, 0.14, 0.30, "›", 15, DIM, True,
           align=PP_ALIGN.CENTER)

rect(s, 0.55, 3.42, 5.95, 2.10, CARD, rounded=True)
card_title(s, 0.80, 3.60, 5.4, "关键参数（教学可复现）")
bullets(s, 0.82, 4.02, 5.35, 1.40, [
    ("动态阈值：", "μ + 0.8σ，随窗口自适应，不写死幅值。"),
    ("双约束：", "峰谷差 ≥ 0.6 g；与上峰间隔 250–2000 ms。"),
    ("抗干扰：", "马达震动期间屏蔽 300 ms，防“震”出步。"),
], size=11.3, gap=8)

rect(s, 6.78, 3.42, 6.00, 2.10, CARD, rounded=True)
card_title(s, 7.03, 3.60, 5.4, "采样率取舍：课堂对比实验", accent=BLUE)
table(s, 7.03, 4.00, 5.50, 1.32, [
    ["采样率", "现象", "结论"],
    ["25 Hz", "快走峰值被削平，漏计明显", "不可用"],
    ["50 Hz", "覆盖 0.5–5 Hz 步态主频，I2C 不过载", "采用"],
    ["100 Hz", "精度提升有限，功耗与总线占用翻倍", "不推荐"],
], col_w=[1.00, 3.30, 1.20], fs=10.0, hs=10.2, head_color=BLUE)

rect(s, 0.55, 5.72, 12.23, 1.20, CARD2, rounded=True)
card_title(s, 0.80, 5.86, 11.7, "流水线设计（对应四层架构）", accent=AMBER, size=12.5)
tb(s, 0.82, 6.24, 11.7, 0.58,
   "采集与算法用队列解耦削峰 → 跨任务传值不传指针 → UI 1 s 节流刷新 → RecordStore 每分钟聚合写盘；"
   "数据回放可脱离硬件复现调参过程。",
   11.2, TXT, spacing=1.22)

# ============================================================ 11 久坐状态机
s = new_slide(prs, 11, "核心算法 ②：久坐提醒状态机", "表 3-4 · 六态迁移，把“提醒”做成可验证的工程闭环")
states = [("WearOff", DIM), ("Active", TEAL), ("Sitting", BLUE),
          ("Alerting", RED), ("Snoozed", AMBER), ("Quiet", DIM)]
for i, (st, c) in enumerate(states):
    chip(s, 0.55 + i * 2.06, 1.56, 1.90, 0.46, st, color=c, size=11)
tb(s, 0.55, 2.10, 12.2, 0.28,
   "六态：未佩戴 → 活动 → 久坐 → 提醒中 → 稍后 → 免打扰", 11, SUB)

card_title(s, 0.55, 2.48, 8.0, "状态迁移表（表 3-4 精简）")
table(s, 0.55, 2.86, 8.35, 3.30, [
    ["迁移", "触发条件", "动作"],
    ["WearOff → Active", "PPG 直流恢复 + 姿态稳定", "开始计时"],
    ["Active → Sitting", "连续 45 min 无有效活动", "启动久坐计时，更新倒计时"],
    ["Sitting → Active", "检测到行走 / 抬腕", "清零计时"],
    ["Sitting → Alerting", "计时达阈值", "马达 2 s × 3 + 全屏提醒"],
    ["Alerting → Active", "“已起身”或 5 min 内 ≥ 30 步", "清零并记录"],
    ["Alerting → Snoozed", "“稍后”或 60 s 无操作", "10 min 后再提醒，最多 3 次"],
    ["Snoozed → Alerting", "10 min 到期", "再次震动 + 弹窗"],
], col_w=[2.35, 3.30, 2.70], fs=10.2, hs=10.4)

rect(s, 9.16, 2.48, 3.62, 3.68, CARD, rounded=True)
card_title(s, 9.40, 2.64, 3.2, "计时前提", accent=AMBER, size=13)
tb(s, 9.42, 3.06, 3.10, 1.20,
   "仅当三条件同时满足才累加：\n① 佩戴中\n② 非免打扰时段\n③ 连续无有效活动",
   11.2, SUB, spacing=1.40)
card_title(s, 9.40, 4.30, 3.2, "可演示性", accent=BLUE, size=13)
tb(s, 9.42, 4.70, 3.10, 1.30,
   "DemoPolicy：60 s 触发，便于课堂演示；必须显式标注，避免与 StandardPolicy 混淆。",
   11.2, SUB, spacing=1.35)

rect(s, 0.55, 6.32, 12.23, 0.60, CARD2, rounded=True)
tb(s, 0.80, 6.32, 11.8, 0.60,
   "策略模式解耦判定：virtual bool shouldAlert(const ActivityWindow& w, const Settings& s) const；"
   "StandardPolicy → w.sittingMinutes >= s.thresholdMin && w.stepsInLast5Min < 30",
   11.0, TXT, anchor=MSO_ANCHOR.MIDDLE)

# ============================================================ 12 六个设计模式
s = new_slide(prs, 12, "实现方案：六个设计模式", "第四章 · 每个模式都对应一个真实工程痛点，并给出 C++ 代码与教学要点")
patterns = [
    ("单例", "I2CBus 全局唯一 + 互斥锁", "多任务共享总线，保证 I2C 事务原子性", TEAL),
    ("简单工厂", "SensorFactory::create(type)", "switch 创建 Mpu6050 / Max30102 / Tmp117", BLUE),
    ("策略", "SedentaryPolicy 接口", "Standard / Demo 两套提醒判定可替换", AMBER),
    ("状态", "enum class SitState", "onEvent + transition 收敛 if-else 分支", RED),
    ("观察者", "EventBus 事件分发", "switch(EventType) 解耦发布与订阅", TEAL),
    ("外观", "HealthBand 门面", "对外统一 start()/loop()，隐藏四层装配", BLUE),
]
for i, (n, impl, why, c) in enumerate(patterns):
    col, row = i % 3, i // 3
    x, y = 0.55 + col * 4.15, 1.62 + row * 2.05
    rect(s, x, y, 3.92, 1.82, CARD, rounded=True)
    rect(s, x, y, 0.055, 1.82, c)
    tb(s, x + 0.26, y + 0.18, 3.4, 0.34, n, 16, c, True)
    tb(s, x + 0.26, y + 0.62, 3.40, 0.34, impl, 10.8, TXT, True)
    tb(s, x + 0.26, y + 1.02, 3.40, 0.66, why, 10.5, SUB, spacing=1.20)

rect(s, 0.55, 5.86, 12.23, 1.06, CARD2, rounded=True)
card_title(s, 0.80, 5.98, 11.7, "教学要点：模式不是目的，可替换性才是", accent=AMBER, size=12.5)
tb(s, 0.82, 6.34, 11.7, 0.48,
   "换 MPU6050 为 ICM-42688 只改驱动 + 工厂（SRP/OCP）· 新增传感器需改工厂 switch，是 OCP 的已知缺口 · "
   "MockSensor 可替换真实传感器做单元测试（LSP/DIP）",
   11.0, TXT)

# ============================================================ 13 存储/同步/可靠性
s = new_slide(prs, 13, "存储、同步与可靠性", "16 字节定长记录 · 断点续传 · 看门狗与降级策略")
cols = [
    ("本地存储", TEAL, [
        "16 字节定长 MinuteRecord（小端）",
        "Unix 秒 UTC / 温度 0.01 ℃ 整数",
        "LittleFS 环形缓冲，先数据后游标",
        "NVS 持久化阈值与设置",
    ]),
    ("通信与同步", BLUE, [
        "TCP + JSON Lines，每行 ≤ 64 KB",
        "type: AUTH / SYNC / ACK / PING / ALERT",
        "幂等唯一键 (device_id, ts) 去重",
        "指数退避 1 / 2 / 4 / 8 s；4xx 客户端 / 5xx 服务端",
    ]),
    ("可靠性设计", AMBER, [
        "任务看门狗 8 s 未喂 → 打印回溯并重启",
        "I2C 连续 NACK 3 次 → 总线重初始化",
        "屏初始化失败 → 进最小系统模式",
        "断网 2 h 补传无重复；五类故障可恢复",
    ]),
]
for i, (t, c, its) in enumerate(cols):
    x = 0.55 + i * 4.15
    rect(s, x, 1.62, 3.92, 3.24, CARD, rounded=True)
    rect(s, x, 1.62, 3.92, 0.06, c)
    tb(s, x + 0.26, 1.88, 3.4, 0.38, t, 15, c, True)
    bullets(s, x + 0.26, 2.42, 3.40, 2.30, its, size=11.0, gap=10, dot=c)

rect(s, 0.55, 5.06, 12.23, 1.86, CARD2, rounded=True)
card_title(s, 0.80, 5.20, 11.7, "附录 B · 16 字节分钟记录布局（表 B-1）", accent=RED, size=12.5)
fields = [("ts", "4 B"), ("steps", "2 B"), ("active_s", "2 B"),
          ("hr", "1 B"), ("spo2", "1 B"), ("temp", "2 B"),
          ("flags", "1 B"), ("rsv", "3 B")]
for i, (n, b) in enumerate(fields):
    x = 0.82 + i * 1.49
    rect(s, x, 5.66, 1.38, 0.62, CARD, rounded=True)
    tb(s, x, 5.70, 1.38, 0.28, n, 10.5, RED, True, align=PP_ALIGN.CENTER)
    tb(s, x, 5.98, 1.38, 0.26, b, 10, SUB, align=PP_ALIGN.CENTER)
tb(s, 0.82, 6.44, 11.7, 0.34,
   "合计 16 B；“先数据后游标”保证断电时最多丢一条未确认记录，不破坏已确认数据。",
   11.0, TXT)

# ============================================================ 14 界面规范
s = new_slide(prs, 14, "界面与交互规范", "表 3-16 · 深色主题、可触达、局部刷新")
rect(s, 0.55, 1.58, 5.95, 2.62, CARD, rounded=True)
card_title(s, 0.80, 1.76, 5.4, "配色令牌")
sw = [
    ("背景", "#0B1220", BG), ("卡片", "#152238", CARD),
    ("主文字", "#F8FAFC", TXT), ("次要文字", "#94A3B8", SUB),
    ("运动 · 青绿", "#2DD4BF", TEAL), ("心率 · 红", "#FF5A6E", RED),
    ("体温/久坐 · 琥珀", "#FBBF24", AMBER),
]
for i, (n, hexv, c) in enumerate(sw):
    col, row = i % 2, i // 2
    x, y = 0.82 + col * 2.78, 2.22 + row * 0.47
    rect(s, x, y + 0.03, 0.32, 0.30, c, rounded=True)
    tb(s, x + 0.42, y + 0.02, 1.35, 0.30, n, 10.5, TXT)
    tb(s, x + 1.78, y + 0.02, 0.90, 0.30, hexv, 10, SUB)

rect(s, 6.78, 1.58, 6.00, 2.62, CARD, rounded=True)
card_title(s, 7.03, 1.76, 5.4, "可用性与排版", accent=BLUE)
bullets(s, 7.05, 2.20, 5.40, 1.90, [
    ("可触达：", "触摸目标 ≥ 48 px，按钮 ≥ 52 px。"),
    ("字号：", "正文 18–20 px，关键数值 40–112 px。"),
    ("动效：", "页面切换与反馈动画 180 ms。"),
    ("布局：", "六页 = 表盘 / 运动 / 心率血氧 / 体温 / 久坐 / 设置 + 提醒弹窗。"),
], size=11.3, gap=9, dot=BLUE)

rect(s, 0.55, 4.40, 12.23, 2.52, CARD2, rounded=True)
card_title(s, 0.80, 4.56, 11.7, "显示带宽约束与局部刷新", accent=AMBER, size=13)
tb(s, 0.82, 5.00, 11.7, 0.62,
   "4.3″ 800×480 RGB565 全屏帧缓冲 = 768 KB；SPI 80 MHz DMA 下整屏刷新约 90–110 ms（< 11 fps），"
   "直接全屏刷新会拖垮 ui_task 的 16–33 ms 节拍。",
   11.5, TXT, spacing=1.25)
rect(s, 0.82, 5.78, 5.55, 0.82, CARD, rounded=True)
tb(s, 1.02, 5.88, 5.15, 0.62,
   "方案：2 × 800 × 40 行局部缓冲 + DMA\n只重绘数值区与变化卡片",
   11.2, TEAL, True, spacing=1.22, anchor=MSO_ANCHOR.MIDDLE)
rect(s, 6.60, 5.78, 5.90, 0.82, CARD, rounded=True)
tb(s, 6.80, 5.88, 5.50, 0.62,
   "背光：LEDC 20 kHz PWM 避开可听噪声\n熄屏只关背光，不进低功耗 / 休眠",
   11.2, BLUE, True, spacing=1.22, anchor=MSO_ANCHOR.MIDDLE)

# ============================================================ 15 测试与验收
s = new_slide(prs, 15, "测试方案与验收指标", "表 4-1 / 4-2 · T-01～T-10 关键用例 · 可量化、可复现")
rect(s, 0.55, 1.58, 5.95, 4.20, CARD, rounded=True)
card_title(s, 0.80, 1.76, 5.4, "测试类型与关键用例")
bullets(s, 0.82, 2.20, 5.35, 2.00, [
    ("单元：", "计步峰值判定、久坐计时、JSON 编解码。"),
    ("集成：", "I2C 三器件共存与总线互斥、队列传值。"),
    ("系统：", "六页流转 + 提醒闭环 + 断电重启。"),
    ("回归 / 验收：", "T-01～T-10 覆盖计步精度、45 min 计时、断网补传。"),
], size=11.5, gap=10)
card_title(s, 0.80, 4.36, 5.4, "故障演练（五类）", accent=RED, size=12.5)
tb(s, 0.82, 4.76, 5.35, 0.88,
   "断电 · 断网 · 文件损坏 · 传感器 NACK · 3 台 Wi-Fi 干扰\n均要求可恢复且数据不重复、设置不丢失。",
   11.2, SUB, spacing=1.32)

metrics = [
    ("< 3 %", "100 步误计率", TEAL), ("< 5 %", "100 步漏计率", TEAL),
    ("≤ 10 s", "45 min 计时误差", BLUE), ("8 h", "连续运行无死机", BLUE),
    ("± 5 bpm", "心率误差（标定后）", AMBER), ("± 3 %", "SpO₂ 误差", AMBER),
    ("90 s", "非专业用户完成核心操作", RED), ("2 h", "断网补传无重复", RED),
]
for i, (v, n, c) in enumerate(metrics):
    col, row = i % 2, i // 2
    x, y = 6.78 + col * 3.05, 1.58 + row * 1.05
    rect(s, x, y, 2.88, 0.90, CARD, rounded=True)
    rect(s, x, y, 0.05, 0.90, c)
    tb(s, x + 0.20, y + 0.10, 2.55, 0.36, v, 17, c, True)
    tb(s, x + 0.20, y + 0.50, 2.55, 0.30, n, 10.2, SUB)

rect(s, 0.55, 5.98, 12.23, 0.84, CARD2, rounded=True)
tb(s, 0.80, 5.98, 11.8, 0.84,
   "验收口径：每条指标都有可复现的操作步骤与判定阈值；答辩现场可随机抽取用例重跑，"
   "而非只演示“理想路径”。",
   11.5, TXT, anchor=MSO_ANCHOR.MIDDLE)

# ============================================================ 16 课程设计
s = new_slide(prs, 16, "实训课程设计", "表 5-2 80 学时分配 · 表 5-4 十天计划 · 表 5-3 角色分工")
rect(s, 0.55, 1.58, 6.20, 4.36, CARD, rounded=True)
card_title(s, 0.80, 1.76, 5.7, "80 学时分配（P0～P7）")
phases = [
    ("P0 项目导入与需求分析", 4, TEAL),
    ("P1 硬件搭建与环境配置", 8, TEAL),
    ("P2 驱动与传感器", 12, BLUE),
    ("P3 计步与久坐", 14, AMBER),
    ("P4 完整界面与选做测量", 12, BLUE),
    ("P5 交互完善与功能验证", 12, BLUE),
    ("P6 联调测试与部署", 10, RED),
    ("P7 答辩互评与总结", 8, RED),
]
for i, (n, v, c) in enumerate(phases):
    hbar(s, 0.82, 2.24 + i * 0.44, 5.65, n, v, 14, c, lw=2.42)
rect(s, 0.82, 5.82, 5.65, 0.012, LINE)
tb(s, 0.82, 5.88, 5.65, 0.28, "合计 80 学时 = 讲授 + 讨论 / 实操 / 实操 + 评审 / 答辩",
   10.5, SUB)

rect(s, 6.98, 1.58, 5.80, 4.36, CARD, rounded=True)
card_title(s, 7.23, 1.76, 5.3, "十天推进计划（表 5-4）", accent=BLUE)
days = [
    ("D1", "P0 导入 / P1 点屏", "能烧录能点屏"),
    ("D2", "P1 接线扫描 / P2 I2C 封装", "三传感器识别"),
    ("D3", "P2 MAX30102 + TMP117", "驱动验收"),
    ("D4", "P3 计步调参 / 久坐状态机", "计步闭环"),
    ("D5", "P3 提醒弹窗 / 测试用例", "久坐闭环"),
    ("D6", "P4 界面框架 / 表盘运动久坐页", "主页面可用"),
    ("D7", "P4 六页 / 可用性测试", "界面验收"),
    ("D8", "P5 NVS 存储 / 看门狗", "可独立运行"),
    ("D9", "P6 故障演练 / 测试报告", "测试验收"),
    ("D10", "P7 答辩互评 / 改进清单", "最终交付"),
]
for i, (d, t, g) in enumerate(days):
    yy = 2.22 + i * 0.365
    tb(s, 7.25, yy, 0.48, 0.28, d, 10.5, TEAL, True)
    tb(s, 7.78, yy, 3.30, 0.28, t, 10.2, TXT)
    tb(s, 11.12, yy, 1.42, 0.28, g, 9.8, SUB)

rect(s, 0.55, 6.10, 12.23, 0.82, CARD2, rounded=True)
tb(s, 0.80, 6.10, 11.8, 0.82,
   "小组角色（3～4 人）：组长 · 硬件负责人 · 算法负责人 · 界面负责人 · 测试/文档（0～1）；"
   "Git 规范按阶段打 tag，阶段验收通过才进入下一阶段。",
   11.3, TXT, anchor=MSO_ANCHOR.MIDDLE)

# ============================================================ 17 考核评价
s = new_slide(prs, 17, "考核评价", "表 5-7 总评构成 · 表 5-8 评分细则（70 分）· 过程与结果并重")
scores = [("40", "功能实现", "必做功能 / 精度 / 稳定性", "教师 + 助教", TEAL),
          ("30", "过程评价", "站会 / 提交 / 代码评审 / 阶段验收", "教师 + 组内互评", BLUE),
          ("15", "文档", "需求 / 设计 / 测试 / 部署手册", "教师", AMBER),
          ("15", "答辩", "演示 / 讲解 / 问答 / 改进清单", "教师 + 组间互评", RED)]
for i, (v, n, d, who, c) in enumerate(scores):
    x = 0.55 + i * 3.10
    rect(s, x, 1.62, 2.92, 2.02, CARD, rounded=True)
    rect(s, x, 1.62, 2.92, 0.06, c)
    tb(s, x + 0.22, 1.86, 2.5, 0.62, v, 34, c, True)
    tb(s, x + 0.22, 2.54, 2.5, 0.32, n, 13, TXT, True)
    tb(s, x + 0.22, 2.92, 2.50, 0.42, d, 10.2, SUB, spacing=1.15)
    tb(s, x + 0.22, 3.36, 2.50, 0.24, who, 9.8, DIM)

rect(s, 0.55, 3.86, 12.23, 0.34, ROW_A, rounded=True)
cum = 0
segs = [(40, TEAL), (30, BLUE), (15, AMBER), (15, RED)]
for v, c in segs:
    w = 12.23 * v / 100.0
    rect(s, 0.55 + 12.23 * cum / 100.0, 3.86, w, 0.34, c)
    cum += v
tb(s, 0.55, 4.28, 12.23, 0.28, "0", 9.5, DIM)
tb(s, 11.90, 4.28, 0.88, 0.28, "100 分", 9.5, DIM, align=PP_ALIGN.RIGHT)

card_title(s, 0.55, 4.66, 11.7, "评分细则要点（表 5-8，合计 70 分）", accent=AMBER, size=12.5)
table(s, 0.55, 5.04, 12.23, 1.88, [
    ["维度", "验收点", "分值", "维度", "验收点", "分值"],
    ["功能", "驱动与自检", "8", "文档", "需求 / 设计 / 测试", "5 + 5 + 5"],
    ["功能", "计步与久坐闭环", "12", "答辩", "演示 / 讲解 / 问答", "6 + 6 + 3"],
    ["功能", "界面 / 界面与选做测量 / 稳定性", "8 + 8 + 4", "说明", "过程评价 30 分另按站会与阶段验收计", "—"],
], col_w=[1.00, 4.30, 1.55, 1.00, 3.30, 1.08], fs=10.3, hs=10.4,
   align_left_cols=(0, 1, 3, 4))

# ============================================================ 18 设计原则与难易度
s = new_slide(prs, 18, "设计原则自评与难易度分析", "第六章 / 第八章 · 知道哪里做得好，也知道哪里还欠账")
rect(s, 0.55, 1.58, 6.20, 4.30, CARD, rounded=True)
card_title(s, 0.80, 1.76, 5.7, "七条设计原则逐条自评")
tb(s, 0.82, 2.18, 5.65, 1.30,
   "SRP 四层拆分 · OCP 策略/工厂/显示抽象 · LSP Screen 子类可替换\n"
   "DIP 依赖 IStore / PpgSensor 等抽象 · ISP 拆分读写与数据源接口\n"
   "合成复用 HealthBand 组合而非继承 · 迪米特 界面只与 AppController 交互",
   11.2, SUB, spacing=1.45)
card_title(s, 0.80, 3.60, 5.7, "五点不足与改进", accent=RED, size=12.5)
bullets(s, 0.82, 4.00, 5.60, 1.72, [
    "AppController 编排过多，需进一步拆分用例",
    "EventBus 同步 std::function 回调可能阻塞",
    "全局单例降低可测性",
    "同步协议无版本协商与加密（需 TLS + 签名）",
    "页面切换轻微掉帧（双缓冲 + DMA 改进）",
], size=10.8, gap=6, dot=RED)

rect(s, 6.98, 1.58, 5.80, 4.30, CARD, rounded=True)
card_title(s, 7.23, 1.76, 5.3, "难易度：整体“中等偏难”", accent=BLUE)
legend = [("● 必须通过", TEAL), ("◐ 按能力打分", AMBER), ("○ 加分", DIM)]
for i, (t, c) in enumerate(legend):
    chip(s, 7.25 + i * 1.83, 2.20, 1.68, 0.38, t, color=c, size=10)
bullets(s, 7.25, 2.78, 5.25, 1.60, [
    ("四大难点：", "计步调参 / 久坐边界 / PPG 信号质量 / 4.3″ 屏刷新带宽。"),
    ("分层教学：", "弱组只做 P0–P3；强组加 P4 PPG 精度、P5 连续运行、P6 PC 端同步。"),
    ("统一口径：", "无论哪一档，都必须通过四条底线。"),
], size=11.0, gap=9, dot=BLUE)
card_title(s, 7.23, 4.48, 5.3, "教学价值", accent=TEAL, size=12.5)
tb(s, 7.25, 4.88, 5.25, 0.84,
   "把“能跑—好用—可靠”拆成可验收的阶段产物，让不同基础的学生都能获得完整工程体验。",
   11.2, SUB, spacing=1.35)

rect(s, 0.55, 6.06, 12.23, 0.86, CARD2, rounded=True)
tb(s, 0.80, 6.06, 11.8, 0.86,
   "答辩追问重点：为什么选 50 Hz？为什么不做低功耗？断电时最多丢什么？"
   "四条底线如何逐条验证？——这些问题在文档中都有明确答案。",
   11.3, TXT, anchor=MSO_ANCHOR.MIDDLE)

# ============================================================ 19 总结与展望
s = new_slide(prs, 19, "总结与展望", "从课程设计到工程能力的迁移")
rect(s, 0.55, 1.58, 3.92, 2.30, CARD, rounded=True)
rect(s, 0.55, 1.58, 3.92, 0.06, TEAL)
tb(s, 0.82, 1.86, 3.4, 0.40, "项目价值", 16, TEAL, True)
tb(s, 0.82, 2.38, 3.40, 1.30,
   "每一行代码都有直接反馈：\n屏幕亮起、步数跳动、马达震动，\n而不是串口里的一句 success。",
   11.5, SUB, spacing=1.42)
rect(s, 4.70, 1.58, 3.92, 2.30, CARD, rounded=True)
rect(s, 4.70, 1.58, 3.92, 0.06, BLUE)
tb(s, 4.97, 1.86, 3.4, 0.40, "能力三台阶", 16, BLUE, True)
tb(s, 4.97, 2.38, 3.40, 1.30,
   "能跑 → 好用 → 可靠\n分层架构、量化验收、故障演练，\n把“做完”变成“做对”。",
   11.5, SUB, spacing=1.42)
rect(s, 8.85, 1.58, 3.93, 2.30, CARD, rounded=True)
rect(s, 8.85, 1.58, 3.93, 0.06, AMBER)
tb(s, 9.12, 1.86, 3.4, 0.40, "教学落地", 16, AMBER, True)
tb(s, 9.12, 2.38, 3.40, 1.30,
   "80 学时 · 10 天 · 8 阶段\n40 功能 + 30 过程 + 15 文档 + 15 答辩\n可裁剪、可分层、可复现。",
   11.5, SUB, spacing=1.42)

card_title(s, 0.55, 4.10, 11.7, "改进空间", accent=RED, size=13)
gaps = [
    ("形态", "4.3″ 屏偏离消费级手环，向小型化与低功耗演进"),
    ("精度", "心率血氧依赖个体标定；跌倒检测误报率待优化"),
    ("闭环", "PC 端归档后可选做大模型健康问答 + RAG 检索，形成“设备—数据—解释”链路"),
    ("工程", "同步协议补版本协商与加密；界面改双缓冲；鼠标映射选做默认关"),
]
for i, (k, v) in enumerate(gaps):
    col, row = i % 2, i // 2
    x, y = 0.55 + col * 6.20, 4.50 + row * 0.86
    rect(s, x, y, 5.95, 0.72, CARD, rounded=True)
    tb(s, x + 0.22, y + 0.10, 0.85, 0.28, k, 11.5, RED, True)
    tb(s, x + 1.10, y + 0.08, 4.65, 0.56, v, 10.8, SUB, spacing=1.18,
       anchor=MSO_ANCHOR.MIDDLE)

rect(s, 0.55, 6.34, 12.23, 0.58, CARD2, rounded=True)
tb(s, 0.80, 6.34, 11.8, 0.58,
   "致谢 · 欢迎提问与评审指正   |   四条底线：提醒能触发 · 数据能显示 · 断电不丢设置 · 连续运行不死机",
   11.5, TEAL, True, anchor=MSO_ANCHOR.MIDDLE)

prs.save(OUT)
print("SAVED", OUT, "slides=", len(prs.slides._sldIdLst))
