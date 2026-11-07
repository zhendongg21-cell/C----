# -*- coding: utf-8 -*-
"""生成《六人分工表》Word + Excel 两个交付物（单脚本、共享数据）。"""
import math, os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE = os.path.dirname(os.path.abspath(__file__))
FONT = 'Microsoft YaHei'
TEAL = '2DD4BF'
NAVY = '152238'
DARK = '0B1220'
WHITE = 'F8FAFC'

# ===================== 共享数据 =====================
NOTE = ('文档原文为 3~4 人 / 组、五类角色（表 5-3），本表扩编为「6 人单组」：'
        '1 名 PM + 5 名模块负责人；选做项挂靠已有模块负责人，不设第 7 人，'
        '保证必做闭环（计步 + 久坐提醒 + 表盘/提醒交互）不被稀释。'
        '对应文档：《基于 ESP32-S3 的智能健康手环 项目开发文档说明书（C++版）》，'
        '80 学时 / 10 天集中实训 / 表 5-2 八阶段 P0~P7。')

MAIN_HEADERS = ['序号', '角色', '主责模块（对应文档）', '主责阶段', '交付物', '验收要点']
MAIN_ROWS = [
    ['1', '组长 / 系统集成（PM）',
     '需求冻结（表 3-2 / 3-3）、四层装配、AppController / HealthBand 门面、看门狗 8 s、最小系统模式、部署拓扑 192.168.10.10:9527、Git 与阶段验收门、答辩统筹',
     'P0(4) · P6(10) · P7(8)',
     '需求 / 设计文档、Git 历史与 P0~P7 tag、部署手册、答辩材料与改进清单',
     '8 个阶段验收门逐门通过；四条底线全绿；密码不明文存 NVS'],
    ['2', '硬件与结构',
     '表 3-6~3-13 + 附录 A：接线焊接、USB-C VBUS 加 1 A 保险丝、3V3 / GND、I2C 400 kHz（GPIO8/9）、SPI 屏 GPIO10~16 / 80 MHz DMA、按键 GPIO4（10 kΩ + 100 nF）、马达 GPIO5（2N7002 + 1N4148 / 4 kHz）、表体 110×70×12 mm ≤ 85 g / 24 mm 表带、装调验收表 3-13、电气安全',
     'P1(8)；P2 支撑',
     '接线图、焊接 / 结构件记录、排查日志',
     '三传感器自检 ID 全绿（0x68 / 0x15 / 0x0117）、触摸 0x38；无短路虚焊；先测 3.3 V 与 GND'],
    ['3', '驱动与 HAL',
     'I2CBus 单例互斥锁、SpiBus / Gpio / Ledc / Nvs；MPU6050 / MAX30102 / TMP117 / FT6236 / ST7262 寄存器封装；WHO_AM_I · Part ID · Device ID 自检；20 kHz 背光 PWM；数据回放数据集',
     'P2(12)',
     'hal/ + driver/ 代码、自检日志、回放数据',
     '400 kHz I2C 多任务不冲突；连续 NACK 3 次自动重初始化；越层检查三条（ui 无 i2c_master_*、driver 无 lv_*、service 不碰 GPIO）'],
    ['4', '算法 · 计步与久坐（必做核心）',
     '50 Hz → SVM → 5 s 滑窗 → μ+0.8σ → 峰谷差 ≥ 0.6 g / 间隔 250–2000 ms / 震动屏蔽 300 ms；久坐六态（表 3-4）、SedentaryPolicy 与 Standard / DemoPolicy、SitState 状态模式、阈值 30/45/60 min 写 NVS、马达 2 s×3 + 全屏弹窗',
     'P3(14)；P4 支撑',
     'StepService / SedentaryService、T-01~T-05 用例、调参记录',
     '100 步误计 < 3% / 漏计 < 5%；45 min 计时误差 ≤ 10 s；“已起身 / 稍后 10 min（最多 3 次）”闭环；30 步复位'],
    ['5', '界面与交互',
     'LVGL 8.3 六页（表盘 / 运动 / 心率血氧 / 体温 / 久坐 / 设置）+ 提醒弹窗；表 3-16 配色 #0B1220 / #152238 / #F8FAFC / #94A3B8 与青绿 / 红 / 琥珀；触摸 ≥ 48 px、按钮 ≥ 52 px、动画 180 ms；2×800×40 局部刷新（全屏 768 KB ≈ 90–110 ms）；20 ms 消抖、短按翻页 / 长按 2 s 熄屏 / 5 s 回表盘 / 10 s 恢复出厂',
     'P4(12) · P5(12)',
     'app/ui 六页、可用性测试记录',
     '非专业用户 90 s 内完成“看步数—设阈值—查看提醒”；页面切换不阻塞；熄屏只关背光'],
    ['6', '测试 · 存储同步 · 文档',
     '16 字节定长 MinuteRecord（小端）、LittleFS 环形缓冲、“先数据后游标”、NVS；TCP + JSON Lines（每行 ≤ 64 KB，AUTH / SYNC / ACK / PING / ALERT，(device_id, ts) 幂等，退避 1/2/4/8 s）；T-01~T-10、8 项量化指标、五类故障演练、排障指南表 5-9、文档四件套',
     'P5 · P6；P7 支撑',
     '测试报告、量化验收表、部署手册、答辩数据',
     '连续 8 h 无死机、断电不丢设置、断网 2 h 补传无重复、五类故障可恢复；选做生理心率 ≤ ±5 bpm / SpO₂ ≤ ±3%'],
]

PHASE_HEADERS = ['阶段', '学时', '主责', '协同']
PHASE_ROWS = [
    ['P0 项目导入与需求分析', 4, '1', '全员'],
    ['P1 硬件搭建与环境配置', 8, '2', '3 辅、1 验收'],
    ['P2 驱动与传感器', 12, '3', '2 辅'],
    ['P3 计步与久坐', 14, '4', '3 / 6 写用例'],
    ['P4 完整界面与选做测量', 12, '5', '4 提供生理数据'],
    ['P5 交互完善与功能验证', 12, '6', '5 辅'],
    ['P6 联调测试与部署', 10, '1 + 6', '全员联调'],
    ['P7 答辩互评与总结', 8, '1', '全员答辩'],
    ['合计', 80, '—', '—'],
]

OPTIONAL_HEADERS = ['选做项', '难度分级', '归属人', '协同 / 说明']
OPTIONAL_ROWS = [
    ['PPG 心率血氧、体温', '◐ 按能力', '3 + 4',
     '3 出 MAX30102 / TMP117 驱动自检；4 做 0.5–5 Hz 带通、SpO₂≈110−25R、PI ≥ 0.5%；5 出页面'],
    ['跌倒检测', '◐ 按能力', '4', 'IMU 数据已在手，复用计步流水线'],
    ['Wi-Fi 同步 / PC 端 Qt 5', '○ 加分', '6 + 1',
     '6 归口存储同步与 SQLite WAL；1 负责部署拓扑 192.168.10.10:9527'],
    ['PC 端大模型健康问答（FR-19，新增选做）', '○ 加分', '4 + 6',
     '4 做本地/局域网推理与 Prompt；6 提供 PC 端归档数据；5 出 UI-PC-02 问答面板；回答须附来源、标注非医疗建议'],
    ['RAG 检索（FR-20，新增选做）', '○ 加分', '6 + 4',
     '6 用 SQLite records / events 已归档数据建本地索引并按 (device_id, ts) 溯源；4 接 Top-K 片段；前置＝Wi-Fi 同步 / PC 端'],
    ['手环操控鼠标（FR-21，新增选做）', '○ 加分', '5 + 2',
     '5 做手势映射与开关（默认关 / 二次确认 / 一键断开）；2 验证 BLE HID 或自定义协议；不进入计步与久坐主流程'],
    ['裁剪说明', '—', '—',
     '新增三项均为 P2 加分选做（FR-19/20/21，见 SRS §5.2、NFR-11/12）；裁剪掉以上任何一项都不影响四条底线'],
]

DEFENSE_HEADERS = ['人', '环节', '时长 / 内容']
DEFENSE_ROWS = [
    ['1', '开场', '3 min：定位、四层架构、改进清单'],
    ['4', '现场演示', '计步 + 久坐闭环（走 30 步复位 / DemoPolicy 60 s 触发）'],
    ['5', '界面演示', '六页界面与提醒弹窗'],
    ['2', '硬件实物', '排障顺序：3.3 V / GND → 总线波形 → 代码'],
    ['3', '代码质量', '自检 ID 与越层检查三条'],
    ['6', '数据验收', '量化验收数据与五类故障演练记录'],
]

RULES = [
    ('接口先行', 'P0 由 1 + 3 冻结 ImuFrame / PpgSample / EventBus / MinuteRecord 数据结构，跨任务传值不传指针，避免 P3 返工'),
    ('Git 规范', 'main 受保护、feature/<模块> 分支、P0~P7 打 tag，合并由 PM + 相邻模块负责人评审'),
    ('节奏', '每日站会 15 min；阶段验收门未通过不进下一阶段'),
    ('并发三不要', '不在中断调 LVGL 或动态分配内存、不持 I2C 锁等队列、不在 UI 任务做阻塞式文件写入'),
    ('共担', '四条底线（提醒能触发 / 数据能显示 / 断电不丢设置 / 连续运行不死机）是全员共担，不是 6 号一个人的验收项'),
    ('备选方案', '若必须严格按文档 3~4 人 / 组，可拆成 2 组 × 3 人——终端组（2 硬件 / 3 驱动 / 4 算法）+ 应用组（1 PM / 5 界面 / 6 测试），共用统一接口与四条底线验收；但六人单组集成成本更低，推荐用主表'),
]

DOCX_PATH = os.path.join(BASE, 'ESP32-S3智能健康手环_六人分工表.docx')
XLSX_PATH = os.path.join(BASE, 'ESP32-S3智能健康手环_六人分工表.xlsx')


# ===================== Word =====================
def _set_font(run, size=10.5, bold=False, color=None):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:eastAsia'), FONT)
    rFonts.set(qn('w:ascii'), FONT)
    rFonts.set(qn('w:hAnsi'), FONT)


def _shade(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hexcolor)
    tcPr.append(shd)


def _cell_text(cell, text, size=9, bold=False, color='1F2937'):
    cell.text = ''
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.18
    r = p.add_run(str(text))
    _set_font(r, size=size, bold=bold, color=color)


def _add_table(doc, headers, rows, widths_cm):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    for j, h in enumerate(headers):
        c = t.rows[0].cells[j]
        c.width = Cm(widths_cm[j])
        _shade(c, NAVY)
        _cell_text(c, h, size=9.5, bold=True, color=WHITE)
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    for i, row in enumerate(rows):
        cells = t.add_row().cells
        for j, val in enumerate(row):
            c = cells[j]
            c.width = Cm(widths_cm[j])
            if i % 2 == 1:
                _shade(c, 'F1F5F9')
            _cell_text(c, val, size=8.5)
    doc.add_paragraph()
    return t


def build_docx():
    doc = Document()
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = Cm(29.7), Cm(21.0)
    sec.left_margin = sec.right_margin = Cm(1.6)
    sec.top_margin = sec.bottom_margin = Cm(1.6)

    # 默认样式
    normal = doc.styles['Normal']
    normal.font.name = FONT
    normal.font.size = Pt(10.5)
    rPr = normal.element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts'); rPr.append(rFonts)
    rFonts.set(qn('w:eastAsia'), FONT)

    h = doc.add_heading(level=0)
    r = h.add_run('基于 ESP32-S3 的智能健康手环 · 六人分工表')
    _set_font(r, size=20, bold=True, color=DARK)

    sub = doc.add_paragraph()
    r = sub.add_run('答辩 / 评审汇报配套｜80 学时 · 10 天集中实训 · 对应文档表 5-2 / 表 5-3')
    _set_font(r, size=10.5, color='5B6B85')
    sub.paragraph_format.space_after = Pt(10)

    p = doc.add_paragraph()
    r = p.add_run(NOTE)
    _set_font(r, size=10, color='334155')
    p.paragraph_format.space_after = Pt(12)

    doc.add_heading('一 · 六人分工总表', level=1)
    _add_table(doc, MAIN_HEADERS, MAIN_ROWS, [1.2, 3.4, 9.6, 2.9, 4.4, 5.0])

    doc.add_heading('二 · 阶段—主责矩阵（80 学时）', level=1)
    _add_table(doc, PHASE_HEADERS, PHASE_ROWS, [7.0, 2.2, 3.0, 6.0])

    doc.add_heading('三 · 选做项归属（◐ 按能力 / ○ 加分，可整块裁剪）', level=1)
    _add_table(doc, OPTIONAL_HEADERS, OPTIONAL_ROWS, [5.4, 3.0, 3.0, 12.0])

    doc.add_heading('四 · 答辩现场分工', level=1)
    _add_table(doc, DEFENSE_HEADERS, DEFENSE_ROWS, [2.0, 4.0, 16.0])

    doc.add_heading('五 · 协作约定', level=1)
    for k, v in RULES:
        p = doc.add_paragraph(style='List Bullet')
        r = p.add_run(k + '：')
        _set_font(r, size=10, bold=True, color=TEAL and '0F766E')
        r2 = p.add_run(v)
        _set_font(r2, size=10, color='1F2937')

    doc.add_paragraph()
    foot = doc.add_paragraph()
    r = foot.add_run('四条底线：提醒能触发 · 数据能显示 · 断电不丢设置 · 连续运行不死机　｜　'
                     '平台：ESP32-S3 · FreeRTOS · LVGL · C++17')
    _set_font(r, size=9, color='5B6B85')

    doc.save(DOCX_PATH)
    return DOCX_PATH


# ===================== Excel =====================
THIN = Side(style='thin', color='CBD5E1')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _est_height(row, widths):
    lines = 1
    for val, w in zip(row, widths):
        units = sum(2 if ord(ch) > 127 else 1 for ch in str(val))
        lines = max(lines, math.ceil(units / max(6, w - 2)))
    return max(26, lines * 15 + 8)


def _sheet(wb, name, title, headers, rows, widths, first=False):
    ws = wb.active if first else wb.create_sheet()
    ws.title = name
    n = len(headers)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n)
    c = ws.cell(1, 1, title)
    c.font = Font(name=FONT, size=14, bold=True, color=DARK)
    c.alignment = Alignment(vertical='center', horizontal='left')
    ws.row_dimensions[1].height = 32

    for j, hname in enumerate(headers, 1):
        c = ws.cell(2, j, hname)
        c.fill = PatternFill('solid', fgColor=NAVY)
        c.font = Font(name=FONT, size=10.5, bold=True, color=WHITE)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[2].height = 28

    for i, row in enumerate(rows, 3):
        for j, val in enumerate(row, 1):
            c = ws.cell(i, j, val)
            c.font = Font(name=FONT, size=10, color='1F2937')
            c.alignment = Alignment(vertical='top', wrap_text=True)
            c.border = BORDER
            if (i - 3) % 2 == 1:
                c.fill = PatternFill('solid', fgColor='F1F5F9')
        ws.row_dimensions[i].height = _est_height(row, widths)

    for j, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = 'A3'
    ws.sheet_view.showGridLines = False
    return ws


def build_xlsx():
    wb = Workbook()

    # 说明页
    ws = wb.active
    ws.title = '说明'
    ws.sheet_view.showGridLines = False
    ws.merge_cells('A1:B1')
    c = ws.cell(1, 1, '基于 ESP32-S3 的智能健康手环 · 六人分工表')
    c.font = Font(name=FONT, size=16, bold=True, color=DARK)
    ws.row_dimensions[1].height = 36
    ws.merge_cells('A2:B2')
    c = ws.cell(2, 1, '答辩 / 评审汇报配套｜80 学时 · 10 天集中实训 · 对应文档表 5-2 / 表 5-3')
    c.font = Font(name=FONT, size=10.5, color='5B6B85')
    ws.row_dimensions[2].height = 24

    items = [
        ('编制说明', NOTE),
        ('组织形式', '6 人单组 = 1 名 PM + 5 名模块负责人；选做项挂靠已有负责人，不设第 7 人'),
        ('必做闭环', '计步 + 久坐提醒 + 表盘 / 提醒交互'),
        ('四条底线', '提醒能触发 · 数据能显示 · 断电不丢设置 · 连续运行不死机（全员共担）'),
        ('工作表', '六人分工总表 / 阶段主责矩阵 / 选做项归属 / 答辩现场分工 / 协作约定'),
        ('平台', 'ESP32-S3 · FreeRTOS · LVGL · C++17'),
        ('来源文档', 'ESP32-S3智能健康手环课程设计方案.docx（八章 + 三附录）'),
    ]
    for i, (k, v) in enumerate(items, 4):
        a = ws.cell(i, 1, k)
        a.font = Font(name=FONT, size=11, bold=True, color='0F766E')
        a.fill = PatternFill('solid', fgColor='E2F5F1')
        a.alignment = Alignment(vertical='top')
        a.border = BORDER
        b = ws.cell(i, 2, v)
        b.font = Font(name=FONT, size=10.5, color='1F2937')
        b.alignment = Alignment(vertical='top', wrap_text=True)
        b.border = BORDER
        ws.row_dimensions[i].height = _est_height([k, v], [18, 96])
    ws.column_dimensions['A'].width = 18
    ws.column_dimensions['B'].width = 96

    _sheet(wb, '六人分工总表', '一 · 六人分工总表',
           MAIN_HEADERS, MAIN_ROWS, [7, 22, 62, 18, 30, 34])
    wp = _sheet(wb, '阶段主责矩阵', '二 · 阶段—主责矩阵（80 学时，P0~P7 合计 80）',
                PHASE_HEADERS, PHASE_ROWS, [30, 10, 12, 26])
    # 合计行用公式，保持可核验（B3:B10 为 8 个阶段的数值学时）
    last = wp.max_row
    wp.cell(last, 2, '=SUM(B3:B%d)' % (last - 1))
    tc = wp.cell(last, 2)
    tc.font = Font(name=FONT, size=10.5, bold=True, color='0F766E')
    tc.alignment = Alignment(vertical='center', horizontal='center')
    tc.border = BORDER
    for col in (1, 3, 4):
        c = wp.cell(last, col)
        c.font = Font(name=FONT, size=10, bold=True, color='1F2937')
        c.alignment = Alignment(vertical='center', wrap_text=True)
        c.border = BORDER
    _sheet(wb, '选做项归属', '三 · 选做项归属（◐ 按能力 / ○ 加分，可整块裁剪）',
           OPTIONAL_HEADERS, OPTIONAL_ROWS, [26, 14, 14, 70])
    _sheet(wb, '答辩现场分工', '四 · 答辩现场分工',
           DEFENSE_HEADERS, DEFENSE_ROWS, [8, 18, 66])

    rule_rows = [[k, v] for k, v in RULES]
    _sheet(wb, '协作约定', '五 · 协作约定', ['条目', '内容'], rule_rows, [18, 100])

    wb.save(XLSX_PATH)
    return XLSX_PATH


if __name__ == '__main__':
    print('DOCX', build_docx())
    print('XLSX', build_xlsx())
    print('DONE')
