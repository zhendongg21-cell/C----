# -*- coding: utf-8 -*-
"""生成《项目进程安排策划书》Word + 配套《项目进程台账》Excel（单脚本、共享数据）。

数据来源：《基于 ESP32-S3 的智能健康手环 项目开发文档说明书（C++版）》
表 5-2（80 学时 P0~P7）、表 5-4（10 天计划）、表 5-7（总评）、
表 5-9（排障）、表 8-1（难易度 / 分层教学）与第三章验收指标。
"""
import math
import os

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
NAVY = '152238'
DARK = '0B1220'
WHITE = 'F8FAFC'
GREEN = '0F766E'
SUB = '5B6B85'
BODY = '1F2937'
ZEBRA = 'F1F5F9'
MILE = 'FDE68A'          # 里程碑底色（浅琥珀）
OPTIONAL = 'FFF3D6'      # 选做支线底色
CONTINUOUS = 'E2E8F0'    # 贯穿项底色

DOCX_PATH = os.path.join(BASE, 'ESP32-S3智能健康手环_项目进程安排策划书.docx')
XLSX_PATH = os.path.join(BASE, 'ESP32-S3智能健康手环_项目进程台账.xlsx')

# ===================== 共享数据 =====================
OVERVIEW = [
    ('项目名称', '基于 ESP32-S3 的智能健康手环（C++ 版）软硬结合综合实训'),
    ('策划书名称', '项目进程安排策划书（进度 / 里程碑 / 验收门 / 裁剪预案 / 风险应对）'),
    ('编制依据', '《基于 ESP32-S3 的智能健康手环 项目开发文档说明书（C++版）》'
                 '表 5-2、表 5-4、表 5-7、表 5-9、表 8-1 及第三章验收指标'),
    ('周期与学时', '10 天集中实训 / 80 学时（按 8 学时·天⁻¹ 折算，可裁剪至 64 学时 / 8 天）'),
    ('组织形式', '3~4 人 / 组（表 5-3）；若采用 6 人单组，见配套《六人分工表》'),
    ('核心平台', 'ESP32-S3 + FreeRTOS + LVGL 8.3 + C++17；PC 端 Qt 5（选做）'),
    ('必做闭环', '计步 + 久坐提醒 + 表盘 / 提醒交互，对应核心闭环“察觉—提醒—反馈”'),
    ('选做范围', '心率血氧 / 体温 / 跌倒检测 / Wi-Fi 同步 / PC 管理端；新增选做：PC 端大模型健康问答（FR-19）、'
                 '基于 PC 端已归档数据的 RAG 检索（FR-20）、手环操控鼠标（FR-21）（◐ 按能力 / ○ 加分，可整块裁剪）'),
    ('新增选做依赖', 'FR-14/15（Wi-Fi 同步 + PC 端只读归档）→ FR-20（RAG 索引）→ FR-19（健康问答）；'
                  'FR-21 手环操控鼠标为独立链路。三项均 P2 加分排在主线门 G3/G6 之后，裁剪任一不影响必做闭环。'),
    ('新增选做边界', '模型与检索仅本机/局域网、不出公网、不作医疗诊断；回答须附检索来源；鼠标映射默认关闭、'
                  '需二次确认、PC 端一键断开（见 SRS §5.2、NFR-11/12）。'),
    ('四条底线', '提醒能触发 · 数据能显示 · 断电不丢设置 · 连续运行不死机（全员共担，任一不绿不得答辩）'),
    ('总评构成', '功能实现 40 + 过程评价 30 + 文档 15 + 答辩 15（表 5-7）'),
    ('部署环境', '教师机 192.168.10.10:9527、2.4 GHz Wi-Fi、TCP + JSON Lines'),
]

PRINCIPLES = [
    ('接口先行',
     'P0 由 PM 与驱动负责人冻结 ImuFrame / PpgSample / EventBus / MinuteRecord 数据结构，'
     '跨任务传值不传指针；接口未冻结不放行 G0，避免 P3 阶段返工。'),
    ('阶段验收门',
     '设 G0~G7 共 8 道门，逐门放行，未通过不进下一阶段；退出标准全部可量化，'
     '不接受“差不多”“能跑就行”这类主观结论。'),
    ('关键路径优先',
     '主线 8 个阶段学时合计 4+8+12+14+12+12+10+8 = 80，恰好等于全周期，'
     '即零浮动时间。缓冲只能来自选做支线裁剪与日内复盘效率，绝不能挤占必做闭环。'),
    ('选做后置',
     'PPG / 体温 / 跌倒 / Wi-Fi 同步 / PC 端只在对应主线验收门通过后启动；'
     '整块裁剪不影响四条底线，也不降低必做闭环的验收标准。'),
    ('每日闭环',
     '15 min 站会 + 当日 Git 提交 + 阶段门复盘；进度偏差当天暴露、当天决策，'
     '阻塞项 30 min 未解决即升级 PM。'),
    ('分层不降底线',
     '弱组按表 8-1 裁剪到 P0~P3 + 最小界面；强组加 P4 PPG 精度、P5 连续运行、P6 PC 端同步。'
     '无论哪一档，四条底线与必做闭环的量化指标不得下调。'),
    ('文档同步',
     '需求 / 设计 / 测试 / 部署四件套按阶段门当日归档，禁止第 10 天集中补写；'
     '文档是验收项（15 分）而非收尾作业。'),
]

MILE_HEADERS = ['验收门', '阶段', '对应日', '学时', '进入条件', '退出交付物', '退出验收标准', '主责']
MILE_ROWS = [
    ['G0', 'P0 项目导入与需求分析', '第 1 天', 4,
     '器材到位、分组完成、Git 仓库与接口模板就绪',
     '需求说明（表 3-2 / 3-3）、接口数据结构冻结、Git main + tag P0、四条底线验收表',
     '接口先行（跨任务传值不传指针）；四条底线写入验收表；全员确认必做 / 选做范围',
     '1 PM'],
    ['G1', 'P1 硬件搭建与环境配置', '第 1–2 天', 8,
     'G0 通过；ESP-IDF 5.1+ / VS Code / LVGL 8.3 环境就绪',
     '接线图、焊接与结构件记录、装调验收表 3-13、烧录点屏可运行',
     '能烧录能点屏；三传感器自检 ID 0x68 / 0x15 / 0x0117 全绿、触摸 0x38；'
     '先测 3.3 V 与 GND，无短路虚焊；USB-C VBUS 加 1 A 保险丝',
     '2 硬件（3 辅、1 验收）'],
    ['G2', 'P2 驱动与传感器', '第 2–3 天', 12,
     'G1 通过；I2C 400 kHz（GPIO8/9）总线可扫描',
     'hal/ + driver/ 代码、三传感器自检日志、数据回放数据集',
     '多任务 I2C 不冲突（互斥锁）；连续 NACK 3 次自动重初始化；'
     '越层检查三条通过（ui 无 i2c_master_*、driver 无 lv_*、service 不碰 GPIO）',
     '3 驱动'],
    ['G3', 'P3 计步与久坐（关键路径最长段）', '第 4–5 天', 14,
     'G2 通过；回放数据集与离线调参环境可用',
     'StepService / SedentaryService、提醒弹窗、T-01~T-05 用例、调参记录',
     '100 步误计 < 3% / 漏计 < 5%；45 min 计时误差 ≤ 10 s；30 步复位；'
     '“已起身 / 稍后 10 min（最多 3 次）”闭环；DemoPolicy 60 s 可课堂演示',
     '4 算法（3 / 6 写用例）'],
    ['G4', 'P4 完整界面与选做测量', '第 6–7 天', 12,
     'G3 通过；六页视觉规范冻结（表 3-16）',
     'app/ui 六页 + 提醒弹窗、可用性测试记录、选做生理页面',
     '非专业用户 90 s 内完成“看步数—设阈值—查看提醒”；'
     '2×800×40 局部刷新（全屏 768 KB ≈ 90–110 ms）；熄屏只关背光',
     '5 界面（4 提供生理数据）'],
    ['G5', 'P5 交互完善与功能验证', '第 8 天', 12,
     'G4 通过；NVS / LittleFS / 看门狗方案冻结',
     '设置页与 NVS 持久化、16 字节 MinuteRecord 环形缓冲、看门狗 8 s、连续运行记录',
     '连续 8 h 无死机、断电不丢设置、阈值 30/45/60 min 掉电保持；“先数据后游标”',
     '6 测试（5 辅）'],
    ['G6', 'P6 联调测试与部署', '第 9 天', 10,
     'G5 通过；教师机 192.168.10.10:9527 与 2.4 GHz Wi-Fi 就绪',
     '测试报告、8 项量化验收表、部署手册、五类故障演练记录',
     '断网 2 h 补传无重复；五类故障（断电 / 断网 / 文件损坏 / 传感器 NACK / '
     '3 台 Wi-Fi 干扰）可恢复；四条底线全绿',
     '1 PM + 6 测试'],
    ['G7', 'P7 答辩互评与总结', '第 10 天', 8,
     'G6 通过；四条底线全绿',
     '答辩演示脚本、PPT、组间互评表、改进清单、文档四件套归档包',
     '演示 / 讲解 / 问答 / 改进清单齐全；总评按功能 40 + 过程 30 + 文档 15 + 答辩 15 评定',
     '1 PM（全员答辩）'],
]

DAY_HEADERS = ['日', '上午（4 学时）', '下午（4 学时）', '当日里程碑', '当日交付物']
DAY_ROWS = [
    ['第 1 天', 'P0 项目导入与需求冻结、分组与 Git 仓库', 'P1 环境搭建、烧录点屏',
     'G0 需求冻结 · 能烧录能点屏', '需求说明、接口数据结构、tag P0'],
    ['第 2 天', 'P1 接线焊接、I2C 总线扫描', 'P2 I2CBus 单例互斥封装、三传感器识别',
     '三传感器识别通过', '接线图、自检日志'],
    ['第 3 天', 'P2 MAX30102 / TMP117 驱动', '数据回放数据集、离线调参环境',
     'G2 驱动验收', 'hal/ + driver/、回放数据集'],
    ['第 4 天', 'P3 计步算法调参（50 Hz / μ+0.8σ）', '久坐状态机六态（表 3-4）',
     '计步闭环', 'StepService、调参记录'],
    ['第 5 天', 'P3 提醒弹窗与阈值写 NVS', 'T-01~T-05 用例与 DemoPolicy 60 s',
     'G3 计步 + 久坐闭环', 'SedentaryService、测试用例'],
    ['第 6 天', 'P4 LVGL 框架、表盘 / 运动 / 久坐页', '选做支线启动：PPG / 体温 / 跌倒',
     '主页面可用', '三页界面、可用性初测'],
    ['第 7 天', 'P4 六页补齐、配色与触摸规范', '非专业用户 90 s 可用性测试',
     'G4 界面验收', '六页 + 弹窗、可用性记录'],
    ['第 8 天', 'P5 设置页 / NVS / 16 字节记录', '看门狗 8 s、连续 8 h 运行、选做同步支线',
     'G5 可独立运行', '存储模块、连续运行记录'],
    ['第 9 天', 'P6 全员联调、五类故障演练', '测试报告与部署手册、量化验收表',
     'G6 测试验收', '测试报告、部署手册'],
    ['第 10 天', 'P7 答辩彩排、组间互评', '总结、改进清单、文档四件套归档',
     'G7 最终交付', 'PPT、改进清单、归档包'],
]

DEP_HEADERS = ['序号', '依赖关系', '类型', '说明 / 缓冲']
DEP_ROWS = [
    ['1', 'P0 → P1（接口数据结构冻结）', '串行 · 硬依赖',
     '接口未冻结则 P3 返工；G0 不放行不进 P1'],
    ['2', 'P1 → P2（3.3 V / GND / 总线波形）', '串行 · 硬依赖',
     '排障顺序先电源后代码；第 2 天预留半天缓冲'],
    ['3', 'P2 → P3（回放数据集）', '串行 · 硬依赖',
     '驱动不稳时改用离线回放调算法，避免占用在线硬件时间'],
    ['4', 'P3 → P4（步数 / 久坐事件）', '串行 · 硬依赖',
     '14 学时为全程最长段、关键路径瓶颈，PM 每日盯进度'],
    ['5', 'P4 → P5（页面事件）', '串行 · 硬依赖',
     '界面未稳定不做持久化联调'],
    ['6', 'P5 → P6（可独立运行）', '串行 · 硬依赖',
     '连续 8 h 无死机才进入故障演练'],
    ['7', 'P6 → P7（四条底线全绿）', '串行 · 硬依赖',
     '底线不全绿不允许答辩'],
    ['8', 'PPG / 体温 / 跌倒（第 6–8 天）', '并行 · 可裁剪',
     '仅在 G3 通过后启动；整块裁剪不影响四条底线'],
    ['9', 'Wi-Fi 同步 / PC 端 Qt 5（第 8–9 天）', '并行 · 可裁剪',
     '归口测试 + PM；裁剪后由本地 LittleFS 验收替代'],
    ['10', '文档四件套（第 1–10 天贯穿）', '并行 · 持续',
     '每阶段当日归档，禁止最后一天补写'],
]

RHYTHM_HEADERS = ['时段', '时长', '内容', '产出']
RHYTHM_ROWS = [
    ['08:30–08:45', '15 min',
     '每日站会：昨日完成 / 今日计划 / 阻塞项；阻塞 30 min 未解升级 PM',
     '站会记录、阻塞清单'],
    ['08:45–12:00', '4 学时', '上午实操（含 10 min 短休）：主线模块开发', '当日代码、自检日志'],
    ['13:30–16:45', '4 学时', '下午实操：联调、用例执行、可用性测试', '用例记录、联调日志'],
    ['16:45–17:15', '30 min', '当日提交：feature 分支合并、tag、验收台账更新', 'Git 提交、验收台账'],
    ['17:15–17:45', '30 min（仅阶段门日）', '阶段验收评审：对照退出标准逐条核验', '放行 / 不放行结论'],
    ['阶段门 / 每周', '—', '复盘：进度偏差、风险清单刷新、裁剪决策', '复盘纪要、裁剪决议'],
]

TRIM_HEADERS = ['档位', '适用对象', '保留范围', '裁剪范围', '学时', '四条底线']
TRIM_ROWS = [
    ['A · 完整版', '强组 / 评优',
     'P0~P7 全主线 + 全部选做支线（PPG / 体温 / 跌倒 / Wi-Fi / PC 端）+ '
     '新增三项选做（PC 大模型健康问答 / RAG 检索 / 手环操控鼠标）',
     '无', 80, '全绿，另获选做加分'],
    ['B · 标准版（推荐）', '绝大多数组',
     'P0~P7 全主线（必做闭环 + 六页 + 稳定性 + 联调 + 答辩）',
     '选做支线整块裁剪（含新增三项选做）', 80, '全绿'],
    ['C · 压缩版', '弱组 / 课时受限',
     'P0~P3 + P4 六页基础 + P5 必做稳定性 + P6 联调 + P7 答辩',
     '选做支线、界面精修与动画、PC 端、跌倒检测、新增 AI/HID 选做、部分扩展用例（约 16 学时 / 2 天）',
     64, '全绿（必做闭环不裁剪）'],
    ['D · 新增选做说明', 'G3/G6 门通过后按能力加做',
     'FR-19 PC 大模型健康问答 / FR-20 基于 PC 端归档数据的 RAG 检索 / FR-21 手环操控鼠标（P2 加分）',
     '可整块裁剪；FR-19 依赖 FR-20，FR-20 依赖 FR-14/15 的 Wi-Fi 同步与 PC 端归档；FR-21 独立',
     0, '不占用必做学时，裁剪不影响四条底线'],
]

RISK_HEADERS = ['编号', '进度风险', '触发信号', '应对 / 缓冲', '责任']
RISK_ROWS = [
    ['R1', '硬件接线短路 / 虚焊导致点屏失败', '自检 ID 不全绿、3.3 V 异常',
     '按“先 3.3 V 与 GND → 再总线波形 → 最后怀疑代码”排障；第 2 天预留半天缓冲',
     '2 硬件'],
    ['R2', 'I2C 多任务冲突 / 连续 NACK', '400 kHz 波形异常、传感器掉线',
     'I2CBus 互斥锁、连续 NACK 3 次自动重初始化；启用回放数据集离线调算法',
     '3 驱动'],
    ['R3', '计步调参不收敛（漏计 / 误计超标）', '100 步误计 ≥ 3% 或漏计 ≥ 5%',
     '25 / 50 / 100 Hz 对比实验 + 离线回放调参；不占用在线硬件时间；'
     '必要时收紧峰谷差 ≥ 0.6 g 与间隔 250–2000 ms 门槛', '4 算法'],
    ['R4', '4.3 寸屏刷新掉帧（< 11 fps）', '全屏刷新 90–110 ms、页面卡顿',
     '2×800×40 局部刷新缓冲、双缓冲 + DMA、UI 1 s 节流刷新', '5 界面'],
    ['R5', '选做支线挤占必做闭环', 'G3 未通过却已启动 PPG / PC 端',
     '选做支线仅在对应主线门通过后启动；整块裁剪不影响四条底线', '1 PM'],
    ['R6', 'PPG 信号质量不达标', 'PI < 0.5%、心率偏差 > ±5 bpm',
     '降为选做、不作为 G4 放行条件；标注“未标定”后仅展示趋势', '4 算法 / 6 测试'],
    ['R7', '主线进度整体延迟', '任一阶段门较计划晚 ≥ 半天',
     '启用 64 学时压缩版（C 档）；砍选做支线；阶段门当日复盘决定', '1 PM'],
    ['R8', '联调故障不可恢复', '断电 / 断网 / 文件损坏 / NACK / Wi-Fi 干扰 任一演练失败',
     '按五类故障演练清单逐项修复；看门狗 8 s、先数据后游标、指数退避 1/2/4/8 s',
     '6 测试'],
    ['R9', '新增 AI / HID 选做越界', '问答出公网、无归档数据仍编造、鼠标映射默认开启或侵入计步主流程',
     '按 SRS §5.2、NFR-11/12：模型与检索仅本机/局域网、回答附来源、默认关 + 二次确认 + 一键断开；'
     '整块裁剪该项即恢复，不影响四条底线',
     '1 PM / 4 算法 / 5 界面'],
]

DELIV_HEADERS = ['类别', '交付物', '归档时点', '归属']
DELIV_ROWS = [
    ['代码', 'health-band/（hal / driver / service / app + test / tools / docs）、'
             'Git 历史与 P0~P7 tag', '每日 16:45 提交、阶段门打 tag', '全员 / 1 归档'],
    ['硬件', '接线图、焊接与结构件记录、装调验收表 3-13、电气安全检查记录',
     '第 2 天（G1）', '2 硬件'],
    ['数据', '三传感器自检日志、回放数据集、计步调参记录、8 项量化验收表',
     '第 3 / 5 / 9 天', '3 驱动 / 4 算法 / 6 测试'],
    ['文档四件套', '需求说明（表 3-2 / 3-3，含新增 FR-19~21 / NFR-11~12 / UC-09~11）、'
                   '设计说明（架构 / 状态机 / 协议 / 选做扩展链路图）、'
                   '测试报告（T-01~T-17，含 T-15~17 新增选做用例）、部署手册',
     '阶段门当日归档，禁止第 10 天补写', '6 测试 / 1 审核'],
    ['答辩', '演示脚本、答辩 PPT、组间互评表、改进清单',
     '第 10 天上午', '1 PM（全员）'],
]

# 甘特：阶段、学时、起始日、结束日、颜色
PHASES = [
    ('P0 项目导入', 4, 1, 1, '64748B', 'G0'),
    ('P1 硬件与环境', 8, 1, 2, '3B82F6', 'G1'),
    ('P2 驱动与传感器', 12, 2, 3, '06B6D4', 'G2'),
    ('P3 计步与久坐', 14, 4, 5, '0D9488', 'G3'),
    ('P4 界面与选做测量', 12, 6, 7, 'D97706', 'G4'),
    ('P5 交互与验证', 12, 8, 8, 'EA580C', 'G5'),
    ('P6 联调测试与部署', 10, 9, 9, 'E11D48', 'G6'),
    ('P7 答辩与总结', 8, 10, 10, '7C3AED', 'G7'),
]

# Excel 细粒度甘特：编号、任务、起始日、结束日、阶段、类型、主责
TASKS = [
    ('T01', '项目导入 / 需求与接口冻结', 1, 1, 'P0', '主线', '1'),
    ('T02', '环境搭建 / 烧录点屏', 1, 2, 'P1', '主线', '2'),
    ('T03', '接线焊接 / 结构与电气验收', 2, 2, 'P1', '主线', '2'),
    ('T04', 'I2CBus HAL + 三传感器自检', 2, 3, 'P2', '主线', '3'),
    ('T05', 'MAX30102 / TMP117 驱动 + 回放数据', 3, 3, 'P2', '主线', '3'),
    ('T06', '计步算法调参（50 Hz / μ+0.8σ）', 4, 4, 'P3', '主线', '4'),
    ('T07', '久坐状态机六态 + 提醒弹窗', 4, 5, 'P3', '主线', '4'),
    ('T08', 'T-01~T-05 用例 / DemoPolicy 60 s', 5, 5, 'P3', '主线', '6'),
    ('T09', 'LVGL 框架 + 表盘 / 运动 / 久坐页', 6, 6, 'P4', '主线', '5'),
    ('T10', '六页补齐 + 非专业用户 90 s 测试', 7, 7, 'P4', '主线', '5'),
    ('T11', '选做：PPG 心率血氧 / 体温', 6, 8, 'P4', '选做', '3+4'),
    ('T12', '选做：跌倒检测', 7, 8, 'P4', '选做', '4'),
    ('T13', 'NVS / 16 字节记录 / 看门狗 8 s', 8, 8, 'P5', '主线', '6'),
    ('T14', '选做：Wi-Fi 同步 / PC 端 Qt 5', 8, 9, 'P5', '选做', '6+1'),
    ('T15', '全员联调 + 五类故障演练 + 测试报告', 9, 9, 'P6', '主线', '1+6'),
    ('T16', '答辩彩排 / 互评 / 改进清单', 10, 10, 'P7', '主线', '1'),
    ('T17', '文档四件套（贯穿）', 1, 10, '全程', '持续', '6'),
    ('T18', '选做：RAG 检索 / 本地索引（FR-20，前置）', 9, 9, 'P6', '选做', '6+4'),
    ('T19', '选做：PC 大模型健康问答（FR-19）', 9, 10, 'P6', '选做', '4+6'),
    ('T20', '选做：手环操控鼠标 BLE HID（FR-21）', 8, 9, 'P5', '选做', '5+2'),
]

NOTE = ('本策划书以文档表 5-2 的 80 学时 / P0~P7 八阶段为骨架，以表 5-4 的 10 天日历为时间轴，'
        '把每阶段展开为「进入条件 → 退出交付物 → 退出验收标准」的验收门 G0~G7；'
        '并给出关键路径、并行支线、每日节奏、64 学时裁剪预案与进度风险应对。'
        '甘特图按表 5-4 的日历归属绘制，学时按阶段汇总（表 5-2），阶段可跨半日，日内以 8 学时折算。'
        '配套《项目进程台账.xlsx》提供可直接填写的里程碑、甘特与跟踪台账。')


# ===================== Word 辅助 =====================
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


def _cell_text(cell, text, size=9, bold=False, color=BODY, align=None):
    cell.text = ''
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.16
    r = p.add_run(str(text))
    _set_font(r, size=size, bold=bold, color=color)
    if align is not None:
        p.alignment = align


def _add_table(doc, headers, rows, widths_cm, size=8.5):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    for j, h in enumerate(headers):
        c = t.rows[0].cells[j]
        c.width = Cm(widths_cm[j])
        _shade(c, NAVY)
        _cell_text(c, h, size=9.5, bold=True, color=WHITE,
                   align=WD_ALIGN_PARAGRAPH.CENTER)
    for i, row in enumerate(rows):
        cells = t.add_row().cells
        for j, val in enumerate(row):
            c = cells[j]
            c.width = Cm(widths_cm[j])
            if i % 2 == 1:
                _shade(c, ZEBRA)
            _cell_text(c, val, size=size)
    doc.add_paragraph()
    return t


def _add_bullets(doc, pairs):
    for k, v in pairs:
        p = doc.add_paragraph(style='List Bullet')
        r = p.add_run(k + '：')
        _set_font(r, size=10, bold=True, color=GREEN)
        r2 = p.add_run(v)
        _set_font(r2, size=10, color=BODY)


def _add_gantt(doc):
    headers = ['阶段 / 学时', 'D1', 'D2', 'D3', 'D4', 'D5',
               'D6', 'D7', 'D8', 'D9', 'D10']
    widths = [4.6] + [2.0] * 10
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    for j, h in enumerate(headers):
        c = t.rows[0].cells[j]
        c.width = Cm(widths[j])
        _shade(c, NAVY)
        _cell_text(c, h, size=9, bold=True, color=WHITE,
                   align=WD_ALIGN_PARAGRAPH.CENTER)
    for name, hours, d0, d1, color, gate in PHASES:
        cells = t.add_row().cells
        c = cells[0]
        c.width = Cm(widths[0])
        _cell_text(c, '%s（%d）' % (name, hours), size=8.5, bold=True, color=DARK)
        for d in range(1, 11):
            c = cells[d]
            c.width = Cm(widths[d])
            if d0 <= d <= d1:
                _shade(c, color)
                _cell_text(c, '●' if d == d1 else ' ',
                           size=9, bold=True, color=WHITE,
                           align=WD_ALIGN_PARAGRAPH.CENTER)
            else:
                _cell_text(c, '', size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_paragraph()
    p = doc.add_paragraph()
    r = p.add_run('　● = 该阶段验收门（G0~G7）所在日；色块 = 阶段占用日。'
                  '主线八阶段合计 80 学时，与 10 天 × 8 学时完全对齐，无浮动时间。')
    _set_font(r, size=9, color=SUB)
    doc.add_paragraph()
    return t


def build_docx():
    doc = Document()
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = Cm(29.7), Cm(21.0)
    sec.left_margin = sec.right_margin = Cm(1.6)
    sec.top_margin = sec.bottom_margin = Cm(1.6)

    normal = doc.styles['Normal']
    normal.font.name = FONT
    normal.font.size = Pt(10.5)
    rPr = normal.element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:eastAsia'), FONT)

    h = doc.add_heading(level=0)
    r = h.add_run('基于 ESP32-S3 的智能健康手环 · 项目进程安排策划书')
    _set_font(r, size=20, bold=True, color=DARK)

    sub = doc.add_paragraph()
    r = sub.add_run('进度 · 里程碑 · 验收门 · 关键路径 · 裁剪预案｜80 学时 / 10 天集中实训｜'
                    '对应文档表 5-2 / 表 5-4 / 表 5-7 / 表 5-9 / 表 8-1')
    _set_font(r, size=10.5, color=SUB)
    sub.paragraph_format.space_after = Pt(10)

    p = doc.add_paragraph()
    r = p.add_run(NOTE)
    _set_font(r, size=10, color='334155')
    p.paragraph_format.space_after = Pt(12)

    doc.add_heading('一 · 项目概况与编制依据', level=1)
    _add_table(doc, ['项目', '内容'], OVERVIEW, [4.4, 22.0], size=9)

    doc.add_heading('二 · 进度安排总原则', level=1)
    _add_bullets(doc, PRINCIPLES)

    doc.add_heading('三 · 总体里程碑与阶段验收门（G0~G7）', level=1)
    p = doc.add_paragraph()
    r = p.add_run('每道门均设「进入条件 → 退出交付物 → 退出验收标准」三段；'
                  '未通过不放行下一阶段，验收标准全部可量化。')
    _set_font(r, size=9.5, color=SUB)
    _add_table(doc, MILE_HEADERS, MILE_ROWS,
               [1.4, 3.5, 2.0, 1.2, 4.3, 5.6, 6.2, 2.2], size=8)

    doc.add_heading('四 · 阶段甘特图（按表 5-4 日历归属）', level=1)
    _add_gantt(doc)

    doc.add_heading('五 · 十日日历进度', level=1)
    _add_table(doc, DAY_HEADERS, DAY_ROWS, [2.0, 7.0, 7.0, 4.6, 5.8], size=8.5)

    doc.add_heading('六 · 关键路径与并行支线', level=1)
    p = doc.add_paragraph()
    r = p.add_run('关键路径：P0 → P1 → P2 → P3 → P4 → P5 → P6 → P7，'
                  '共 4+8+12+14+12+12+10+8 = 80 学时，等于全周期，零浮动时间。'
                  '其中 P3（14 学时）为最长段与首要瓶颈。缓冲只能来自选做支线裁剪与日内复盘效率。')
    _set_font(r, size=10, bold=True, color=GREEN)
    p.paragraph_format.space_after = Pt(8)
    _add_table(doc, DEP_HEADERS, DEP_ROWS, [1.4, 7.4, 3.4, 14.2], size=8.5)

    doc.add_heading('七 · 每日节奏与进度跟踪机制', level=1)
    _add_table(doc, RHYTHM_HEADERS, RHYTHM_ROWS, [3.4, 3.2, 12.4, 7.4], size=8.5)
    _add_bullets(doc, [
        ('Git 规范', 'main 受保护、feature/<模块> 分支、P0~P7 打 tag；合并由 PM + 相邻模块负责人评审'),
        ('验收台账', '配套《项目进程台账.xlsx》逐日记录计划 / 实际 / 偏差 / 阻塞 / 处理 / 状态'),
        ('量化看板', '8 项量化指标（100 步误计 < 3%、漏计 < 5%、45 min 误差 ≤ 10 s、连续 8 h 无死机、'
                     '断电不丢设置、断网 2 h 补传无重复、心率 ≤ ±5 bpm / SpO₂ ≤ ±3%、非专业用户 90 s 完成三步）'),
        ('放行规则', '四条底线全绿 + 该门退出标准逐条通过 + 文档四件套归档 = 放行；否则当日复盘并决定裁剪'),
    ])

    doc.add_heading('八 · 缓冲与 64 学时裁剪预案', level=1)
    _add_table(doc, TRIM_HEADERS, TRIM_ROWS, [3.4, 3.2, 7.6, 7.4, 1.4, 3.4], size=8.5)
    p = doc.add_paragraph()
    r = p.add_run('分层教学（表 8-1）：弱组只做 P0~P3，强组加 P4 PPG 精度 / P5 存储连续运行 / '
                  'P6 PC 端同步；无论哪一档都须通过四条底线，且主线 80 学时不可压缩到 64 以下'
                  '而不裁剪选做与界面精修。')
    _set_font(r, size=9.5, color=SUB)

    doc.add_heading('九 · 进度风险与应对', level=1)
    _add_table(doc, RISK_HEADERS, RISK_ROWS, [1.4, 5.0, 4.6, 12.6, 2.8], size=8)

    doc.add_heading('十 · 交付物清单与答辩收尾', level=1)
    _add_table(doc, DELIV_HEADERS, DELIV_ROWS, [2.8, 12.6, 5.6, 5.4], size=8.5)

    doc.add_paragraph()
    foot = doc.add_paragraph()
    r = foot.add_run('四条底线：提醒能触发 · 数据能显示 · 断电不丢设置 · 连续运行不死机　｜　'
                     '配套：ESP32-S3智能健康手环_项目进程台账.xlsx、ESP32-S3智能健康手环_六人分工表.docx / .xlsx、'
                     'ESP32-S3智能健康手环_答辩汇报.pptx、ESP32-S3智能健康手环_思维导图.html')
    _set_font(r, size=8.5, color=SUB)

    doc.save(DOCX_PATH)
    return DOCX_PATH


# ===================== Excel 辅助 =====================
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
            c.font = Font(name=FONT, size=10, color=BODY)
            c.alignment = Alignment(vertical='top', wrap_text=True)
            c.border = BORDER
            if (i - 3) % 2 == 1:
                c.fill = PatternFill('solid', fgColor=ZEBRA)
        ws.row_dimensions[i].height = _est_height(row, widths)

    for j, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = 'A3'
    ws.sheet_view.showGridLines = False
    return ws


def build_xlsx():
    wb = Workbook()

    # ---- 说明 ----
    ws = wb.active
    ws.title = '说明'
    ws.sheet_view.showGridLines = False
    ws.merge_cells('A1:B1')
    c = ws.cell(1, 1, '基于 ESP32-S3 的智能健康手环 · 项目进程台账')
    c.font = Font(name=FONT, size=16, bold=True, color=DARK)
    ws.row_dimensions[1].height = 36
    ws.merge_cells('A2:B2')
    c = ws.cell(2, 1, '《项目进程安排策划书》配套｜80 学时 / 10 天 / 8 道验收门 G0~G7')
    c.font = Font(name=FONT, size=10.5, color=SUB)
    ws.row_dimensions[2].height = 24

    items = [
        ('编制说明', NOTE),
        ('工作表', '里程碑计划 / 十日甘特 / 每日日历 / 裁剪预案 / 跟踪台账'),
        ('主线学时', '4 + 8 + 12 + 14 + 12 + 12 + 10 + 8 = 80（零浮动时间，缓冲来自选做裁剪）'),
        ('关键路径', 'P0 → P1 → P2 → P3 → P4 → P5 → P6 → P7；P3（14 学时）为瓶颈'),
        ('验收门', 'G0~G7 逐门放行，未通过不进下一阶段'),
        ('四条底线', '提醒能触发 · 数据能显示 · 断电不丢设置 · 连续运行不死机（全员共担）'),
        ('甘特图例', '深色实心 = 主线阶段；浅琥珀 = 选做支线；浅灰 = 贯穿项；● = 验收门所在日'),
        ('跟踪台账', '逐日填写计划 / 实际 / 偏差 / 阻塞 / 处理 / 状态，阶段门日追加复盘结论'),
        ('来源文档', 'ESP32-S3智能健康手环课程设计方案.docx（表 5-2 / 5-4 / 5-7 / 5-9 / 8-1）'),
        ('配套文件', '项目进程安排策划书.docx、六人分工表.docx/.xlsx、答辩汇报.pptx、思维导图.html'),
    ]
    for i, (k, v) in enumerate(items, 4):
        a = ws.cell(i, 1, k)
        a.font = Font(name=FONT, size=11, bold=True, color=GREEN)
        a.fill = PatternFill('solid', fgColor='E2F5F1')
        a.alignment = Alignment(vertical='top')
        a.border = BORDER
        b = ws.cell(i, 2, v)
        b.font = Font(name=FONT, size=10.5, color=BODY)
        b.alignment = Alignment(vertical='top', wrap_text=True)
        b.border = BORDER
        ws.row_dimensions[i].height = _est_height([k, v], [18, 96])
    ws.column_dimensions['A'].width = 18
    ws.column_dimensions['B'].width = 96

    # ---- 里程碑计划（含 SUM 公式） ----
    mile_rows = [[r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7]] for r in MILE_ROWS]
    wm = _sheet(wb, '里程碑计划',
                '一 · 总体里程碑与阶段验收门（G0~G7，学时合计 80）',
                MILE_HEADERS, mile_rows, [9, 30, 12, 9, 38, 46, 54, 18])
    last = wm.max_row + 1
    wm.cell(last, 1, '合计')
    wm.cell(last, 4, '=SUM(D3:D%d)' % (last - 1))
    for j in range(1, len(MILE_HEADERS) + 1):
        c = wm.cell(last, j)
        c.fill = PatternFill('solid', fgColor='E2F5F1')
        c.font = Font(name=FONT, size=10.5, bold=True,
                      color=GREEN if j == 4 else DARK)
        c.alignment = Alignment(vertical='center', wrap_text=True,
                                horizontal='center' if j == 4 else 'left')
        c.border = BORDER
    wm.row_dimensions[last].height = 26

    # ---- 十日甘特 ----
    wg = wb.create_sheet()
    wg.title = '十日甘特'
    wg.sheet_view.showGridLines = False
    g_headers = ['编号', '任务', '起始日', '结束日', '阶段', '类型', '主责'] + \
                ['D%d' % d for d in range(1, 11)]
    wg.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(g_headers))
    c = wg.cell(1, 1, '二 · 十日甘特（深色 = 主线，浅琥珀 = 选做，浅灰 = 贯穿；● = 验收门日）')
    c.font = Font(name=FONT, size=13, bold=True, color=DARK)
    c.alignment = Alignment(vertical='center', horizontal='left')
    wg.row_dimensions[1].height = 30
    for j, h in enumerate(g_headers, 1):
        c = wg.cell(2, j, h)
        c.fill = PatternFill('solid', fgColor=NAVY)
        c.font = Font(name=FONT, size=10, bold=True, color=WHITE)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = BORDER
    wg.row_dimensions[2].height = 26

    phase_color = {p[0].split()[0]: p[4] for p in PHASES}
    gate_days = {p[5]: p[3] for p in PHASES}
    for i, (tid, name, d0, d1, phase, kind, owner) in enumerate(TASKS, 3):
        vals = [tid, name, d0, d1, phase, kind, owner]
        for j, v in enumerate(vals, 1):
            c = wg.cell(i, j, v)
            c.font = Font(name=FONT, size=9.5, color=BODY)
            c.alignment = Alignment(vertical='center', wrap_text=True,
                                    horizontal='center' if j in (1, 3, 4, 5, 6, 7) else 'left')
            c.border = BORDER
            if (i - 3) % 2 == 1:
                c.fill = PatternFill('solid', fgColor=ZEBRA)
        for d in range(1, 11):
            col = 7 + d
            c = wg.cell(i, col, '')
            c.border = BORDER
            c.alignment = Alignment(horizontal='center', vertical='center')
            if not (d0 <= d <= d1):
                continue
            if kind == '选做':
                c.fill = PatternFill('solid', fgColor=OPTIONAL)
                c.font = Font(name=FONT, size=9, bold=True, color='92400E')
                c.value = '○'
            elif kind == '持续':
                c.fill = PatternFill('solid', fgColor=CONTINUOUS)
                c.font = Font(name=FONT, size=9, color='475569')
                c.value = '–'
            else:
                c.fill = PatternFill('solid', fgColor=phase_color.get(phase, '0D9488'))
                c.font = Font(name=FONT, size=9, bold=True, color=WHITE)
                # 主线阶段末日且为验收门日 → ●
                c.value = '●' if d == d1 and gate_days.get(
                    next((p[5] for p in PHASES if p[0].startswith(phase)), ''), -1) == d else '■'
        wg.row_dimensions[i].height = 28
    for d in range(1, 11):
        wg.column_dimensions[get_column_letter(7 + d)].width = 7
    for j, w in enumerate([9, 40, 9, 9, 9, 9, 9], 1):
        wg.column_dimensions[get_column_letter(j)].width = w
    wg.freeze_panes = 'H3'

    # ---- 每日日历 ----
    _sheet(wb, '每日日历', '三 · 十日日历进度（表 5-4 展开）',
           DAY_HEADERS, DAY_ROWS, [11, 40, 40, 28, 34])

    # ---- 裁剪预案 ----
    _sheet(wb, '裁剪预案', '四 · 缓冲与 64 学时裁剪预案',
           TRIM_HEADERS, TRIM_ROWS, [20, 18, 44, 44, 10, 22])

    # ---- 跟踪台账 ----
    track_headers = ['日期', '阶段 / 门', '计划内容', '实际完成', '偏差（学时）',
                     '阻塞项', '处理 / 裁剪决议', '责任人', '状态']
    track_rows = [
        ['第 1 天', 'P0 / G0', '需求冻结、接口数据结构、能烧录能点屏', '', '', '', '', '1 PM', ''],
        ['第 2 天', 'P1 / P2', '接线焊接、I2C 总线扫描、三传感器识别', '', '', '', '', '2 / 3', ''],
        ['第 3 天', 'P2 / G2', '驱动验收、回放数据集', '', '', '', '', '3 驱动', ''],
        ['第 4 天', 'P3', '计步算法调参、久坐状态机', '', '', '', '', '4 算法', ''],
        ['第 5 天', 'P3 / G3', '提醒弹窗、测试用例、计步 + 久坐闭环', '', '', '', '', '4 算法', ''],
        ['第 6 天', 'P4', 'LVGL 框架、表盘 / 运动 / 久坐页', '', '', '', '', '5 界面', ''],
        ['第 7 天', 'P4 / G4', '六页补齐、90 s 可用性测试', '', '', '', '', '5 界面', ''],
        ['第 8 天', 'P5 / G5', 'NVS / 16 字节记录 / 看门狗、连续 8 h', '', '', '', '', '6 测试', ''],
        ['第 9 天', 'P6 / G6', '联调、五类故障演练、测试报告与部署手册', '', '', '', '', '1 + 6', ''],
        ['第 10 天', 'P7 / G7', '答辩彩排、互评、改进清单、文档归档', '', '', '', '', '1 PM', ''],
    ]
    wt = _sheet(wb, '跟踪台账', '五 · 每日进度跟踪台账（逐日填写，阶段门日追加复盘结论）',
                track_headers, track_rows, [11, 13, 40, 30, 13, 26, 34, 12, 12])
    for i in range(3, wt.max_row + 1):
        for j in (4, 5, 6, 7, 9):
            c = wt.cell(i, j)
            c.fill = PatternFill('solid', fgColor='FFFDF3')
            c.border = BORDER
            c.alignment = Alignment(vertical='top', wrap_text=True)

    wb.save(XLSX_PATH)
    return XLSX_PATH


if __name__ == '__main__':
    print('DOCX', build_docx())
    print('XLSX', build_xlsx())
    print('DONE')
