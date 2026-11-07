# Git 仓库与提交规范

> P0 阶段交付物 · 任务 ④「建立 Git 仓库、目录结构与提交规范」· 依据原方案 §4.1、§5.4（1109 行）、§5.1

## 1. 仓库信息

| 项 | 值 |
| --- | --- |
| 仓库根目录 | `health-band/`（本目录） |
| 默认分支 | `main` |
| 远程 | GitLab / GitHub（课程指定其一；本地已可完整离线工作） |
| 设备编号 | `HB-2026-G03-01` |
| 提交者 | `gao <zhendongg21@gmail.com>`（本机 git 配置） |

## 2. 分支模型

```
main                  ← 始终保持可运行；受保护，禁止直接 push
 ├── feature/srs      ← 需求规格说明书
 ├── feature/diagrams ← 用例图 / 状态图 / 原型 / 五链路草图
 ├── feature/hal      ← I2CBus / SpiBus / Gpio / Ledc / Nvs
 ├── feature/driver   ← mpu6050 / max30102 / tmp117 / ft6236 / st7262
 ├── feature/step     ← 计步与活动强度
 ├── feature/sedentary← 久坐状态机与提醒
 ├── feature/ui       ← LVGL 六页面与主题
 ├── feature/record   ← RecordStore / SyncClient
 ├── feature/llm-rag  ← 选做：PC 端大模型健康问答 + 基于 PC 端数据的 RAG 检索（FR-19/20）
 ├── feature/hid-mouse← 选做：手环操控鼠标（BLE HID / 自定义协议，FR-21）
 └── feature/test     ← 测试用例与记录
```

**规则**：
1. `main` 保持可运行，任何时刻都能 `idf.py build` 通过（对应「连续运行不死机」底线）。
2. 所有功能在 `feature/xxx` 分支开发，分支名用小写 + 连字符，语义化。
3. 通过 **Merge Request（MR / PR）** 合入，禁止直接向 `main` 提交。
4. 每次合入前 **至少一名同组同学评审**，检查：编译、日志、异常处理、命名规范。
5. 一个 MR 只做一件事，便于回滚与追溯。

## 3. 提交信息规范（原方案 1109 行原文口径）

格式：**动词 + 对象 + 原因**，采用 Conventional Commits 前缀 + 中文说明。

```
<type>: <动词><对象>，<原因/目的>
```

| type | 用于 |
| --- | --- |
| `feat` | 新增功能 / 新文档 |
| `fix` | 修复缺陷 |
| `docs` | 文档变更 |
| `refactor` | 重构（不改行为） |
| `test` | 测试用例与测试记录 |
| `chore` | 构建、脚本、配置、目录骨架 |
| `perf` | 性能优化 |

**示例**（可直接照抄）：

```
feat: 增加久坐状态机的稍后提醒，避免连续三次提醒后不再提醒
fix: 修复马达震动期间误计步的问题，因合加速度峰值被 PWM 干扰
docs: 补充 P0 需求规格说明书，使每条需求可被测试用例覆盖
test: 增加 T-01 平地 100 步×5 组计步基准用例
chore: 建立 hal/driver/service/app 四组件目录骨架，落实「目录即架构」
```

**禁止**：`update`、`修改`、`111`、`asdf`、空信息、一次提交塞入多个不相关改动。

## 4. 评审检查清单（MR 模板见 `CONTRIBUTING.md`）

- [ ] `idf.py build` 通过，无新增警告
- [ ] 关键路径有日志（含时间戳），异常分支有重试或降级
- [ ] 命名符合分层：不出现 `i2c_master_` 于 `ui/`、不出现 `lv_` 于 `driver/`
- [ ] 未改坏 `main` 可运行性；未提交构建产物 / 密钥 / 临时文件
- [ ] 若涉及需求，已同步更新 `docs/04_需求追踪矩阵.md`
- [ ] 至少 1 名同组同学 Approve

## 5. 目录结构（目录即架构，§4.1）

```
health-band/
├── CMakeLists.txt                 # 顶层：注册 hal/driver/service/app 四组件
├── README.md
├── CONTRIBUTING.md                # 提交与评审速查
├── .gitignore
├── .gitattributes
├── main/
│   ├── CMakeLists.txt
│   ├── app_main.cpp
│   ├── hal/          # I2CBus / SpiBus / Gpio / Ledc / Nvs（互斥锁保证 I2C 原子性）
│   ├── driver/       # mpu6050 / max30102 / tmp117 / ft6236 / st7262
│   ├── service/      # sensor_hub / step / sedentary / vital / fall
│   │                 # record_store.cpp（LittleFS + NVS 游标）
│   │                 # sync_client.cpp（Wi-Fi + JSON Lines）
│   ├── app/          # app_controller / screen_manager / screens
│   ├── ui/           # LVGL 页面与主题、字体子集
│   └── board/        # 引脚定义、板级初始化、版本号
├── test/             # 主机侧单元测试
├── tools/            # 数据回放、CSV 转二进制、打包脚本
└── docs/             # 需求、协议、测试记录、图集
    ├── 00_需求规格说明书.md
    ├── 01_角色痛点与需求映射.md
    ├── 02_分组与角色分工.md
    ├── 03_Git仓库与提交规范.md
    ├── 04_需求追踪矩阵.md
    ├── 05_P0验收对照表.md
    └── diagrams/P0_图集.drawio
```

**CMake 强制分层**：四个组件 `hal` / `driver` / `service` / `app` 用 `target_link_libraries` 表达「上层依赖下层」，编译期即阻止反向依赖。

**三条越层红线**：
1. `ui/` 不得出现 `i2c_master_` API
2. `driver/` 不得出现 `lv_` API
3. `service/` 不得直接操作 GPIO

> P2 阶段起要求「目录即架构」；P0 已先把骨架建好，后续按目录填代码即可。

## 6. 忽略规则（.gitignore 要点）

- 构建产物：`build/`、`sdkconfig`、`sdkconfig.old`、`*.o`、`*.elf`、`*.bin`、`*.map`
- IDE / 系统：`.vscode/`（除共享配置）、`.DS_Store`、`Thumbs.db`
- 临时与锁文件：`~$*`、`*.tmp`、`*.swp`
- 密钥与本地凭据：`*.pem`、`secrets.*`、Wi-Fi 密码明文文件（协议要求「不保存 Wi-Fi 密码明文」）
- 大数据：原始 CSV 回放数据放 `tools/data/` 并按需 LFS 或忽略，只提交样例

## 7. 常用命令速查

```bash
git init                                   # 初始化（P0 已执行）
git checkout -b feature/srs                # 建功能分支
git add docs/00_需求规格说明书.md
git commit -m "docs: 补充 P0 需求规格说明书，使每条需求可被测试用例覆盖"
git checkout main
git merge --no-ff feature/srs -m "Merge MR !1: P0 需求规格说明书"
git log --oneline --graph --decorate       # 查看历史
```

## 8. P0 验收对应

| P0 要求 | 本文件落点 | 核查命令 |
| --- | --- | --- |
| 建立 Git 仓库 | §1 + 根目录 `.git/` | `git rev-parse --is-inside-work-tree` |
| 目录结构 | §5 + `main/` 等实际目录 | `find . -type d | sort` |
| 提交规范 | §3 + `CONTRIBUTING.md` | `git log -1 --format=%s` |
| main 可运行 / 评审 | §2 §4 | MR 记录、分支保护设置 |

## 9. 关联文件

- `CONTRIBUTING.md`（根目录，评审速查）
- `.gitignore`、`.gitattributes`
- `docs/02_分组与角色分工.md`
- `docs/04_需求追踪矩阵.md`
