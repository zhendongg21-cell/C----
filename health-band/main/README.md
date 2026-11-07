# main/ 源码骨架说明（P0）

> P0 阶段只建目录骨架，不写固件逻辑；P1 起在此填充代码。
> 分层与强制依赖见 [`docs/03_Git仓库与提交规范.md`](../docs/03_Git仓库与提交规范.md) §5。

| 目录 | 职责 | 依赖方向 |
| --- | --- | --- |
| `hal/` | I2CBus / SpiBus / Gpio / Ledc / Nvs；互斥锁保证 I2C 原子性 | 最底层，不依赖任何上层 |
| `driver/` | mpu6050 / max30102 / tmp117 / ft6236 / st7262 | → `hal` |
| `service/` | sensor_hub / step / sedentary / vital / fall / record_store / sync_client | → `driver`, `hal` |
| `app/` | app_controller / screen_manager / screens | → `service` |
| `ui/` | LVGL 页面、主题、字体子集 | → `app`（不得出现 `i2c_master_`） |
| `board/` | 引脚定义、板级初始化、版本号 | 被各层引用 |

**三条越层红线**：
1. `ui/` 不得出现 `i2c_master_` API
2. `driver/` 不得出现 `lv_` API
3. `service/` 不得直接操作 GPIO

**引脚/地址备忘**（原方案表 1-2）：I2C 0x68 (MPU6050) / 0x57 (MAX30102) / 0x48 (TMP117)；触摸 FT6236 0x38 或 XPT2046 SPI；屏 800×480 16 位色 SPI DMA。

**新增三项选做的预留落点**（P2 加分，均为选做，P0 不写固件逻辑）：
- FR-19 / FR-20（PC 端大模型健康问答 + 基于 PC 归档数据的 RAG 检索）：主跑在 PC 端，手环侧仅需 `FR-14/15` 的 `sync_client` 把记录补传归档；对应分支 `feature/llm-rag`。
- FR-21（手环操控鼠标 BLE HID）：预留 `service/` 的 `hid_mouse` 与 `hal/` 的 BLE 通道，默认关闭、二次确认、一键断开；对应分支 `feature/hid-mouse`。
- 三项不进入计步 / 久坐提醒主流程，可整块裁剪，不影响四条底线。判据见 [`docs/00_需求规格说明书.md`](../docs/00_需求规格说明书.md) §5.2 与 NFR-11/12。
