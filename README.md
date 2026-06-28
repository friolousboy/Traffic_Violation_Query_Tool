# 违章查询工具 / Traffic Violation Query Tool

[中文](#中文) | [English](#english)

---

## 中文

### 简介

基于 Selenium 的四川省交通违章批量查询工具，自动从 [sc.122.gov.cn](https://sc.122.gov.cn) 查询车辆违章记录，并将结果导出为 Excel 文件。

### 功能特性

- **批量查询** — 从 Excel 文件中读取车牌号列表，逐一自动查询
- **智能车型匹配** — 自动识别小型汽车、小型新能源汽车、大型汽车
- **多页结果支持** — 单辆车多条违章时自动翻页抓取
- **断点保护** — 每处理 20 辆车自动保存中间结果，防止数据丢失
- **容错重试** — 页面元素失效时自动重试（最多 3 次）
- **日志记录** — 自动将控制台输出保存到 `log.txt`

### 环境要求

| 依赖 | 版本 |
|------|------|
| Python | 3.7+ |
| Microsoft Edge | 最新版 |
| selenium | 4.x |
| pandas | 1.x |
| openpyxl | 3.x |

### 安装

```bash
pip install selenium pandas openpyxl
```

> Edge 浏览器驱动（msedgedriver）由 Selenium 自动管理，无需手动安装。

### 使用方法

#### 1. 准备输入文件

在项目根目录下放置 `example.xlsx`，包含一个名为 **"车牌"** 的列：

| 车牌 |
|------|
| 川AXN557 |
| 川A3EL30 |
| 川AD12345 |
| ... |

#### 2. 运行脚本

```bash
python main.py
```

#### 3. 操作流程

1. 脚本自动打开 Edge 浏览器，跳转到 122 统一身份认证页面
2. 页面显示二维码后，使用 **交管12123 APP** 扫码登录
3. 登录成功后，点击单位入口进入
4. 脚本自动导航至违章查询页面，逐条处理车牌
5. 查询结果实时输出到控制台，并最终保存到 `output.xlsx`

#### 4. 查看结果

结果保存在 `output.xlsx` 中，包含以下字段：

| 字段 | 说明 |
|------|------|
| 车牌 | 车牌号 |
| 时间 | 违章时间 |
| 地点 | 违章地点 |
| 违章行为 | 具体违章内容 |
| 罚款 | 罚款金额（元） |
| 记分 | 扣分 |
| 处理状态 | 已处理 / 未处理 |

### 配置说明

编辑 `main.py` 顶部的常量即可调整行为：

```python
HEADLESS = False          # 设为 True 启用无头模式（需有效 cookies）
SAVE_INTERVAL = 20        # 每处理 N 辆车保存一次
DEFAULT_TIMEOUT = 5       # 页面元素等待超时（秒）
LOGIN_TIMEOUT = 300       # 扫码登录超时（秒）
PAGE_STABLE_WAIT = 5      # 查询后等待页面稳定（秒）
MAX_ROW_RETRIES = 3       # 单条记录重试次数
MAX_PLATE_RETRIES = 3     # 单车重试次数
```

### 注意事项

- 首次使用需要手动扫码登录，登录态有效期有限
- 查询频率过高可能触发反爬机制，建议适当增加 `PAGE_STABLE_WAIT`
- 仅支持四川省（sc.122.gov.cn）的违章查询
- 运行期间请勿关闭浏览器窗口或手动操作页面

### 目录结构

```
selenium/
├── main.py              # 主程序
├── example.xlsx         # 输入文件（车牌列表）
├── output.xlsx          # 输出文件（查询结果）
├── cookies.json         # 登录凭证缓存
└── log.txt              # 运行日志
```

---

## English

### Overview

A Selenium-based tool for batch querying traffic violation records from the Sichuan province official website [sc.122.gov.cn](https://sc.122.gov.cn) (China). Reads license plates from an Excel file, queries each one, and exports the results.

### Features

- **Batch Processing** — Query multiple license plates from an Excel input file
- **Smart Vehicle Type Detection** — Automatically tries passenger car, new energy vehicle, or large vehicle
- **Multi-page Results** — Handles pagination when a vehicle has multiple violations
- **Checkpoint Saving** — Saves intermediate results every 20 plates to prevent data loss
- **Automatic Retries** — Retries up to 3 times on stale elements or failed clicks
- **Console Logging** — Duplicates all console output to `log.txt`

### Requirements

| Dependency | Version |
|------------|---------|
| Python | 3.7+ |
| Microsoft Edge | Latest |
| selenium | 4.x |
| pandas | 1.x |
| openpyxl | 3.x |

### Installation

```bash
pip install selenium pandas openpyxl
```

> The Edge WebDriver (msedgedriver) is auto-managed by Selenium — no manual setup needed.

### Usage

#### 1. Prepare Input File

Place `example.xlsx` in the project root with a column named **"车牌"**:

| 车牌 |
|------|
| 川AXN557 |
| 川A3EL30 |
| 川AD12345 |
| ... |

#### 2. Run

```bash
python main.py
```

#### 3. Workflow

1. The script launches Edge and navigates to the 122 unified authentication page
2. Scan the QR code with the **交管12123 (Traffic Management 12123)** mobile app
3. After login, click your organization entry
4. The script auto-navigates to the violation query page and processes each plate
5. Results are streamed to the console and saved to `output.xlsx`

#### 4. Output

Results are written to `output.xlsx` with the following columns:

| Column | Description |
|--------|-------------|
| 车牌 | License plate number |
| 时间 | Violation time |
| 地点 | Violation location |
| 违章行为 | Offense description |
| 罚款 | Fine amount (CNY) |
| 记分 | Demerit points |
| 处理状态 | Status (processed / unprocessed) |

### Configuration

Edit the constants at the top of `main.py`:

```python
HEADLESS = False          # Set to True for headless mode (requires valid cookies)
SAVE_INTERVAL = 20        # Save results every N plates
DEFAULT_TIMEOUT = 5       # Element wait timeout (seconds)
LOGIN_TIMEOUT = 300       # QR code scan timeout (seconds)
PAGE_STABLE_WAIT = 5      # Wait time after query (seconds)
MAX_ROW_RETRIES = 3       # Per-row retry limit
MAX_PLATE_RETRIES = 3     # Per-plate retry limit
```

### Notes

- Manual QR code login is required on first run; sessions expire after a period
- Excessive query frequency may trigger anti-bot measures — increase `PAGE_STABLE_WAIT` if needed
- Only supports Sichuan province (sc.122.gov.cn)
- Do not close the browser window or interact with the page while the script is running

### Directory Structure

```
selenium/
├── main.py              # Main script
├── example.xlsx         # Input file (plate list)
├── output.xlsx          # Output file (query results)
├── cookies.json         # Session cookie cache
└── log.txt              # Runtime log
```
