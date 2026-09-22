# 物流运营监控看板（monitor）— 部署与维护手册

> 本看板已按新范式改造：**PC 端左侧导航布局 + 每日 17:00 自动从最新《产品跟踪表》刷新，不再手动上传双份对比表；仓配模块已移除。**

---

## 1. 本次改造要点（对比旧版）

| 维度 | 旧版 | 新版 |
|------|------|------|
| 布局 | 手机版 tab 横滑 | **PC 端左侧固定导航**（日度监控 / 查验监控 两个分组） |
| 数据源 | 浏览器上传 **2 份** xlsx（今日+昨日），实时解析并对比 | **服务端 Python 预生成 `data.json`**，前端 `fetch` 后渲染；**单表快照，不做双日对比** |
| 更新方式 | 人工上传 | **每日 17:00 计划任务**跑 `run_daily_monitor.py` 自动生成并推送 |
| 仓配模块 | 有（warehouse.js） | **已砍掉**（daily.js 无仓配耦合，可安全移除） |
| 其余模块 | 日度 6 项 + 查验 5 项 + 滞留 | **全部平移**，计算口径与旧版一致 |

业务模块清单（均保留）：
- **日度监控**：日度总览 / 在途 / 时效 SLA / 异常 / 渠道占比 / 明日预警
- **查验监控**：查验总览 / 渠道 / 代理 / 下钻 / 预警 / 🚚 滞留

---

## 2. 架构

```
D:/素芸/跟踪表/  (最新《产品跟踪表》xlsx)
        │  (每日 17:00 计划任务)
        ▼
run_daily_monitor.py
   ├─ find_latest_excel()        按修改时间取最新
   ├─ build_monitor_json.py      openpyxl 读取 → 原始行透传 data.json
   ├─ 校验（记录数 ≥ 100）       异常则退出并写日志
   ├─ 同步快照 source/current.xlsx
   └─ git push → GitHub Pages
        │
        ▼
monitor/index.html  ──fetch('data.json')──▶  chayan.js + daily.js 完毕派生计算
```

### 关键设计：原始行透传
- Python **只做「读取 + 列头原样保留 + 值清洗」**，把每一行输出为
  `{"列头": 清洗值, ...}` 的数组，外加 `meta`（来源文件 / sheet / 生成时间 / 行数 / 数据基准日）。
- 复杂物流算法（双维查验、事件链、滞留、周月聚合、SLA）**仍由前端 `chayan.js` / `daily.js` 完成**，与旧版完全一致，避免 Python 侧重写导致口径漂移。
- 日期统一清洗为 `YYYY-MM-DD`，整数浮点 → `int`，空 / `NaN` → `''`。

---

## 3. 目录文件

| 文件 | 作用 |
|------|------|
| `index.html` | PC 端页面（左侧导航 + 主内容区），自托管 echarts / xlsx 库 |
| `chayan.js` | 查验 / 滞留 / 总览主引擎；`loadFromServer()` 拉取 data.json；应急上传兜底 |
| `daily.js` | 日度 6 模块渲染（单表快照模式） |
| `chayan.css` | 样式（含 PC 侧边栏布局） |
| `build_monitor_json.py` | ETL：xlsx → `data.json`（原始行透传），仅依赖 openpyxl |
| `run_daily_monitor.py` | 每日编排：定位最新表 → 生成 → 校验 → 同步快照 → git 推送 |
| `source/current.xlsx` | 当日所用 Excel 快照（自动生成，供追溯） |
| `last_run.json` / `daily_log.txt` | 运行结果 / 日志（自动生成） |
| `data.json` | **运行时生成**，不入库历史（每次刷新覆盖） |
| `.nojekyll` | GitHub Pages 跳过 Jekyll 处理 |

---

## 4. 每日 17:00 自动刷新部署（Windows 计划任务）

### 4.1 依赖安装（仅首次）
```bat
:: 用户机器若无 openpyxl（脚本不依赖 pandas）
python -m pip install openpyxl
```
> 本机测试用的隔离 venv：`C:\Users\yanmi\.workbuddy\binaries\python\envs\monitor_etl\`（openpyxl 3.1.5）。
> 生产机请确认 `run_daily_monitor.py` 调用的 `python` 能 `import openpyxl`。

### 4.2 配置计划任务
1. `Win + R` → `taskschd.msc` → 创建**基本任务** / 或「创建任务」(进阶)。
2. 触发器：**每日 17:00**。
3. 操作 → 启动程序：
   - 程序：`python`（或绝对路径 `C:\...\python.exe`）
   - 参数：`C:\...\checklist\monitor\run_daily_monitor.py`
   - 起始于：`C:\...\checklist\monitor`
4. 条件：取消「只有在计算机使用交流电源时才启动此任务」（笔记本防漏跑）。
5. 设置：勾选「如果任务失败，重新启动」，最长重试 3 次。

### 4.3 手动运行 / 调试
```bat
:: 自动找最新表
python run_daily_monitor.py
:: 指定某天表
python run_daily_monitor.py --input "D:/素芸/跟踪表/0905 产品跟踪表.xlsx"
:: 单独只生成 data.json（不推送）
python build_monitor_json.py --input "D:/素芸/跟踪表/0905 产品跟踪表.xlsx"
```

### 4.4 退出码
`0`=成功 · `1`=未找到 Excel · `2`=build 失败 · `3`=校验失败(<100行) · `4`=git 推送失败

### 4.5 首次部署（必须手动 commit 一次）
`run_daily_monitor.py` 的自动推送**只含运行时文件**（`index.html` `data.json` `chayan.js` `daily.js` `chayan.css` `.nojekyll`），
**不会**推送本次新增/改写的**源文件**（`build_monitor_json.py` `run_daily_monitor.py` `README_monitor.md`）。
因此改版后需先在仓库里手动提交一次全部变更，之后每日脚本才只需推送运行时文件：

```bat
git add build_monitor_json.py run_daily_monitor.py README_monitor.md .nojekyll ^
        index.html chayan.js daily.js chayan.css
git commit -m "monitor: PC 端改造 + 每日自动 ETL + 砍仓配"
git push origin HEAD
```

> 注意：`data.json` 由每日脚本生成并自动推送；首次手动 commit 时仓库里可以没有 data.json（页面会显示错误横幅 + 应急上传，待首次 17:00 脚本跑出后正常）。

---

## 5. data.json 字段约定（原始行透传）

`data.json` 结构：
```json
{
  "meta": {
    "sourceFile": "0905 产品跟踪表.xlsx",
    "sheet": "空运+快递+陆运",
    "generatedAt": "2026-09-05 17:01:23",
    "rowCount": 1234,
    "dataDate": "2026-09-05"
  },
  "rows": [
    {"分出仓单号":"...", "客户":"...", "国家":"美国", "类型":"海运",
     "素芸物流渠道":"美国海运专线", "代理":"明捷", "仓库出货日期":"2026-09-01",
     "实际签收时间\n（当地时间）":"2026-09-12", "状态备注":"...", ...},
    ...
  ]
}
```

- **列头 = Excel 表头原样**（含多行表头如 `实际签收时间\n（当地时间）` 也会保留，前端按同一字符串 lookup）。
- **前端依赖的关键列**（缺失会被容错为「未知」/空，不报错）：
  `分出仓单号` `主出仓单号` `客户` `国家` `类型` `素芸物流渠道` `代理` `代理渠道`
  `产品属性` `仓库出货日期` `起运日期` `入承运商仓日期` `到港日期` `末端提取日`
  `实际签收时间\n（当地时间）` `货物状态` `状态备注` `方数CBM` `毛重` `件数` `箱数` `操作负责人` `销售名`
- `meta.dataDate` 取「仓库出货日期」列最大日期，作为看板基准日展示。

---

## 6. 页面交互

- **自动加载**：打开页面即 `fetch('data.json', {cache:'no-store'})`；失败显示错误横幅。
- **🔄 重新加载**：侧栏按钮，重新拉取 data.json。
- **📥 导出**：按当前数据导出 Excel。
- **🔗 分享链接**：生成只读快照链接（客户/领导免上传查看）。
- **应急单文件上传**：当 `data.json` 自动加载失败时，横幅下方可直接选最新《产品跟踪表》xlsx
  应急加载（**单文件，无昨日对比**），用于脚本未跑或推送异常时的临时补救。

---

## 7. 运维 / 排错

| 现象 | 原因 / 处理 |
|------|------|
| 页面提示「无法自动加载 data.json」 | 17 点脚本未跑 / 推送未成功 / 文件名不在匹配清单。先用应急上传兜底，再查 `daily_log.txt`。 |
| `run_daily_monitor.py` 退出码 1 | 未在 `D:/素芸/跟踪表` 找到含「产品跟踪表/跟踪表」的 xlsx。确认目录与文件名。 |
| 退出码 3（校验失败） | data.json 行数 < 100，疑似 Excel 结构异常或表头错乱，检查源表。 |
| git 推送失败（退出码 4） | 本机用 WorkBuddy 自带 PortableGit（`C:\Users\yanmi\.workbuddy\vendor\PortableGit\cmd\git.exe`）。需 SSH key / 凭证可用；沙箱网络不可达时请在受限网络外机器执行推送。 |
| 找不到 git 仓库 | 脚本会跳过推送、仅本地生成 data.json，不影响本地预览。 |
| 列头改名导致某模块空白 | 原始行透传依赖固定列头文本；若《产品跟踪表》改表头，需同步 `chayan.js`/`daily.js` 的 lookup 字符串。 |

---

## 8. 已从旧版移除

- `warehouse.js`（仓配模块）及 `index.html` 中对应 tab / 面板 / `<script>` 引用。
- 浏览器「上传两份对比表」上传面板（`#uploadPanel`）及双日对比逻辑（`Daily.setData(rows, null)` 不再传昨日）。
- 首页手机版 tab 横滑布局 → 改为 PC 左侧导航。

---

_改造完成于 2026-09-22，对齐 `logistics-dashboard` 范式的 PC 端 + 每日自动刷新模式。_
