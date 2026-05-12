# 询价转让投资备忘录 Workflow（A股）

一个用 [Claude Code](https://claude.com/claude-code) 自动化生成 A 股**询价转让**投资备忘录的 workflow。

输入公司名+股票代码（+可选的卖方底价、Wind数据、研报PDF），自动产出一份 5 页以内的结构化 .docx 备忘录，覆盖估值、市场表现、季度业绩、股东资金面、舆情面、公司财务、亮点风险。

## 适用场景

- 创业板/科创板**询价转让**询价过程中的快速估值判断
- 卖方底价 vs 当前价的折扣率分析（场景A/B 见下）
- 转让价档位（折10% / 折20% / 折30% / 卖方底价）的 P/E、EV/EBITDA 推算
- 多平台舆情面分析（股吧 + 韭研公社 + 财富号 + 互动易）

## 仓库布局

```
xunjiazhuanrang-project/
├── CLAUDE.md             # 主提示词模板 (tracked, Claude 自动加载)
├── README.md             # 本文件 (tracked)
├── package.json          # node 依赖清单 (tracked)
├── scripts/              # 可复用模板脚本 (tracked) — 不含任何项目数据
│   ├── build_charts.js   # 图1/2/3 生成: PE Band / 同业对比 / 个股价格
│   ├── build_pdf.js      # PDF 生成 (puppeteer, 主交付物)
│   └── build_docx.js     # .docx 生成 (备选)
└── projects/             # 每个标的一个子目录, 内容全部 gitignored
    └── <公司名>/
        ├── data/         # 原始上传 (年报PDF + Wind Excel) — ignored
        ├── extracts/     # Claude 提取的中间分析 — ignored
        ├── build_data.js # 项目专用 — 含 inline 财务原值 (ignored)
        ├── data.json     # 单一数据真相源 (ignored)
        ├── chart*.png    # 三张图 (ignored)
        └── <公司>_<代码>_询价转让投资备忘录.{pdf,docx,html} (ignored)
```

**核心原则**：
- `CLAUDE.md` + `scripts/` 是**模板**，永远不含任何标的数据，可被 push
- `projects/<标的>/` 整个子目录**全部 ignored**——内含许可数据 + 内部投资判断，绝不公开

## 如何使用（手动模式）

1. 把本仓库 clone 到本地（或 `git clone --bare` 之后只拷贝模板）
2. `npm install`（安装 docx / puppeteer / chart.js 等依赖）
3. 启动 [Claude Code](https://claude.com/claude-code)
4. 在 `projects/` 下新建 `<公司名>/data/` 文件夹，把原始资料放进去：
   - 近三年年报 PDF
   - 最新季报/半年报 PDF
   - Wind 日度估值+股价+CSC可比 Excel
   - （可选）Wind consensus 截图、rating 截图、研报 PDF
5. 在 Claude Code 中输入：`公司名 股票代码 [卖方底价（可选）]`
   - 例：`华海诚科 688535 22.50`（场景A — 已拿到卖方底价）
   - 例：`华海诚科 688535`（场景B — 无底价，用当前价 10%/20%/30% 折扣档）
6. Claude 按三阶段流程执行：
   - **阶段一**：数据采集（PDF提取、Excel解析、Web 检索）
   - **阶段1.5**：写 `projects/<标的>/build_data.js`，执行后生成 `data.json` 并跑 23 项校验
   - **阶段二**：用 `scripts/build_charts.js`、`scripts/build_pdf.js`、`scripts/build_docx.js` 生成三张图 + PDF + docx

## 命令速查（阶段二）

```bash
# 数据 (项目专用)
node projects/<标的>/build_data.js

# 模板 (可复用)
node scripts/build_charts.js --project projects/<标的>
node scripts/build_pdf.js    --project projects/<标的>   # 主交付
node scripts/build_docx.js   --project projects/<标的>   # 备选
```

更新底价或补充数据后，只需重跑 `build_data.js` + `build_pdf.js` 即可。

## 核心设计

### 1. 数据来源分层

| 类别 | 来源 | 规则 |
|---|---|---|
| 财务数据 | 公司年报/季报原文PDF | 一手公告唯一来源，不允许从研报/新闻凑数 |
| 市场/交易数据 | 财经网站（北向、融资融券、股价等） | 可以从财经网站获取 |
| 舆情/情绪数据 | 股吧、韭研公社、财富号、互动易 | 严禁进入财务表，仅作情绪面参考 |

### 2. 估值表两个场景（v13.3）

- **场景A**（已拿到卖方底价）：估值表展示「当前价 / 折10% / 折20% / 折30% / 卖方底价」，**跳过任何低于底价的折扣档**（这些价格不可能成交）
- **场景B**（无底价）：固定三档「当前价 / 折10% / 折20% / 折30%」

监管层面的"机制底价"（20日均价×70%）在 v13.3 已弃用——实际成交折扣集中在 3%-25%，30%+ 极少出现，机制底价从未 binding。

### 3. 数据中间层（data.json）

财务数字由项目专用脚本（`projects/<标的>/build_data.js`）统一计算并输出 `data.json`。文档生成模板（`scripts/build_pdf.js` / `build_docx.js` / `build_charts.js`）从 JSON 读取，**禁止硬编码任何财务数字**——杜绝手工复制粘贴错误。

> 注：本仓库的 build 脚本使用 **Node.js**（不是 Python），原因是 Windows 默认无 Python 解释器但 Node 通用，且 docx/PDF/Chart 生态在 Node 上更成熟。

### 4. 自动化校验

阶段1.5 包含：
- 毛利率口径一致性（三年同口径）
- EBITDA 公式一致性（营业利润 + D&A + 利息支出）
- P/E 口径分离（FY(N)A 用年报EPS，FY(N+1)E 用consensus）
- 三年趋势合理性检查

### 5. 输出格式

主交付物为 **PDF**（puppeteer 渲染，绕过 Word 字体/兼容问题），同时保留 **.docx** 作为可编辑备选。

- A4 单页排版，5 页以内
- Calibri 字体 → 中文 fallback 到微软雅黑/苹方
- 4 模块：A 估值 / B 季度+News Run+舆情 / C 股东资金面 / D 公司+财务+亮点风险
- 投资亮点 vs 风险表格化对照

### 6. 数据安全（git 角度）

`.gitignore` 已配置为**所有项目专用文件全部排除**：
- `projects/*/data/`（含 Wind/CSC 许可数据）
- `projects/*/extracts/`（含内部分析）
- `projects/*/build_data.js`（含 inline 原始财报数据）
- `projects/*/data.json`、`chart*.png`
- `projects/*/*.{docx,pdf,html}`（含投资判断输出）

`git status` 在新增标的后**应保持干净**（只有 `scripts/` 或 `CLAUDE.md` 的改动可见）。push 时不会泄露任何标的级数据。

## 当前版本

**v13.2**——Module A 新增图3：标的个股单独价格图（绝对价格+均线+成交量+关键价位标注）

历史版本：v13.1 舆情多源采集+自适应输出 / v13 web_search主路径 / v12 双底价+亮点风险表格化

历史迭代见 `CLAUDE.md` 尾部的版本特性清单。

## 致谢/灵感来源

- [Anthropic Claude Code](https://claude.com/claude-code)：执行环境
- [OpenCLI](https://github.com/jackwener/OpenCLI)：曾尝试集成的舆情采集工具（实测不适合本场景，已降级为 web_search 主路径）

## License

[MIT](./LICENSE)

## 免责声明

本 workflow 仅供 A 股二级市场投研流程参考。**不构成任何投资建议**。使用本 workflow 产出的备忘录所做的投资决策，由使用者自行承担风险。
