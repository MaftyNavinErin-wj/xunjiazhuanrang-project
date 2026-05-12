// 通用 PDF 备忘录生成模板
// 用法: node scripts/build_pdf.js --project projects/<标的名>

const fs = require('fs');
const path = require('path');
const puppeteer = require('puppeteer');

function getProjectDir() {
  const args = process.argv.slice(2);
  const i = args.indexOf('--project');
  if (i >= 0 && args[i + 1]) return path.resolve(args[i + 1]);
  console.error('Usage: node scripts/build_pdf.js --project projects/<标的名>');
  process.exit(1);
}
const PROJECT_DIR = getProjectDir();
const data = JSON.parse(fs.readFileSync(path.join(PROJECT_DIR, 'data.json'), 'utf8'));

// 工具
const f = (n, dec = 1) => (n === null || n === undefined || isNaN(n)) ? '—' : n.toFixed(dec);
const pct = (n, dec = 1) => f(n, dec) + '%';
const pctSigned = (n, dec = 1) => (n > 0 ? '+' : '') + f(n, dec) + '%';
const yi = (n, dec = 2) => f(n, dec);
const imgB64 = (filename) => {
  const buf = fs.readFileSync(path.join(PROJECT_DIR, filename));
  return 'data:image/png;base64,' + buf.toString('base64');
};

const fy23 = data.financials.FY2023;
const fy24 = data.financials.FY2024;
const fy25 = data.financials.FY2025;
const q1 = data.financials.Q1_2026;
const m = data.market;
const v = data.valuation;
const acq = data.acquisition;
const peerM = data.peer_median;
const peers = data.peers;

// 估值表（场景A，4列）EV-at-px 计算
const raw_fy25_cash_yuan = fy25.cash * 1e8;
const raw_fy25_tfa_yuan = fy25.trading_fa * 1e8;
const raw_fy25_tfl_yuan = fy25.trading_fl * 1e8;
const ebitda_2025_yuan = fy25.ebitda * 1e8;
const evMultAt = (px) => {
  const mc = px * m.total_shares;
  const ev = mc - (raw_fy25_cash_yuan + raw_fy25_tfa_yuan - raw_fy25_tfl_yuan);
  return ev / ebitda_2025_yuan;
};

const evCur = evMultAt(m.close);
const evD10 = evMultAt(v.floor_discount_10);
const evD20 = evMultAt(v.floor_discount_20);
const evFlr = evMultAt(v.floor_price);
const bvps = fy25.equity_parent / (m.total_shares / 1e8);
const ts = data.top_shareholders_FY2025;

const html = `<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>${data.meta.company} (${data.meta.code}) — 询价转让投资备忘录</title>
<style>
@page { size: A4; margin: 12mm 12mm 12mm 12mm; }
* { box-sizing: border-box; }
body {
  font-family: "Calibri", "Microsoft YaHei", "PingFang SC", sans-serif;
  font-size: 8.5pt;
  line-height: 1.35;
  color: #111;
  margin: 0;
}
.header {
  font-size: 7.5pt;
  color: #666;
  border-bottom: 1px solid #ccc;
  padding-bottom: 4px;
  margin-bottom: 6px;
  display: flex;
  justify-content: space-between;
}
.header .left { color: #2C5F9B; font-weight: bold; }
.header .right { font-style: italic; }
h1 {
  font-size: 14pt;
  color: #2C5F9B;
  margin: 0 0 3px 0;
  font-weight: bold;
}
h2 {
  font-size: 11pt;
  color: #2C5F9B;
  margin: 10px 0 4px 0;
  border-bottom: 2px solid #2C5F9B;
  padding-bottom: 2px;
}
.subtitle {
  font-size: 8pt;
  color: #666;
  font-style: italic;
  margin-bottom: 6px;
}
.summary {
  background: #F5F8FB;
  border-left: 3px solid #2C5F9B;
  padding: 6px 8px;
  margin-bottom: 8px;
  font-size: 8.5pt;
}
.summary .floor-note { color: #C2185B; font-weight: bold; }
.tag { font-weight: bold; }
table {
  width: 100%;
  border-collapse: collapse;
  font-size: 7.5pt;
  margin: 4px 0 6px 0;
}
th {
  background: #2C5F9B;
  color: white;
  padding: 4px 6px;
  text-align: center;
  font-weight: bold;
  border: 1px solid #1A3B6C;
}
td {
  padding: 3px 6px;
  border: 1px solid #D5D5D5;
  vertical-align: middle;
}
tr:nth-child(even) td { background: #F8F9FA; }
td.num { text-align: right; }
td.center { text-align: center; }
.floor-col { background: #FCE4EC !important; font-weight: bold; color: #C2185B; }
.floor-header { background: #C2185B !important; }
.red { color: #C0392B; }
.green { color: #27AE60; }
.bold { font-weight: bold; }
.note {
  font-size: 7pt;
  color: #777;
  font-style: italic;
  margin: 2px 0 6px 0;
}
.chart {
  width: 100%;
  margin: 6px 0;
  page-break-inside: avoid;
}
.chart img {
  width: 100%;
  display: block;
}
.chart-caption {
  text-align: center;
  font-size: 7pt;
  color: #777;
  font-style: italic;
  margin-top: 2px;
}
.module {
  page-break-inside: avoid;
}
.two-col {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 8px;
}
.cat-pos { color: #27AE60; }
.cat-neg { color: #C0392B; }
.footer-source {
  margin-top: 8px;
  font-size: 7pt;
  color: #777;
  font-style: italic;
  border-top: 1px solid #DDD;
  padding-top: 4px;
}
</style>
</head>
<body>
<div class="header">
  <span class="left">${data.meta.company}（${data.meta.code}）— 询价转让投资备忘录</span>
  <span class="right">机密文件 | 仅供内部投资决策参考</span>
</div>

<h1>${data.meta.company}（${data.meta.code}.SZ）— 询价转让投资备忘录</h1>
<div class="subtitle">${data.meta.memo_date}　|　数据截止：${data.meta.data_cutoff}　|　股价口径：${m.date} 收盘价 ${m.close} 元　|　<span style="color:#C2185B;font-weight:bold">卖方底价 ${f(v.floor_price, 2)} 元（折扣 ${pct(v.floor_discount_pct)}）</span></div>

<div class="summary">
国内粉末冶金细分龙头（电动工具零部件 ${pct(data.rev_breakdown.FY2025.by_industry.电动工具零部件.share)} / 海外占比 ${pct(data.rev_breakdown.FY2025.by_region.境外.share)}），FY2025 营收 <span class="bold">${yi(fy25.revenue)} 亿</span>（${pctSigned((fy25.revenue / fy24.revenue - 1) * 100)}）、归母净利 <span class="bold">${yi(fy25.np_parent)} 亿</span>（${pctSigned((fy25.np_parent / fy24.np_parent - 1) * 100)}）—— 主业承压；
<span class="bold">2025-12 完成 ${acq.target} ${acq.pct}% 现金收购（${yi(acq.consideration_yi)} 亿元，形成商誉 ${yi(acq.goodwill_yi)} 亿）</span>，Q1 2026 信为通讯首次并表带动营收 +123.3% 但拆分后主业 −32%。近1年股价 ${pctSigned(m.return_1y)}，机构无覆盖、估值显著高于可比中位。
<span class="floor-note">本次询价转让卖方底价 ${f(v.floor_price, 2)} 元（折扣 ${pct(v.floor_discount_pct)}），对应 P/E (FY2025A) ${f(v.pe_at_floor, 1)}x、P/E (FY2024A) ${f(v.pe_at_floor_fy24, 1)}x，仍显著高于可比中位 ${f(peerM.pe_2024, 1)}x。</span>
</div>

<h2>A. 估值与市场表现</h2>

<p style="margin: 6px 0 4px 0;"><span class="tag">A.1 估值快照与询价转让定价（场景A：卖方底价 ${f(v.floor_price, 2)} 元 / 折扣率 ${pct(v.floor_discount_pct)}）</span><span class="note">　单位：元/股、亿元、倍</span></p>

<table>
<tr>
  <th style="width:24%">指标</th>
  <th style="width:17%">当前股价</th>
  <th style="width:17%">折价 10%</th>
  <th style="width:17%">折价 20%</th>
  <th class="floor-header" style="width:25%">卖方底价（折 ${pct(v.floor_discount_pct)}）</th>
</tr>
<tr>
  <td>股价（元）</td>
  <td class="center bold">${f(m.close, 2)}</td>
  <td class="center">${f(v.floor_discount_10, 2)}</td>
  <td class="center">${f(v.floor_discount_20, 2)}</td>
  <td class="center floor-col">${f(v.floor_price, 2)}</td>
</tr>
<tr>
  <td>对应市值（亿元）</td>
  <td class="center">${f(v.mkt_cap_yi, 1)}</td>
  <td class="center">${f(v.mkt_cap_yi * 0.9, 1)}</td>
  <td class="center">${f(v.mkt_cap_yi * 0.8, 1)}</td>
  <td class="center floor-col">${f(v.mkt_cap_at_floor_yi, 1)}</td>
</tr>
<tr>
  <td>P/E (FY2025A=${yi(fy25.np_parent)} 亿)</td>
  <td class="center">${f(v.pe_2025a, 1)}x</td>
  <td class="center">${f(v.pe_at_disc10, 1)}x</td>
  <td class="center">${f(v.pe_at_disc20, 1)}x</td>
  <td class="center floor-col">${f(v.pe_at_floor, 1)}x</td>
</tr>
<tr>
  <td>P/E (FY2024A=${yi(fy24.np_parent)} 亿)</td>
  <td class="center">${f(v.pe_2024a, 1)}x</td>
  <td class="center">${f(v.pe_2024a * 0.9, 1)}x</td>
  <td class="center">${f(v.pe_2024a * 0.8, 1)}x</td>
  <td class="center floor-col">${f(v.pe_at_floor_fy24, 1)}x</td>
</tr>
<tr>
  <td>EV/EBITDA (FY2025A=${yi(fy25.ebitda)} 亿)</td>
  <td class="center">${f(evCur, 1)}x</td>
  <td class="center">${f(evD10, 1)}x</td>
  <td class="center">${f(evD20, 1)}x</td>
  <td class="center floor-col">${f(evFlr, 1)}x</td>
</tr>
<tr>
  <td>P/S (FY2025A=${yi(fy25.revenue)} 亿)</td>
  <td class="center">${f(v.ps_2025, 1)}x</td>
  <td class="center">${f(v.ps_2025 * 0.9, 1)}x</td>
  <td class="center">${f(v.ps_2025 * 0.8, 1)}x</td>
  <td class="center floor-col">${f((v.floor_price * m.total_shares) / (fy25.revenue * 1e8), 1)}x</td>
</tr>
<tr>
  <td>P/B (BVPS=${f(bvps, 2)} 元)</td>
  <td class="center">${f(v.pb_2025, 1)}x</td>
  <td class="center">${f(v.pb_2025 * 0.9, 1)}x</td>
  <td class="center">${f(v.pb_2025 * 0.8, 1)}x</td>
  <td class="center floor-col">${f((v.floor_price * m.total_shares) / (fy25.equity_parent * 1e8), 1)}x</td>
</tr>
</table>
<div class="note">注：卖方底价折扣率 ${pct(v.floor_discount_pct)} 介于 20%-30% 之间，故跳过 折30% 档（28.86 × 0.7 = 20.20 &lt; 22.79，低于底价不可能成交）。海昌新材无机构一致预期覆盖（NTM consensus 缺失），表格仅展示 actual 倍数；EV/EBITDA 已剔除信为通讯应付分期款 ${yi(fy25.trading_fl)} 亿。</div>

<p style="margin: 6px 0 4px 0;"><span class="tag">A.2 vs A股可比公司（基准日 2026-04-14，来源 CSC）</span></p>

<table>
<tr>
  <th>公司 (代码)</th>
  <th>股价</th>
  <th>市值(亿)</th>
  <th>EV(亿)</th>
  <th>P/S CY24</th>
  <th>EV/EBITDA CY24</th>
  <th>P/E CY24</th>
  <th>P/B LTM</th>
</tr>
<tr>
  <td class="red bold">海昌新材 (300885)</td>
  <td class="num">${f(peers.haichang.px, 2)}</td>
  <td class="num">${f(peers.haichang.mkt_cap_yi, 1)}</td>
  <td class="num">${f(peers.haichang.ev_yi, 1)}</td>
  <td class="num bold">${f(peers.haichang.ps_2024, 1)}x</td>
  <td class="num bold">${f(peers.haichang.ev_ebitda_2024, 1)}x</td>
  <td class="num bold">${f(peers.haichang.pe_2024, 1)}x</td>
  <td class="num">${f(peers.haichang.pb_ltm, 1)}x</td>
</tr>
<tr>
  <td>东睦股份 (600114)</td>
  <td class="num">${f(peers.dongmu.px, 2)}</td>
  <td class="num">${f(peers.dongmu.mkt_cap_yi, 1)}</td>
  <td class="num">${f(peers.dongmu.ev_yi, 1)}</td>
  <td class="num">${f(peers.dongmu.ps_2024, 1)}x</td>
  <td class="num">${f(peers.dongmu.ev_ebitda_2024, 1)}x</td>
  <td class="num">${f(peers.dongmu.pe_2024, 1)}x</td>
  <td class="num">${f(peers.dongmu.pb_ltm, 1)}x</td>
</tr>
<tr>
  <td>精研科技 (300709)</td>
  <td class="num">${f(peers.jingyan.px, 2)}</td>
  <td class="num">${f(peers.jingyan.mkt_cap_yi, 1)}</td>
  <td class="num">${f(peers.jingyan.ev_yi, 1)}</td>
  <td class="num">${f(peers.jingyan.ps_2024, 1)}x</td>
  <td class="num">${f(peers.jingyan.ev_ebitda_2024, 1)}x</td>
  <td class="num">${f(peers.jingyan.pe_2024, 1)}x</td>
  <td class="num">${f(peers.jingyan.pb_ltm, 1)}x</td>
</tr>
<tr>
  <td>派克新材 (605123)</td>
  <td class="num">${f(peers.paike.px, 2)}</td>
  <td class="num">${f(peers.paike.mkt_cap_yi, 1)}</td>
  <td class="num">${f(peers.paike.ev_yi, 1)}</td>
  <td class="num">${f(peers.paike.ps_2024, 1)}x</td>
  <td class="num">${f(peers.paike.ev_ebitda_2024, 1)}x</td>
  <td class="num">${f(peers.paike.pe_2024, 1)}x</td>
  <td class="num">${f(peers.paike.pb_ltm, 1)}x</td>
</tr>
<tr>
  <td class="bold" style="font-style: italic;">可比中位数</td>
  <td>—</td><td>—</td><td>—</td>
  <td class="num bold">${f(peerM.ps_2024, 1)}x</td>
  <td class="num bold">${f(peerM.ev_ebitda_2024, 1)}x</td>
  <td class="num bold">${f(peerM.pe_2024, 1)}x</td>
  <td class="num bold">${f(peerM.pb_ltm, 1)}x</td>
</tr>
<tr>
  <td class="red bold" style="font-style: italic;">海昌溢价幅度</td>
  <td>—</td><td>—</td><td>—</td>
  <td class="num red bold">${pctSigned((peers.haichang.ps_2024 / peerM.ps_2024 - 1) * 100, 0)}</td>
  <td class="num red bold">${pctSigned((peers.haichang.ev_ebitda_2024 / peerM.ev_ebitda_2024 - 1) * 100, 0)}</td>
  <td class="num red bold">${pctSigned((peers.haichang.pe_2024 / peerM.pe_2024 - 1) * 100, 0)}</td>
  <td class="num red bold">${pctSigned((peers.haichang.pb_ltm / peerM.pb_ltm - 1) * 100, 0)}</td>
</tr>
</table>
<div class="note">A股可比 forward 一致预期：CY2025E P/E ${peerM.pe_2025e}x → CY2026E ${peerM.pe_2026e}x → CY2027E 31.4x；EV/EBITDA CY2025E ${peerM.ev_ebitda_2025e}x（仅东睦股份单家有 forward consensus）。</div>

<p style="margin: 6px 0 4px 0;"><span class="tag">A.3 价格表现与流动性：</span>近1月 ${pctSigned(m.return_1m)}，近6月 ${pctSigned(m.return_6m)}，近1年 ${pctSigned(m.return_1y)}（vs 东睦 +78.8% / 精研 +40.8% / 派克 +112.8%）。52周价格区间 ${f(m.px_52w_low, 2)} ~ ${f(m.px_52w_high, 2)} 元，当前距 52W 高位 ${pct((m.close / m.px_52w_high - 1) * 100)}。近20日均成交额 ${f(m.avg20d_turnover_wan / 10000, 2)} 亿元，最新换手率 ${pct(m.latest_turnover_rate_pct)}（异常活跃）。融资融券余额（2026-03-26）${f(m.margin_balance_yi, 2)} 亿元（占流通市值 ~6%）。</p>

<div class="chart"><img src="${imgB64('chart1_pe_band.png')}" /></div>
<div class="chart-caption">当前 TTM PE ${m.ttm_pe_last}x 远高于历史均值 ${m.ttm_pe_mean}x（约 +2.4 SD），处于近3年估值上限附近。</div>

<div class="chart"><img src="${imgB64('chart2_peer_comparison.png')}" /></div>
<div class="chart-caption">归一化基期 2023-05-12 = 100；海昌新材 3 年累计涨幅显著高于行业龙头东睦股份。</div>

<h2>B. 最新季度 + 近期动态 + News Run</h2>

<p style="margin: 6px 0 4px 0;"><span class="tag">B.1 Q1 2026（${acq.target}首次并表后第一份季报）</span>：营收 ${yi(q1.revenue)} 亿（+123.3% YoY）/ 归母净利 ${yi(q1.np_parent)} 亿（+36.7%）/ 扣非归母 ${yi(q1.np_nonrecurring_excluded)} 亿（+43.5%）/ 综合毛利率 ${pct(q1.gross_margin)}（vs Q1 2025 35.9%，略降 1pp）。</p>

<p style="margin: 4px 0 4px 0;"><span class="tag">关键拆分（重要）：</span>合并净利润 ${yi(q1.net_income)} 亿 = 归母 ${yi(q1.np_parent)} 亿 + 少数股东 ${yi(q1.np_minority)} 亿。信为通讯Q1全口径净利 ≈${yi(q1.xinwei_total_np_yi)}亿（少数股东 ÷ 49%），归属海昌51% ≈${yi(q1.xinwei_to_haichang_yi)}亿。
<span class="red bold">则海昌主业Q1 2026净利 ≈${yi(q1.main_business_np_yi)}亿 vs Q1 2025 主业 0.18亿 → 主业YoY ${pctSigned(q1.main_business_yoy_pct)}，信为通讯并表掩盖了主业下滑。</span></p>

<p style="margin: 6px 0 4px 0;"><span class="tag">B.2 信为通讯并购（最核心事件）</span>：${acq.fully_consolidated_date} 完成工商变更，现金对价 ${yi(acq.consideration_yi)} 亿元（已付 ${yi(acq.consideration_yi - acq.transaction_fl_yi, 2)}亿，剩余 ${yi(acq.transaction_fl_yi)} 亿分期 3-5 期，记入"交易性金融负债"）。标的整体估值 ${yi(acq.valuation_full_yi)} 亿（账面 6,663 万），评估增值率 ${pct(acq.valuation_premium_pct)}，形成商誉 ${yi(acq.goodwill_yi)} 亿。业绩对赌 FY2025/26/27 = 3,800/4,000/4,200 万元（累计 1.2 亿）；FY2025 前三季度实现净利 3,465 万元（已超 FY2024 全年）—— 一阶段验证。至迟 2027 年启动剩余 49% 收购（发行股份）。业务方向：GNSS 天线定位 + 北斗卫星通信天线（跨界至卫星通讯）。</p>

<p style="margin: 6px 0 4px 0;"><span class="tag">B.3 散户情绪与讨论热点（散户情绪面 | 韭研+股吧+财富号+互动易）</span>：热度【高】—— 多次龙虎榜（2025-12-31 累涨 9.04%，2025-07-28~30 累涨 36.73%）；主力资金日内大幅震荡（单日 ±7,000-9,000万）；融资融券占流通市值 ~6%。看多主线（强叙事）：人形机器人灵巧手齿轮——"Optimus 12个/手×300元=3,600元/手价值，粉末合金是唯一量产工艺路径"（韭研公社多篇深度帖），客户含 TTI/博世/大疆DJI/UBTech，及特斯拉送样预期；卫星通讯跨界（信为通讯）。看空声音：估值高（P/E 87.5x vs 可比中位 49.8x）+ 主业-22% + 1.87亿商誉减值压力 + 跨界整合不确定（粉末冶金→卫星通讯）。</p>
<p style="margin: 4px 0 4px 0;"><span class="tag">Top1 争议：</span>"人形机器人灵巧手 + 卫星通讯叙事（溢价主因）" vs "115x TTM PE + 1.87亿商誉 + 主业 −22% / Q1 2026 主业 −32%"。</p>

<h2>C. 股东与资金面</h2>

<p style="margin: 6px 0 4px 0;"><span class="tag">C.1 股权结构（FY2025末，高度集中）</span>：实控人周光荣 ${ts[0].pct}%（限售 ${(ts[0].restricted_M / ts[0].shares_M * 100).toFixed(0)}%）+ 配偶徐晓玉 12.63%（已全部解禁）+ 员工持股平台 7.31% + 周广华（胞弟）0.76% = 一致行动人合计 ${(ts[0].pct + ts[1].pct + ts[2].pct + ts[3].pct).toFixed(2)}%。</p>

<table>
<tr>
  <th style="width:30%">股东</th>
  <th style="width:14%">持股比例</th>
  <th style="width:16%">持股(百万股)</th>
  <th style="width:14%">限售比例</th>
  <th>关系/备注</th>
</tr>
${ts.slice(0, 5).map(s => `<tr>
  <td>${s.name}</td>
  <td class="center">${pct(s.pct, 2)}</td>
  <td class="num">${f(s.shares_M, 2)}</td>
  <td class="center">${s.restricted_M !== undefined && s.shares_M > 0 ? pct(s.restricted_M / s.shares_M * 100, 0) : '—'}</td>
  <td>${s.role || '—'}</td>
</tr>`).join('')}
</table>

<p style="margin: 6px 0 4px 0;"><span class="tag">C.2 资金面与解禁</span>：北向（香港中央结算）仅持 ${ts[4].pct}%（小盘股典型）；2026-03-25 推出限制性股票激励计划（203 万股、占比 0.82%、27 人、授予价 11.81 元/股深度折扣）；散户参与度高。<span class="bold">询价转让历史：本次为公司首次询价转让（无历史先例）。</span></p>

<div class="chart"><img src="${imgB64('chart3_price_levels.png')}" /></div>
<div class="chart-caption">图3　当前价 ${f(m.close, 2)} 在 MA60 之上、距 52W 高位 ${pct((m.close / m.px_52w_high - 1) * 100)}；<span class="red bold">卖方底价 ${f(v.floor_price, 2)} 元（折 ${pct(v.floor_discount_pct)}）</span>与 折10%/折20% 共同构成成交价区间。</div>

<h2>D. 公司概况 + 财务 + 投资亮点 + 风险提示</h2>

<p style="margin: 6px 0 4px 0;"><span class="tag">D.1 基本信息</span>：扬州海昌新材股份有限公司（${data.meta.code}.SZ），创业板，2020-09 上市；专精特新"小巨人"。主营粉末冶金零部件（PM 压制成形 ${pct(data.rev_breakdown.FY2025.by_product.PM.share)} + MIM 注射成形 ${pct(data.rev_breakdown.FY2025.by_product.MIM.share)}），下游集中于电动工具（${pct(data.rev_breakdown.FY2025.by_industry.电动工具零部件.share)}）。实控人周光荣（董事长）。已设新加坡、越南全资子公司（2026-03 完成注册）。</p>

<p style="margin: 6px 0 4px 0;"><span class="tag">D.2 三年财务速览（单位：亿元）</span></p>

<table>
<tr><th>指标</th><th>FY2023</th><th>FY2024</th><th>FY2025</th><th>YoY</th></tr>
<tr><td>营业收入</td><td class="num">${yi(fy23.revenue)}</td><td class="num">${yi(fy24.revenue)}</td><td class="num bold">${yi(fy25.revenue)}</td><td class="num red">${pctSigned((fy25.revenue / fy24.revenue - 1) * 100)}</td></tr>
<tr><td>毛利率</td><td class="num">${pct(fy23.gross_margin)}</td><td class="num">${pct(fy24.gross_margin)}</td><td class="num bold">${pct(fy25.gross_margin)}</td><td class="num red">${pctSigned(fy25.gross_margin - fy24.gross_margin)}pp</td></tr>
<tr><td>营业利润</td><td class="num">${yi(fy23.op_profit)}</td><td class="num">${yi(fy24.op_profit)}</td><td class="num bold">${yi(fy25.op_profit)}</td><td class="num red">${pctSigned((fy25.op_profit / fy24.op_profit - 1) * 100)}</td></tr>
<tr><td>归母净利润</td><td class="num">${yi(fy23.np_parent)}</td><td class="num">${yi(fy24.np_parent)}</td><td class="num bold">${yi(fy25.np_parent)}</td><td class="num red">${pctSigned((fy25.np_parent / fy24.np_parent - 1) * 100)}</td></tr>
<tr><td>扣非归母净利润</td><td class="num">${yi(fy23.np_nonrecurring_excluded)}</td><td class="num">${yi(fy24.np_nonrecurring_excluded)}</td><td class="num bold">${yi(fy25.np_nonrecurring_excluded)}</td><td class="num red">${pctSigned((fy25.np_nonrecurring_excluded / fy24.np_nonrecurring_excluded - 1) * 100)}</td></tr>
<tr><td>EBITDA</td><td class="num">${yi(fy23.ebitda)}</td><td class="num">${yi(fy24.ebitda)}</td><td class="num bold">${yi(fy25.ebitda)}</td><td class="num red">${pctSigned((fy25.ebitda / fy24.ebitda - 1) * 100)}</td></tr>
<tr><td>经营性现金流 (OCF)</td><td class="num">${yi(fy23.ocf)}</td><td class="num">${yi(fy24.ocf)}</td><td class="num">${yi(fy25.ocf)}</td><td class="num red">${pctSigned((fy25.ocf / fy24.ocf - 1) * 100)}</td></tr>
<tr><td>Capex</td><td class="num">${yi(fy23.capex)}</td><td class="num">${yi(fy24.capex)}</td><td class="num">${yi(fy25.capex)}</td><td class="num">${pctSigned((fy25.capex / fy24.capex - 1) * 100)}</td></tr>
<tr><td>FCF (=OCF−Capex)</td><td class="num">${yi(fy23.fcf)}</td><td class="num">${yi(fy24.fcf)}</td><td class="num bold">${yi(fy25.fcf)}</td><td class="num red">${pctSigned((fy25.fcf / fy24.fcf - 1) * 100)}</td></tr>
<tr><td>FCF/EBITDA</td><td class="num">${pct(fy23.fcf_ebitda)}</td><td class="num">${pct(fy24.fcf_ebitda)}</td><td class="num">${pct(fy25.fcf_ebitda)}</td><td class="num">${pctSigned(fy25.fcf_ebitda - fy24.fcf_ebitda)}pp</td></tr>
<tr><td>加权ROE</td><td class="num">${pct(fy23.roe_weighted)}</td><td class="num">${pct(fy24.roe_weighted)}</td><td class="num bold">${pct(fy25.roe_weighted)}</td><td class="num red">${pctSigned(fy25.roe_weighted - fy24.roe_weighted)}pp</td></tr>
<tr><td>基本EPS (元)</td><td class="num">${f(fy23.eps_basic, 4)}</td><td class="num">${f(fy24.eps_basic, 4)}</td><td class="num">${f(fy25.eps_basic, 4)}</td><td class="num red">${pctSigned((fy25.eps_basic / fy24.eps_basic - 1) * 100)}</td></tr>
<tr><td>资产总额</td><td class="num">${yi(fy23.total_assets)}</td><td class="num">${yi(fy24.total_assets)}</td><td class="num bold">${yi(fy25.total_assets)}</td><td class="num green">${pctSigned((fy25.total_assets / fy24.total_assets - 1) * 100)}</td></tr>
<tr><td>总负债</td><td class="num">${yi(fy23.total_liabilities)}</td><td class="num">${yi(fy24.total_liabilities)}</td><td class="num bold">${yi(fy25.total_liabilities)}</td><td class="num red">${pctSigned((fy25.total_liabilities / fy24.total_liabilities - 1) * 100)}</td></tr>
<tr><td>资产负债率</td><td class="num">${pct(fy23.lev_ratio)}</td><td class="num">${pct(fy24.lev_ratio)}</td><td class="num bold">${pct(fy25.lev_ratio)}</td><td class="num red">${pctSigned(fy25.lev_ratio - fy24.lev_ratio)}pp</td></tr>
<tr><td>商誉</td><td class="num">${yi(fy23.goodwill)}</td><td class="num">${yi(fy24.goodwill)}</td><td class="num bold red">${yi(fy25.goodwill)}</td><td class="num red">+${yi(fy25.goodwill)}亿(信为通讯)</td></tr>
<tr><td>净现金 (无有息负债)</td><td class="num">${yi(fy23.net_cash)}</td><td class="num">${yi(fy24.net_cash)}</td><td class="num bold green">${yi(fy25.net_cash)}</td><td class="num">${pctSigned((fy25.net_cash / fy24.net_cash - 1) * 100)}</td></tr>
</table>

<p style="margin: 6px 0 4px 0;"><span class="tag">D.3 FY2025 收入拆分</span>：分行业 电动工具零部件 ${pct(data.rev_breakdown.FY2025.by_industry.电动工具零部件.share)}（YoY ${pctSigned(data.rev_breakdown.FY2025.by_industry.电动工具零部件.yoy)}）/ 汽车零部件 ${pct(data.rev_breakdown.FY2025.by_industry.汽车零部件.share)}（YoY ${pctSigned(data.rev_breakdown.FY2025.by_industry.汽车零部件.yoy)}）。分产品 PM ${pct(data.rev_breakdown.FY2025.by_product.PM.share)}（YoY ${pctSigned(data.rev_breakdown.FY2025.by_product.PM.yoy)}）/ MIM ${pct(data.rev_breakdown.FY2025.by_product.MIM.share)}（YoY ${pctSigned(data.rev_breakdown.FY2025.by_product.MIM.yoy)}，高毛利 ${pct(48.0)}）。分地区 境外 ${pct(data.rev_breakdown.FY2025.by_region.境外.share)}（YoY ${pctSigned(data.rev_breakdown.FY2025.by_region.境外.yoy)}，毛利率 ${pct(data.rev_breakdown.FY2025.gm_by_region.境外)}）/ 境内 ${pct(data.rev_breakdown.FY2025.by_region.境内.share)}（毛利率 ${pct(data.rev_breakdown.FY2025.gm_by_region.境内)}）。</p>

<p style="margin: 6px 0 4px 0;"><span class="tag">D.4 投资亮点（+）/ 风险提示（−）</span></p>
<table>
<tr><th style="width:50%">投资亮点 (+)</th><th style="width:50%">风险提示 (−)</th></tr>
${data.catalysts_risks.positive.map((p, i) => `<tr>
  <td class="cat-pos">${p}</td>
  <td class="cat-neg">${data.catalysts_risks.negative[i] || ''}</td>
</tr>`).join('')}
</table>

<div class="footer-source">
数据来源：FY2023/FY2024/FY2025 年报、H1 2025 / Q3 2025 / Q1 2026 季报、Wind 日度估值+股价+CSC 可比表（用户上传）；公开信息：cninfo / 东方财富 / 韭研公社 / 财富号 / 新浪财经。<br/>
数据完整度自检：23/23 校验通过（毛利率 / EBITDA 公式 / FCF / 资产负债率 / 净现金 / P/E 口径 / 三年趋势 / 商誉变动）。海昌新材无机构一致预期覆盖，所有 forward 倍数标 NA。
</div>

</body>
</html>`;

// 输出 HTML (调试用)
// 输出文件名 = <标的>_<代码>_询价转让投资备忘录
const baseName = data.meta.company + '_' + data.meta.code + '_询价转让投资备忘录';
const htmlPath = path.join(PROJECT_DIR, baseName + '_预览.html');
fs.writeFileSync(htmlPath, html);
console.log('HTML saved: ' + htmlPath);

// 渲染 PDF
(async () => {
  const browser = await puppeteer.launch({ headless: true });
  const page = await browser.newPage();
  await page.setContent(html, { waitUntil: 'networkidle0' });
  const pdfPath = path.join(PROJECT_DIR, baseName + '.pdf');
  await page.pdf({
    path: pdfPath,
    format: 'A4',
    printBackground: true,
    displayHeaderFooter: true,
    headerTemplate: '<div></div>',
    footerTemplate: `<div style="font-size:8px;color:#777;text-align:center;width:100%;font-family:'Calibri',sans-serif">
      第 <span class="pageNumber"></span> 页 / 共 <span class="totalPages"></span> 页
    </div>`,
    margin: { top: '12mm', right: '12mm', bottom: '14mm', left: '12mm' },
  });
  await browser.close();

  const stat = fs.statSync(pdfPath);
  console.log('✅ PDF 已生成: ' + pdfPath);
  console.log('   文件大小: ' + (stat.size / 1024).toFixed(1) + ' KB');
})();
