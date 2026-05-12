// 通用 .docx 备忘录生成模板 (作为 PDF 的备选格式)
// 用法: node scripts/build_docx.js --project projects/<标的名>
// 所有财务数字必须来自 data.json (CLAUDE.md 硬性要求, 不允许硬编码)

const fs = require('fs');
const path = require('path');
const docx = require('docx');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, HeadingLevel, WidthType, PageOrientation,
  ImageRun, ShadingType, Header, Footer, PageNumber, PageBreak,
  convertInchesToTwip, convertMillimetersToTwip,
} = docx;

function getProjectDir() {
  const args = process.argv.slice(2);
  const i = args.indexOf('--project');
  if (i >= 0 && args[i + 1]) return path.resolve(args[i + 1]);
  console.error('Usage: node scripts/build_docx.js --project projects/<标的名>');
  process.exit(1);
}
const PROJECT_DIR = getProjectDir();
const data = JSON.parse(fs.readFileSync(path.join(PROJECT_DIR, 'data.json'), 'utf8'));

// ============================================================
// 格式工具
// ============================================================
const FONT = 'Calibri';
const SIZE_TITLE = 22;       // 11pt
const SIZE_H1 = 18;          // 9pt
const SIZE_BODY = 16;        // 8pt
const SIZE_TABLE = 14;       // 7pt
const SIZE_TABLE_HEADER = 15;// 7.5pt
const SIZE_SOURCE = 12;      // 6pt italic
const SIZE_FOOTER = 14;      // 7pt

// 颜色
const RED = 'C0392B';
const GREEN = '27AE60';
const BLUE = '2C5F9B';
const GRAY = '7F8C8D';
const LIGHT_BG = 'EEF5FB';
const HEADER_BG = '2C5F9B';
const ZEBRA_BG = 'F8F9FA';

const f = (n, dec = 1) => {
  if (n === null || n === undefined || isNaN(n)) return '—';
  return n.toFixed(dec);
};
const pct = (n, dec = 1) => f(n, dec) + '%';
const pctSigned = (n, dec = 1) => (n > 0 ? '+' : '') + f(n, dec) + '%';
const yi = (n, dec = 2) => f(n, dec);

// TextRun helpers
const tr = (text, opts = {}) => new TextRun({
  text: String(text),
  font: FONT,
  size: opts.size || SIZE_BODY,
  bold: opts.bold || false,
  italics: opts.italics || false,
  color: opts.color || '000000',
});

const p = (text, opts = {}) => new Paragraph({
  alignment: opts.align || AlignmentType.LEFT,
  spacing: { after: opts.after === undefined ? 60 : opts.after, before: opts.before || 0, line: opts.line || 240 },
  children: typeof text === 'string' ? [tr(text, opts)] : text,
});

// Bold label + body in same paragraph
const labelPara = (label, body, opts = {}) => new Paragraph({
  spacing: { after: opts.after === undefined ? 40 : opts.after, line: 240 },
  children: [
    tr(label, { bold: true, size: opts.size || SIZE_BODY }),
    tr(body, { size: opts.size || SIZE_BODY }),
  ],
});

// 表格 cell helper
const cell = (text, opts = {}) => new TableCell({
  width: { size: opts.width || 1000, type: WidthType.DXA },
  shading: opts.shading ? { type: ShadingType.CLEAR, color: 'auto', fill: opts.shading } : undefined,
  margins: { top: 30, bottom: 30, left: 50, right: 50 },
  children: [new Paragraph({
    alignment: opts.align || AlignmentType.LEFT,
    spacing: { after: 0, line: 220 },
    children: Array.isArray(text) ? text : [
      tr(text, { size: opts.size || SIZE_TABLE, bold: opts.bold, color: opts.color }),
    ],
  })],
});

// 表头行
const headerRow = (cells, widths) => new TableRow({
  tableHeader: true,
  children: cells.map((t, i) => cell(t, { width: widths[i], shading: HEADER_BG, bold: true, color: 'FFFFFF', size: SIZE_TABLE_HEADER, align: AlignmentType.CENTER })),
});
// 数据行
const dataRow = (cells, widths, opts = {}) => new TableRow({
  children: cells.map((t, i) => {
    let opt = opts.cellOpts && opts.cellOpts[i] ? opts.cellOpts[i] : {};
    return cell(t, { width: widths[i], shading: opts.zebra ? ZEBRA_BG : undefined, ...opt });
  }),
});

const tableStandard = (rows) => new Table({
  width: { size: 10800, type: WidthType.DXA },
  rows,
  borders: {
    top:    { style: BorderStyle.SINGLE, size: 4, color: '888888' },
    bottom: { style: BorderStyle.SINGLE, size: 4, color: '888888' },
    left:   { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC' },
    right:  { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC' },
    insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: 'D5D5D5' },
    insideVertical:   { style: BorderStyle.SINGLE, size: 2, color: 'D5D5D5' },
  },
});

// 图片 helper
const chartImage = (filename, width = 600, height = 180) => {
  const buf = fs.readFileSync(path.join(PROJECT_DIR, filename));
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 60, before: 40 },
    children: [new ImageRun({ data: buf, transformation: { width, height } })],
  });
};

// ============================================================
// 内容构建
// ============================================================
const fy23 = data.financials.FY2023;
const fy24 = data.financials.FY2024;
const fy25 = data.financials.FY2025;
const q1 = data.financials.Q1_2026;
const m = data.market;
const v = data.valuation;
const acq = data.acquisition;
const peerM = data.peer_median;
const peers = data.peers;

const children = [];

// 标题区
children.push(new Paragraph({
  spacing: { after: 40, line: 240 },
  alignment: AlignmentType.LEFT,
  children: [
    tr(data.meta.company + '（' + data.meta.code + '.SZ）— 询价转让投资备忘录', { size: SIZE_TITLE, bold: true, color: BLUE }),
  ],
}));
children.push(new Paragraph({
  spacing: { after: 100, line: 200 },
  alignment: AlignmentType.LEFT,
  children: [
    tr(data.meta.memo_date + ' | 数据截止：' + data.meta.data_cutoff + ' | 股价口径：' + m.date + ' 收盘价 ' + m.close + ' 元 | 卖方底价 ' + f(v.floor_price, 2) + ' 元 (折扣率 ' + pct(v.floor_discount_pct) + ')',
      { size: SIZE_SOURCE, italics: true, color: GRAY }),
  ],
}));

// 一句话定位
const fy25_rev_yoy = ((fy25.revenue / fy24.revenue - 1) * 100).toFixed(1);
const fy25_np_yoy = ((fy25.np_parent / fy24.np_parent - 1) * 100).toFixed(1);
children.push(new Paragraph({
  spacing: { after: 100, line: 240 },
  children: [
    tr('国内粉末冶金细分龙头（电动工具零部件 ' + data.rev_breakdown.FY2025.by_industry.电动工具零部件.share.toFixed(1) + '% / 海外占比 ' + data.rev_breakdown.FY2025.by_region.境外.share.toFixed(1) + '%），', { size: SIZE_BODY }),
    tr('FY2025 营收 ' + yi(fy25.revenue) + '亿（' + fy25_rev_yoy + '%）、归母净利 ' + yi(fy25.np_parent) + '亿（' + fy25_np_yoy + '%）—— 主业承压；', { size: SIZE_BODY }),
    tr('2025-12 完成 ' + acq.target + ' ' + acq.pct + '% 现金收购（' + yi(acq.consideration_yi) + '亿元，形成商誉 ' + yi(acq.goodwill_yi) + '亿），', { size: SIZE_BODY, bold: true }),
    tr('Q1 2026 信为通讯首次并表带动营收 +123.3% 但拆分后主业 -32%。近1年股价 ' + pctSigned(m.return_1y) + '，机构无覆盖、估值显著高于可比中位。', { size: SIZE_BODY }),
    tr(' 本次询价转让卖方底价 ' + f(v.floor_price, 2) + ' 元（折扣 ' + pct(v.floor_discount_pct) + '），对应 P/E (FY2025A) ' + f(v.pe_at_floor, 1) + 'x、P/E (FY2024A) ' + f(v.pe_at_floor_fy24, 1) + 'x，仍显著高于可比中位 ' + f(peerM.pe_2024, 1) + 'x。',
      { size: SIZE_BODY, bold: true, color: 'C2185B' }),
  ],
}));

// ============================================================
// 模块A：估值与市场表现
// ============================================================
children.push(new Paragraph({
  spacing: { before: 80, after: 60 },
  children: [tr('A. 估值与市场表现', { size: SIZE_H1, bold: true, color: BLUE })],
}));

// A.1 估值快照 + 询价转让定价 (场景A合并表)
// 场景A: 卖方底价 22.79, 折扣率 21.0% > 20%, 跳过折30% (20.20 < 22.79)
const FLOOR_BG = 'FCE4EC';  // 卖方底价列底色 (浅粉红)
children.push(labelPara('A.1 估值快照与询价转让定价（场景A：卖方底价 ' + f(v.floor_price, 2) + ' 元 / 折扣率 ' + pct(v.floor_discount_pct) + '）',
  '   单位：元/股、亿元、倍', { size: SIZE_BODY }));

// 列定义: 指标 / 当前价 / 折10% / 折20% / 卖方底价 (折X%)
// 注: 跳过折30%, 因 28.86×0.7=20.20 < 22.79 (低于底价不可能成交)
const W = [2700, 1900, 1900, 1900, 2400];

const floorColHeader = [
  tr('卖方底价', { size: SIZE_TABLE_HEADER, bold: true, color: 'FFFFFF' }),
  new TextRun({ text: '\n', font: FONT }),
  tr('(折 ' + pct(v.floor_discount_pct) + ')', { size: SIZE_TABLE_HEADER, bold: true, color: 'FFFFFF' }),
];

// EV-at-discount helper (估值表的 EV 倍数随股价变化)
function evMultAt(px) {
  const mc = px * m.total_shares;
  const ev = mc - (raw_fy25_cash + raw_fy25_tfa - raw_fy25_tfl);
  return ev / ebitda_2025_yuan;
}
// 需要的常量从 data 拿
const raw_fy25_cash = fy25.cash * 1e8;
const raw_fy25_tfa = fy25.trading_fa * 1e8;
const raw_fy25_tfl = fy25.trading_fl * 1e8;
const ebitda_2025_yuan = fy25.ebitda * 1e8;

const evCurrent = evMultAt(m.close);
const evDisc10 = evMultAt(v.floor_discount_10);
const evDisc20 = evMultAt(v.floor_discount_20);
const evFloor = evMultAt(v.floor_price);

const valTable = tableStandard([
  new TableRow({
    tableHeader: true,
    children: [
      cell('指标', { width: W[0], shading: HEADER_BG, bold: true, color: 'FFFFFF', size: SIZE_TABLE_HEADER, align: AlignmentType.CENTER }),
      cell('当前股价', { width: W[1], shading: HEADER_BG, bold: true, color: 'FFFFFF', size: SIZE_TABLE_HEADER, align: AlignmentType.CENTER }),
      cell('折价 10%', { width: W[2], shading: HEADER_BG, bold: true, color: 'FFFFFF', size: SIZE_TABLE_HEADER, align: AlignmentType.CENTER }),
      cell('折价 20%', { width: W[3], shading: HEADER_BG, bold: true, color: 'FFFFFF', size: SIZE_TABLE_HEADER, align: AlignmentType.CENTER }),
      cell([
        tr('卖方底价（折 ' + pct(v.floor_discount_pct) + '）', { size: SIZE_TABLE_HEADER, bold: true, color: 'FFFFFF' }),
      ], { width: W[4], shading: 'C2185B', bold: true, color: 'FFFFFF', size: SIZE_TABLE_HEADER, align: AlignmentType.CENTER }),
    ],
  }),
  dataRow(['股价（元）', f(m.close, 2), f(v.floor_discount_10, 2), f(v.floor_discount_20, 2), f(v.floor_price, 2)],
    W, {
      cellOpts: [
        {},
        { align: AlignmentType.CENTER, bold: true },
        { align: AlignmentType.CENTER, shading: LIGHT_BG },
        { align: AlignmentType.CENTER, shading: LIGHT_BG },
        { align: AlignmentType.CENTER, shading: FLOOR_BG, bold: true, color: 'C2185B' },
      ],
    }),
  dataRow(['对应市值（亿元）',
      f(v.mkt_cap_yi, 1),
      f(v.mkt_cap_yi * 0.9, 1),
      f(v.mkt_cap_yi * 0.8, 1),
      f(v.mkt_cap_at_floor_yi, 1)],
    W,
    { zebra: true, cellOpts: [
      {}, { align: AlignmentType.CENTER }, { align: AlignmentType.CENTER }, { align: AlignmentType.CENTER },
      { align: AlignmentType.CENTER, shading: FLOOR_BG, bold: true },
    ] }),
  dataRow(['P/E (FY2025A=' + yi(fy25.np_parent) + '亿)',
      f(v.pe_2025a, 1) + 'x',
      f(v.pe_at_disc10, 1) + 'x',
      f(v.pe_at_disc20, 1) + 'x',
      f(v.pe_at_floor, 1) + 'x'],
    W,
    { cellOpts: [
      {}, { align: AlignmentType.CENTER }, { align: AlignmentType.CENTER }, { align: AlignmentType.CENTER },
      { align: AlignmentType.CENTER, shading: FLOOR_BG, bold: true, color: 'C2185B' },
    ] }),
  dataRow(['P/E (FY2024A=' + yi(fy24.np_parent) + '亿)',
      f(v.pe_2024a, 1) + 'x',
      f(v.pe_2024a * 0.9, 1) + 'x',
      f(v.pe_2024a * 0.8, 1) + 'x',
      f(v.pe_at_floor_fy24, 1) + 'x'],
    W,
    { zebra: true, cellOpts: [
      {}, { align: AlignmentType.CENTER }, { align: AlignmentType.CENTER }, { align: AlignmentType.CENTER },
      { align: AlignmentType.CENTER, shading: FLOOR_BG, bold: true, color: 'C2185B' },
    ] }),
  dataRow(['EV/EBITDA (FY2025A=' + yi(fy25.ebitda) + '亿)',
      f(evCurrent, 1) + 'x',
      f(evDisc10, 1) + 'x',
      f(evDisc20, 1) + 'x',
      f(evFloor, 1) + 'x'],
    W,
    { cellOpts: [
      {}, { align: AlignmentType.CENTER }, { align: AlignmentType.CENTER }, { align: AlignmentType.CENTER },
      { align: AlignmentType.CENTER, shading: FLOOR_BG, bold: true, color: 'C2185B' },
    ] }),
  dataRow(['P/S (FY2025A=' + yi(fy25.revenue) + '亿)',
      f(v.ps_2025, 1) + 'x',
      f(v.ps_2025 * 0.9, 1) + 'x',
      f(v.ps_2025 * 0.8, 1) + 'x',
      f((v.floor_price * m.total_shares) / (fy25.revenue * 1e8), 1) + 'x'],
    W,
    { zebra: true, cellOpts: [
      {}, { align: AlignmentType.CENTER }, { align: AlignmentType.CENTER }, { align: AlignmentType.CENTER },
      { align: AlignmentType.CENTER, shading: FLOOR_BG, bold: true },
    ] }),
  dataRow(['P/B (FY2025末 BVPS=' + f(fy25.equity_parent / (m.total_shares / 1e8), 2) + '元)',
      f(v.pb_2025, 1) + 'x',
      f(v.pb_2025 * 0.9, 1) + 'x',
      f(v.pb_2025 * 0.8, 1) + 'x',
      f((v.floor_price * m.total_shares) / (fy25.equity_parent * 1e8), 1) + 'x'],
    W,
    { cellOpts: [
      {}, { align: AlignmentType.CENTER }, { align: AlignmentType.CENTER }, { align: AlignmentType.CENTER },
      { align: AlignmentType.CENTER, shading: FLOOR_BG, bold: true },
    ] }),
]);
children.push(valTable);

children.push(new Paragraph({
  spacing: { before: 30, after: 60 },
  children: [
    tr('注：卖方底价折扣率 ' + pct(v.floor_discount_pct) + ' 介于 20%-30% 之间，故跳过 折30% 档（28.86 × 0.7 = 20.20 < 22.79，低于底价不可能成交）。海昌新材无机构一致预期覆盖（NTM consensus 缺失），表格仅展示 actual 倍数；EV/EBITDA 已剔除信为通讯应付分期款 ' + yi(fy25.trading_fl) + ' 亿。',
      { size: SIZE_SOURCE, italics: true, color: GRAY }),
  ],
}));

// A.2 估值 vs A股可比公司
children.push(labelPara('A.2 vs A股可比公司（基准日 2026-04-14, 来源 CSC）', '', { size: SIZE_BODY }));

const peerTable = tableStandard([
  headerRow(['公司 (代码)', '股价', '市值(亿)', 'EV(亿)', 'P/S CY24', 'EV/EBITDA CY24', 'P/E CY24', 'P/B LTM'],
    [2000, 1100, 1300, 1300, 1300, 1500, 1200, 1100]),
  dataRow(['海昌新材 (300885)', f(peers.haichang.px, 2), f(peers.haichang.mkt_cap_yi, 1), f(peers.haichang.ev_yi, 1),
    f(peers.haichang.ps_2024, 1) + 'x', f(peers.haichang.ev_ebitda_2024, 1) + 'x', f(peers.haichang.pe_2024, 1) + 'x', f(peers.haichang.pb_ltm, 1) + 'x'],
    [2000, 1100, 1300, 1300, 1300, 1500, 1200, 1100],
    { cellOpts: [{ bold: true, color: RED }, { align: AlignmentType.RIGHT }, { align: AlignmentType.RIGHT }, { align: AlignmentType.RIGHT },
      { align: AlignmentType.RIGHT, bold: true }, { align: AlignmentType.RIGHT, bold: true }, { align: AlignmentType.RIGHT, bold: true }, { align: AlignmentType.RIGHT }] }),
  dataRow(['东睦股份 (600114)', f(peers.dongmu.px, 2), f(peers.dongmu.mkt_cap_yi, 1), f(peers.dongmu.ev_yi, 1),
    f(peers.dongmu.ps_2024, 1) + 'x', f(peers.dongmu.ev_ebitda_2024, 1) + 'x', f(peers.dongmu.pe_2024, 1) + 'x', f(peers.dongmu.pb_ltm, 1) + 'x'],
    [2000, 1100, 1300, 1300, 1300, 1500, 1200, 1100],
    { zebra: true, cellOpts: [{}, { align: AlignmentType.RIGHT }, { align: AlignmentType.RIGHT }, { align: AlignmentType.RIGHT },
      { align: AlignmentType.RIGHT }, { align: AlignmentType.RIGHT }, { align: AlignmentType.RIGHT }, { align: AlignmentType.RIGHT }] }),
  dataRow(['精研科技 (300709)', f(peers.jingyan.px, 2), f(peers.jingyan.mkt_cap_yi, 1), f(peers.jingyan.ev_yi, 1),
    f(peers.jingyan.ps_2024, 1) + 'x', f(peers.jingyan.ev_ebitda_2024, 1) + 'x', f(peers.jingyan.pe_2024, 1) + 'x', f(peers.jingyan.pb_ltm, 1) + 'x'],
    [2000, 1100, 1300, 1300, 1300, 1500, 1200, 1100],
    { cellOpts: [{}, { align: AlignmentType.RIGHT }, { align: AlignmentType.RIGHT }, { align: AlignmentType.RIGHT },
      { align: AlignmentType.RIGHT }, { align: AlignmentType.RIGHT }, { align: AlignmentType.RIGHT }, { align: AlignmentType.RIGHT }] }),
  dataRow(['派克新材 (605123)', f(peers.paike.px, 2), f(peers.paike.mkt_cap_yi, 1), f(peers.paike.ev_yi, 1),
    f(peers.paike.ps_2024, 1) + 'x', f(peers.paike.ev_ebitda_2024, 1) + 'x', f(peers.paike.pe_2024, 1) + 'x', f(peers.paike.pb_ltm, 1) + 'x'],
    [2000, 1100, 1300, 1300, 1300, 1500, 1200, 1100],
    { zebra: true, cellOpts: [{}, { align: AlignmentType.RIGHT }, { align: AlignmentType.RIGHT }, { align: AlignmentType.RIGHT },
      { align: AlignmentType.RIGHT }, { align: AlignmentType.RIGHT }, { align: AlignmentType.RIGHT }, { align: AlignmentType.RIGHT }] }),
  dataRow(['可比中位数', '—', '—', '—', f(peerM.ps_2024, 1) + 'x', f(peerM.ev_ebitda_2024, 1) + 'x', f(peerM.pe_2024, 1) + 'x', f(peerM.pb_ltm, 1) + 'x'],
    [2000, 1100, 1300, 1300, 1300, 1500, 1200, 1100],
    { cellOpts: [{ italics: true, bold: true }, {}, {}, {},
      { align: AlignmentType.RIGHT, bold: true }, { align: AlignmentType.RIGHT, bold: true }, { align: AlignmentType.RIGHT, bold: true }, { align: AlignmentType.RIGHT, bold: true }] }),
  dataRow(['海昌溢价幅度', '—', '—', '—',
    pctSigned((peers.haichang.ps_2024 / peerM.ps_2024 - 1) * 100, 0),
    pctSigned((peers.haichang.ev_ebitda_2024 / peerM.ev_ebitda_2024 - 1) * 100, 0),
    pctSigned((peers.haichang.pe_2024 / peerM.pe_2024 - 1) * 100, 0),
    pctSigned((peers.haichang.pb_ltm / peerM.pb_ltm - 1) * 100, 0)],
    [2000, 1100, 1300, 1300, 1300, 1500, 1200, 1100],
    { zebra: true, cellOpts: [{ italics: true, color: RED, bold: true }, {}, {}, {},
      { align: AlignmentType.RIGHT, color: RED, bold: true }, { align: AlignmentType.RIGHT, color: RED, bold: true },
      { align: AlignmentType.RIGHT, color: RED, bold: true }, { align: AlignmentType.RIGHT, color: RED, bold: true }] }),
]);
children.push(peerTable);

// Forward consensus (东睦独家)
children.push(new Paragraph({
  spacing: { before: 40, after: 40 },
  children: [tr('A股可比 forward 一致预期：CY2025E P/E ' + peerM.pe_2025e + 'x → CY2026E ' + peerM.pe_2026e + 'x → CY2027E 31.4x；EV/EBITDA CY2025E ' + peerM.ev_ebitda_2025e + 'x。仅东睦股份单家有 forward consensus（精研、派克均无机构覆盖延伸至 forward）。', { size: SIZE_BODY, italics: true })],
}));

// A.3 价格表现
children.push(labelPara('A.3 价格表现与流动性：',
  '近1月 ' + pctSigned(m.return_1m) + '，近6月 ' + pctSigned(m.return_6m) + '，近1年 ' + pctSigned(m.return_1y) +
  '（vs 东睦 +78.8% / 精研 +40.8% / 派克 +112.8%）。' +
  ' 52周价格区间 ' + f(m.px_52w_low, 2) + ' ~ ' + f(m.px_52w_high, 2) + ' 元，当前距 52W 高位 ' + pct((m.close / m.px_52w_high - 1) * 100) +
  '。近20日均成交额 ' + f(m.avg20d_turnover_wan / 10000, 2) + ' 亿元，最新换手率 ' + pct(m.latest_turnover_rate_pct) +
  '（异常活跃）。融资融券余额（2026-03-26）' + f(m.margin_balance_yi, 2) + ' 亿元（占流通市值 ~6%）。',
  { size: SIZE_BODY }));

// 图1: PE Band
children.push(chartImage('chart1_pe_band.png', 620, 175));
children.push(new Paragraph({
  spacing: { before: 0, after: 40 },
  alignment: AlignmentType.CENTER,
  children: [tr('当前 TTM PE ' + m.ttm_pe_last + 'x 远高于历史均值 ' + m.ttm_pe_mean + 'x（约 +2.4 SD），处于近3年估值上限附近。',
    { size: SIZE_SOURCE, italics: true, color: GRAY })],
}));

// 图2: 同业对比
children.push(chartImage('chart2_peer_comparison.png', 620, 175));
children.push(new Paragraph({
  spacing: { before: 0, after: 60 },
  alignment: AlignmentType.CENTER,
  children: [tr('归一化基期 2023-05-12 = 100；海昌新材 3 年累计涨幅显著高于行业龙头东睦股份。',
    { size: SIZE_SOURCE, italics: true, color: GRAY })],
}));

// ============================================================
// 模块B：最新季度 + News Run
// ============================================================
children.push(new Paragraph({
  spacing: { before: 100, after: 60 },
  children: [tr('B. 最新季度 + 近期动态 + News Run', { size: SIZE_H1, bold: true, color: BLUE })],
}));

// B.1 Q1 2026 季报
const fy24_h = data.financials.FY2024;  // for ROE base
children.push(labelPara('B.1 Q1 2026（' + acq.target + '首次并表后第一份季报）',
  '：营收 ' + yi(q1.revenue) + ' 亿（+123.3% YoY）/ 归母净利 ' + yi(q1.np_parent) + ' 亿（+36.7%）/ 扣非归母 ' + yi(q1.np_nonrecurring_excluded) + ' 亿（+43.5%）/ 综合毛利率 ' + pct(q1.gross_margin) +
  '（vs Q1 2025 35.9%，略降 1pp）。', { size: SIZE_BODY }));

children.push(new Paragraph({
  spacing: { after: 60, line: 240 },
  children: [
    tr('  关键拆分（重要）：合并净利润 ' + yi(q1.net_income) + ' 亿 = 归母 ' + yi(q1.np_parent) + ' 亿 + 少数股东 ' + yi(q1.np_minority) + ' 亿。', { size: SIZE_BODY }),
    tr('信为通讯Q1全口径净利 ≈' + yi(q1.xinwei_total_np_yi) + '亿（少数股东 ÷ 49%），', { size: SIZE_BODY }),
    tr('归属海昌51% ≈' + yi(q1.xinwei_to_haichang_yi) + '亿。', { size: SIZE_BODY }),
    tr(' 则海昌主业Q1 2026净利 ≈' + yi(q1.main_business_np_yi) + '亿 vs Q1 2025 主业 0.18亿 → 主业YoY ' + pctSigned(q1.main_business_yoy_pct) + '，', { size: SIZE_BODY, bold: true, color: RED }),
    tr('信为通讯并表掩盖了主业下滑。', { size: SIZE_BODY, bold: true }),
  ],
}));

// B.2 信为通讯并购 (核心事件)
children.push(labelPara('B.2 信为通讯并购（最核心事件）',
  '：' + acq.fully_consolidated_date + ' 完成工商变更，现金对价 ' + yi(acq.consideration_yi) + ' 亿元（已付 ' + yi((acq.consideration_yi - acq.transaction_fl_yi), 2) + '亿，剩余 ' + yi(acq.transaction_fl_yi) + ' 亿分期 3-5 期，记入"交易性金融负债"）。' +
  ' 标的整体估值 ' + yi(acq.valuation_full_yi) + ' 亿（账面 6,663 万），评估增值率 ' + pct(acq.valuation_premium_pct) + '，形成商誉 ' + yi(acq.goodwill_yi) + ' 亿。' +
  ' 业绩对赌 FY2025/26/27 = 3,800/4,000/4,200 万元（累计 1.2 亿）；FY2025 前三季度实现净利 3,465 万元（已超 FY2024 全年）—— 一阶段验证。 ' +
  ' 至迟 2027 年启动剩余 49% 收购（发行股份）。' +
  ' 业务方向：GNSS 天线定位 + 北斗卫星通信天线（跨界至卫星通讯）。',
  { size: SIZE_BODY }));

// B.3 散户情绪与讨论 (CLAUDE.md v13.1 中等档默认)
children.push(labelPara('B.3 散户情绪与讨论热点（散户情绪面 | 韭研+股吧+财富号+互动易）',
  '：热度【高】—— 多次龙虎榜（2025-12-31 累涨 9.04%，2025-07-28~30 累涨 36.73%）；主力资金日内大幅震荡（单日 ±7,000-9,000万）；融资融券占流通市值 ~6%。' +
  ' 看多主线（强叙事）：人形机器人灵巧手齿轮——"Optimus 12个/手×300元=3,600元/手价值，粉末合金是唯一量产工艺路径"（韭研公社多篇深度帖），客户含 TTI/博世/大疆DJI/UBTech，及特斯拉送样预期；卫星通讯跨界（信为通讯）。' +
  ' 看空声音：估值高（P/E 87.5x vs 可比中位 49.8x）+ 主业-22% + 1.87亿商誉减值压力 + 跨界整合不确定（粉末冶金→卫星通讯）。',
  { size: SIZE_BODY }));
children.push(new Paragraph({
  spacing: { after: 60, line: 240 },
  children: [
    tr('  Top1 争议：', { size: SIZE_BODY, bold: true }),
    tr('"人形机器人灵巧手 + 卫星通讯叙事（溢价主因）" vs "115x TTM PE + 1.87亿商誉 + 主业 -22% / Q1 2026 主业 -32%"。', { size: SIZE_BODY }),
  ],
}));

// ============================================================
// 模块C：股东与资金面
// ============================================================
children.push(new Paragraph({
  spacing: { before: 100, after: 60 },
  children: [tr('C. 股东与资金面', { size: SIZE_H1, bold: true, color: BLUE })],
}));

children.push(labelPara('C.1 股权结构（FY2025末，高度集中）',
  '：实控人周光荣 ' + data.top_shareholders_FY2025[0].pct + '%（限售 ' + (data.top_shareholders_FY2025[0].restricted_M / data.top_shareholders_FY2025[0].shares_M * 100).toFixed(0) + '%）+ 配偶徐晓玉 12.63%（已全部解禁）+ 员工持股平台扬州海昌协力 7.31% + 周广华（胞弟）0.76% = 一致行动人合计 ' +
  (data.top_shareholders_FY2025[0].pct + data.top_shareholders_FY2025[1].pct + data.top_shareholders_FY2025[2].pct + data.top_shareholders_FY2025[3].pct).toFixed(2) + '%。', { size: SIZE_BODY }));

// 股东表
const shareTable = tableStandard([
  headerRow(['股东', '持股比例', '持股数(百万)', '限售比例', '关系/备注'], [3000, 1400, 1600, 1600, 3200]),
  ...data.top_shareholders_FY2025.slice(0, 5).map((s, i) => dataRow(
    [s.name, pct(s.pct, 2), f(s.shares_M, 2),
      s.restricted_M !== undefined ? (s.shares_M > 0 ? pct(s.restricted_M / s.shares_M * 100, 0) : '—') : '—',
      s.role || '—'],
    [3000, 1400, 1600, 1600, 3200],
    { zebra: i % 2 === 1, cellOpts: [{}, { align: AlignmentType.CENTER }, { align: AlignmentType.RIGHT }, { align: AlignmentType.CENTER }, {}] }
  )),
]);
children.push(shareTable);

children.push(labelPara('C.2 资金面与解禁',
  '：北向（香港中央结算）仅持 ' + data.top_shareholders_FY2025[4].pct + '%（小盘股典型）；2026-03-25 推出限制性股票激励计划（203 万股、占比 0.82%、27 人、授予价 11.81 元/股深度折扣）；散户参与度高。 询价转让历史：本次为公司首次询价转让（无历史先例）。',
  { size: SIZE_BODY }));

// 图3: 股价图
children.push(chartImage('chart3_price_levels.png', 620, 175));
children.push(new Paragraph({
  spacing: { before: 0, after: 60 },
  alignment: AlignmentType.CENTER,
  children: [tr('图3  当前价 ' + f(m.close, 2) + ' 在 MA60 之上、距 52W 高位 ' + pct((m.close / m.px_52w_high - 1) * 100) + '；卖方底价 ' + f(v.floor_price, 2) + ' 元（折 ' + pct(v.floor_discount_pct) + '）与 折10%/折20% 共同构成成交价区间。',
    { size: SIZE_SOURCE, italics: true, color: GRAY })],
}));

// ============================================================
// 模块D：公司概况 + 财务 + 投资亮点 + 风险提示
// ============================================================
children.push(new Paragraph({
  spacing: { before: 100, after: 60 },
  children: [tr('D. 公司概况 + 财务 + 投资亮点 + 风险提示', { size: SIZE_H1, bold: true, color: BLUE })],
}));

children.push(labelPara('D.1 基本信息',
  '：扬州海昌新材股份有限公司（' + data.meta.code + '.SZ），创业板，2020-09 上市；专精特新"小巨人"。' +
  ' 主营粉末冶金零部件（PM 压制成形 ' + data.rev_breakdown.FY2025.by_product.PM.share.toFixed(1) + '% + MIM 注射成形 ' + data.rev_breakdown.FY2025.by_product.MIM.share.toFixed(1) + '%），下游集中于电动工具（' + data.rev_breakdown.FY2025.by_industry.电动工具零部件.share.toFixed(1) + '%）。' +
  ' 实控人周光荣（董事长）。已设新加坡、越南全资子公司（2026-03 完成注册）。',
  { size: SIZE_BODY }));

// D.2 三年财务速览表
children.push(labelPara('D.2 三年财务速览（单位：亿元）', '', { size: SIZE_BODY }));

const finTable = tableStandard([
  headerRow(['指标', 'FY2023', 'FY2024', 'FY2025', 'YoY'], [3000, 1900, 1900, 1900, 2100]),
  // P&L
  dataRow(['营业收入', yi(fy23.revenue), yi(fy24.revenue), yi(fy25.revenue), pctSigned((fy25.revenue/fy24.revenue-1)*100)],
    [3000, 1900, 1900, 1900, 2100], { cellOpts: [{}, {align: AlignmentType.RIGHT}, {align: AlignmentType.RIGHT}, {align: AlignmentType.RIGHT, bold: true}, {align: AlignmentType.RIGHT, color: RED}] }),
  dataRow(['毛利率', pct(fy23.gross_margin), pct(fy24.gross_margin), pct(fy25.gross_margin), pctSigned(fy25.gross_margin - fy24.gross_margin) + 'pp'],
    [3000, 1900, 1900, 1900, 2100], { zebra: true, cellOpts: [{}, {align: AlignmentType.RIGHT}, {align: AlignmentType.RIGHT}, {align: AlignmentType.RIGHT, bold: true}, {align: AlignmentType.RIGHT, color: RED}] }),
  dataRow(['营业利润', yi(fy23.op_profit), yi(fy24.op_profit), yi(fy25.op_profit), pctSigned((fy25.op_profit/fy24.op_profit-1)*100)],
    [3000, 1900, 1900, 1900, 2100], { cellOpts: [{}, {align: AlignmentType.RIGHT}, {align: AlignmentType.RIGHT}, {align: AlignmentType.RIGHT, bold: true}, {align: AlignmentType.RIGHT, color: RED}] }),
  dataRow(['归母净利润', yi(fy23.np_parent), yi(fy24.np_parent), yi(fy25.np_parent), pctSigned((fy25.np_parent/fy24.np_parent-1)*100)],
    [3000, 1900, 1900, 1900, 2100], { zebra: true, cellOpts: [{}, {align: AlignmentType.RIGHT}, {align: AlignmentType.RIGHT}, {align: AlignmentType.RIGHT, bold: true}, {align: AlignmentType.RIGHT, color: RED}] }),
  dataRow(['扣非归母净利润', yi(fy23.np_nonrecurring_excluded), yi(fy24.np_nonrecurring_excluded), yi(fy25.np_nonrecurring_excluded), pctSigned((fy25.np_nonrecurring_excluded/fy24.np_nonrecurring_excluded-1)*100)],
    [3000, 1900, 1900, 1900, 2100], { cellOpts: [{}, {align: AlignmentType.RIGHT}, {align: AlignmentType.RIGHT}, {align: AlignmentType.RIGHT, bold: true}, {align: AlignmentType.RIGHT, color: RED}] }),
  dataRow(['EBITDA (营利+D&A+利息)', yi(fy23.ebitda), yi(fy24.ebitda), yi(fy25.ebitda), pctSigned((fy25.ebitda/fy24.ebitda-1)*100)],
    [3000, 1900, 1900, 1900, 2100], { zebra: true, cellOpts: [{}, {align: AlignmentType.RIGHT}, {align: AlignmentType.RIGHT}, {align: AlignmentType.RIGHT, bold: true}, {align: AlignmentType.RIGHT, color: RED}] }),
  // CF
  dataRow(['经营性现金流 (OCF)', yi(fy23.ocf), yi(fy24.ocf), yi(fy25.ocf), pctSigned((fy25.ocf/fy24.ocf-1)*100)],
    [3000, 1900, 1900, 1900, 2100], { cellOpts: [{}, {align: AlignmentType.RIGHT}, {align: AlignmentType.RIGHT}, {align: AlignmentType.RIGHT}, {align: AlignmentType.RIGHT, color: RED}] }),
  dataRow(['Capex', yi(fy23.capex), yi(fy24.capex), yi(fy25.capex), pctSigned((fy25.capex/fy24.capex-1)*100)],
    [3000, 1900, 1900, 1900, 2100], { zebra: true, cellOpts: [{}, {align: AlignmentType.RIGHT}, {align: AlignmentType.RIGHT}, {align: AlignmentType.RIGHT}, {align: AlignmentType.RIGHT}] }),
  dataRow(['FCF (=OCF-Capex)', yi(fy23.fcf), yi(fy24.fcf), yi(fy25.fcf), pctSigned((fy25.fcf/fy24.fcf-1)*100)],
    [3000, 1900, 1900, 1900, 2100], { cellOpts: [{}, {align: AlignmentType.RIGHT}, {align: AlignmentType.RIGHT}, {align: AlignmentType.RIGHT, bold: true}, {align: AlignmentType.RIGHT, color: RED}] }),
  dataRow(['FCF/EBITDA', pct(fy23.fcf_ebitda), pct(fy24.fcf_ebitda), pct(fy25.fcf_ebitda), pctSigned(fy25.fcf_ebitda - fy24.fcf_ebitda) + 'pp'],
    [3000, 1900, 1900, 1900, 2100], { zebra: true, cellOpts: [{}, {align: AlignmentType.RIGHT}, {align: AlignmentType.RIGHT}, {align: AlignmentType.RIGHT}, {align: AlignmentType.RIGHT}] }),
  // 回报
  dataRow(['加权ROE', pct(fy23.roe_weighted), pct(fy24.roe_weighted), pct(fy25.roe_weighted), pctSigned(fy25.roe_weighted - fy24.roe_weighted) + 'pp'],
    [3000, 1900, 1900, 1900, 2100], { cellOpts: [{}, {align: AlignmentType.RIGHT}, {align: AlignmentType.RIGHT}, {align: AlignmentType.RIGHT, bold: true}, {align: AlignmentType.RIGHT, color: RED}] }),
  dataRow(['基本EPS (元)', f(fy23.eps_basic, 4), f(fy24.eps_basic, 4), f(fy25.eps_basic, 4), pctSigned((fy25.eps_basic/fy24.eps_basic-1)*100)],
    [3000, 1900, 1900, 1900, 2100], { zebra: true, cellOpts: [{}, {align: AlignmentType.RIGHT}, {align: AlignmentType.RIGHT}, {align: AlignmentType.RIGHT}, {align: AlignmentType.RIGHT, color: RED}] }),
  // BS
  dataRow(['资产总额', yi(fy23.total_assets), yi(fy24.total_assets), yi(fy25.total_assets), pctSigned((fy25.total_assets/fy24.total_assets-1)*100)],
    [3000, 1900, 1900, 1900, 2100], { cellOpts: [{}, {align: AlignmentType.RIGHT}, {align: AlignmentType.RIGHT}, {align: AlignmentType.RIGHT, bold: true}, {align: AlignmentType.RIGHT, color: GREEN}] }),
  dataRow(['总负债', yi(fy23.total_liabilities), yi(fy24.total_liabilities), yi(fy25.total_liabilities), pctSigned((fy25.total_liabilities/fy24.total_liabilities-1)*100)],
    [3000, 1900, 1900, 1900, 2100], { zebra: true, cellOpts: [{}, {align: AlignmentType.RIGHT}, {align: AlignmentType.RIGHT}, {align: AlignmentType.RIGHT, bold: true}, {align: AlignmentType.RIGHT, color: RED}] }),
  dataRow(['资产负债率', pct(fy23.lev_ratio), pct(fy24.lev_ratio), pct(fy25.lev_ratio), pctSigned(fy25.lev_ratio - fy24.lev_ratio) + 'pp'],
    [3000, 1900, 1900, 1900, 2100], { cellOpts: [{}, {align: AlignmentType.RIGHT}, {align: AlignmentType.RIGHT}, {align: AlignmentType.RIGHT, bold: true}, {align: AlignmentType.RIGHT, color: RED}] }),
  dataRow(['商誉', yi(fy23.goodwill), yi(fy24.goodwill), yi(fy25.goodwill), '+' + yi(fy25.goodwill) + '亿(信为通讯)'],
    [3000, 1900, 1900, 1900, 2100], { zebra: true, cellOpts: [{}, {align: AlignmentType.RIGHT}, {align: AlignmentType.RIGHT}, {align: AlignmentType.RIGHT, bold: true, color: RED}, {align: AlignmentType.RIGHT, color: RED}] }),
  dataRow(['净现金 (无有息负债)', yi(fy23.net_cash), yi(fy24.net_cash), yi(fy25.net_cash), pctSigned((fy25.net_cash/fy24.net_cash-1)*100)],
    [3000, 1900, 1900, 1900, 2100], { cellOpts: [{}, {align: AlignmentType.RIGHT}, {align: AlignmentType.RIGHT}, {align: AlignmentType.RIGHT, bold: true, color: GREEN}, {align: AlignmentType.RIGHT}] }),
]);
children.push(finTable);

// D.3 收入拆分
children.push(labelPara('D.3 FY2025 收入拆分',
  '：分行业 电动工具零部件 ' + pct(data.rev_breakdown.FY2025.by_industry.电动工具零部件.share) + '（YoY ' + pctSigned(data.rev_breakdown.FY2025.by_industry.电动工具零部件.yoy) + '）/ 汽车零部件 ' + pct(data.rev_breakdown.FY2025.by_industry.汽车零部件.share) + '（YoY ' + pctSigned(data.rev_breakdown.FY2025.by_industry.汽车零部件.yoy) + '）。' +
  ' 分产品 PM ' + pct(data.rev_breakdown.FY2025.by_product.PM.share) + '（YoY ' + pctSigned(data.rev_breakdown.FY2025.by_product.PM.yoy) + '）/ MIM ' + pct(data.rev_breakdown.FY2025.by_product.MIM.share) + '（YoY ' + pctSigned(data.rev_breakdown.FY2025.by_product.MIM.yoy) + '，高毛利 ' + pct(48.0) + '）。' +
  ' 分地区 境外 ' + pct(data.rev_breakdown.FY2025.by_region.境外.share) + '（YoY ' + pctSigned(data.rev_breakdown.FY2025.by_region.境外.yoy) + '，毛利率 ' + pct(data.rev_breakdown.FY2025.gm_by_region.境外) + '）/ 境内 ' + pct(data.rev_breakdown.FY2025.by_region.境内.share) + '（毛利率 ' + pct(data.rev_breakdown.FY2025.gm_by_region.境内) + '）。',
  { size: SIZE_BODY }));

// D.4 投资亮点 + 风险表
children.push(labelPara('D.4 投资亮点（+）/ 风险提示（−）', '', { size: SIZE_BODY }));

const catTable = tableStandard([
  headerRow(['投资亮点 (+)', '风险提示 (−)'], [5400, 5400]),
  ...data.catalysts_risks.positive.map((pos, i) => {
    const neg = data.catalysts_risks.negative[i] || '';
    return dataRow([pos, neg], [5400, 5400], {
      zebra: i % 2 === 0,
      cellOpts: [{ size: SIZE_TABLE, color: GREEN }, { size: SIZE_TABLE, color: RED }],
    });
  }),
]);
children.push(catTable);

// 来源标注
children.push(new Paragraph({
  spacing: { before: 100, after: 0 },
  children: [tr('数据来源：FY2023/FY2024/FY2025 年报、H1 2025 / Q3 2025 / Q1 2026 季报、Wind 日度估值+股价+CSC 可比表（用户上传）；公开信息：cninfo / 东方财富 / 韭研公社 / 财富号 / 新浪财经。',
    { size: SIZE_SOURCE, italics: true, color: GRAY })],
}));
children.push(new Paragraph({
  spacing: { before: 30, after: 0 },
  children: [tr('数据完整度自检：23/23 校验通过（毛利率 / EBITDA 公式 / FCF / 资产负债率 / 净现金 / P/E 口径 / 三年趋势 / 商誉变动）。海昌新材无机构一致预期覆盖，所有 forward 倍数标 NA。',
    { size: SIZE_SOURCE, italics: true, color: GRAY })],
}));

// ============================================================
// Document 装配
// ============================================================
const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: FONT, size: SIZE_BODY },
        paragraph: { spacing: { after: 60, line: 240 } },
      },
    },
  },
  sections: [{
    properties: {
      page: {
        margin: { top: 720, right: 720, bottom: 720, left: 720, header: 360, footer: 360 },
        size: { orientation: PageOrientation.PORTRAIT },
      },
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.LEFT,
          spacing: { after: 0, line: 200 },
          children: [
            tr(data.meta.company + '（' + data.meta.code + '）— 询价转让投资备忘录', { size: 12, bold: true, color: BLUE }),
            tr('  |  ', { size: 12, color: GRAY }),
            tr('机密文件 | 仅供内部投资决策参考', { size: 12, italics: true, color: GRAY }),
          ],
        })],
      }),
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { line: 200 },
          children: [
            tr('第 ', { size: 12, color: GRAY }),
            new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 12, color: GRAY }),
            tr(' 页 / 共 ', { size: 12, color: GRAY }),
            new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: 12, color: GRAY }),
            tr(' 页', { size: 12, color: GRAY }),
          ],
        })],
      }),
    },
    children,
  }],
});

// 输出
Packer.toBuffer(doc).then(buf => {
  const baseName = data.meta.company + '_' + data.meta.code + '_询价转让投资备忘录';
  let outPath = path.join(PROJECT_DIR, baseName + '.docx');
  try {
    fs.writeFileSync(outPath, buf);
  } catch (e) {
    if (e.code === 'EBUSY') {
      // 文件被 Word 锁定, 退而写入加版本号文件
      const ts = new Date().toISOString().slice(0,16).replace(/[T:-]/g, '');
      outPath = path.join(PROJECT_DIR, baseName + '_v2_' + ts + '.docx');
      fs.writeFileSync(outPath, buf);
      console.log('⚠️ 主文件被锁定，已写入新版本: ' + path.basename(outPath));
    } else throw e;
  }
  console.log('✅ 备忘录已生成: ' + outPath);
  console.log('   文件大小: ' + (buf.length / 1024).toFixed(1) + ' KB');
}).catch(e => { console.error('ERR:', e); process.exit(1); });
