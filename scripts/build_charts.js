// 通用图表生成模板 (可复用于所有询价转让标的)
// 用法: node scripts/build_charts.js --project projects/<标的名>
// 输入: <project>/data/<...Wind...>.xlsx + <project>/data.json
// 输出: <project>/chart{1,2,3}.png

const fs = require('fs');
const path = require('path');
const { ChartJSNodeCanvas } = require('chartjs-node-canvas');
const ExcelJS = require('exceljs');

// 解析 --project <path>
function getProjectDir() {
  const args = process.argv.slice(2);
  const i = args.indexOf('--project');
  if (i >= 0 && args[i + 1]) return path.resolve(args[i + 1]);
  console.error('Usage: node scripts/build_charts.js --project projects/<标的名>');
  process.exit(1);
}
const PROJECT_DIR = getProjectDir();
const dataJsonPath = path.join(PROJECT_DIR, 'data.json');
if (!fs.existsSync(dataJsonPath)) {
  console.error('data.json 不存在: ' + dataJsonPath + '\n请先运行 build_data.js');
  process.exit(1);
}
const data = JSON.parse(fs.readFileSync(dataJsonPath, 'utf8'));
const dataSubdir = path.join(PROJECT_DIR, 'data');
const xlsxFiles = fs.readdirSync(dataSubdir).filter(f => f.endsWith('.xlsx'));
const FILE = path.join(dataSubdir, xlsxFiles[0]);
console.log('Project: ' + PROJECT_DIR);
console.log('Excel:   ' + path.basename(FILE));

function getVal(cell) {
  let v = cell.value;
  if (v === null || v === undefined) return null;
  if (typeof v === 'object') {
    if (v.text !== undefined) v = v.text;
    else if (v.result !== undefined) v = v.result;
  }
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  return v;
}
function num(v) {
  if (v === null || v === undefined || v === '') return null;
  const n = typeof v === 'number' ? v : parseFloat(String(v).replace(/,/g, ''));
  return isNaN(n) ? null : n;
}

async function main() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(FILE);

  // 读取日度估值
  const wsVal = wb.getWorksheet('日度估值和交易信息时间序列');
  const valDataRaw = [];
  for (let r = 3; r <= wsVal.rowCount; r++) {
    const row = wsVal.getRow(r);
    const date = getVal(row.getCell(1));
    if (!date) continue;
    valDataRaw.push({
      date,
      close: num(getVal(row.getCell(2))),
      eps_ttm: num(getVal(row.getCell(3))),
      ev_excl_cash_wan: num(getVal(row.getCell(5))),
      ebitda_ttm_wan: num(getVal(row.getCell(6))),
      volume: num(getVal(row.getCell(8))),
    });
  }
  const valData = valDataRaw.reverse();

  // 读取日度股价 — peer 列顺序由 data.peers 决定 (Wind 表 column 6 起为 [target, peer1, peer2, peer3])
  const peerKeys = Object.keys(data.peers);  // 第一个key是target本身
  const wsPx = wb.getWorksheet('日度股价');
  const pxRaw = [];
  for (let r = 3; r <= wsPx.rowCount; r++) {
    const row = wsPx.getRow(r);
    const date = getVal(row.getCell(1));
    if (!date) continue;
    const rec = { date };
    peerKeys.forEach((k, i) => { rec[k] = num(getVal(row.getCell(6 + i))); });
    pxRaw.push(rec);
  }
  const pxData = pxRaw.reverse();

  console.log('日度估值: ' + valData.length + ' rows, ' + valData[0].date + ' → ' + valData[valData.length - 1].date);
  console.log('日度股价: ' + pxData.length + ' rows');

  // 画布配置 - 高质量
  const chart = new ChartJSNodeCanvas({
    width: 1200,
    height: 360,
    backgroundColour: 'white',
    chartCallback: (ChartJS) => {
      ChartJS.defaults.font.family = 'Arial, Helvetica, sans-serif';
      ChartJS.defaults.font.size = 11;
    },
  });

  // ============ Chart 1: TTM PE Band ============
  const peSeries = valData.map(d => {
    if (d.close && d.eps_ttm && d.eps_ttm > 0) return d.close / d.eps_ttm;
    return null;
  });
  const peValid = peSeries.filter(x => x !== null);
  const peMean = peValid.reduce((s, x) => s + x, 0) / peValid.length;
  const peSd = Math.sqrt(peValid.reduce((s, x) => s + (x - peMean) ** 2, 0) / peValid.length);

  const labels1 = valData.map((d, i) => i % 60 === 0 ? d.date.slice(0, 7) : '');
  const meanLine = peSeries.map(_ => peMean);
  const plus1sd = peSeries.map(_ => peMean + peSd);
  const minus1sd = peSeries.map(_ => peMean - peSd);

  const config1 = {
    type: 'line',
    data: {
      labels: labels1,
      datasets: [
        { label: 'TTM PE', data: peSeries, borderColor: '#1E5BBA', backgroundColor: 'rgba(30,91,186,0.08)', borderWidth: 1.8, pointRadius: 0, fill: true, tension: 0.1 },
        { label: '均值 (' + peMean.toFixed(1) + 'x)', data: meanLine, borderColor: '#666', borderWidth: 1.2, borderDash: [6,4], pointRadius: 0 },
        { label: '+1SD (' + (peMean + peSd).toFixed(1) + 'x)', data: plus1sd, borderColor: '#D32F2F', borderWidth: 1, borderDash: [4,3], pointRadius: 0 },
        { label: '-1SD (' + (peMean - peSd).toFixed(1) + 'x)', data: minus1sd, borderColor: '#388E3C', borderWidth: 1, borderDash: [4,3], pointRadius: 0 },
      ],
    },
    options: {
      responsive: false,
      plugins: {
        title: { display: true, text: '图1  ' + data.meta.company + ' TTM PE 历史Band (近3年, n=' + peValid.length + ')', font: { size: 13, weight: 'bold' } },
        legend: { position: 'top', labels: { font: { size: 10 } } },
      },
      scales: {
        x: { ticks: { font: { size: 9 }, maxRotation: 0, autoSkip: false } },
        y: { title: { display: true, text: 'TTM PE (x)' }, ticks: { font: { size: 10 } } },
      },
    },
  };

  const png1 = await chart.renderToBuffer(config1);
  fs.writeFileSync(path.join(PROJECT_DIR, 'chart1_pe_band.png'), png1);
  console.log('✅ Chart 1 (PE Band) saved');

  // ============ Chart 2: 股价对比 (归一化) ============
  const base = pxData[0];
  const normalized = pxData.map(d => {
    const out = { date: d.date };
    peerKeys.forEach(k => { out[k] = (d[k] / base[k]) * 100; });
    return out;
  });
  const labels2 = normalized.map((d, i) => i % 60 === 0 ? d.date.slice(0, 7) : '');
  const peerColors = ['#D32F2F', '#1976D2', '#388E3C', '#F57C00', '#7B1FA2', '#FFA000'];

  const config2 = {
    type: 'line',
    data: {
      labels: labels2,
      datasets: peerKeys.map((k, i) => ({
        label: data.peers[k].name + ' (' + data.peers[k].code + ')',
        data: normalized.map(d => d[k]),
        borderColor: peerColors[i % peerColors.length],
        borderWidth: i === 0 ? 2.5 : 1.5,
        pointRadius: 0,
        tension: 0.1,
      })),
    },
    options: {
      responsive: false,
      plugins: {
        title: { display: true, text: '图2  股价走势对比 (归一化, ' + pxData[0].date + ' = 100, ~3年)', font: { size: 13, weight: 'bold' } },
        legend: { position: 'top', labels: { font: { size: 10 } } },
      },
      scales: {
        x: { ticks: { font: { size: 9 }, maxRotation: 0, autoSkip: false } },
        y: { title: { display: true, text: '归一化价格 (起点=100)' }, ticks: { font: { size: 10 } } },
      },
    },
  };
  const png2 = await chart.renderToBuffer(config2);
  fs.writeFileSync(path.join(PROJECT_DIR, 'chart2_peer_comparison.png'), png2);
  console.log('✅ Chart 2 (peer comparison) saved');

  // ============ Chart 3: 海昌单独股价 + 均线 + 关键价位 ============
  // 计算 MA20, MA60, MA120
  const targetPx = pxData.map(d => d[peerKeys[0]]);  // 第一个key是标的本身
  function MA(arr, win) {
    return arr.map((_, i) => {
      if (i < win - 1) return null;
      const slice = arr.slice(i - win + 1, i + 1);
      return slice.reduce((s, x) => s + x, 0) / win;
    });
  }
  const ma20 = MA(targetPx, 20);
  const ma60 = MA(targetPx, 60);
  const ma120 = MA(targetPx, 120);

  // 关键价位 — 从 data.json 读取卖方底价与场景判定
  const currentPx = targetPx[targetPx.length - 1];
  const floorPx = data.valuation.floor_price;  // 场景A时存在; 场景B时通常为null/undefined
  const floorDiscPct = floorPx ? (1 - floorPx / currentPx) * 100 : null;
  const disc10 = currentPx * 0.9;
  const disc20 = currentPx * 0.8;
  const disc30 = currentPx * 0.7;
  // 场景A: 跳过低于卖方底价的折扣档; 场景B: 三档都显示
  const showDisc30 = floorPx ? floorDiscPct > 30 : true;
  // 52周高低
  const last250 = targetPx.slice(-250);
  const w52h = Math.max(...last250);
  const w52l = Math.min(...last250);

  const horizontalLine = (n, value) => new Array(n).fill(value);
  const N = pxData.length;
  const labels3 = pxData.map((d, i) => i % 60 === 0 ? d.date.slice(0, 7) : '');

  const config3 = {
    type: 'line',
    data: {
      labels: labels3,
      datasets: [
        { label: '收盘价', data: targetPx, borderColor: '#1A1A1A', borderWidth: 1.8, pointRadius: 0, tension: 0.1 },
        { label: 'MA20', data: ma20, borderColor: '#1976D2', borderWidth: 1, pointRadius: 0 },
        { label: 'MA60', data: ma60, borderColor: '#388E3C', borderWidth: 1, pointRadius: 0 },
        { label: 'MA120', data: ma120, borderColor: '#9E9E9E', borderWidth: 1, pointRadius: 0 },
        { label: '当前 ' + currentPx.toFixed(2), data: horizontalLine(N, currentPx), borderColor: '#D32F2F', borderWidth: 1, borderDash: [4,3], pointRadius: 0 },
        { label: '52W高 ' + w52h.toFixed(2), data: horizontalLine(N, w52h), borderColor: '#7B1FA2', borderWidth: 0.8, borderDash: [2,2], pointRadius: 0 },
        { label: '52W低 ' + w52l.toFixed(2), data: horizontalLine(N, w52l), borderColor: '#7B1FA2', borderWidth: 0.8, borderDash: [2,2], pointRadius: 0 },
        { label: '折10% ' + disc10.toFixed(2), data: horizontalLine(N, disc10), borderColor: '#F57C00', borderWidth: 1, borderDash: [6,3], pointRadius: 0 },
        { label: '折20% ' + disc20.toFixed(2), data: horizontalLine(N, disc20), borderColor: '#FFA000', borderWidth: 1, borderDash: [6,3], pointRadius: 0 },
        ...(showDisc30 ? [{ label: '折30% ' + disc30.toFixed(2), data: horizontalLine(N, disc30), borderColor: '#FFD54F', borderWidth: 1, borderDash: [6,3], pointRadius: 0 }] : []),
        ...(floorPx ? [{ label: '卖方底价 ' + floorPx.toFixed(2) + ' (-' + floorDiscPct.toFixed(1) + '%)', data: horizontalLine(N, floorPx), borderColor: '#C2185B', borderWidth: 2, borderDash: [], pointRadius: 0 }] : []),
      ],
    },
    options: {
      responsive: false,
      plugins: {
        title: { display: true, text: '图3  ' + data.meta.company + ' (' + data.meta.code + ') 绝对股价 + 均线 + 关键价位' + (floorPx ? '（含卖方底价 ' + floorPx.toFixed(2) + '）' : '（场景B无卖方底价）') + '，近3年', font: { size: 13, weight: 'bold' } },
        legend: { position: 'top', labels: { font: { size: 9 }, boxWidth: 18 } },
      },
      scales: {
        x: { ticks: { font: { size: 9 }, maxRotation: 0, autoSkip: false } },
        y: { title: { display: true, text: '股价 (元)' }, ticks: { font: { size: 10 } } },
      },
    },
  };
  const png3 = await chart.renderToBuffer(config3);
  fs.writeFileSync(path.join(PROJECT_DIR, 'chart3_price_levels.png'), png3);
  console.log('✅ Chart 3 (price levels) saved');

  console.log('\n所有图表已生成');
}

main().catch(e => { console.error('ERR:', e.message, e.stack); process.exit(1); });
