/**
 * 查验率监控核心引擎
 * 数据口径：以发货日为基准，票数按主出仓单号换行分割
 */

// ===================== 常量 =====================
const COUNTRY_MAP = {
  '德国': '欧盟',
  '法国': '欧盟',
  '阿拉伯联合酋长国': '阿联酋'
};

const SETTINGS = {
  highRisk: 5.0,   // 综合查验率 >5% 为高风险
  midRisk: 3.0,    // 综合查验率 3%~5% 为中风险
  minSample: 10    // 统计最小样本数（低于此样本不预警）
};

// 中文颜色配置（涨红跌绿）
const COLORS = {
  primary: '#07c160',
  danger: '#d62929',
  warning: '#e6a23c',
  info: '#1764e8',
  text: '#222',
  textSecondary: '#666',
  bg: '#f5f5f5',
  card: '#fff',
  domestic: '#d62929',   // 起运港 - 红色（即时风险）
  foreign: '#1764e8',    // 目的港 - 蓝色（滞后风险）
  overall: '#07c160'     // 综合 - 绿色
};

// ===================== 全局状态 =====================
let rawData = [];       // 原始行数据
let records = [];       // 展开后的票记录
let maxShipDate = null; // 数据中最大发货日期
let aggByTime = {};     // 时间聚合
let aggByChannel = {};  // 渠道大类聚合
let aggByAgent = {};    // 代理聚合
let aggByUsSub = {};    // 美国子渠道聚合
let aggByCustomer = {}; // 客户聚合
let aggByProduct = {};  // 产品属性聚合
let aggByLogistic = {}; // 素芸物流渠道聚合
let aggByAgentWeekly = {};  // 代理周度拆解 {agent: {weekKey: {total, inspected, domestic}}}
let aggByChannelMonth = {}; // 渠道月度拆解
let aggByChannelWeekly = {}; // 渠道周度拆解 {channel: {weekKey: {total, inspected, domestic}}}
let allWeekKeys = [];    // 所有周Key（排序后）
let allMonthKeys = [];   // 所有月份Key（排序后）
let selectedChannels = []; // 用户选中的渠道（用于趋势图）
let channelTrendMetric = 'overall'; // 趋势图指标: 'overall' | 'domestic' | 'foreign'

// ===================== 工具函数 =====================

function parseDate(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  if (typeof val === 'number') {
    // Excel 日期序列号 -> JS Date (1900/1904 兼容由 SheetJS 处理)
    // 这里假设 SheetJS 已经转换，如果未转换则手动处理
    return new Date((val - 25569) * 86400 * 1000);
  }
  const d = new Date(val);
  return isNaN(d.getTime()) ? null : d;
}

function formatDate(d) {
  if (!d) return '';
  const dt = d instanceof Date ? d : parseDate(d);
  if (!dt) return String(d);
  return `${dt.getFullYear()}-${String(dt.getMonth()+1).padStart(2,'0')}-${String(dt.getDate()).padStart(2,'0')}`;
}

function formatDateShort(d) {
  if (!d) return '';
  const dt = d instanceof Date ? d : parseDate(d);
  if (!dt) return String(d);
  return `${dt.getMonth()+1}/${dt.getDate()}`;
}

function formatDateYM(d) {
  if (!d) return '';
  const dt = d instanceof Date ? d : parseDate(d);
  if (!dt) return String(d);
  return `${dt.getFullYear()}-${String(dt.getMonth()+1).padStart(2,'0')}`;
}

function getWeekKey(d) {
  const dt = d instanceof Date ? d : parseDate(d);
  if (!dt) return '';
  const day = dt.getDay() || 7;
  const mon = new Date(dt);
  mon.setDate(dt.getDate() - day + 1);
  const y = mon.getFullYear();
  const m = mon.getMonth() + 1;
  const dayOfYear = Math.floor((mon - new Date(y, 0, 1)) / 86400000) + 1;
  const w = Math.ceil(dayOfYear / 7);
  return `${y}-W${String(w).padStart(2,'0')}`;
}

function safeStr(val) {
  if (val === undefined || val === null) return '';
  return String(val).trim();
}

function safeNum(val) {
  const n = parseFloat(val);
  return isNaN(n) ? 0 : n;
}

// ===================== Excel 读取 =====================

function readExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error('读取文件失败'));
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });
        // 查找主数据 sheet（名称可能带空格）
        let sheetName = workbook.SheetNames.find(s => s.trim().startsWith('空运+快递+陆运'));
        if (!sheetName) sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
        resolve(rows);
      } catch (err) {
        reject(err);
      }
    };
    reader.readAsArrayBuffer(file);
  });
}

// ===================== 数据清洗与展开 =====================

function processData(rows) {
  const list = [];
  let maxD = null;

  for (const row of rows) {
    // B列 主出仓单号 -> 按换行分割票数
    const ticketStr = safeStr(row['主出仓单号']);
    const tickets = ticketStr.split(/\r?\n/).map(s => s.trim()).filter(s => s);
    const ticketCount = tickets.length > 0 ? tickets.length : 1; // 空则计1票

    // Q列 状态备注 -> 查验判定（金标准）
    const remark = safeStr(row['状态备注']);
    const isDomestic = remark.includes('国内查验');
    const isForeign = remark.includes('国外查验');
    const isInspected = isDomestic || isForeign;

    // G列 国家 -> 映射
    let country = safeStr(row['国家']);
    if (COUNTRY_MAP[country]) country = COUNTRY_MAP[country];

    // AW列 类型
    const type = safeStr(row['类型']);

    // K列 素芸物流渠道
    const logisticChannel = safeStr(row['素芸物流渠道']);

    // 渠道大类 = 国家+类型
    const channel = country && type ? (country + type) : '';

    // 美国子渠道（仅美国海运）
    let usSubChannel = '';
    if (country === '美国' && type === '海运') {
      if (logisticChannel.includes('美森')) usSubChannel = '美森';
      else if (logisticChannel.includes('快船')) usSubChannel = '快船';
      else usSubChannel = '慢船';
    }

    // T列 仓库出货日期
    const shipDate = parseDate(row['仓库出货日期']);

    if (shipDate) {
      if (!maxD || shipDate > maxD) maxD = shipDate;
    }

    // 查验时效（BI/BJ列，数值型为天数）
    const destInspectTime = safeNum(row['目的地查验时间']);
    const domInspectTime = safeNum(row['国内查验时间']);
    const avgInspectTime = (destInspectTime > 0 && domInspectTime > 0)
      ? (destInspectTime + domInspectTime) / 2
      : (destInspectTime > 0 ? destInspectTime : (domInspectTime > 0 ? domInspectTime : 0));

    list.push({
      ticketCount,
      tickets,
      isDomestic,
      isForeign,
      isInspected,
      country,
      type,
      channel,
      usSubChannel,
      logisticChannel,
      agent: safeStr(row['代理']),
      agentChannel: safeStr(row['代理渠道']),
      customer: safeStr(row['客户']),
      productAttr: safeStr(row['产品属性']),
      shipDate,
      remark,
      avgInspectTime,
      boxes: safeNum(row['箱数']),
      pieces: safeNum(row['件数']),
      weight: safeNum(row['毛重']),
      volume: safeNum(row['方数CBM'])
    });
  }

  return { records: list, maxDate: maxD };
}

// ===================== 聚合计算 =====================

function createAgg() {
  return {
    totalTickets: 0,
    totalBoxes: 0,
    totalPieces: 0,
    totalWeight: 0,
    totalVolume: 0,
    domesticTickets: 0,
    foreignTickets: 0,
    inspectedTickets: 0,
    inspectTimeSum: 0,
    inspectTimeCount: 0,
    // 用于环比的月度/周度数据
    timeSeries: {} // key: weekKey/monthKey -> {total, inspected}
  };
}

function addToAgg(agg, rec) {
  agg.totalTickets += rec.ticketCount;
  agg.totalBoxes += rec.boxes * rec.ticketCount;
  agg.totalPieces += rec.pieces * rec.ticketCount;
  agg.totalWeight += rec.weight * rec.ticketCount;
  agg.totalVolume += rec.volume * rec.ticketCount;

  if (rec.isDomestic) agg.domesticTickets += rec.ticketCount;
  if (rec.isForeign) agg.foreignTickets += rec.ticketCount;
  if (rec.isInspected) agg.inspectedTickets += rec.ticketCount;

  if (rec.avgInspectTime > 0) {
    agg.inspectTimeSum += rec.avgInspectTime * rec.ticketCount;
    agg.inspectTimeCount += rec.ticketCount;
  }
}

function finalizeAgg(agg) {
  const r = { ...agg };
  r.domesticRate = r.totalTickets > 0 ? (r.domesticTickets / r.totalTickets * 100) : 0;
  r.foreignRate = r.totalTickets > 0 ? (r.foreignTickets / r.totalTickets * 100) : 0;
  r.overallRate = r.totalTickets > 0 ? (r.inspectedTickets / r.totalTickets * 100) : 0;
  r.avgInspectTime = r.inspectTimeCount > 0 ? (r.inspectTimeSum / r.inspectTimeCount) : 0;
  return r;
}

function aggregateBy(records, getKey, opts = {}) {
  const map = {};
  const { timeGranularity = null } = opts;

  for (const rec of records) {
    if (!rec.shipDate) continue;

    const key = getKey(rec);
    if (!key) continue;
    if (!map[key]) map[key] = createAgg();
    addToAgg(map[key], rec);

    if (timeGranularity) {
      const tKey = timeGranularity === 'week' ? getWeekKey(rec.shipDate) : formatDateYM(rec.shipDate);
      if (!map[key].timeSeries[tKey]) {
        map[key].timeSeries[tKey] = { totalTickets: 0, inspectedTickets: 0, domesticTickets: 0, foreignTickets: 0 };
      }
      map[key].timeSeries[tKey].totalTickets += rec.ticketCount;
      if (rec.isInspected) map[key].timeSeries[tKey].inspectedTickets += rec.ticketCount;
      if (rec.isDomestic) map[key].timeSeries[tKey].domesticTickets += rec.ticketCount;
      if (rec.isForeign) map[key].timeSeries[tKey].foreignTickets += rec.ticketCount;
    }
  }

  // finalize
  const result = {};
  for (const key in map) {
    result[key] = finalizeAgg(map[key]);
  }
  return result;
}

function aggregateByTime(records, granularity) {
  const map = {};
  for (const rec of records) {
    if (!rec.shipDate) continue;
    const key = granularity === 'week' ? getWeekKey(rec.shipDate) : formatDateYM(rec.shipDate);
    if (!key) continue;
    if (!map[key]) map[key] = createAgg();
    addToAgg(map[key], rec);
  }
  const result = {};
  for (const key in map) {
    result[key] = finalizeAgg(map[key]);
  }
  return result;
}

// 按实体(代理/渠道)×周 拆解，返回 {entityKey: {weekKey: {total, inspected, domestic}}}
function aggregateWeeklyByEntity(records, entityField) {
  const map = {};
  for (const rec of records) {
    if (!rec.shipDate) continue;
    const entityKey = rec[entityField];
    if (!entityKey) continue;
    const wk = getWeekKey(rec.shipDate);
    if (!wk) continue;
    if (!map[entityKey]) map[entityKey] = {};
    if (!map[entityKey][wk]) map[entityKey][wk] = { totalTickets: 0, inspectedTickets: 0, domesticTickets: 0 };
    map[entityKey][wk].totalTickets += rec.ticketCount;
    if (rec.isInspected) map[entityKey][wk].inspectedTickets += rec.ticketCount;
    if (rec.isDomestic) map[entityKey][wk].domesticTickets += rec.ticketCount;
  }
  return map;
}

// 获取指定实体的最近N周查验率序列
function getEntityRateSeries(entityWeekly, entityKey, weekKeys, rateType) {
  const wm = entityWeekly[entityKey];
  return weekKeys.map(wk => {
    const d = wm && wm[wk];
    if (!d || d.totalTickets === 0) return null;
    const key = rateType === 'domestic' ? 'domesticTickets' : 'inspectedTickets';
    return parseFloat((d[key] / d.totalTickets * 100).toFixed(1));
  });
}

// ===================== 预警规则 =====================

function generateAlerts() {
  const alerts = [];

  // 代理预警（全量数据）
  for (const [agent, data] of Object.entries(aggByAgent)) {
    if (data.totalTickets < SETTINGS.minSample) continue;
    if (data.overallRate > SETTINGS.highRisk) {
      alerts.push({ type: 'high', category: '代理', name: agent, rate: data.overallRate, total: data.totalTickets });
    } else if (data.overallRate > SETTINGS.midRisk) {
      alerts.push({ type: 'mid', category: '代理', name: agent, rate: data.overallRate, total: data.totalTickets });
    }
  }

  // 渠道大类预警
  for (const [channel, data] of Object.entries(aggByChannel)) {
    if (data.totalTickets < SETTINGS.minSample) continue;
    if (data.overallRate > SETTINGS.highRisk) {
      alerts.push({ type: 'high', category: '渠道大类', name: channel, rate: data.overallRate, total: data.totalTickets });
    } else if (data.overallRate > SETTINGS.midRisk) {
      alerts.push({ type: 'mid', category: '渠道大类', name: channel, rate: data.overallRate, total: data.totalTickets });
    }
  }

  // 素芸渠道预警
  for (const [channel, data] of Object.entries(aggByLogistic)) {
    if (data.totalTickets < SETTINGS.minSample) continue;
    if (data.overallRate > SETTINGS.highRisk) {
      alerts.push({ type: 'high', category: '素芸渠道', name: channel, rate: data.overallRate, total: data.totalTickets });
    } else if (data.overallRate > SETTINGS.midRisk) {
      alerts.push({ type: 'mid', category: '素芸渠道', name: channel, rate: data.overallRate, total: data.totalTickets });
    }
  }

  // 代理周环比恶化预警（最近两周对比）
  if (allWeekKeys.length >= 2) {
    const lastWk = allWeekKeys[allWeekKeys.length - 1];
    const prevWk = allWeekKeys[allWeekKeys.length - 2];
    for (const [agent, wm] of Object.entries(aggByAgentWeekly)) {
      const cur = wm[lastWk];
      const prev = wm[prevWk];
      if (!cur || !prev || cur.totalTickets < 5 || prev.totalTickets < 5) continue;
      const curRate = cur.inspectedTickets / cur.totalTickets * 100;
      const prevRate = prev.inspectedTickets / prev.totalTickets * 100;
      if (curRate > SETTINGS.midRisk && curRate > prevRate * 1.5) {
        alerts.push({ type: 'trend', category: '代理恶化', name: agent, rate: curRate, total: cur.totalTickets, change: ((curRate - prevRate) / Math.max(prevRate, 0.1) * 100).toFixed(0) + '%' });
      }
    }
    // 渠道环比恶化
    for (const [ch, wm] of Object.entries(aggByChannelWeekly)) {
      const cur = wm[lastWk];
      const prev = wm[prevWk];
      if (!cur || !prev || cur.totalTickets < 5 || prev.totalTickets < 5) continue;
      const curRate = cur.inspectedTickets / cur.totalTickets * 100;
      const prevRate = prev.inspectedTickets / prev.totalTickets * 100;
      if (curRate > SETTINGS.midRisk && curRate > prevRate * 1.5) {
        alerts.push({ type: 'trend', category: '渠道恶化', name: ch, rate: curRate, total: cur.totalTickets, change: ((curRate - prevRate) / Math.max(prevRate, 0.1) * 100).toFixed(0) + '%' });
      }
    }
  }

  // 排序：高风险在前，中风险，趋势
  alerts.sort((a, b) => {
    const order = { high: 0, trend: 1, mid: 2 };
    return order[a.type] - order[b.type];
  });
  return alerts;
}

// ===================== 图表配置生成器 =====================

function createLineChartOption(title, timeKeys, seriesData) {
  return {
    title: { text: title, left: 'center', textStyle: { fontSize: 14 } },
    tooltip: { trigger: 'axis', formatter: params => {
      let s = params[0].axisValue + '<br/>';
      params.forEach(p => {
        s += `${p.marker} ${p.seriesName}: ${p.value.toFixed(2)}%<br/>`;
      });
      return s;
    }},
    legend: { data: seriesData.map(s => s.name), bottom: 0 },
    grid: { left: '3%', right: '4%', bottom: '15%', top: '15%', containLabel: true },
    xAxis: { type: 'category', data: timeKeys, axisLabel: { rotate: 30, fontSize: 10 } },
    yAxis: { type: 'value', name: '查验率%', axisLabel: { formatter: '{value}%' } },
    series: seriesData.map(s => ({
      name: s.name,
      type: 'line',
      data: s.data,
      smooth: true,
      lineStyle: { width: 2 },
      symbol: 'circle',
      symbolSize: 4
    }))
  };
}

function createBarChartOption(title, categories, seriesData, color) {
  return {
    title: { text: title, left: 'center', textStyle: { fontSize: 14 } },
    tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' }, formatter: params => {
      let s = params[0].axisValue + '<br/>';
      params.forEach(p => {
        s += `${p.marker} ${p.seriesName}: ${p.value.toFixed(2)}% (票数:${p.data?.count || ''})<br/>`;
      });
      return s;
    }},
    legend: { data: seriesData.map(s => s.name), bottom: 0 },
    grid: { left: '3%', right: '4%', bottom: '15%', top: '15%', containLabel: true },
    xAxis: { type: 'category', data: categories, axisLabel: { rotate: 45, fontSize: 10 } },
    yAxis: { type: 'value', name: '查验率%', axisLabel: { formatter: '{value}%' } },
    series: seriesData.map(s => ({
      name: s.name,
      type: 'bar',
      data: s.data,
      itemStyle: { color: s.color || color },
      barMaxWidth: 20
    }))
  };
}

function createHeatmapOption(title, xLabels, yLabels, data) {
  return {
    title: { text: title, left: 'center', textStyle: { fontSize: 14 } },
    tooltip: { position: 'top', formatter: p => `${p.name}: ${p.data[2].toFixed(2)}%` },
    grid: { left: '22%', right: '15%', top: '12%', bottom: '18%' },
    xAxis: { type: 'category', data: xLabels, splitArea: { show: true }, axisLabel: { rotate: 45, fontSize: 10 } },
    yAxis: { type: 'category', data: yLabels, splitArea: { show: true }, axisLabel: { fontSize: 10 } },
    visualMap: { min: 0, max: 10, calculable: false, orient: 'vertical', right: '2%', top: 'center', itemHeight: 120, inRange: { color: ['#e8f7ef', '#fff3cd', '#d62929'] } },
    series: [{ name: '查验率', type: 'heatmap', data, label: { show: true, fontSize: 10, formatter: p => p.data[2] > 0 ? p.data[2].toFixed(1) : '' } }]
  };
}

// ===================== 核心处理流程 =====================

async function loadAndProcess(file) {
  try {
    showLoading(true);
    rawData = await readExcel(file);
    const res = processData(rawData);
    records = res.records;
    maxShipDate = res.maxDate;

    // 全量聚合
    aggByTime = { week: aggregateByTime(records, 'week'), month: aggregateByTime(records, 'month') };
    aggByChannel = aggregateBy(records, r => r.channel, { timeGranularity: 'week' });
    aggByChannelMonth = aggregateBy(records, r => r.channel, { timeGranularity: 'month' });
    aggByAgent = aggregateBy(records, r => r.agent, { timeGranularity: 'week' });
    aggByUsSub = aggregateBy(records, r => r.usSubChannel);
    aggByLogistic = aggregateBy(records, r => r.logisticChannel);
    aggByCustomer = aggregateBy(records, r => r.customer);
    aggByProduct = aggregateBy(records, r => r.productAttr);

    // 周度拆解（用于环比对比）
    aggByAgentWeekly = aggregateWeeklyByEntity(records, 'agent');
    aggByChannelWeekly = aggregateWeeklyByEntity(records, 'channel');
    // 收集所有周Key
    const weekSet = new Set();
    Object.values(aggByAgentWeekly).forEach(wm => Object.keys(wm).forEach(k => weekSet.add(k)));
    allWeekKeys = [...weekSet].sort();
    // 收集所有月份Key
    const monthSet = new Set();
    Object.values(aggByChannelMonth).forEach(d => Object.keys(d.timeSeries).forEach(k => monthSet.add(k)));
    allMonthKeys = [...monthSet].sort();
    // 初始化选中渠道：综合查验率 Top6
    selectedChannels = Object.entries(aggByChannel)
      .filter(([k, v]) => k && v.totalTickets >= 20)
      .sort((a, b) => b[1].overallRate - a[1].overallRate)
      .slice(0, 6)
      .map(([k]) => k);

    // 隐藏上传面板，显示主内容
    const uploadPanel = document.getElementById('uploadPanel');
    const mainContent = document.getElementById('mainContent');
    if (uploadPanel) uploadPanel.style.display = 'none';
    if (mainContent) mainContent.style.display = 'block';

    // 填充下钻下拉选项
    const ddSelect = document.getElementById('drilldownSelect');
    if (ddSelect) {
      ddSelect.innerHTML = '<option value="">— 请选择渠道大类 —</option>' +
        Object.entries(aggByChannel)
          .filter(([k, v]) => k && v.totalTickets >= 10)
          .sort((a, b) => b[1].totalTickets - a[1].totalTickets)
          .map(([k, v]) => `<option value="${k}">${k} (${v.totalTickets}票)</option>`)
          .join('');
    }

    // 更新 UI
    updateAllTabs();
    renderChannelSelector();
    showLoading(false);
    return true;
  } catch (err) {
    console.error(err);
    alert('数据处理失败: ' + err.message);
    showLoading(false);
    return false;
  }
}

function showLoading(show) {
  const el = document.getElementById('loadingMask');
  if (el) el.style.display = show ? 'flex' : 'none';
}

// ===================== UI 更新 =====================

let currentGranularity = 'week'; // week | month
let currentView = 'dashboard';   // dashboard | channel | agent | alert
let chartInstances = {};         // 缓存图表实例用于 resize

function updateAllTabs() {
  updateDashboardTab();
  updateChannelTab();
  updateAgentTab();
  updateDrilldownTab();
  updateAlertTab();
}

// 获取或创建图表实例（只初始化一次，后续复用，避免 dispose/reinit 渲染空白）
function setChart(domId, domEl) {
  let chart = chartInstances[domId];
  if (chart && !chart.isDisposed()) return chart;
  chart = echarts.init(domEl);
  chartInstances[domId] = chart;
  return chart;
}

function resizeCharts() {
  Object.values(chartInstances).forEach(c => { try { c.resize(); } catch(e) {} });
}

// 总览
function updateDashboardTab() {
  const total = records.reduce((s, r) => s + r.ticketCount, 0);
  const dom = records.reduce((s, r) => s + (r.isDomestic ? r.ticketCount : 0), 0);
  const foreign = records.reduce((s, r) => s + (r.isForeign ? r.ticketCount : 0), 0);
  const inspected = records.reduce((s, r) => s + (r.isInspected ? r.ticketCount : 0), 0);

  document.getElementById('kpiTotal').textContent = total.toLocaleString();
  document.getElementById('kpiDomestic').textContent = (total > 0 ? (dom / total * 100).toFixed(2) : '0.00') + '%';
  document.getElementById('kpiForeign').textContent = (total > 0 ? (foreign / total * 100).toFixed(2) : '0.00') + '%';
  document.getElementById('kpiOverall').textContent = (total > 0 ? ((dom + foreign) / total * 100).toFixed(2) : '0.00') + '%';

  // KPI: 上周（倒数第2周，本周刚开始数据不全）
  if (allWeekKeys.length >= 2) {
    const prevWk = allWeekKeys[allWeekKeys.length - 2];
    const timeData = aggByTime['week'];
    if (timeData[prevWk]) {
      document.getElementById('kpiWeekRate').textContent = timeData[prevWk].overallRate.toFixed(1) + '%';
      document.getElementById('kpiWeekTotal').textContent = timeData[prevWk].totalTickets.toLocaleString();
    }
  }

  // 趋势图（总体）
  const timeData = aggByTime[currentGranularity];
  const keys = Object.keys(timeData).sort();
  const domesticData = keys.map(k => timeData[k].domesticRate);
  const foreignData = keys.map(k => timeData[k].foreignRate);
  const overallData = keys.map(k => timeData[k].overallRate);

  const chartDom = document.getElementById('trendChart');
  if (chartDom) {
    const chart = setChart('trendChart', chartDom);
    const option = createLineChartOption(
      `总体查验率趋势 (${currentGranularity === 'week' ? '周度' : '月度'})`,
      keys,
      [
        { name: '起运港查验率', data: domesticData },
        { name: '目的港查验率', data: foreignData },
        { name: '综合查验率', data: overallData }
      ]
    );
    option.color = [COLORS.domestic, COLORS.foreign, COLORS.overall];
    chart.setOption(option, true);
  }

  // 渠道趋势：Top 6 渠道的周度查验率对比线图
  updateChannelTrendChart();
}

// 渠道趋势线图（支持自选渠道 + 指标切换 + 周度/月度）
function updateChannelTrendChart() {
  const chartDom = document.getElementById('channelTrendChart');
  if (!chartDom) return;

  const isWeek = currentGranularity === 'week';
  const dataMap = isWeek ? aggByChannel : aggByChannelMonth;
  const timeKeys = isWeek ? allWeekKeys : allMonthKeys;

  if (timeKeys.length < 2) return;

  // 过滤用户选中的渠道（确保数据存在）
  const activeChannels = selectedChannels.filter(ch => dataMap[ch]);
  if (activeChannels.length === 0) {
    // 如果没有选中任何渠道，默认选 Top6
    activeChannels.push(...Object.entries(dataMap)
      .filter(([k, v]) => k && v.totalTickets >= 20)
      .sort((a, b) => b[1].overallRate - a[1].overallRate)
      .slice(0, 6)
      .map(([k]) => k));
  }

  const series = activeChannels.map((ch, idx) => {
    const ts = dataMap[ch].timeSeries || {};
    const data = timeKeys.map(tk => {
      const d = ts[tk];
      if (!d || d.totalTickets === 0) return null;
      let num = 0;
      if (channelTrendMetric === 'domestic') num = d.domesticTickets;
      else if (channelTrendMetric === 'foreign') num = d.foreignTickets;
      else num = d.inspectedTickets;
      return parseFloat((num / d.totalTickets * 100).toFixed(1));
    });
    return { name: ch, data, lineStyle: { width: 2 } };
  });

  const metricName = channelTrendMetric === 'domestic' ? '起运港' : (channelTrendMetric === 'foreign' ? '目的港' : '综合');
  const option = createLineChartOption(`渠道查验率${isWeek ? '周' : '月'}趋势（${metricName}）`, timeKeys, series);
  const palette = ['#d62929', '#e6a23c', '#1764e8', '#07c160', '#8b5cf6', '#ec4899', '#f59e0b', '#10b981', '#6366f1', '#ef4444', '#06b6d4', '#84cc16'];
  option.color = palette.slice(0, series.length);
  option.legend = { data: series.map(s => s.name), bottom: 0, textStyle: { fontSize: 10 } };
  option.grid = { left: '3%', right: '4%', bottom: '20%', top: '15%', containLabel: true };

  const chart = setChart('channelTrendChart', chartDom);
  chart.setOption(option, true);
}

// 渲染渠道选择器标签
function renderChannelSelector() {
  const container = document.getElementById('channelTags');
  if (!container) return;
  const allChannels = Object.entries(aggByChannel)
    .filter(([k, v]) => k && v.totalTickets >= 20)
    .sort((a, b) => b[1].overallRate - a[1].overallRate)
    .map(([k]) => k);
  container.innerHTML = allChannels.map(ch => {
    const isSelected = selectedChannels.includes(ch);
    return `<span class="channel-tag ${isSelected ? 'active' : ''}" onclick="toggleChannel('${ch}')">${ch}</span>`;
  }).join('');
}

// 切换渠道选中状态
function toggleChannel(channel) {
  if (selectedChannels.includes(channel)) {
    selectedChannels = selectedChannels.filter(c => c !== channel);
  } else {
    selectedChannels.push(channel);
  }
  renderChannelSelector();
  updateChannelTrendChart();
}

// 设置趋势图指标
function setChannelMetric(metric) {
  channelTrendMetric = metric;
  document.querySelectorAll('.metric-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.metric === metric);
  });
  updateChannelTrendChart();
}

// 渠道分析
function updateChannelTab() {
  // 渠道大类柱状图
  const channels = Object.entries(aggByChannel)
    .filter(([k, v]) => v.totalTickets >= 10)
    .sort((a, b) => b[1].overallRate - a[1].overallRate)
    .slice(0, 20);
  const catNames = channels.map(([k]) => k);
  const domData = channels.map(([k, v]) => ({ value: v.domesticRate, count: v.totalTickets }));
  const foreignData = channels.map(([k, v]) => ({ value: v.foreignRate, count: v.totalTickets }));
  const overallData = channels.map(([k, v]) => ({ value: v.overallRate, count: v.totalTickets }));

  const chartDom = document.getElementById('channelChart');
  if (chartDom) {
    const chart = setChart('channelChart', chartDom);
    const option = createBarChartOption('渠道大类查验率 Top20', catNames, [
      { name: '起运港', data: domData, color: COLORS.domestic },
      { name: '目的港', data: foreignData, color: COLORS.foreign },
      { name: '综合', data: overallData, color: COLORS.overall }
    ]);
    chart.setOption(option, true);
  }

  // 美国子渠道
  const usSubs = Object.entries(aggByUsSub)
    .filter(([k, v]) => k && v.totalTickets >= 10)
    .sort((a, b) => b[1].overallRate - a[1].overallRate);
  const usNames = usSubs.map(([k]) => k);
  const usDom = usSubs.map(([k, v]) => ({ value: v.domesticRate, count: v.totalTickets }));
  const usForeign = usSubs.map(([k, v]) => ({ value: v.foreignRate, count: v.totalTickets }));
  const usOverall = usSubs.map(([k, v]) => ({ value: v.overallRate, count: v.totalTickets }));

  const usChartDom = document.getElementById('usSubChart');
  if (usChartDom) {
    const chart = setChart('usSubChart', usChartDom);
    const option = createBarChartOption('美国海运子渠道查验率', usNames, [
      { name: '起运港', data: usDom, color: COLORS.domestic },
      { name: '目的港', data: usForeign, color: COLORS.foreign },
      { name: '综合', data: usOverall, color: COLORS.overall }
    ]);
    chart.setOption(option, true);
  }

  // 素芸渠道表格
  renderTable('channelTable', Object.entries(aggByLogistic)
    .filter(([k, v]) => k && v.totalTickets >= 5)
    .sort((a, b) => b[1].overallRate - a[1].overallRate)
    .map(([k, v]) => ({
      name: k,
      total: v.totalTickets,
      domestic: v.domesticRate.toFixed(2) + '%',
      foreign: v.foreignRate.toFixed(2) + '%',
      overall: v.overallRate.toFixed(2) + '%',
      avgTime: v.domesticTickets + v.foreignTickets  // 复用 avgTime 字段存查验票数，renderTable 映射用第6列
    })),
    ['渠道名', '票数', '起运港', '目的港', '综合', '查验票数']
  );

  // 渠道大类周期对比表
  buildChannelComparisonTable();
}

// 渠道周期对比表
function buildChannelComparisonTable() {
  const tbody = document.getElementById('channelCompareBody');
  if (!tbody) return;

  if (allWeekKeys.length < 3) {
    tbody.innerHTML = '<tr><td colspan="6" style="text-align:center;color:#999;padding:16px">需要至少三周数据</td></tr>';
    return;
  }

  // 上上周 vs 上周（本周刚开始数据不全，不看本周）
  const prevWk = allWeekKeys[allWeekKeys.length - 2];
  const prev2Wk = allWeekKeys[allWeekKeys.length - 3];

  const rows = [];
  for (const [ch, wm] of Object.entries(aggByChannelWeekly)) {
    const cur = wm[prevWk];   // 上周
    const prev = wm[prev2Wk]; // 上上周
    if (!cur || cur.totalTickets < 10) continue;
    const curRate = cur.domesticTickets / cur.totalTickets * 100;  // 起运港即期查验率
    const prevRate = prev && prev.totalTickets >= 5 ? prev.domesticTickets / prev.totalTickets * 100 : null;
    const delta = prevRate !== null ? (curRate - prevRate).toFixed(1) : null;

    rows.push({ channel: ch, curTotal: cur.totalTickets, curRate, prevRate, delta });
  }

  rows.sort((a, b) => b.curRate - a.curRate);

  tbody.innerHTML = rows.map(r => {
    let deltaHtml = '';
    if (r.delta !== null) {
      const d = parseFloat(r.delta);
      if (d > 1) deltaHtml = `<span style="color:#d62929">↑${r.delta}%</span>`;
      else if (d < -1) deltaHtml = `<span style="color:#07c160">↓${Math.abs(d).toFixed(1)}%</span>`;
      else deltaHtml = `<span style="color:#999">→${r.delta}%</span>`;
    } else { deltaHtml = '<span style="color:#aaa">新</span>'; }
    return `<tr>
      <td>${r.channel}</td><td>${r.curTotal}</td>
      <td style="font-weight:600">${r.curRate.toFixed(1)}%</td>
      <td>${r.prevRate !== null ? r.prevRate.toFixed(1)+'%' : '-'}</td>
      <td>${deltaHtml}</td>
    </tr>`;
  }).join('');
}

// 代理分析
function updateAgentTab() {
  const agents = Object.entries(aggByAgent)
    .filter(([k, v]) => k && v.totalTickets >= 10)
    .sort((a, b) => b[1].overallRate - a[1].overallRate)
    .slice(0, 20);
  const names = agents.map(([k]) => k);
  const domData = agents.map(([k, v]) => ({ value: v.domesticRate, count: v.totalTickets }));
  const foreignData = agents.map(([k, v]) => ({ value: v.foreignRate, count: v.totalTickets }));
  const overallData = agents.map(([k, v]) => ({ value: v.overallRate, count: v.totalTickets }));

  const chartDom = document.getElementById('agentChart');
  if (chartDom) {
    const chart = setChart('agentChart', chartDom);
    const option = createBarChartOption('代理查验率 Top20', names, [
      { name: '起运港', data: domData, color: COLORS.domestic },
      { name: '目的港', data: foreignData, color: COLORS.foreign },
      { name: '综合', data: overallData, color: COLORS.overall }
    ]);
    chart.setOption(option, true);
  }

  // 代理绩效对比表（最近两周环比）
  buildAgentComparisonTable();

  // 代理 × 渠道热力图
  buildAgentChannelHeatmap();
}

// 代理绩效对比表：总货量 + 国内查验周期对比
function buildAgentComparisonTable() {
  const tbody = document.getElementById('agentCompareBody');
  if (!tbody) return;

  // 全局总票数（所有历史）
  const globalTotal = records.reduce((s, r) => s + r.ticketCount, 0);

  // 全局上周、上上周总票数
  const lastWk = allWeekKeys[allWeekKeys.length - 2];  // 上周
  const prevWk = allWeekKeys[allWeekKeys.length - 3]; // 上上周
  let globalLastWkTotal = 0, globalPrevWkTotal = 0;
  if (lastWk) {
    for (const wm of Object.values(aggByAgentWeekly)) {
      if (wm[lastWk]) globalLastWkTotal += wm[lastWk].totalTickets;
    }
  }
  if (prevWk) {
    for (const wm of Object.values(aggByAgentWeekly)) {
      if (wm[prevWk]) globalPrevWkTotal += wm[prevWk].totalTickets;
    }
  }

  const rows = [];
  for (const [agent, data] of Object.entries(aggByAgent)) {
    if (!agent || data.totalTickets < 5) continue;

    const totalTickets = data.totalTickets;
    const totalShare = globalTotal > 0 ? totalTickets / globalTotal * 100 : 0;

    // 全量查验率
    const domesticRate = data.domesticRate;
    const foreignRate = data.foreignRate;
    const overallRate = data.overallRate;

    // 上周 / 上上周数据
    const wm = aggByAgentWeekly[agent];
    let lastWkTotal = 0, lastWkDom = 0, lastWkRate = null;
    let prevWkTotal = 0, prevWkDom = 0, prevWkRate = null;
    let lastWkShare = 0, delta = null;

    if (wm && lastWk && wm[lastWk]) {
      lastWkTotal = wm[lastWk].totalTickets;
      lastWkDom = wm[lastWk].domesticTickets;
      lastWkRate = lastWkTotal > 0 ? lastWkDom / lastWkTotal * 100 : null;
      lastWkShare = globalLastWkTotal > 0 ? lastWkTotal / globalLastWkTotal * 100 : 0;
    }
    if (wm && prevWk && wm[prevWk]) {
      prevWkTotal = wm[prevWk].totalTickets;
      prevWkDom = wm[prevWk].domesticTickets;
      prevWkRate = prevWkTotal > 0 ? prevWkDom / prevWkTotal * 100 : null;
    }
    if (lastWkRate !== null && prevWkRate !== null) {
      delta = lastWkRate - prevWkRate;
    }

    // 分配建议：基于国内查验率（即期风险）+ 变化趋势
    let suggestion = '';
    if (lastWkRate !== null && lastWkRate > SETTINGS.highRisk) suggestion = '🚫 规避';
    else if (lastWkRate !== null && lastWkRate > SETTINGS.midRisk) suggestion = '⚠️ 减量';
    else if (domesticRate <= 1 && overallRate <= SETTINGS.midRisk) suggestion = '✅ 推荐';
    else suggestion = '👌 正常';

    rows.push({
      agent, totalTickets, totalShare,
      domesticRate, foreignRate, overallRate,
      lastWkTotal, lastWkShare, lastWkRate,
      prevWkRate, delta,
      suggestion
    });
  }

  // 排序：上周国内查验率从高到低（风险优先）
  rows.sort((a, b) => {
    const ra = a.lastWkRate !== null ? a.lastWkRate : a.domesticRate;
    const rb = b.lastWkRate !== null ? b.lastWkRate : b.domesticRate;
    return rb - ra;
  });

  tbody.innerHTML = rows.map(r => {
    let deltaHtml = '';
    if (r.delta !== null) {
      const d = r.delta;
      if (d > 1) deltaHtml = `<span style="color:#d62929">↑${d.toFixed(1)}%</span>`;
      else if (d < -1) deltaHtml = `<span style="color:#07c160">↓${Math.abs(d).toFixed(1)}%</span>`;
      else deltaHtml = `<span style="color:#999">→${d.toFixed(1)}%</span>`;
    } else {
      deltaHtml = '<span style="color:#aaa">—</span>';
    }
    const sugClass = r.suggestion.includes('规避') ? 'sug-avoid' : (r.suggestion.includes('减量') ? 'sug-warn' : (r.suggestion.includes('推荐') ? 'sug-good' : ''));
    return `<tr>
      <td>${r.agent}</td>
      <td>${r.totalTickets.toLocaleString()}<br><span style="color:#999;font-size:10px">${r.totalShare.toFixed(1)}%</span></td>
      <td>${r.lastWkTotal ? r.lastWkTotal.toLocaleString() : '—'}<br><span style="color:#999;font-size:10px">${r.lastWkShare ? r.lastWkShare.toFixed(1)+'%' : '—'}</span></td>
      <td style="color:#d62929;font-weight:500">${r.domesticRate.toFixed(1)}%</td>
      <td style="color:#1764e8;font-weight:500">${r.foreignRate.toFixed(1)}%</td>
      <td style="font-weight:600">${r.overallRate.toFixed(1)}%</td>
      <td>${r.prevWkRate !== null ? r.prevWkRate.toFixed(1)+'%' : '—'}</td>
      <td>${r.lastWkRate !== null ? r.lastWkRate.toFixed(1)+'%' : '—'}</td>
      <td>${deltaHtml}</td>
      <td class="${sugClass}">${r.suggestion}</td>
    </tr>`;
  }).join('');
}

// 代理明细表
function buildAgentChannelHeatmap() {
  // 取 Top 15 代理和 Top 15 渠道（热力图空间已加大）
  const topAgents = Object.entries(aggByAgent)
    .filter(([k, v]) => k && v.totalTickets >= 20)
    .sort((a, b) => b[1].overallRate - a[1].overallRate)
    .slice(0, 15)
    .map(([k]) => k);
  const topChannels = Object.entries(aggByChannel)
    .filter(([k, v]) => k && v.totalTickets >= 20)
    .sort((a, b) => b[1].overallRate - a[1].overallRate)
    .slice(0, 15)
    .map(([k]) => k);

  const heatData = [];
  for (let i = 0; i < topAgents.length; i++) {
    for (let j = 0; j < topChannels.length; j++) {
      const subset = records.filter(r => r.agent === topAgents[i] && r.channel === topChannels[j]);
      const total = subset.reduce((s, r) => s + r.ticketCount, 0);
      const inspected = subset.reduce((s, r) => s + (r.isInspected ? r.ticketCount : 0), 0);
      const rate = total > 0 ? (inspected / total * 100) : 0;
      heatData.push([j, i, rate, total]);
    }
  }

  const chartDom = document.getElementById('agentChannelHeatmap');
  if (chartDom) {
    const chart = setChart('agentChannelHeatmap', chartDom);
    const option = createHeatmapOption('代理×渠道综合查验率', topChannels, topAgents, heatData);
    chart.setOption(option, true);
  }
}

// 渠道下钻：选择渠道大类后，看该渠道上周各代理的表现
function updateDrilldownTab() {
  const select = document.getElementById('drilldownSelect');
  const resultDiv = document.getElementById('drilldownResult');
  if (!select || !resultDiv) return;

  const channel = select.value;
  if (!channel) {
    resultDiv.style.display = 'none';
    return;
  }

  resultDiv.style.display = 'block';
  const titleEl = document.getElementById('drilldownTitle');
  if (titleEl) titleEl.textContent = `${channel} — 上周渠道代理明细`;

  // 筛选该渠道下的记录
  const subset = records.filter(r => r.channel === channel);

  // 上周和上上周
  const lastWk = allWeekKeys[allWeekKeys.length - 2];  // 上周
  const prevWk = allWeekKeys[allWeekKeys.length - 3]; // 上上周

  // 按代理+周聚合（起运港 + 重量）
  const weeklyAgentMap = {};      // {weekKey: {agent: {total, domestic}}}
  const weeklyAgentWeightMap = {}; // {weekKey: {agent: totalWeight}}
  const weeklyTotalMap = {};       // {weekKey: totalTickets}
  const weeklyWeightMap = {};      // {weekKey: totalWeight}
  for (const r of subset) {
    if (!r.shipDate) continue;
    const wk = getWeekKey(r.shipDate);
    if (!wk) continue;
    const a = r.agent || '(未登记)';
    if (!weeklyAgentMap[wk]) weeklyAgentMap[wk] = {};
    if (!weeklyAgentMap[wk][a]) weeklyAgentMap[wk][a] = { total: 0, domestic: 0 };
    weeklyAgentMap[wk][a].total += r.ticketCount;
    if (r.isDomestic) weeklyAgentMap[wk][a].domestic += r.ticketCount;

    // 重量聚合
    if (!weeklyAgentWeightMap[wk]) weeklyAgentWeightMap[wk] = {};
    if (!weeklyAgentWeightMap[wk][a]) weeklyAgentWeightMap[wk][a] = 0;
    weeklyAgentWeightMap[wk][a] += r.weight * r.ticketCount;

    weeklyTotalMap[wk] = (weeklyTotalMap[wk] || 0) + r.ticketCount;
    weeklyWeightMap[wk] = (weeklyWeightMap[wk] || 0) + r.weight * r.ticketCount;
  }
  const sortedWeeks = Object.keys(weeklyAgentMap).sort();

  // 上周代理数据
  const lastWkTotal = weeklyTotalMap[lastWk] || 0;
  const lastWkWeight = weeklyWeightMap[lastWk] || 0;
  const prevWkTotal = weeklyTotalMap[prevWk] || 0;
  const lastWkAgents = weeklyAgentMap[lastWk] || {};
  const prevWkAgents = weeklyAgentMap[prevWk] || {};
  const lastWkAgentWeights = weeklyAgentWeightMap[lastWk] || {};

  const rows = [];
  for (const [agent, d] of Object.entries(lastWkAgents)) {
    if (d.total === 0) continue;
    const lastRate = d.domestic / d.total * 100;
    const lastShare = lastWkTotal > 0 ? d.total / lastWkTotal * 100 : 0;
    const lastWeight = lastWkAgentWeights[agent] || 0;
    const weightShare = lastWkWeight > 0 ? lastWeight / lastWkWeight * 100 : 0;

    // 上上周数据
    let prevRate = null, prevTotal = 0;
    if (prevWkAgents[agent]) {
      prevTotal = prevWkAgents[agent].total;
      if (prevTotal >= 5) {
        prevRate = prevWkAgents[agent].domestic / prevTotal * 100;
      }
    }

    let delta = null;
    if (prevRate !== null) delta = lastRate - prevRate;

    // 建议：基于上周国内查验率
    let suggestion = '';
    if (lastRate > SETTINGS.highRisk) suggestion = '🚫 规避';
    else if (lastRate > SETTINGS.midRisk) suggestion = '⚠️ 减量';
    else if (lastRate <= 1) suggestion = '✅ 推荐';
    else suggestion = '👌 正常';

    rows.push({ agent, total: d.total, share: lastShare, weightShare, lastRate, prevRate, delta, suggestion });
  }
  // 按上周国内查验率排序（高->低）
  rows.sort((a, b) => b.lastRate - a.lastRate);

  // 图表：上周代理起运港查验率对比
  const chartDom = document.getElementById('drilldownChart');
  if (chartDom) {
    const names = rows.map(r => r.agent);
    const domData = rows.map(r => ({ value: r.lastRate, count: r.total }));
    const chart = setChart('drilldownChart', chartDom);
    const option = createBarChartOption(`${channel} 上周代理起运港查验率`, names, [
      { name: '起运港', data: domData, color: COLORS.domestic }
    ]);
    chart.setOption(option, true);
  }

  // 趋势图：当前渠道下各代理的起运港查验率周趋势
  // 美国海运代理多时取 Top6（按上周票数），其余展示全部
  const trendDom = document.getElementById('drilldownTrendChart');
  const trendTitleEl = document.getElementById('drilldownTrendTitle');
  if (trendDom && sortedWeeks.length >= 2) {
    // 确定要展示的代理列表（按上周票数降序）
    const allAgents = [...rows].sort((a, b) => b.total - a.total).map(r => r.agent);
    const isUsSea = channel === '美国海运';
    let showAgents = allAgents;
    if (isUsSea && allAgents.length > 6) {
      showAgents = allAgents.slice(0, 6);
    }
    const showTopN = showAgents.length < allAgents.length;

    if (trendTitleEl) {
      trendTitleEl.textContent = `📈 ${channel}代理起运港查验率周趋势${showTopN ? '（Top' + showAgents.length + '）' : ''}`;
    }

    const palette = ['#d62929', '#e6a23c', '#1764e8', '#07c160', '#8b5cf6', '#ec4899', '#f59e0b', '#10b981', '#6366f1', '#ef4444', '#06b6d4', '#84cc16'];
    const series = showAgents.map((agent, idx) => {
      const data = sortedWeeks.map(wk => {
        const d = weeklyAgentMap[wk] && weeklyAgentMap[wk][agent];
        if (!d || d.total === 0) return null;
        return parseFloat((d.domestic / d.total * 100).toFixed(1));
      });
      return { name: agent, data, lineStyle: { width: 2 }, itemStyle: { color: palette[idx % palette.length] } };
    });
    const chart = setChart('drilldownTrendChart', trendDom);
    const option = createLineChartOption(`${channel}代理起运港查验率周趋势${showTopN ? '（Top' + showAgents.length + '）' : ''}`, sortedWeeks, series);
    option.color = palette.slice(0, series.length);
    option.legend = { data: series.map(s => s.name), bottom: 0, textStyle: { fontSize: 10 } };
    option.grid = { left: '3%', right: '4%', bottom: '20%', top: '15%', containLabel: true };
    chart.setOption(option, true);
  }

  // 表格：上周渠道代理明细
  const tbody = document.getElementById('drilldownTableBody');
  if (tbody) {
    tbody.innerHTML = rows.map(r => {
      const sugClass = r.suggestion.includes('规避') ? 'sug-avoid' : (r.suggestion.includes('减量') ? 'sug-warn' : (r.suggestion.includes('推荐') ? 'sug-good' : ''));
      let deltaHtml = '';
      if (r.delta !== null) {
        const d = r.delta;
        if (d > 1) deltaHtml = `<span style="color:#d62929">↑${d.toFixed(1)}%</span>`;
        else if (d < -1) deltaHtml = `<span style="color:#07c160">↓${Math.abs(d).toFixed(1)}%</span>`;
        else deltaHtml = `<span style="color:#999">→${d.toFixed(1)}%</span>`;
      } else {
        deltaHtml = '<span style="color:#aaa">—</span>';
      }
      return `<tr>
        <td>${r.agent}</td>
        <td>${r.total.toLocaleString()}</td>
        <td>${r.share.toFixed(1)}%</td>
        <td>${r.weightShare.toFixed(1)}%</td>
        <td style="color:#d62929;font-weight:500">${r.lastRate.toFixed(1)}%</td>
        <td>${r.prevRate !== null ? r.prevRate.toFixed(1)+'%' : '—'}</td>
        <td>${deltaHtml}</td>
        <td class="${sugClass}">${r.suggestion}</td>
      </tr>`;
    }).join('');
  }
}

// 预警与明细
function updateAlertTab() {
  const alerts = generateAlerts();
  const tbody = document.getElementById('alertTableBody');
  if (!tbody) return;

  if (alerts.length === 0) {
    tbody.innerHTML = '<tr><td colspan="5" style="text-align:center;color:#999;padding:20px">暂无预警</td></tr>';
  } else {
    tbody.innerHTML = alerts.map(a => {
      const typeClass = a.type === 'high' ? 'alert-high' : (a.type === 'mid' ? 'alert-mid' : 'alert-trend');
      const typeText = a.type === 'high' ? '🔴 高风险' : (a.type === 'mid' ? '🟡 中风险' : '⚠️ 趋势恶化');
      return `<tr class="${typeClass}">
        <td>${typeText}</td>
        <td>${a.category}</td>
        <td>${a.name}</td>
        <td>${a.rate.toFixed(2)}%</td>
        <td>${a.total ? a.total.toLocaleString() + '票' : (a.change || '')}</td>
      </tr>`;
    }).join('');
  }

  // 产品属性
  renderTable('productTable', Object.entries(aggByProduct)
    .filter(([k, v]) => k && v.totalTickets >= 5)
    .sort((a, b) => b[1].overallRate - a[1].overallRate)
    .map(([k, v]) => ({
      name: k,
      total: v.totalTickets,
      domestic: v.domesticRate.toFixed(2) + '%',
      foreign: v.foreignRate.toFixed(2) + '%',
      overall: v.overallRate.toFixed(2) + '%'
    })),
    ['产品属性', '票数', '起运港', '目的港', '综合']
  );

  // 客户
  renderTable('customerTable', Object.entries(aggByCustomer)
    .filter(([k, v]) => k && v.totalTickets >= 5)
    .sort((a, b) => b[1].overallRate - a[1].overallRate)
    .map(([k, v]) => ({
      name: k,
      total: v.totalTickets,
      domestic: v.domesticRate.toFixed(2) + '%',
      foreign: v.foreignRate.toFixed(2) + '%',
      overall: v.overallRate.toFixed(2) + '%'
    })),
    ['客户', '票数', '起运港', '目的港', '综合']
  );
}

// 通用表格渲染
function renderTable(tableId, rows, headers) {
  const table = document.getElementById(tableId);
  if (!table) return;

  let html = '<thead><tr>' + headers.map(h => `<th>${h}</th>`).join('') + '</tr></thead><tbody>';
  if (rows.length === 0) {
    html += `<tr><td colspan="${headers.length}" style="text-align:center;color:#999;padding:20px">暂无数据</td></tr>`;
  } else {
    html += rows.map(r => `<tr>${headers.map((h, i) => {
      const key = ['name', 'total', 'domestic', 'foreign', 'overall', 'avgTime'][i] || 'name';
      return `<td>${r[key] !== undefined ? r[key] : ''}</td>`;
    }).join('')}</tr>`).join('');
  }
  html += '</tbody>';
  table.innerHTML = html;
}

// 通用表格渲染（自定义key映射）
function renderTableCustom(tableId, rows, headers, keys) {
  const table = document.getElementById(tableId);
  if (!table) return;

  let html = '<thead><tr>' + headers.map(h => `<th>${h}</th>`).join('') + '</tr></thead><tbody>';
  if (rows.length === 0) {
    html += `<tr><td colspan="${headers.length}" style="text-align:center;color:#999;padding:20px">暂无数据</td></tr>`;
  } else {
    html += rows.map(r => `<tr>${keys.map(k => `<td>${r[k] !== undefined ? r[k] : ''}</td>`).join('')}</tr>`).join('');
  }
  html += '</tbody>';
  table.innerHTML = html;
}

// ===================== 事件绑定 =====================

function initEvents() {
  // 文件上传
  const dropZone = document.getElementById('dropZone');
  const fileInput = document.getElementById('fileInput');
  if (!dropZone || !fileInput) return;

  dropZone.addEventListener('click', () => fileInput.click());
  dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('dragover'); });
  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
  dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    const file = e.dataTransfer.files[0];
    if (file && file.name.endsWith('.xlsx')) loadAndProcess(file);
  });
  fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) loadAndProcess(file);
  });

  // 粒度切换
  document.querySelectorAll('.granularity-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.granularity-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      currentGranularity = btn.dataset.granularity;
      updateDashboardTab();
    });
  });

  // 阈值设置
  const highInput = document.getElementById('highRiskInput');
  const midInput = document.getElementById('midRiskInput');
  if (highInput) {
    highInput.value = SETTINGS.highRisk;
    highInput.addEventListener('change', () => { SETTINGS.highRisk = parseFloat(highInput.value) || 5; updateAlertTab(); });
  }
  if (midInput) {
    midInput.value = SETTINGS.midRisk;
    midInput.addEventListener('change', () => { SETTINGS.midRisk = parseFloat(midInput.value) || 3; updateAlertTab(); });
  }

  // 搜索过滤
  const searchInput = document.getElementById('detailSearch');
  if (searchInput) {
    searchInput.addEventListener('input', () => filterDetailTable(searchInput.value));
  }
}

function filterDetailTable(keyword) {
  const kw = keyword.trim().toLowerCase();
  const tbody = document.getElementById('detailTableBody');
  if (!tbody) return;
  const rows = records.filter(r => {
    if (!kw) return true;
    const text = `${r.tickets.join(' ')} ${r.agent} ${r.channel} ${r.logisticChannel} ${r.country} ${r.type} ${r.customer}`.toLowerCase();
    return text.includes(kw);
  });

  // 只展示前 200 条避免卡顿
  const displayRows = rows.slice(0, 200);
  tbody.innerHTML = displayRows.map(r => {
    const status = r.isInspected ? (r.isDomestic && r.isForeign ? '双港查验' : (r.isDomestic ? '起运港查验' : '目的港查验')) : '正常';
    const statusClass = r.isInspected ? 'status-inspected' : 'status-normal';
    return `<tr>
      <td>${r.tickets.slice(0, 2).join('<br>')}${r.tickets.length > 2 ? '...' : ''}</td>
      <td>${formatDate(r.shipDate)}</td>
      <td>${r.country}</td>
      <td>${r.type}</td>
      <td>${r.channel}</td>
      <td>${r.agent}</td>
      <td>${r.logisticChannel}</td>
      <td class="${statusClass}">${status}</td>
    </tr>`;
  }).join('');

  if (rows.length > 200) {
    tbody.innerHTML += `<tr><td colspan="8" style="text-align:center;color:#999;padding:10px">还有 ${rows.length - 200} 条，请缩小搜索范围</td></tr>`;
  }
}

// ===================== Tab 切换 =====================

function switchTab(tabName) {
  document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
  document.querySelector(`.tab-btn[data-tab="${tabName}"]`)?.classList.add('active');
  document.querySelectorAll('.tab-pane').forEach(pane => pane.classList.remove('active'));
  document.getElementById('pane-' + tabName)?.classList.add('active');
  currentView = tabName;
  // 先 resize 恢复尺寸（切换回来时容器可能刚从 display:none 恢复）
  setTimeout(() => resizeCharts(), 50);
  // 重新渲染图表
  if (tabName === 'dashboard') updateDashboardTab();
  if (tabName === 'channel') updateChannelTab();
  if (tabName === 'agent') updateAgentTab();
  if (tabName === 'drilldown') updateDrilldownTab();
  if (tabName === 'alert') {
    updateAlertTab();
    // 同时刷新明细表格
    const searchInput = document.getElementById('detailSearch');
    filterDetailTable(searchInput ? searchInput.value : '');
  }
  // 再次 resize 确保渲染后尺寸正确
  setTimeout(resizeCharts, 150);
}

// ===================== 初始化 =====================
document.addEventListener('DOMContentLoaded', () => {
  initEvents();
  // 默认显示上传页面，等待用户上传数据
});
