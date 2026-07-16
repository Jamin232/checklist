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
  matureDays: 30,  // 成熟期窗口 30 天
  minSample: 20    // 统计最小样本数（低于此样本不预警）
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

function isMature(date, maxDate, days) {
  if (!date || !maxDate) return false;
  const d = date instanceof Date ? date : parseDate(date);
  const m = maxDate instanceof Date ? maxDate : parseDate(maxDate);
  if (!d || !m) return false;
  const diff = (m - d) / (1000 * 60 * 60 * 24);
  return diff >= days;
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
  const { filterMature = false, timeGranularity = null } = opts;

  for (const rec of records) {
    if (!rec.shipDate) continue;
    if (filterMature && !isMature(rec.shipDate, maxShipDate, SETTINGS.matureDays)) continue;

    const key = getKey(rec);
    if (!key) continue;
    if (!map[key]) map[key] = createAgg();
    addToAgg(map[key], rec);

    if (timeGranularity) {
      const tKey = timeGranularity === 'week' ? getWeekKey(rec.shipDate) : formatDateYM(rec.shipDate);
      if (!map[key].timeSeries[tKey]) {
        map[key].timeSeries[tKey] = { totalTickets: 0, inspectedTickets: 0 };
      }
      map[key].timeSeries[tKey].totalTickets += rec.ticketCount;
      if (rec.isInspected) map[key].timeSeries[tKey].inspectedTickets += rec.ticketCount;
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

// ===================== 预警规则 =====================

function generateAlerts() {
  const alerts = [];

  // 代理预警（基于成熟数据）
  const matureAgents = aggregateBy(records, r => r.agent, { filterMature: true });
  for (const [agent, data] of Object.entries(matureAgents)) {
    if (data.totalTickets < SETTINGS.minSample) continue;
    if (data.overallRate > SETTINGS.highRisk) {
      alerts.push({ type: 'high', category: '代理', name: agent, rate: data.overallRate, total: data.totalTickets });
    } else if (data.overallRate > SETTINGS.midRisk) {
      alerts.push({ type: 'mid', category: '代理', name: agent, rate: data.overallRate, total: data.totalTickets });
    }
  }

  // 渠道大类预警
  const matureChannels = aggregateBy(records, r => r.channel, { filterMature: true });
  for (const [channel, data] of Object.entries(matureChannels)) {
    if (data.totalTickets < SETTINGS.minSample) continue;
    if (data.overallRate > SETTINGS.highRisk) {
      alerts.push({ type: 'high', category: '渠道大类', name: channel, rate: data.overallRate, total: data.totalTickets });
    } else if (data.overallRate > SETTINGS.midRisk) {
      alerts.push({ type: 'mid', category: '渠道大类', name: channel, rate: data.overallRate, total: data.totalTickets });
    }
  }

  // 素芸渠道预警
  const matureLogistics = aggregateBy(records, r => r.logisticChannel, { filterMature: true });
  for (const [channel, data] of Object.entries(matureLogistics)) {
    if (data.totalTickets < SETTINGS.minSample) continue;
    if (data.overallRate > SETTINGS.highRisk) {
      alerts.push({ type: 'high', category: '素芸渠道', name: channel, rate: data.overallRate, total: data.totalTickets });
    } else if (data.overallRate > SETTINGS.midRisk) {
      alerts.push({ type: 'mid', category: '素芸渠道', name: channel, rate: data.overallRate, total: data.totalTickets });
    }
  }

  // 趋势恶化（环比）- 基于时间聚合的成熟数据
  // 按周计算
  const weeklyAgg = aggregateByTime(records.filter(r => isMature(r.shipDate, maxShipDate, SETTINGS.matureDays)), 'week');
  const weeks = Object.keys(weeklyAgg).sort();
  if (weeks.length >= 2) {
    const last = weeklyAgg[weeks[weeks.length - 1]];
    const prev = weeklyAgg[weeks[weeks.length - 2]];
    if (prev.overallRate > 0 && last.overallRate > 0) {
      const change = ((last.overallRate - prev.overallRate) / prev.overallRate * 100);
      if (change > 50) {
        alerts.push({ type: 'trend', category: '总体', name: '综合查验率周环比', rate: last.overallRate, change: change.toFixed(1), prevRate: prev.overallRate });
      }
    }
  }

  // 排序：高风险在前，然后中风险，然后趋势
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
    grid: { left: '15%', right: '10%', top: '15%', bottom: '15%' },
    xAxis: { type: 'category', data: xLabels, splitArea: { show: true }, axisLabel: { rotate: 45, fontSize: 10 } },
    yAxis: { type: 'category', data: yLabels, splitArea: { show: true }, axisLabel: { fontSize: 10 } },
    visualMap: { min: 0, max: 10, calculable: true, orient: 'horizontal', left: 'center', bottom: '0%', inRange: { color: ['#e8f7ef', '#fff3cd', '#d62929'] } },
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
    aggByChannel = aggregateBy(records, r => r.channel);
    aggByAgent = aggregateBy(records, r => r.agent);
    aggByUsSub = aggregateBy(records, r => r.usSubChannel);
    aggByLogistic = aggregateBy(records, r => r.logisticChannel);
    aggByCustomer = aggregateBy(records, r => r.customer);
    aggByProduct = aggregateBy(records, r => r.productAttr);

    // 隐藏上传面板，显示主内容
    const uploadPanel = document.getElementById('uploadPanel');
    const mainContent = document.getElementById('mainContent');
    if (uploadPanel) uploadPanel.style.display = 'none';
    if (mainContent) mainContent.style.display = 'block';

    // 更新 UI
    updateAllTabs();
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
  updateAlertTab();
}

function setChart(domId, chart) {
  if (chartInstances[domId]) chartInstances[domId].dispose();
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
  document.getElementById('kpiOverall').textContent = (total > 0 ? (inspected / total * 100).toFixed(2) : '0.00') + '%';

  // 趋势图
  const timeData = aggByTime[currentGranularity];
  const keys = Object.keys(timeData).sort();
  const domesticData = keys.map(k => timeData[k].domesticRate);
  const foreignData = keys.map(k => timeData[k].foreignRate);
  const overallData = keys.map(k => timeData[k].overallRate);

  // 成熟度标记（最后N个标记为虚线）
  const matureThreshold = SETTINGS.matureDays;
  const matureKeys = keys.map(k => {
    // 简单判断：最近2个周期标记为未成熟（基于周/月粒度）
    return true; // 简化，实际按日期判断
  });

  const chartDom = document.getElementById('trendChart');
  if (!chartDom) return;
  const chart = setChart('trendChart', echarts.init(chartDom));
  const option = createLineChartOption(
    `查验率趋势 (${currentGranularity === 'week' ? '周度' : '月度'})`,
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
    const chart = setChart('channelChart', echarts.init(chartDom));
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
    const chart = setChart('usSubChart', echarts.init(usChartDom));
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
      avgTime: v.avgInspectTime > 0 ? v.avgInspectTime.toFixed(1) + '天' : '-'
    })),
    ['渠道名', '票数', '起运港', '目的港', '综合', '平均时效']
  );
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
    const chart = setChart('agentChart', echarts.init(chartDom));
    const option = createBarChartOption('代理查验率 Top20', names, [
      { name: '起运港', data: domData, color: COLORS.domestic },
      { name: '目的港', data: foreignData, color: COLORS.foreign },
      { name: '综合', data: overallData, color: COLORS.overall }
    ]);
    chart.setOption(option, true);
  }

  // 代理表格
  renderTable('agentTable', Object.entries(aggByAgent)
    .filter(([k, v]) => k && v.totalTickets >= 5)
    .sort((a, b) => b[1].overallRate - a[1].overallRate)
    .map(([k, v]) => ({
      name: k,
      total: v.totalTickets,
      domestic: v.domesticRate.toFixed(2) + '%',
      foreign: v.foreignRate.toFixed(2) + '%',
      overall: v.overallRate.toFixed(2) + '%',
      avgTime: v.avgInspectTime > 0 ? v.avgInspectTime.toFixed(1) + '天' : '-'
    })),
    ['代理名', '票数', '起运港', '目的港', '综合', '平均时效']
  );

  // 代理 × 渠道热力图
  buildAgentChannelHeatmap();
}

function buildAgentChannelHeatmap() {
  // 取 Top 10 代理和 Top 10 渠道
  const topAgents = Object.entries(aggByAgent)
    .filter(([k, v]) => k && v.totalTickets >= 20)
    .sort((a, b) => b[1].overallRate - a[1].overallRate)
    .slice(0, 10)
    .map(([k]) => k);
  const topChannels = Object.entries(aggByChannel)
    .filter(([k, v]) => k && v.totalTickets >= 20)
    .sort((a, b) => b[1].overallRate - a[1].overallRate)
    .slice(0, 10)
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
    const chart = setChart('agentChannelHeatmap', echarts.init(chartDom));
    const option = createHeatmapOption('代理×渠道综合查验率', topChannels, topAgents, heatData);
    chart.setOption(option, true);
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
        <td>${a.total || a.change || ''}</td>
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
  // 重新渲染图表（因为隐藏状态下 echarts 尺寸不对）
  if (tabName === 'dashboard') updateDashboardTab();
  if (tabName === 'channel') updateChannelTab();
  if (tabName === 'agent') updateAgentTab();
  if (tabName === 'alert') {
    updateAlertTab();
    // 同时刷新明细表格
    const searchInput = document.getElementById('detailSearch');
    filterDetailTable(searchInput ? searchInput.value : '');
  }
  // 延迟 resize 确保 DOM 已更新
  setTimeout(resizeCharts, 100);
}

// ===================== 初始化 =====================
document.addEventListener('DOMContentLoaded', () => {
  initEvents();
  // 默认显示上传页面，等待用户上传数据
});
