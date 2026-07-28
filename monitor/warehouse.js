// ============================================================
// warehouse.js — 仓配发运监控（独立数据源：仓配每日报表 xlsx）
// 与产品跟踪表(日度/查验)完全解耦：有自己的上传区、自己的解析口径、自己的看板。
// 图表通过 chayan.js 的全局 setChart 创建，自动纳入 resizeCharts 统一缩放。
// ============================================================
(function () {
  const FIELDS = {
    date: "处理日期", weekday: "星期", group: "事业群",
    receipt: "当日来货签收数量", prev_inbound: "上日未入 库数量", qc_done: "当日待入库数量（质检完成数量）",
    inbound_plan: "当日应入库数量", inbound_actual: "当日实际入库数量", inbound_unfinish: "当日未入库数量",
    prev_transfer: "上日未完成调拨数量", transfer_plan: "当日调拨计划数量", transfer_total: "调拨合计数量",
    boxed: "当日完成装箱数量", transfer_unfinish: "当日调拨未完成数量",
    prev_unshipped: "上日未交运总调拨数量", ship_out: "当日交运数量", unshipped: "当日未交运数量",
    cutoff_sea: "当周截单未完成装箱数量-美国海运（周六出数据）",
    cutoff_rail: "当周截单未完成装箱数量-铁陆运（周六出数据）",
    cutoff_air: "当周截单未完成装箱数量-空运+快递（周六、周一出数据）",
    boxed_addr: "截单前完成装箱未交运数量（未有地址标+等拼）",
    boxed_other: "截单前完成装箱未交运数量（其他原因）",
    boxed_legacy: "截单前完成装箱未交运数量", remark: "备注"
  };
  let appData = null;

  const $ = id => document.getElementById(id);
  const fmtNum = n => (n === null || n === undefined || isNaN(n)) ? '—' : Number(n).toLocaleString('zh-CN', { maximumFractionDigits: 0 });
  const fmtPct = (a, b) => b ? ((a / b) * 100).toFixed(1) + '%' : '—';
  const formatDate = d => {
    if (!d) return null;
    if (typeof d === 'string') return d.split(' ')[0];
    if (typeof d === 'number' && window.XLSX && window.XLSX.SSF) {
      const pd = window.XLSX.SSF.parse_date_code(d);
      if (pd) return `${pd.y}-${String(pd.m).padStart(2, '0')}-${String(pd.d).padStart(2, '0')}`;
    }
    const dt = new Date(d);
    if (isNaN(dt)) return String(d);
    return `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, '0')}-${String(dt.getDate()).padStart(2, '0')}`;
  };
  const num = v => {
    if (v === null || v === undefined || v === '') return null;
    const n = Number(v); return isNaN(n) ? null : (n === Math.floor(n) ? Math.floor(n) : n);
  };

  // ---------------- 文件加载 ----------------
  function loadFile(file) {
    const dz = $('whDropZone');
    if (dz) dz.textContent = '正在解析：' + file.name;
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: 'array' });
        appData = processWorkbook(wb);
        if (!Object.keys(appData).length) throw new Error('未识别到月份数据（请确认 sheet 含 处理日期/事业群 等列）');
        populateMonths();
        showDashboard();
        render();
      } catch (err) {
        alert('解析失败：' + err.message);
        if (dz) dz.textContent = '点击或拖拽「仓配每日报表」xlsx 到此处导入';
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function processWorkbook(wb) {
    const out = {};
    wb.SheetNames.forEach(name => {
      if (name === '取值说明' || name === '模板') return;
      const ws = wb.Sheets[name];
      const json = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
      if (!json.length) return;
      const parsed = parseSheet(json, name);
      if (parsed && parsed.daily.length) out[name] = parsed;
    });
    return out;
  }

  function parseSheet(rows, sheetName) {
    const header = rows[0];
    const idx = {};
    header.forEach((h, i) => {
      if (!h) return;
      for (const k in FIELDS) if (h === FIELDS[k]) idx[k] = i;
    });
    const hasSplit = idx.boxed_addr !== undefined && idx.boxed_other !== undefined;
    const rawDays = {};
    let lastDate = null;
    const getVal = (row, k) => { const i = idx[k]; return (i !== undefined && i < row.length) ? row[i] : null; };

    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      const d = getVal(row, 'date');
      if (d) lastDate = formatDate(d);
      const g = getVal(row, 'group');
      if (!g || !lastDate) continue;
      if (!rawDays[lastDate]) rawDays[lastDate] = { date: lastDate, totals: [], details: [] };

      const rec = {
        receipt: num(getVal(row, 'receipt')), prev_inbound: num(getVal(row, 'prev_inbound')), qc_done: num(getVal(row, 'qc_done')),
        inbound_plan: num(getVal(row, 'inbound_plan')), inbound_actual: num(getVal(row, 'inbound_actual')), inbound_unfinish: num(getVal(row, 'inbound_unfinish')),
        prev_transfer: num(getVal(row, 'prev_transfer')), transfer_plan: num(getVal(row, 'transfer_plan')), transfer_total: num(getVal(row, 'transfer_total')),
        boxed: num(getVal(row, 'boxed')), transfer_unfinish: num(getVal(row, 'transfer_unfinish')),
        prev_unshipped: num(getVal(row, 'prev_unshipped')), ship_out: num(getVal(row, 'ship_out')), unshipped: num(getVal(row, 'unshipped')),
        cutoff_sea: num(getVal(row, 'cutoff_sea')), cutoff_rail: num(getVal(row, 'cutoff_rail')), cutoff_air: num(getVal(row, 'cutoff_air'))
      };
      if (hasSplit) {
        rec.boxed_addr = num(getVal(row, 'boxed_addr')); rec.boxed_other = num(getVal(row, 'boxed_other'));
        rec.boxed_unshipped = (rec.boxed_addr || 0) + (rec.boxed_other || 0);
      } else {
        rec.boxed_unshipped = num(getVal(row, 'boxed_legacy')); rec.boxed_addr = null; rec.boxed_other = null;
      }

      if (g.endsWith('合计：') || g.endsWith('合计')) {
        rawDays[lastDate].totals.push({ name: g, ...rec });
      } else {
        rawDays[lastDate].details.push({ group: g, ...rec, remark: getVal(row, 'remark') });
      }
    }

    const daily = Object.values(rawDays).sort((a, b) => a.date.localeCompare(b.date));
    daily.forEach(day => {
      const names = day.totals.map(t => t.name);
      if (names.includes('两仓合计：')) {
        day.combined = { ...day.totals.find(t => t.name === '两仓合计：') };
      } else {
        day.combined = sumRecords(day.totals.filter(t => !t.name.includes('两仓')));
        day.combined.name = '合计';
      }
      day.warehouses = day.totals.filter(t => !t.name.includes('两仓'));
    });
    return { daily, sheetName };
  }

  function sumRecords(list) {
    const keys = ['receipt', 'prev_inbound', 'qc_done', 'inbound_plan', 'inbound_actual', 'inbound_unfinish', 'prev_transfer', 'transfer_plan', 'transfer_total', 'boxed', 'transfer_unfinish', 'prev_unshipped', 'ship_out', 'unshipped', 'cutoff_sea', 'cutoff_rail', 'cutoff_air', 'boxed_unshipped', 'boxed_addr', 'boxed_other'];
    const out = {};
    keys.forEach(k => out[k] = list.reduce((s, r) => (s + (r[k] || 0)), 0));
    return out;
  }

  // ---------------- 选择器 ----------------
  function populateMonths() {
    const sel = $('whMonthSelect'); if (!sel) return;
    sel.innerHTML = '';
    const names = Object.keys(appData).sort();
    names.forEach(n => { const o = document.createElement('option'); o.value = n; o.textContent = n; sel.appendChild(o); });
    sel.value = names[names.length - 1];
    sel.onchange = () => { populateWarehouses(); render(); };
    populateWarehouses();
  }

  function populateWarehouses() {
    const m = $('whMonthSelect').value;
    const wh = $('whWarehouseSelect'); if (!wh || !m) return;
    const current = wh.value;
    wh.innerHTML = '<option value="combined">两仓合计</option>';
    const seen = new Set();
    appData[m].daily.forEach(d => d.warehouses.forEach(w => seen.add(w.name)));
    seen.forEach(name => { const o = document.createElement('option'); o.value = name; o.textContent = name; wh.appendChild(o); });
    wh.value = seen.has(current) ? current : 'combined';
    wh.onchange = render;
  }

  // ---------------- 计算辅助 ----------------
  function getSeries(month, warehouseKey) {
    return month.daily.map(d => {
      let r;
      if (warehouseKey === 'combined') r = d.combined;
      else r = d.warehouses.find(x => x.name === warehouseKey) || d.combined;
      return { ...r, date: d.date };
    });
  }
  function findLastReported(daily) {
    for (let i = daily.length - 1; i >= 0; i--) {
      if (daily[i].details.some(r => r.ship_out !== null || r.unshipped !== null)) return i;
    }
    return daily.length - 1;
  }
  function avg(arr) { const a = arr.filter(v => v !== null && !isNaN(v)); return a.length ? a.reduce((s, v) => s + v, 0) / a.length : null; }
  function slope(values) {
    const pts = values.map((v, i) => ({ x: i, y: v })).filter(p => p.y !== null);
    if (pts.length < 3) return null;
    const n = pts.length, mx = avg(pts.map(p => p.x)), my = avg(pts.map(p => p.y));
    const num_ = pts.reduce((s, p) => s + (p.x - mx) * (p.y - my), 0);
    const den = pts.reduce((s, p) => s + (p.x - mx) ** 2, 0);
    return den ? num_ / den : null;
  }
  function trendInfo(arr, key) {
    const vals = arr.map(r => r[key]);
    const last3 = vals.slice(-3), prev3 = vals.slice(-6, -3);
    const a = avg(last3), b = avg(prev3);
    let dir = 'flat', pct = 0;
    if (a !== null && b !== null && b !== 0) {
      pct = ((a - b) / Math.abs(b)) * 100;
      dir = pct > 5 ? 'up' : (pct < -5 ? 'down' : 'flat');
    }
    const sl = slope(vals);
    return { dir, pct: Math.abs(pct).toFixed(1), sl, latest: vals[vals.length - 1], avg7: avg(vals.slice(-7)), max: Math.max(...vals.filter(v => v !== null)) };
  }

  // ---------------- 渲染 ----------------
  function render() {
    if (!appData) { showUpload(); return; }
    showDashboard();
    const mName = $('whMonthSelect').value;
    if (!mName || !appData[mName]) return;
    const month = appData[mName];
    const whKey = $('whWarehouseSelect').value || 'combined';
    const fullSeries = getSeries(month, whKey);
    const lastIdx = findLastReported(month.daily);
    const series = fullSeries.slice(0, lastIdx + 1);
    const latest = series[series.length - 1];
    const prev = series.length > 1 ? series[series.length - 2] : latest;
    const latestDay = month.daily[lastIdx];
    const rawLastDay = month.daily[month.daily.length - 1];
    const whName = whKey === 'combined' ? '两仓合计' : whKey;
    const hasPartial = lastIdx < month.daily.length - 1;

    const us = trendInfo(series, 'unshipped');
    const ib = trendInfo(series, 'inbound_unfinish');
    const tf = trendInfo(series, 'transfer_unfinish');
    const so = trendInfo(series, 'ship_out');

    const cutoffDays = month.daily.filter(d => d.combined.cutoff_sea > 0 || d.combined.cutoff_rail > 0 || d.combined.cutoff_air > 0);
    const latestCutoff = cutoffDays.length ? cutoffDays[cutoffDays.length - 1] : null;
    const whNames = latestDay.warehouses.map(w => w.name);
    const unitRows = latestDay.details.slice().sort((a, b) => (b.unshipped || 0) - (a.unshipped || 0));
    const partialNote = hasPartial ? `（注：${rawLastDay.date} 为今日，物流侧数据尚未填报；当前分析截至最新已填报日 ${latestDay.date}）` : '';

    $('conclusionTag').textContent = mName + ' · ' + whName + (hasPartial ? ' · 截至' + latestDay.date : '');
    let html = `<p class="small" style="color:#c2410c">${partialNote}</p>`;
    html += `<p><b>【物流交运】</b>截至 ${latestDay.date}，${whName} 当日未交运 ${fmtNum(latest.unshipped)} 件，较昨日 ${fmtNum(prev.unshipped)} 件 ${us.dir === 'up' ? '上升' : '下降'}。`;
    if (us.dir === 'up') html += `近3日均值较此前3日上升 ${us.pct}%，<b>积压正在扩大</b>，需排查物流运力/地址标/等拼原因。`;
    else if (us.dir === 'down') html += `近3日均值较此前3日下降 ${us.pct}%，<b>积压持续消化</b>。`;
    else html += `近3日均值与此前3日基本持平，<b>积压维持高位震荡</b>。`;
    html += `</p>`;

    html += `<p><b>【入库环节】</b>当日未入库 ${fmtNum(latest.inbound_unfinish)} 件。`;
    if (ib.dir === 'up') html += `近3日入库未完成均值上升 ${ib.pct}%，<b>收货/质检/上架存在卡点</b>。`;
    else if (ib.dir === 'down') html += `近3日入库未完成均值下降 ${ib.pct}%，入库节奏改善。`;
    else html += `入库未完成量相对平稳。`;
    html += `</p>`;

    const tfLatest = tf.latest || 0;
    html += `<p><b>【仓库备货】</b>${tfLatest < 0 ? `当日提前完成调拨 ${fmtNum(-tfLatest)} 件` : `当日调拨未完成 ${fmtNum(tfLatest)} 件`}，当日完成装箱 ${fmtNum(latest.boxed)} 件。`;
    if (tfLatest < 0) html += `当前已提前消化次日计划，仓库备货效率良好。`;
    else if (tf.dir === 'up') html += `调拨未完成近3日上升 ${tf.pct}%，<b>拣货/包装/装箱跟不上调拨计划</b>。`;
    else if (tf.dir === 'down') html += `调拨未完成近3日下降 ${tf.pct}%，仓库备货效率提升。`;
    else html += `仓库备货节奏与前几日基本持平。`;
    html += `</p>`;

    html += `<p><b>【截单/异常】</b>`;
    if (latestCutoff) {
      const c = latestCutoff.combined;
      html += `最近有截单数据的是 ${latestCutoff.date}：美国海运 ${fmtNum(c.cutoff_sea)} / 铁陆运 ${fmtNum(c.cutoff_rail)} / 空运+快递 ${fmtNum(c.cutoff_air)}。`;
      if ((c.cutoff_sea + c.cutoff_rail + c.cutoff_air) > 0) html += `<b>存在截单未完成装箱，需确认是否漏截单。</b>`;
      else html += '截单前装箱已全部完成。';
    } else {
      html += '本月暂无截单未完成装箱数据（通常仅周六/周一出数）。';
    }
    if ((latest.boxed_unshipped || 0) > 0) html += ` 截单前已完成装箱但未交运 ${fmtNum(latest.boxed_unshipped)} 件，箱已备好卡在物流。`;
    html += `</p>`;

    if (whKey === 'combined' && whNames.length > 1) {
      html += `<p><b>【仓库对比】</b>${latestDay.date}：`;
      latestDay.warehouses.forEach(w => {
        const wtf = w.transfer_unfinish || 0;
        const tfText = wtf < 0 ? `提前完成 ${fmtNum(-wtf)}` : `调拨未完成 ${fmtNum(wtf)}`;
        html += `${w.name} 未交运 ${fmtNum(w.unshipped)} / ${tfText} / 未入库 ${fmtNum(w.inbound_unfinish)}；`;
      });
      html += `</p>`;
    }

    if (unitRows.length) {
      const top = unitRows.slice(0, 3).filter(u => (u.unshipped || 0) > 0 || (u.transfer_unfinish || 0) > 0);
      if (top.length) {
        html += `<p><b>【事业部重点】</b>最新日积压靠前：`;
        top.forEach(u => html += `${u.group}（未交运 ${fmtNum(u.unshipped)} / 调拨未完成 ${fmtNum(u.transfer_unfinish)}）；`);
        html += `</p>`;
      }
    }
    $('conclusionText').innerHTML = html;

    // 供应链漏斗
    const flow = [
      { name: '来货签收', val: latest.receipt, next: latest.inbound_actual, key: 'receipt' },
      { name: '实际入库', val: latest.inbound_actual, next: latest.boxed, key: 'inbound_actual' },
      { name: '完成装箱', val: latest.boxed, next: latest.ship_out, key: 'boxed' },
      { name: '当日交运', val: latest.ship_out, next: null, key: 'ship_out' }
    ];
    $('flowBox').innerHTML = flow.map((s, i) => {
      const rate = s.next !== null && (s.val || 0) > 0 ? fmtPct(s.next, s.val) : '';
      const rateLbl = rate && i > 0 ? `${flow[i - 1].name}→${s.name} 转化率 ${rate}` : (rate ? `转化率 ${rate}` : '');
      const backlog = i === 3 ? ((flow[2].val || 0) - (s.val || 0)) : null;
      const backlogHtml = (backlog !== null && backlog > 0) ? `<div class="rate" style="color:var(--danger);font-weight:600;">完成装箱未当日交运 ${fmtNum(backlog)} 件</div>` : '';
      return `<div class="flow-step"><div class="name">${s.name}</div><div class="val">${fmtNum(s.val)}</div>${rateLbl ? `<div class="rate">${rateLbl}</div>` : ''}${backlogHtml}</div>`;
    }).join('');

    const flowCalc = [
      { toKey: 'inbound_actual', fromKey: 'receipt' },
      { toKey: 'boxed', fromKey: 'inbound_actual' },
      { toKey: 'ship_out', fromKey: 'boxed' }
    ].map(({ toKey, fromKey }) => {
      const toName = FIELDS[toKey], fromName = FIELDS[fromKey];
      const toVal = latest[toKey], fromVal = latest[fromKey];
      const rate = fromVal ? ((toVal / fromVal) * 100).toFixed(1) : '—';
      return `<li><b>${toName.replace('数量', '')} ${rate}%</b> = ${toName} ${fmtNum(toVal)} ÷ ${fromName} ${fmtNum(fromVal)}</li>`;
    }).join('');
    $('flowNote').innerHTML = `<div style="margin-bottom:4px;"><b>转化率计算说明：</b></div><ul style="margin:0 0 0 18px;padding:0;line-height:1.7;">${flowCalc}</ul><div style="margin-top:6px;">低于 100% 表示在该节点产生积压。</div>`;

    // KPI
    const kpis = [
      { lbl: '当日来货签收', val: latest.receipt, ok: true },
      { lbl: '当日实际入库', val: latest.inbound_actual, ok: true },
      { lbl: '当日未入库', val: latest.inbound_unfinish, warn: true },
      { lbl: '当日完成装箱', val: latest.boxed, ok: true },
      { lbl: '完成装箱未当日交运', val: (latest.boxed || 0) - (latest.ship_out || 0), warn: true },
      { lbl: '当日调拨未完成', val: latest.transfer_unfinish, warn: true, negOk: true },
      { lbl: '当日交运', val: latest.ship_out, ok: true },
      { lbl: '当日未交运', val: latest.unshipped, warn: true },
      { lbl: '截单前装箱未交运', val: latest.boxed_unshipped, warn: true }
    ];
    $('kpiBox').innerHTML = kpis.map(k => {
      const isNegOk = k.negOk && (k.val || 0) < 0;
      const cls = isNegOk ? 'ok' : (k.warn ? 'warn' : (k.ok ? 'ok' : ''));
      const sub = isNegOk ? '<div style="font-size:11px;color:var(--green);margin-top:4px;">提前完成</div>' : '';
      return `<div class="kpi"><div class="lbl">${k.lbl}</div><div class="val ${cls}">${fmtNum(k.val)}</div>${sub}</div>`;
    }).join('');

    // 图表（复用 setChart，自动纳入全局 resize）
    const dates = series.map(r => r.date);
    setChart('c_backlog', $('c_backlog')).setOption({
      tooltip: { trigger: 'axis' }, grid: { left: 60, right: 20, top: 30, bottom: 60 },
      legend: { data: ['未交运', '未入库', '调拨未完成'], top: 0 },
      xAxis: { type: 'category', data: dates, axisLabel: { rotate: 45, fontSize: 10 } },
      yAxis: { type: 'value', name: '件数' },
      dataZoom: [{ type: 'inside' }, { type: 'slider', height: 16, bottom: 18 }],
      series: [
        { name: '未交运', type: 'line', data: series.map(r => r.unshipped ?? 0), smooth: true, itemStyle: { color: '#e0533d' }, lineStyle: { width: 2 } },
        { name: '未入库', type: 'line', data: series.map(r => r.inbound_unfinish ?? 0), smooth: true, itemStyle: { color: '#f0a500' }, lineStyle: { width: 2 } },
        { name: '调拨未完成', type: 'line', data: series.map(r => r.transfer_unfinish ?? 0), smooth: true, itemStyle: { color: '#7c5cff' }, lineStyle: { width: 2 } }
      ]
    }, true);
    setChart('c_ship', $('c_ship')).setOption({
      tooltip: { trigger: 'axis' }, grid: { left: 60, right: 20, top: 30, bottom: 60 },
      legend: { data: ['当日交运', '当日未交运'], top: 0 },
      xAxis: { type: 'category', data: dates, axisLabel: { rotate: 45, fontSize: 10 } },
      yAxis: { type: 'value', name: '件数' },
      dataZoom: [{ type: 'inside' }, { type: 'slider', height: 16, bottom: 18 }],
      series: [
        { name: '当日交运', type: 'bar', stack: 's', data: series.map(r => r.ship_out ?? 0), itemStyle: { color: '#2b9d6e' } },
        { name: '当日未交运', type: 'bar', stack: 's', data: series.map(r => r.unshipped ?? 0), itemStyle: { color: '#e0533d' } }
      ]
    }, true);
    setChart('c_throughput', $('c_throughput')).setOption({
      tooltip: { trigger: 'axis' }, grid: { left: 60, right: 20, top: 30, bottom: 60 },
      legend: { data: ['实际入库', '完成装箱', '当日交运'], top: 0 },
      xAxis: { type: 'category', data: dates, axisLabel: { rotate: 45, fontSize: 10 } },
      yAxis: { type: 'value', name: '件数' },
      dataZoom: [{ type: 'inside' }, { type: 'slider', height: 16, bottom: 18 }],
      series: [
        { name: '实际入库', type: 'line', data: series.map(r => r.inbound_actual ?? 0), smooth: true, itemStyle: { color: '#f0a500' } },
        { name: '完成装箱', type: 'line', data: series.map(r => r.boxed ?? 0), smooth: true, itemStyle: { color: '#1f7aec' } },
        { name: '当日交运', type: 'line', data: series.map(r => r.ship_out ?? 0), smooth: true, itemStyle: { color: '#2b9d6e' } }
      ]
    }, true);
    const cutoff = month.daily.filter(d => (d.combined.cutoff_sea > 0 || d.combined.cutoff_rail > 0 || d.combined.cutoff_air > 0));
    setChart('c_cutoff', $('c_cutoff')).setOption({
      tooltip: { trigger: 'axis' }, grid: { left: 60, right: 20, top: 30, bottom: 60 },
      legend: { data: ['美国海运', '铁陆运', '空运+快递'], top: 0 },
      xAxis: { type: 'category', data: cutoff.map(d => d.date), axisLabel: { rotate: 45, fontSize: 10 } },
      yAxis: { type: 'value', name: '件数' },
      series: [
        { name: '美国海运', type: 'bar', stack: 'c', data: cutoff.map(d => d.combined.cutoff_sea || 0), itemStyle: { color: '#1f7aec' } },
        { name: '铁陆运', type: 'bar', stack: 'c', data: cutoff.map(d => d.combined.cutoff_rail || 0), itemStyle: { color: '#2b9d6e' } },
        { name: '空运+快递', type: 'bar', stack: 'c', data: cutoff.map(d => d.combined.cutoff_air || 0), itemStyle: { color: '#f0a500' } }
      ]
    }, true);
    setChart('c_wh', $('c_wh')).setOption({
      tooltip: { trigger: 'axis' }, grid: { left: 60, right: 20, top: 30, bottom: 60 },
      legend: { data: ['未交运', '调拨未完成', '未入库'], top: 0 },
      xAxis: { type: 'category', data: latestDay.warehouses.map(w => w.name) },
      yAxis: { type: 'value', name: '件数' },
      series: [
        { name: '未交运', type: 'bar', data: latestDay.warehouses.map(w => w.unshipped || 0), itemStyle: { color: '#e0533d' } },
        { name: '调拨未完成', type: 'bar', data: latestDay.warehouses.map(w => w.transfer_unfinish || 0), itemStyle: { color: '#7c5cff' } },
        { name: '未入库', type: 'bar', data: latestDay.warehouses.map(w => w.inbound_unfinish || 0), itemStyle: { color: '#f0a500' } }
      ]
    }, true);

    // 事业部异常排名
    const utb = $('unitTable').querySelector('tbody');
    utb.innerHTML = unitRows.map(u => {
      const flag = (u.unshipped > 0 || (u.transfer_unfinish || 0) > 0 || u.boxed_unshipped > 0) ? '⚠️' : '';
      return `<tr><td>${u.group}</td><td class="num ${u.unshipped > 0 ? 'warn' : ''}">${fmtNum(u.unshipped)}</td><td class="num ${u.transfer_unfinish > 0 ? 'warn' : (u.transfer_unfinish < 0 ? 'ok' : '')}">${fmtNum(u.transfer_unfinish)}</td><td class="num ${u.boxed_unshipped > 0 ? 'warn' : ''}">${fmtNum(u.boxed_unshipped)}</td><td class="small">${u.remark || ''}</td></tr>`;
    }).join('');

    // 完整 14 字段日报
    const fieldMap = [
      ['当日来货签收数量', latest.receipt, '工厂来货，仓库签收'],
      ['上日未入库数量', latest.prev_inbound, '昨日结转入库 backlog'],
      ['当日待入库数量（质检完成）', latest.qc_done, '质检已完成，等待上架'],
      ['当日应入库数量', latest.inbound_plan, '计划今日上架量'],
      ['当日实际入库数量', latest.inbound_actual, '实际完成上架量'],
      ['当日未入库数量', latest.inbound_unfinish, '仍滞留在收货/质检/上架'],
      ['上日未完成调拨数量', latest.prev_transfer, '昨日结转调拨 backlog'],
      ['当日调拨计划数量', latest.transfer_plan, '销售今日下达调拨计划'],
      ['调拨合计数量', latest.transfer_total, '今日累计需完成调拨'],
      ['当日完成装箱数量', latest.boxed, '仓库已完成装箱'],
      ['当日调拨未完成数量', latest.transfer_unfinish, (latest.transfer_unfinish || 0) < 0 ? '负值表示提前完成次日计划' : '仓库拣/包/装未完成'],
      ['上日未交运总调拨数量', latest.prev_unshipped, '昨日结转物流未交运'],
      ['当日交运数量', latest.ship_out, '物流已交运'],
      ['当日未交运数量', latest.unshipped, '物流待交运（积压）']
    ];
    $('dailyTable').querySelector('tbody').innerHTML = fieldMap.map(([f, v, d]) => `<tr><td>${f}</td><td class="num ${(v > 0 && (f.includes('未') || f.includes('未完成'))) ? 'warn' : ''}">${fmtNum(v)}</td><td class="small">${d}</td></tr>`).join('');
  }

  // ---------------- 视图切换 ----------------
  function showUpload() {
    const u = $('whUploadZone'); const d = $('whDashboard');
    if (u) u.classList.remove('hidden');
    if (d) d.classList.add('hidden');
  }
  function showDashboard() {
    const u = $('whUploadZone'); const d = $('whDashboard');
    if (u) u.classList.add('hidden');
    if (d) d.classList.remove('hidden');
  }

  function exportConclusion() {
    const text = $('conclusionText').innerText;
    const blob = new Blob([text], { type: 'text/plain;charset=utf-8' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = '仓配发运分析结论_' + ($('whMonthSelect').value || '') + '.txt';
    a.click();
  }

  // ---------------- 初始化 ----------------
  function init() {
    const dz = $('whDropZone');
    const fi = $('whFileInput');
    if (!dz || !fi) return;
    dz.addEventListener('click', () => fi.click());
    ['dragenter', 'dragover'].forEach(e => dz.addEventListener(e, ev => { ev.preventDefault(); dz.classList.add('active'); }));
    ['dragleave', 'drop'].forEach(e => dz.addEventListener(e, ev => { ev.preventDefault(); dz.classList.remove('active'); }));
    dz.addEventListener('drop', ev => { const f = ev.dataTransfer.files[0]; if (f) loadFile(f); });
    fi.addEventListener('change', ev => { const f = ev.target.files[0]; if (f) loadFile(f); });
    const ms = $('whMonthSelect'); if (ms) ms.addEventListener('change', () => { populateWarehouses(); render(); });
    const ws = $('whWarehouseSelect'); if (ws) ws.addEventListener('change', render);
    const ex = $('whExportBtn'); if (ex) ex.addEventListener('click', exportConclusion);
  }

  window.Warehouse = { init, render, setData: d => { appData = d; } };
})();
