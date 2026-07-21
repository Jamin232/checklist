// ============================================================
// daily.js — 日度监控看板（在途 / 时效SLA / 异常 / 成本 / 明日预警）
// 与查验看板(chayan.js)共用同一上传入口；今日表由 chayan.js 通过 Daily.setData 注入，
// 昨日表由本文件自行读取，用于异常监控的"较昨日变动"。
// ============================================================

const Daily = (function () {
  // ---------------- 状态 ----------------
  let todayRecs = null;     // 今日表解析后的记录
  let yesterdayRecs = null; // 昨日表解析后的记录
  let todayDate = null;     // 今日表文件名解析出的日期（用于"当日新增查验"基准）
  let yesterdayDate = null; // 昨日表文件名解析出的日期
  let _inited = false;
  let slaPeriod = 'all'; // 'all' | 'halfmonth' | 'month' | 'twomonth'

  let TODAY, TOMORROW; // 由 setData 的文件日期推导；未解析到则取真实今日
  function computeToday(date) {
    let base;
    if (date && !isNaN(date.getTime())) {
      base = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
    } else {
      const n = new Date();
      base = new Date(Date.UTC(n.getUTCFullYear(), n.getUTCMonth(), n.getUTCDate()));
    }
    TODAY = base;
    TOMORROW = new Date(TODAY);
    TOMORROW.setUTCDate(TOMORROW.getUTCDate() + 1);
  }
  computeToday(null); // 默认真实今日

  const STATUS_WORDS = ['查验中', '开查中', '索赔中', '赔付中']; // 异常状态词
  const MILE = ['到港', '清关', '派送'];
  const MILE_PRIORITY = { '到港': 1, '清关': 2, '派送': 3 };

  // ---------------- 工具函数 ----------------
  function safeStr(v) {
    if (v === null || v === undefined) return '';
    if (typeof v === 'string') return v.trim();
    if (typeof v === 'number' && isNaN(v)) return '';
    return String(v).trim();
  }
  function safeNum(v) {
    if (v === null || v === undefined || v === '') return 0;
    if (typeof v === 'number') return isNaN(v) ? 0 : v;
    const n = parseFloat(String(v).replace(/,/g, ''));
    return isNaN(n) ? 0 : n;
  }
  function parseDate(val) {
    if (!val) return null;
    if (val instanceof Date) return new Date(Date.UTC(val.getUTCFullYear(), val.getUTCMonth(), val.getUTCDate()));
    if (typeof val === 'number') {
      // Excel 序列号
      const d = new Date((val - 25569) * 86400 * 1000);
      return isNaN(d.getTime()) ? null : new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()));
    }
    const s = safeStr(val);
    if (!s) return null;
    let m = s.match(/^(\d{4})年(\d{1,2})月(\d{1,2})日?/);
    if (m) return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
    m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
    if (m) return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
    m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})/);
    if (m) return new Date(Date.UTC(+m[3], +m[2] - 1, +m[1]));
    m = s.match(/^(\d{1,2})[-/.](\d{1,2})$/);
    if (m) return makeMDate(+m[1], +m[2]); // 月/日（无年）
    const d = new Date(s);
    return isNaN(d.getTime()) ? null : new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()));
  }
  // 由月/日构造日期（处理跨年）
  function makeMDate(month, day) {
    const y = TODAY.getUTCFullYear();
    let d = new Date(Date.UTC(y, month - 1, day));
    const diff = dayDiff(d, TODAY);
    if (diff > 200) d = new Date(y - 1, month - 1, day);
    else if (diff < -200) d = new Date(y + 1, month - 1, day);
    return d;
  }
  function sameDay(a, b) {
    return a && b && a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
  }
  // DST 安全的整天数差：返回 a - b 的天数（基于 UTC  midnight，避免夏令时 23/25 小时误差）
  function dayDiff(a, b) {
    return Math.round((Date.UTC(a.getFullYear(), a.getMonth(), a.getDate()) - Date.UTC(b.getFullYear(), b.getMonth(), b.getDate())) / 86400000);
  }
  function fmtDate(d) {
    if (!d) return '--';
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  }
  function fmtMD(d) {
    if (!d) return '--';
    return `${d.getMonth() + 1}/${d.getDate()}`;
  }
  function getWeekKey(d) {
    const dt = d instanceof Date ? d : parseDate(d);
    if (!dt) return '';
    const base = new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());
    const day = base.getDay() || 7;
    const mon = new Date(base);
    mon.setDate(base.getDate() - day + 1);
    const y = mon.getFullYear();
    const dayOfYear = Math.floor((mon - new Date(y, 0, 1)) / 86400000) + 1;
    const w = Math.ceil(dayOfYear / 7);
    const mm = String(mon.getMonth() + 1).padStart(2, '0');
    const dd = String(mon.getDate()).padStart(2, '0');
    return `${y}-W${String(w).padStart(2, '0')} (${mm}/${dd})`;
  }
  function fmtYM(d) {
    if (!d) return '';
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
  }

  // ---------------- 解析：把原始行 -> daily 记录 ----------------
  // 按子串匹配列名，避免表头换行符(\n)差异导致整列读不到
  function buildColMap(rows) {
    const m = {};
    if (!rows || !rows.length) return m;
    const keys = Object.keys(rows[0]);
    const find = (...subs) => keys.find(k => subs.every(s => String(k).includes(s))) || '';
    return {
      sub: find('分出仓单号'),
      sign: find('实际签收时间'),
      late: find('物流最晚送达时间'),
      ship: find('仓库出货日期'),
      status: find('货物状态'),
      remark: find('状态备注'),
      dom: find('国内查验时间'),
      dest: find('目的地查验时间'),
      main: find('主出仓单号'),
      type: find('类型'),
      logi: find('素芸物流渠道'),
      agent: find('代理'),
      agentCh: find('代理渠道'),
      cust: find('客户'),
      prod: find('产品属性'),
      country: find('国家'),
      weight: find('毛重'),
      vol: find('方数CBM'),
      ref: find('参考时效')
    };
  }
  function parseDailyRows(rows) {
    const map = buildColMap(rows);
    const seen = new Set();
    const list = [];
    for (const row of rows) {
      // 分出仓单号 -> 换行分割，全局去重
      const subStr = safeStr(row[map.sub]);
      const raw = subStr.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
      const tickets = raw.filter(t => !seen.has(t));
      if (raw.length > 0 && tickets.length === 0) continue; // 整行重复跳过
      const ticketCount = tickets.length > 0 ? tickets.length : 1;
      tickets.forEach(t => seen.add(t));

      const signTime = parseDate(row[map.sign]);
      const latestDeliver = parseDate(row[map.late]);
      const shipDate = parseDate(row[map.ship]);
      const goodsStatus = safeStr(row[map.status]);
      const remark = safeStr(row[map.remark]);
      const domInspectDate = parseDate(row[map.dom]);
      const destInspectDate = parseDate(row[map.dest]);

      const isInspecting = goodsStatus.includes('查验中') || goodsStatus.includes('开查中');
      const isAbnormal = STATUS_WORDS.some(w => goodsStatus.includes(w));

      list.push({
        mainTicket: safeStr(row[map.main]),
        tickets, ticketCount,
        type: safeStr(row[map.type]),
        logisticChannel: safeStr(row[map.logi]),
        // 渠道大类 = 国家 + 类型（与查验看板口径一致）
        channelCategory: (safeStr(row[map.country]) || '') + (safeStr(row[map.type]) || '') || undefined,
        agent: safeStr(row[map.agent]),
        agentChannel: safeStr(row[map.agentCh]),
        customer: safeStr(row[map.cust]),
        productAttr: safeStr(row[map.prod]),
        country: safeStr(row[map.country]),
        weight: safeNum(row[map.weight]),
        volume: safeNum(row[map.vol]),
        shipDate,
        signTime,
        latestDeliver,
        refLead: safeNum(row[map.ref]),
        goodsStatus,
        remark,
        domInspectDate,
        destInspectDate,
        inTransit: signTime === null,
        isInspecting,
        isAbnormal
      });
    }
    return list;
  }

  // ---------------- 数据注入 ----------------
  // date: 由文件名解析出的日期(Date|null)，用于推导 TODAY（当日新增查验 / 明日预警的基准）
  function setData(rawRows, date) {
    computeToday(date);
    todayDate = date || null;
    todayRecs = parseDailyRows(rawRows || []);
  }
  // 昨日表：由 chayan.js 读取后注入（支持多文件按文件名日期自动区分今日/昨日）
  function setYesterday(rawRows, date) {
    yesterdayDate = date || null;
    yesterdayRecs = parseDailyRows(rawRows || []);
    if (typeof showToast === 'function') showToast(`✓ 昨日表已载入（${yesterdayRecs.length} 条），异常监控显示较昨日变动`, 'success');
    if (document.querySelector('.tab-btn[data-tab="d_abnormal"]')?.classList.contains('active')) renderAbnormal();
  }

  // ---------------- 通用聚合 ----------------
  function groupSum(recs, keyFn) {
    const m = {};
    for (const r of recs) {
      const k = keyFn(r);
      if (!k) continue;
      if (!m[k]) m[k] = { tickets: 0, weight: 0, volume: 0 };
      m[k].tickets += r.ticketCount;
      m[k].weight += r.weight * r.ticketCount;
      m[k].volume += r.volume * r.ticketCount;
    }
    return m;
  }
  function sortByTickets(m) {
    return Object.entries(m).map(([k, v]) => ({ key: k, ...v })).sort((a, b) => b.tickets - a.tickets);
  }

  // ---------------- ECharts 管理 ----------------
  const chartMap = {};
  function getChart(id) {
    const el = document.getElementById(id);
    if (!el) return null;
    if (chartMap[id] && !chartMap[id].isDisposed()) return chartMap[id];
    const c = echarts.init(el);
    chartMap[id] = c;
    return c;
  }
  function setOpt(id, opt) {
    const c = getChart(id);
    if (!c) return;
    c.setOption(opt, { notMerge: true, lazyUpdate: true });
    setTimeout(() => c.resize(), 60);
  }
  window.addEventListener('resize', () => {
    Object.values(chartMap).forEach(c => { if (c && !c.isDisposed()) c.resize(); });
  });

  function noData(htmlId) {
    const el = document.getElementById(htmlId);
    if (el) el.innerHTML = '<div style="padding:30px;text-align:center;color:#999">请先上传今日跟踪表（若要看较昨日变动，请一并选择昨日表）</div>';
  }

  // ============================================================
  // ① 日度总览
  // ============================================================
  function renderOverview() {
    if (!todayRecs) { noData('ov-cards'); noData('ov-extra'); return; }
    const inTransit = todayRecs.filter(r => r.inTransit);
    const inT = groupSum(inTransit, () => 'all')['all'] || { tickets: 0, weight: 0, volume: 0 };
    const abnormal = todayRecs.filter(r => r.isAbnormal);
    const inspecting = todayRecs.filter(r => r.isInspecting);
    const newInspect = todayRecs.filter(r =>
      (r.domInspectDate && sameDay(r.domInspectDate, TODAY)) ||
      (r.destInspectDate && sameDay(r.destInspectDate, TODAY))).length;

    // 在途超期(>5天)
    let overdue5 = 0;
    inTransit.forEach(r => {
      if (r.latestDeliver) {
        const d = dayDiff(TODAY, r.latestDeliver);
        if (d > 5) overdue5++;
      }
    });
    // 明日里程碑
    let tom = { 到港: 0, 清关: 0, 派送: 0 };
    todayRecs.forEach(r => {
      const t = tomorrowPrimary(r);
      if (t) tom[t]++;
    });

    const cards = [
      { num: inT.tickets, label: '在途票数', sub: `${(inT.weight / 1000).toFixed(1)}吨 / ${inT.volume.toFixed(1)}方`, cls: 'ov-intransit' },
      { num: inspecting.length, label: '查验进行中', sub: `异常单 ${abnormal.length}`, cls: 'ov-abn' },
      { num: newInspect, label: '当日新增查验', sub: fmtMD(TODAY), cls: 'ov-new' },
      { num: overdue5, label: '在途超期>5天', sub: '未妥投', cls: 'ov-overdue' },
      { num: tom['到港'] + tom['清关'] + tom['派送'], label: '明日到港/清关/派送', sub: `到${tom['到港']}/清${tom['清关']}/派${tom['派送']}`, cls: 'ov-tom' }
    ];
    document.getElementById('ov-cards').innerHTML = cards.map(c => `
      <div class="kpi-card ${c.cls}">
        <div class="kpi-num">${c.num}</div>
        <div class="kpi-label">${c.label}</div>
        <div class="kpi-sub">${c.sub}</div>
      </div>`).join('');

    // 昨日对比（异常相关）
    let extra = `<div class="ov-note">数据基准日：${fmtDate(TODAY)} ｜ 在途定义：实际签收时间为空即视为在途（含赔付中/索赔中/开查中未签收单）。</div>`;
    if (yesterdayRecs) {
      const t = dailyMetrics(todayRecs, TODAY), y = dailyMetrics(yesterdayRecs, yesterdayDate || TODAY);
      const dIns = t.inspecting - y.inspecting;
      const dAbn = t.abnormal - y.abnormal;
      const dNew = t.newInspect - y.newInspect;
      extra += `<div class="ov-delta">
        <span>查验进行中 较昨日 <b class="${dIns >= 0 ? 'up' : 'down'}">${dIns >= 0 ? '+' : ''}${dIns}</b></span>
        <span>异常单 较昨日 <b class="${dAbn >= 0 ? 'up' : 'down'}">${dAbn >= 0 ? '+' : ''}${dAbn}</b></span>
        <span>当日新增查验 较昨日 <b class="${dNew >= 0 ? 'up' : 'down'}">${dNew >= 0 ? '+' : ''}${dNew}</b></span>
      </div>`;
    } else {
      extra += `<div class="ov-note" style="color:#e08e0b">未载入昨日表，异常监控暂不显示"较昨日变动"。</div>`;
    }
    document.getElementById('ov-extra').innerHTML = extra;
  }

  // refDate: 该数据对应的"当日"基准（今日表用 TODAY，昨日表用昨日文件日期），确保"较昨日变动"各自独立
  function dailyMetrics(recs, refDate) {
    const ref = refDate || TODAY;
    let inspecting = 0, abnormal = 0, newInspect = 0;
    for (const r of recs) {
      if (r.isInspecting) inspecting++;
      if (r.isAbnormal) abnormal++;
      if ((r.domInspectDate && sameDay(r.domInspectDate, ref)) || (r.destInspectDate && sameDay(r.destInspectDate, ref))) newInspect++;
    }
    return { inspecting, abnormal, newInspect };
  }

  // ============================================================
  // ② 在途概览
  // ============================================================
  function renderIntransit() {
    if (!todayRecs) { noData('it-transitChart'); noData('it-channelBody'); noData('it-agentBody'); noData('it-custBody'); return; }
    const inTransit = todayRecs.filter(r => r.inTransit);

    // 运输方式构成（饼图）
    const byType = sortByTickets(groupSum(inTransit, r => r.type || '未知'));
    const pie = byType.map(x => ({ name: x.key, value: x.tickets }));
    setOpt('it-transitChart', {
      tooltip: { trigger: 'item', formatter: p => `${p.name}<br>票数：${p.value}（${(p.percent).toFixed(1)}%）<br>吨：${(x_weight(byType, p.name) / 1000).toFixed(1)}　方：${x_vol(byType, p.name).toFixed(1)}` },
      legend: { type: 'scroll', bottom: 0 },
      series: [{ type: 'pie', radius: ['38%', '66%'], center: ['50%', '45%'], data: pie, label: { formatter: '{b}\n{c}' } }]
    });

    renderTable('it-channelBody', groupSum(inTransit, r => r.logisticChannel), '素芸物流渠道');
    renderTable('it-agentBody', groupSum(inTransit, r => r.agent), '代理');
    renderTable('it-custBody', groupSum(inTransit, r => r.customer), '客户');

    // 汇总卡
    const tot = groupSum(inTransit, () => 'all').all || { tickets: 0, weight: 0, volume: 0 };
    const sum = document.getElementById('it-summary');
    if (sum) sum.innerHTML = `在途合计：<b>${tot.tickets}</b> 票 ｜ <b>${(tot.weight / 1000).toFixed(1)}</b> 吨 ｜ <b>${tot.volume.toFixed(1)}</b> 方`;
  }
  function x_weight(arr, name) { const x = arr.find(a => a.key === name); return x ? x.weight : 0; }
  function x_vol(arr, name) { const x = arr.find(a => a.key === name); return x ? x.volume : 0; }

  function renderTable(tbodyId, grouped, dimLabel) {
    const el = document.getElementById(tbodyId);
    if (!el) return;
    const arr = sortByTickets(grouped).slice(0, 60);
    const tot = arr.reduce((s, x) => s + x.tickets, 0) || 1;
    el.innerHTML = `<tr><th>${dimLabel}</th><th>在途票数</th><th>占比</th><th>吨数</th><th>方数</th></tr>` +
      arr.map(x => `<tr><td>${x.key || '—'}</td><td>${x.tickets}</td><td>${(x.tickets / tot * 100).toFixed(1)}%</td><td>${(x.weight / 1000).toFixed(1)}</td><td>${x.volume.toFixed(1)}</td></tr>`).join('') +
      (arr.length === 0 ? '<tr><td colspan="5" style="text-align:center;color:#999">无数据</td></tr>' : '');
  }

  // ============================================================
  // ③ 时效 SLA
  // ============================================================
  // 分位值（升序第 ceil(p*n)-1 个）
  function percentile(arr, p) {
    if (arr.length === 0) return 0;
    const s = [...arr].sort((a, b) => a - b);
    const idx = Math.min(s.length - 1, Math.max(0, Math.ceil(p * s.length) - 1));
    return s[idx];
  }

  function setSlaPeriod(p) { slaPeriod = p; document.querySelectorAll('.sla-period-btn').forEach(b => b.classList.toggle('active', b.dataset.period === p)); renderSLA(); }

  function renderSLA() {
    if (!todayRecs) { noData('sla-chart'); noData('sla-table'); return; }
    // 时间过滤
    let pool = todayRecs;
    if (slaPeriod !== 'all' && TODAY) {
      const days = { halfmonth: 15, month: 30, twomonth: 60 }[slaPeriod] || 99999;
      const cutoff = new Date(TODAY);
      cutoff.setUTCDate(cutoff.getUTCDate() - days);
      pool = todayRecs.filter(r => r.shipDate && r.shipDate >= cutoff);
    }
    // 同事原算法：剔除异常单（查验/开查/索赔/赔付），仅用正常已签收单推导时效基线
    const normalSigned = pool.filter(r => !r.inTransit && !r.isAbnormal && r.shipDate && r.signTime);
    const byCh = {};
    for (const r of normalSigned) {
      const k = r.channelCategory || r.logisticChannel || '未知';
      if (!byCh[k]) byCh[k] = { leads: [], refLeads: [] };
      byCh[k].leads.push(Math.round(dayDiff(r.signTime, r.shipDate))); // 总时效=签收-仓库出货
      if (r.refLead > 0) byCh[k].refLeads.push(r.refLead);
    }
    const MIN_SAMPLE = 5;
    const arr = Object.entries(byCh).map(([k, v]) => {
      const n = v.leads.length;
      const avgLead = n ? v.leads.reduce((a, b) => a + b, 0) / n : 0;
      const p90 = percentile(v.leads, 0.9);                 // 90% 置信时效（区间内）
      const tail = v.leads.filter(l => l > p90);            // 超区间（长尾）单
      const avgExceed = tail.length ? tail.reduce((a, b) => a + (b - p90), 0) / tail.length : 0;
      const suggested = n >= MIN_SAMPLE ? p90 + avgExceed : avgLead; // 建议时效 = P90 + 平均超出
      const refAvg = v.refLeads.length ? v.refLeads.reduce((a, b) => a + b, 0) / v.refLeads.length : 0;
      const over = v.leads.filter(l => l > suggested).length;
      const rate = n ? (n - over) / n * 100 : 0;            // SLA达成率（以建议时效为承诺基线）
      return { key: k, n, avgLead, p90, suggested, refAvg, over, rate, small: n < MIN_SAMPLE };
    }).sort((a, b) => b.n - a.n).slice(0, 20);

    // 图表：平均实际时效 / 建议时效 / 参考时效（柱）+ SLA达成率（线，次轴）
    setOpt('sla-chart', {
      tooltip: { trigger: 'axis' },
      legend: { data: ['平均实际时效(天)', '建议时效(天)', '参考时效(Z列)', 'SLA达成率'], bottom: 0 },
      grid: { left: 50, right: 55, top: 25, bottom: 55 },
      xAxis: { type: 'category', data: arr.map(x => x.key), axisLabel: { interval: 0, rotate: 30, fontSize: 10 } },
      yAxis: [
        { type: 'value', name: '天' },
        { type: 'value', name: '%', max: 100, axisLabel: { formatter: '{value}%' } }
      ],
      series: [
        { name: '平均实际时效(天)', type: 'bar', data: arr.map(x => +x.avgLead.toFixed(1)), itemStyle: { color: '#2b6cb0' } },
        { name: '建议时效(天)', type: 'bar', data: arr.map(x => +x.suggested.toFixed(1)), itemStyle: { color: '#38a169' } },
        { name: '参考时效(Z列)', type: 'bar', data: arr.map(x => +x.refAvg.toFixed(1)), itemStyle: { color: '#e08e0b' } },
        { name: 'SLA达成率', type: 'line', yAxisIndex: 1, data: arr.map(x => +x.rate.toFixed(1)), itemStyle: { color: '#d53f8c' }, symbolSize: 7 }
      ]
    });

    // 表格
    const tb = document.getElementById('sla-table');
    if (tb) {
      const dimLabel = slaPeriod === 'all' ? '全量' : { halfmonth: '近半月', month: '近一月', twomonth: '近两月' }[slaPeriod];
      tb.innerHTML = `<tr><th>渠道大类</th><th>正常已签收</th><th>平均实际时效</th><th>建议时效*</th><th>参考时效(Z)</th><th>超时单数</th><th>SLA达成率</th></tr>` +
        arr.map(x => {
          const rc = x.rate >= 90 ? 'rate-good' : (x.rate >= 80 ? 'rate-mid' : 'rate-bad');
          return `<tr><td>${x.key}</td><td>${x.n}${x.small ? '<span class="pct">样本少</span>' : ''}</td><td>${x.avgLead.toFixed(1)}</td><td><b>${x.suggested.toFixed(1)}</b></td><td>${x.refAvg ? x.refAvg.toFixed(1) : '-'}</td><td class="rate-bad">${x.over}</td><td class="${rc}">${x.rate.toFixed(1)}%</td></tr>`;
        }).join('') +
        `<tr style="font-weight:700;background:#f3f6fa"><td>合计</td><td>${normalSigned.length}</td><td colspan="5"></td></tr>` +
        `<tr><td colspan="7" style="color:#888;font-size:11px">*建议时效 = P90(90%置信区间上限) + 超区间单平均超出天数（已剔除查验/索赔/赔付等异常单）。<br>含义：若将此渠道的参考时效设为"建议时效"，则约 ${arr.length > 0 ? Math.round(arr.filter(x=>x.rate>=90).length/arr.length*100) : 0}% 的渠道可达90%+达成率。<br>数据范围：${dimLabel} | 时效=实际签收时间−仓库出货日期</td></tr>`;
    }
  }

  // ============================================================
  // ④ 异常监控
  // ============================================================
  function renderAbnormal() {
    if (!todayRecs) { noData('ab-cards'); noData('ab-table'); return; }
    const inspecting = todayRecs.filter(r => r.isInspecting);
    const abnormal = todayRecs.filter(r => r.isAbnormal);
    const newInspect = todayRecs.filter(r =>
      (r.domInspectDate && sameDay(r.domInspectDate, TODAY)) ||
      (r.destInspectDate && sameDay(r.destInspectDate, TODAY)));

    let deltaHtml = '';
    if (yesterdayRecs) {
      const t = dailyMetrics(todayRecs, TODAY), y = dailyMetrics(yesterdayRecs, yesterdayDate || TODAY);
      const mk = (cur, prev, label) => {
        const d = cur - prev;
        return `<div class="ab-delta-item">${label}<b class="${d >= 0 ? 'up' : 'down'}">${d >= 0 ? '+' : ''}${d}</b><span>较昨日</span></div>`;
      };
      deltaHtml = `<div class="ab-delta">${mk(t.inspecting, y.inspecting, '查验进行中')}${mk(t.abnormal, y.abnormal, '异常单')}${mk(t.newInspect, y.newInspect, '当日新增查验')}</div>`;
    } else {
      deltaHtml = '<div class="ov-note" style="color:#e08e0b">未载入昨日表，暂不显示"较昨日变动"。</div>';
    }

    document.getElementById('ab-cards').innerHTML =
      `<div class="kpi-card ov-abn"><div class="kpi-num">${inspecting.length}</div><div class="kpi-label">查验进行中(查验中/开查中)</div></div>` +
      `<div class="kpi-card ov-overdue"><div class="kpi-num">${abnormal.length}</div><div class="kpi-label">异常单(含索赔/赔付)</div></div>` +
      `<div class="kpi-card ov-new"><div class="kpi-num">${newInspect.length}</div><div class="kpi-label">当日新增查验(${fmtMD(TODAY)})</div></div>` +
      deltaHtml;

    // 明细表（异常单）— 显示分出仓单号
    const tb = document.getElementById('ab-table');
    if (tb) {
      const rows = abnormal.slice(0, 300).map(r => {
        const t = r.isInspecting ? '查验中' : (r.goodsStatus.match(/索赔中|赔付中/) ? r.goodsStatus.match(/索赔中|赔付中/)[0] : '异常');
        const ticketDisplay = r.tickets.length > 1 ? r.tickets.slice(0, 3).join('<br>') + (r.tickets.length > 3 ? `<br><span style="color:#888;font-size:10px">+${r.tickets.length - 3}更多</span>` : '') : (r.tickets[0] || r.mainTicket);
        return `<tr><td>${ticketDisplay}</td><td>${r.logisticChannel}</td><td>${r.agent}</td><td>${r.customer}</td><td>${r.country}</td><td class="status-inspected">${t}</td><td>${r.goodsStatus}</td></tr>`;
      }).join('');
      tb.innerHTML = `<tr><th>分出仓单号</th><th>渠道</th><th>代理</th><th>客户</th><th>国家</th><th>状态</th><th>货物状态</th></tr>` + (rows || '<tr><td colspan="7" style="text-align:center;color:#999">无异常单</td></tr>');
    }
  }

  // ============================================================
  // ⑤ 成本速览（渠道结构占比）
  // ============================================================
  let costDim = 'week', costMetric = 'tickets';
  function setCostDim(d) { costDim = d; document.querySelectorAll('.cost-dim-btn').forEach(b => b.classList.toggle('active', b.dataset.dim === d)); renderCost(); }
  function setCostMetric(m) { costMetric = m; document.querySelectorAll('.cost-metric-btn').forEach(b => b.classList.toggle('active', b.dataset.metric === m)); renderCost(); }

  function renderCost() {
    if (!todayRecs) { noData('cost-chart'); noData('cost-table'); return; }
    const dimFn = {
      week: r => r.shipDate ? getWeekKey(r.shipDate) : '未知',
      month: r => r.shipDate ? fmtYM(r.shipDate) : '未知',
      customer: r => r.customer || '未知'
    }[costDim];
    const metricFn = {
      tickets: r => r.ticketCount,
      weight: r => r.weight * r.ticketCount,
      volume: r => r.volume * r.ticketCount
    }[costMetric];
    const metricName = { tickets: '票数', weight: '吨数', volume: '方数' }[costMetric];

    // 按渠道大类聚合（非素芸物流渠道）
    const byCh = groupSum(todayRecs, r => r.channelCategory || '未知');
    // 维度拆分：每个维度值下，各渠道大类的 metric 占比
    const dimMap = {};
    for (const r of todayRecs) {
      const dk = dimFn(r);
      if (!dk) continue;
      if (!dimMap[dk]) dimMap[dk] = {};
      const ck = r.channelCategory || '未知';
      dimMap[dk][ck] = (dimMap[dk][ck] || 0) + metricFn(r);
    }
    // 仅展示占比 Top 渠道大类
    const chTotals = {};
    Object.values(dimMap).forEach(m => Object.entries(m).forEach(([ck, v]) => chTotals[ck] = (chTotals[ck] || 0) + v));
    const topCh = Object.entries(chTotals).sort((a, b) => b[1] - a[1]).slice(0, 10).map(x => x[0]);

    // 饼图：全量渠道大类结构
    const pie = topCh.map(ck => ({ name: ck, value: +chTotals[ck].toFixed(1) }));
    const other = Object.entries(chTotals).filter(([ck]) => !topCh.includes(ck)).reduce((s, [, v]) => s + v, 0);
    if (other > 0) pie.push({ name: '其他', value: +other.toFixed(1) });
    setOpt('cost-chart', {
      tooltip: { trigger: 'item', formatter: p => `${p.name}<br>${metricName}：${p.value}（${(p.percent).toFixed(1)}%）` },
      legend: { type: 'scroll', bottom: 0, textStyle: { fontSize: 10 } },
      series: [{ type: 'pie', radius: ['35%', '66%'], center: ['50%', '44%'], data: pie, label: { formatter: '{b}\n{c}' } }]
    });

    // 表格：维度 × 渠道大类 占比
    const tb = document.getElementById('cost-table');
    if (tb) {
      const dims = Object.keys(dimMap).sort();
      let html = `<tr><th>${{ week: '周', month: '月', customer: '客户(事业部)' }[costDim]}</th>`;
      topCh.forEach(ck => html += `<th>${ck}</th>`);
      html += `<th>合计(${metricName})</th></tr>`;
      for (const dk of dims) {
        const m = dimMap[dk];
        const tot = topCh.reduce((s, ck) => s + (m[ck] || 0), 0);
        html += `<tr><td>${dk}</td>`;
        topCh.forEach(ck => {
          const v = m[ck] || 0;
          const pct = tot > 0 ? (v / tot * 100).toFixed(0) : '0';
          html += `<td>${metricName === '票数' ? v : v.toFixed(1)}<br><span class="pct">${pct}%</span></td>`;
        });
        html += `<td><b>${metricName === '票数' ? tot : tot.toFixed(1)}</b></td></tr>`;
      }
      // 总计行
      html += `<tr style="font-weight:700;background:#f3f6fa"><td>全量</td>`;
      topCh.forEach(ck => html += `<td>${metricName === '票数' ? +chTotals[ck].toFixed(0) : chTotals[ck].toFixed(1)}</td>`);
      const gt = topCh.reduce((s, ck) => s + chTotals[ck], 0);
      html += `<td>${metricName === '票数' ? +gt.toFixed(0) : gt.toFixed(1)}</td></tr>`;
      tb.innerHTML = html;
    }
    const note = document.getElementById('cost-note');
    if (note) note.innerHTML = '⚠️ 跟踪表无运费数据，本模块仅展示<b>渠道大类结构占比</b>（按' + metricName + '），用以间接反映成本分布。维度：' +
      ({ week: '中国日历周', month: '月度', customer: '客户(事业部)' }[costDim]);
  }

  // ============================================================
  // ⑥ 明日预警
  // ============================================================
  function parseMilestones(remark) {
    const res = [];
    const lines = String(remark || '').split(/\r?\n/);
    for (const line of lines) {
      for (const m of MILE) {
        let mm = line.match(new RegExp(`(\\d{1,2})[./.](\\d{1,2})\\D{0,4}${m}`));
        if (mm) { res.push({ type: m, date: makeMDate(+mm[1], +mm[2]) }); continue; }
        mm = line.match(new RegExp(`${m}\\D{0,4}(\\d{1,2})[./.](\\d{1,2})`));
        if (mm) { res.push({ type: m, date: makeMDate(+mm[1], +mm[2]) }); }
      }
    }
    return res;
  }
  // 返回该单明日最核心里程碑：以"最新(日期最大)里程碑"为核心，同日再按 派送>清关>到港 优先级
  function tomorrowPrimary(rec) {
    const ms = parseMilestones(rec.remark);
    if (ms.length === 0) return null;
    ms.sort((a, b) => {
      const d = dayDiff(b.date, a.date); // 日期大的在前
      if (d !== 0) return d;
      return MILE_PRIORITY[b.type] - MILE_PRIORITY[a.type];
    });
    const latest = ms[0];
    return sameDay(latest.date, TOMORROW) ? latest.type : null; // 仅当最新里程碑日期=次日才预警
  }

  function renderTomorrow() {
    if (!todayRecs) { noData('tm-overdue'); noData('tm-mile'); return; }

    // Part A：在途超期（按梯度）— 每个梯度独立展示，Tab式分离
    const tiers = [
      { key: '≥10天', test: d => d >= 10, cls: 'tier-10', desc: '严重超期，需立即跟进' },
      { key: '≥7天', test: d => d >= 7 && d < 10, cls: 'tier-7', desc: '明显超期，关注处理' },
      { key: '>5天', test: d => d > 5 && d < 7, cls: 'tier-5', desc: '轻度超期，持续观察' }
    ];
    const list = todayRecs.filter(r => r.inTransit && r.latestDeliver);
    const overdueRows = list.map(r => ({ r, d: dayDiff(TODAY, r.latestDeliver) })).filter(x => x.d > 0);

    // 构建每个梯度的独立卡片+表格
    let oh = '<div class="tm-tier-group">';
    for (const t of tiers) {
      const rows = overdueRows.filter(x => t.test(x.d)).sort((a, b) => b.d - a.d).slice(0, 100);
      oh += `<div class="tm-tier-block">
        <div class="tm-tier-header ${t.cls}"><span class="tm-tier-title">${t.key}</span><span class="tm-tier-count">${rows.length} 单未妥投</span><span class="tm-tier-desc">${t.desc}</span></div>
        <div class="table-wrap" style="max-height:${rows.length > 8 ? 240 : 'auto'}px;overflow:auto">
          <table class="data-table">
            <thead><tr><th>分出仓单号</th><th>渠道大类</th><th>代理</th><th>客户</th><th>超期天数</th><th>最晚送达</th></tr></thead>
            <tbody>${rows.length ? rows.map(x => `<tr>
              <td>${x.r.tickets.length > 1 ? x.r.tickets.slice(0, 2).join('<br>') + '<br><span style="color:#888;font-size:10px">+' + (x.r.tickets.length - 2) + '更多</span>' : (x.r.tickets[0] || x.r.mainTicket)}</td>
              <td>${x.r.channelCategory || x.r.logisticChannel}</td>
              <td>${x.r.agent}</td>
              <td>${x.r.customer}</td>
              <td class="rate-bad">${x.d}天</td>
              <td>${fmtDate(x.r.latestDeliver)}</td>
            </tr>`).join('') : '<tr><td colspan="6" style="text-align:center;color:#999">无</td></tr>'}</tbody>
          </table>
        </div>
      </div>`;
    }
    oh += '</div>';
    document.getElementById('tm-overdue').innerHTML = oh;

    // Part B：明日到港/清关/派送
    const tom = todayRecs.map(r => ({ r, t: tomorrowPrimary(r) })).filter(x => x.t);
    const byType = { 到港: {}, 清关: {}, 派送: {} };
    tom.forEach(x => {
      const ck = x.r.logisticChannel || '未知';
      const cu = x.r.customer || '未知';
      byType[x.t][ck] = (byType[x.t][ck] || 0) + x.r.ticketCount;
      if (!byType[x.t]._cust) byType[x.t]._cust = {};
      byType[x.t]._cust[cu] = (byType[x.t]._cust[cu] || 0) + x.r.ticketCount;
    });
    const cnt = { 到港: 0, 清关: 0, 派送: 0 };
    tom.forEach(x => cnt[x.t] += x.r.ticketCount);

    let mh = `<div class="tier-row">
      <div class="tier-badge tier-5">明日到港：<b>${cnt['到港']}</b> 票</div>
      <div class="tier-badge tier-7">明日清关：<b>${cnt['清关']}</b> 票</div>
      <div class="tier-badge tier-10">明日派送：<b>${cnt['派送']}</b> 票</div>
    </div>`;
    mh += '<div class="mile-cols">';
    ['到港', '清关', '派送'].forEach(t => {
      const ch = Object.entries(byType[t]).filter(([k]) => k !== '_cust').sort((a, b) => b[1] - a[1]);
      const cu = Object.entries(byType[t]._cust || {}).sort((a, b) => b[1] - a[1]);
      mh += `<div class="mile-col"><div class="mile-title">${t}（按渠道）</div>` +
        (ch.length ? ch.map(([k, v]) => `<div class="mile-item">${k}<span>${v}</span></div>`).join('') : '<div class="mile-item">无</div>');
      mh += `<div class="mile-title" style="margin-top:8px">${t}（按客户）</div>` +
        (cu.length ? cu.slice(0, 12).map(([k, v]) => `<div class="mile-item">${k}<span>${v}</span></div>`).join('') : '<div class="mile-item">无</div>') + '</div>';
    });
    mh += '</div>';
    document.getElementById('tm-mile').innerHTML = mh;
  }

  // ---------------- 初始化（占位，选择器用内联 onclick 绑定） ----------------
  function init() { _inited = true; }

  // ---------------- 对外接口 ----------------
  return {
    setData, setYesterday, init,
    setCostDim, setCostMetric, setSlaPeriod,
    renderOverview, renderIntransit, renderSLA, renderAbnormal, renderCost, renderTomorrow
  };
})();
