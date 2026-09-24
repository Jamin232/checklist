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
  let slaDim = 'channel'; // 'channel' | 'agent' | 'customer' | 'logistic' | 'transport' | 'cat' | 'month'
  let slaFilters = { customer: '', channel: '', agent: '', country: '', transport: '', cat: '', month: '' };

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

  const STATUS_WORDS = ['查验中', '开查中', '索赔中', '赔付中']; // 异常状态词（宽口径）
  // 注意："开查中"是快递的开查（非海关查验），不计入查验进行中，但仍属异常单口径
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
    if (val instanceof Date) {
      // SheetJS(cellDates:true) 返回的是浏览器本地时区表示的 Date（如中国 8/5 0:00 存为 2026-08-04T16:00:00Z）。
      // 必须用本地 getter 读年月日，否则 UTC-8 下会把中国日期看早一天。
      return new Date(Date.UTC(val.getFullYear(), val.getMonth(), val.getDate()));
    }
    if (typeof val === 'number') {
      // Excel 序列号转时间戳本质为 UTC，保持 UTC getter
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
    m = s.match(/^(\d{1,2})月(\d{1,2})日?$/);
    if (m) return makeMDate(+m[1], +m[2]); // 月日（中文格式，无年）
    const d = new Date(s);
    return isNaN(d.getTime()) ? null : new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  }
  // 从货物状态文本中提取日期（如 "8/5查验中" / "8月5日查验中"），作为查验开始时间兜底。
  // 仅当状态含"查验"字样时才解析，避免把 "8/12运输中" 等普通状态误判为查验时间。
  function parseDateFromStatus(statusText) {
    const s = safeStr(statusText);
    if (!s || !s.includes('查验')) return null;
    let m = s.match(/^(\d{1,2})[-\/.](\d{1,2})/);
    if (m) return makeMDate(+m[1], +m[2]);
    m = s.match(/^(\d{1,2})月(\d{1,2})日?/);
    if (m) return makeMDate(+m[1], +m[2]);
    return null;
  }
  // 从状态备注中提取"查验发生日期"：定位含关键字的行，解析其中的日期（支持 年/月/日、月/日/年、月/日、月日 多种格式）。
  // 仅当备注含关键字时才解析，返回 UTC 零点日期（year>=2020），无日期返回 null。
  function extractInspectDate(remark, keyword) {
    const s = safeStr(remark);
    if (!s || !s.includes(keyword)) return null;
    const lines = s.split(/\r?\n/);
    for (const line of lines) {
      if (!line.includes(keyword)) continue;
      const m = line.match(/(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})|(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})|(\d{1,2})[-\/.](\d{1,2})|(\d{1,2})月(\d{1,2})日?/);
      if (!m) continue;
      let d = null;
      if (m[1]) d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
      else if (m[4]) d = new Date(Date.UTC(+m[6], +m[5] - 1, +m[4]));
      else if (m[7]) d = makeMDate(+m[7], +m[8]);
      else if (m[9]) d = makeMDate(+m[9], +m[10]);
      if (d && d.getUTCFullYear() >= 2020) return d;
    }
    return null;
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
    // 全程日期均为 UTC 零点构造，用 UTC getter 比较，避免本地时区跨日错位（如 UTC-8）
    return a && b && a.getUTCFullYear() === b.getUTCFullYear() && a.getUTCMonth() === b.getUTCMonth() && a.getUTCDate() === b.getUTCDate();
  }
  // 日期 d 是否在 [ref - n, ref] 的闭区间内（含当天及前 n 天）
  function withinLastNDays(d, ref, n) {
    if (!d || !ref) return false;
    const diff = dayDiff(ref, d);
    return diff >= 0 && diff <= n;
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
  // 美线海运船型：从「素芸物流渠道」名称关键字解析（价格梯度：左贵右廉）
  function shipTypeOf(ch) {
    if (!ch) return '';
    if (/美森正班/.test(ch)) return '美森正班';
    if (/美森/.test(ch)) return '美森加班';
    if (/快船/.test(ch)) return '快船';
    if (/普船/.test(ch)) return '普船';
    return '';
  }
  // 派送类型：海卡→卡派，海派→快递派
  function deliveryTypeOf(ch) {
    if (/海卡/.test(ch)) return '卡派';
    if (/海派/.test(ch)) return '快递派';
    return '';
  }
  function mean(arr) {
    const v = arr.filter(x => (+x || 0) > 0);
    return v.length ? v.reduce((a, b) => a + b, 0) / v.length : 0;
  }
  function transitOf(rs) { return rs.map(r => r.transitDays).filter(x => x > 0); }

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
      dom: find('国内', '查验'),
      dest: find('目的地', '查验'),
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
      ref: find('参考时效'),
      arr: find('到港日期'),
      pay: find('赔付'),
      late: find('物流最晚送达时间')
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
      // 国内外查验判定：以状态备注文本为准（对齐 chayan.js 真实口径）。
      // 原「国内查验时间/目的地查验时间」列实为时长字符串(如"1,23天")而非日期，不可用。
      const isDomInsp = remark.includes('国内查验');
      const isForeignInsp = remark.includes('国外查验');
      // 从备注中提取查验发生日期（用于"今日新增查验"/查验进行中耗时），无日期则为 null
      const domInspectDate = extractInspectDate(remark, '国内查验');
      const destInspectDate = extractInspectDate(remark, '国外查验');

      const isInspecting = goodsStatus.includes('查验中'); // 仅海关查验（不含快递"开查中"）
      const isAbnormal = STATUS_WORDS.some(w => goodsStatus.includes(w));

      // —— 派生字段（支撑 SLA多维 / 异常代理维度 / 延误分析 / 查验日期口径 / 美线海运 等模块）——
      const transport = safeStr(row[map.type]) || '未知';
      const logisticCh = safeStr(row[map.logi]);
      const shipType = shipTypeOf(logisticCh);
      const deliveryType = deliveryTypeOf(logisticCh);
      const arrivalDate = parseDate(row[map.arr]);
      const refLeadV = safeNum(row[map.ref]);
      const intransit = signTime === null;
      const transitDays = (!intransit && signTime && shipDate) ? Math.round(dayDiff(signTime, shipDate)) : null;
      const overRefLead = intransit && refLeadV > 0 && shipDate ? (dayDiff(TODAY, shipDate) > refLeadV) : false;
      const overdue = intransit && latestDeliver ? (TODAY > latestDeliver) : false;
      const onTime = !intransit ? (refLeadV > 0 ? (transitDays <= refLeadV ? '是' : '否') : '已完成') : (overRefLead ? '未完成已超时效' : '在途');
      const domInsp = isDomInsp;       // 备注含"国内查验"（起运港查验）
      const ovsInsp = isForeignInsp;   // 备注含"国外查验"（目的港查验）
      const isKaiCha = goodsStatus.includes('开查');
      const isSuoPei = goodsStatus.includes('索赔');
      const isLiPei = goodsStatus.includes('赔付') || /理赔/.test(goodsStatus) || /理赔|赔付/.test(remark);
      const lossClaim = isKaiCha || isSuoPei || isLiPei;
      const isDelayed = overRefLead;
      const bizMonth = shipDate ? fmtYM(shipDate) : '未知';
      const arrMonth = arrivalDate ? fmtYM(arrivalDate) : null;
      // 月归属口径：起运港查验→出货月；目的港查验→到港月（对齐用户规则）
      const domInspMonth = isDomInsp ? bizMonth : null;
      const destInspMonth = isForeignInsp ? arrMonth : null;

      list.push({
        mainTicket: safeStr(row[map.main]),
        tickets, ticketCount,
        type: safeStr(row[map.type]),
        transport,
        logisticChannel: logisticCh,
        shipType, deliveryType, arrivalDate,
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
        refLead: refLeadV,
        goodsStatus,
        remark,
        domInspectDate,
        destInspectDate,
        inTransit: intransit,
        isInspecting,
        isAbnormal,
        transitDays, overRefLead, overdue, onTime, domInsp, ovsInsp,
        isKaiCha, isSuoPei, isLiPei, lossClaim, isDelayed,
        bizMonth, domInspMonth, destInspMonth, arrMonth
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
    // 分享模式下不覆盖已注入的内容
    if (typeof _shareMode !== 'undefined' && _shareMode) return;
    const el = document.getElementById(htmlId);
    if (el) el.innerHTML = '<div style="padding:30px;text-align:center;color:#999">数据未加载（请确认 data.json 已生成，或用页面底部应急入口上传最新跟踪表）</div>';
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
    // 数据通常 T+1 生成，8/12 文件里 8/12 的查验记录极少；改为统计近两日（当日+昨日）新增
    const newInspect = todayRecs.filter(r =>
      (r.domInspectDate && withinLastNDays(r.domInspectDate, TODAY, 1)) ||
      (r.destInspectDate && withinLastNDays(r.destInspectDate, TODAY, 1))).length;

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
      { num: newInspect, label: '近两日新增查验', sub: `${fmtMD(new Date(Date.UTC(TODAY.getUTCFullYear(), TODAY.getUTCMonth(), TODAY.getUTCDate() - 1)))}-${fmtMD(TODAY)}`, cls: 'ov-new' },
      { num: overdue5, label: '在途超期>5天', sub: '未妥投', cls: 'ov-overdue' },
      { num: tom['到港'] + tom['清关'] + tom['派送'], label: '明日到港/清关/派送', sub: `到${tom['到港']}/清${tom['清关']}/派${tom['派送']}`, cls: 'ov-tom' }
    ];
    document.getElementById('ov-cards').innerHTML = cards.map(c => `
      <div class="kpi-card ${c.cls}">
        <div class="kpi-num">${c.num}</div>
        <div class="kpi-label">${c.label}</div>
        <div class="kpi-sub">${c.sub}</div>
      </div>`).join('');

    // 月度趋势：发运量 + 时效达标率（口径：仓库出货日期按月聚合；达标=已签收且实际时效≤参考时效）
    const mm = {};
    todayRecs.forEach(r => {
      if (!r.shipDate) return;
      const m = '' + r.shipDate.getUTCFullYear() + '-' + String(r.shipDate.getUTCMonth() + 1).padStart(2, '0');
      if (!mm[m]) mm[m] = { n: 0, signed: 0, val: 0, ot: 0 };
      mm[m].n++;
      if (!r.inTransit) {
        mm[m].signed++;
        if (r.refLead > 0) {
          mm[m].val++;
          if (dayDiff(r.signTime, r.shipDate) <= r.refLead) mm[m].ot++;
        }
      }
    });
    const months = Object.keys(mm).sort();
    if (months.length) {
      const vol = months.map(m => mm[m].n);
      const sg = months.map(m => mm[m].signed);
      const rate = months.map(m => { const v = mm[m]; return v.val ? +(v.ot / v.val * 100).toFixed(1) : null; });
      setOpt('ovTrend', {
        tooltip: { trigger: 'axis' },
        legend: { data: ['发运量', '已签收'], bottom: 0 },
        grid: { left: 52, right: 18, top: 18, bottom: 46 },
        xAxis: { type: 'category', data: months, axisLabel: { rotate: 30, fontSize: 10 } },
        yAxis: { type: 'value', name: '票' },
        series: [
          { name: '发运量', type: 'line', data: vol, smooth: true, showSymbol: false, areaStyle: { opacity: .15 }, itemStyle: { color: '#2563eb' } },
          { name: '已签收', type: 'line', data: sg, smooth: true, showSymbol: false, itemStyle: { color: '#10b981' } }
        ]
      });
      setOpt('ovRate', {
        tooltip: { trigger: 'axis' },
        grid: { left: 52, right: 18, top: 18, bottom: 46 },
        xAxis: { type: 'category', data: months, axisLabel: { rotate: 30, fontSize: 10 } },
        yAxis: { type: 'value', name: '%', max: 100 },
        series: [{ name: '时效达标率', type: 'line', data: rate, smooth: true, showSymbol: false, areaStyle: { opacity: .15 }, lineStyle: { width: 2 }, itemStyle: { color: '#06b6d4' } }]
      });
    } else {
      noData('ovTrend'); noData('ovRate');
    }

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
      extra += `<div class="ov-note" style="color:#888">单表快照模式：按需求已取消双日对比，日环比暂不显示。</div>`;
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
      if ((r.domInspectDate && withinLastNDays(r.domInspectDate, ref, 1)) || (r.destInspectDate && withinLastNDays(r.destInspectDate, ref, 1))) newInspect++;
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

    // —— ④ 新增：超时(超参考时效) by 渠道 + 在途账龄分布 + 未完成已超时效清单 ——
    const overList = inTransit.filter(r => r.overRefLead);
    const byChOver = {};
    overList.forEach(r => { const k = r.logisticChannel || '未知'; byChOver[k] = (byChOver[k] || 0) + 1; });
    const chOver = Object.entries(byChOver).sort((a, b) => b[1] - a[1]).slice(0, 12);
    setOpt('it-overdueChart', {
      tooltip: { trigger: 'axis' }, legend: { data: ['超参考时效票', '在途票'], bottom: 0 },
      grid: { left: 50, right: 20, top: 20, bottom: 60 }, xAxis: { type: 'category', data: chOver.map(x => x[0]), axisLabel: { interval: 0, rotate: 30, fontSize: 10 } },
      yAxis: { type: 'value', name: '票' },
      series: [
        { name: '超参考时效票', type: 'bar', data: chOver.map(x => x[1]), itemStyle: { color: '#ef4444' } },
        { name: '在途票', type: 'bar', data: chOver.map(x => inTransit.filter(r => r.logisticChannel === x[0]).length), itemStyle: { color: '#cbd5e1' } }
      ]
    });
    // 账龄分布（出货→今天）
    const ages = inTransit.map(r => r.shipDate ? dayDiff(TODAY, r.shipDate) : 0).filter(x => x > 0);
    const ageBuckets = ['0-15', '15-30', '30-45', '45-60', '60+']; const ageCnt = [0, 0, 0, 0, 0];
    ages.forEach(a => { if (a <= 15) ageCnt[0]++; else if (a <= 30) ageCnt[1]++; else if (a <= 45) ageCnt[2]++; else if (a <= 60) ageCnt[3]++; else ageCnt[4]++; });
    setOpt('it-ageChart', {
      tooltip: { trigger: 'axis' }, grid: { left: 50, right: 20, top: 20, bottom: 40 },
      xAxis: { type: 'category', data: ageBuckets }, yAxis: { type: 'value', name: '票' },
      series: [{ type: 'bar', data: ageCnt, itemStyle: { color: '#06b6d4' } }]
    });
    // 未完成已超时效清单
    const ovTb = document.getElementById('it-overdueTable');
    if (ovTb) {
      const list = overList.map(r => ({ r, over: (r.shipDate ? dayDiff(TODAY, r.shipDate) : 0) - (r.refLead || 0) })).filter(x => x.over > 0).sort((a, b) => b.over - a.over).slice(0, 100);
      ovTb.innerHTML = `<table class="data-table"><thead><tr><th>分出仓单号</th><th>客户(事业部)</th><th>渠道</th><th>代理</th><th>出货日期</th><th>在途天数</th><th>参考时效</th><th>超期天数</th></tr></thead><tbody>` +
        (list.length ? list.map(({ r, over }) => `<tr><td>${r.tickets.length > 1 ? r.tickets.slice(0, 2).join('<br>') + (r.tickets.length > 2 ? `<br><span style="color:#888;font-size:10px">+${r.tickets.length - 2}更多</span>` : '') : (r.tickets[0] || r.mainTicket)}</td><td>${r.customer}</td><td>${r.logisticChannel}</td><td>${r.agent}</td><td>${r.shipDate ? fmtDate(r.shipDate) : '—'}</td><td>${r.shipDate ? dayDiff(TODAY, r.shipDate) : '—'}</td><td>${r.refLead || '—'}</td><td class="rate-bad">${over}</td></tr>`).join('') : '<tr><td colspan="8" style="text-align:center;color:#999">无超时效在途单</td></tr>') +
        `</tbody></table>`;
    }

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
  // 通用 KPI 卡片行（供新增模块复用）
  function setKpiRow(id, arr) {
    const el = document.getElementById(id);
    if (!el) return;
    el.innerHTML = arr.map(c => `<div class="kpi-card ${c.cls || ''}"><div class="kpi-num">${c.num}</div><div class="kpi-label">${c.label}</div><div class="kpi-sub">${c.sub || ''}</div></div>`).join('');
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
  // 中位数（参考时效取同渠道单一规划值，避免平均产生小数）
  function median(arr) {
    if (!arr.length) return 0;
    const s = [...arr].sort((a, b) => a - b);
    const m = Math.floor(s.length / 2);
    return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
  }
  // 时效SLA专用异常界定：仅"查验相关"订单排除出建议时效基线
  // （开查中=快递开查、索赔中/赔付中=售后异常，均非海关查验延误，保留在基线）
  // 注：查验过的单即便后续已签收，因查验时间/备注仍带"查验"标记，仍剔除
  function isSlaAbnormal(r) {
    return r.goodsStatus.includes('查验中') ||
           r.remark.includes('查验') ||
           !!r.domInspectDate || !!r.destInspectDate;
  }

  function setSlaPeriod(p) { slaPeriod = p; document.querySelectorAll('.sla-period-btn').forEach(b => b.classList.toggle('active', b.dataset.period === p)); renderSLA(); }
  function setSlaDim(d) {
    if (!d) d = document.getElementById('slaDimSelect')?.value || 'channel';
    slaDim = d;
    const sel = document.getElementById('slaDimSelect');
    if (sel) sel.value = d;
    renderSLA();
  }
  function resetSlaDim() { setSlaDim('channel'); }

  function setSlaFilter(key, val) {
    if (slaFilters.hasOwnProperty(key)) slaFilters[key] = val || '';
    renderSLA();
  }
  function resetSlaFilters() {
    slaFilters = { customer: '', channel: '', agent: '', country: '', transport: '', cat: '', month: '' };
    ['slaFilterCustomer', 'slaFilterChannel', 'slaFilterAgent', 'slaFilterCountry', 'slaFilterTransport', 'slaFilterCat', 'slaFilterMonth'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.value = '';
    });
    renderSLA();
  }
  function buildSlaFilterOptions() {
    if (!todayRecs) return;
    const unique = fn => [...new Set(todayRecs.map(fn).filter(Boolean))].sort();
    const fill = (id, vals) => {
      const el = document.getElementById(id);
      if (!el) return;
      const cur = el.value;
      el.innerHTML = '<option value="">全部</option>' + vals.map(v => `<option value="${v}">${v}</option>`).join('');
      if (vals.includes(cur)) el.value = cur;
    };
    fill('slaFilterCustomer', unique(r => r.customer));
    fill('slaFilterChannel', unique(r => r.channelCategory));
    fill('slaFilterAgent', unique(r => r.agent));
    fill('slaFilterCountry', unique(r => r.country));
    fill('slaFilterTransport', unique(r => r.transport));
    fill('slaFilterCat', unique(r => r.channelCategory));
    fill('slaFilterMonth', unique(r => r.bizMonth));
  }

  function slaKeyFn(r) {
    switch (slaDim) {
      case 'agent': return r.agent || '未知';
      case 'customer': return r.customer || '未知';
      case 'logistic': return r.logisticChannel || '未知';
      case 'transport': return r.transport || '未知';
      case 'month': return r.bizMonth || '未知';
      case 'cat': return r.channelCategory || '未知';
      default: return r.channelCategory || r.logisticChannel || '未知';
    }
  }

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
    // 多维度筛选（事业部看板风格 + 代理维度）
    buildSlaFilterOptions();
    pool = pool.filter(r => {
      if (slaFilters.customer && r.customer !== slaFilters.customer) return false;
      if (slaFilters.channel && r.channelCategory !== slaFilters.channel) return false;
      if (slaFilters.agent && r.agent !== slaFilters.agent) return false;
      if (slaFilters.country && r.country !== slaFilters.country) return false;
      if (slaFilters.transport && r.transport !== slaFilters.transport) return false;
      if (slaFilters.cat && r.channelCategory !== slaFilters.cat) return false;
      if (slaFilters.month && r.bizMonth !== slaFilters.month) return false;
      return true;
    });
    // 时效SLA基线：仅剔除"查验相关"订单（状态查验中/备注含查验/查验时间有值）；开查中、索赔赔付保留
    const normalSigned = pool.filter(r => !r.inTransit && !isSlaAbnormal(r) && r.shipDate && r.signTime);
    const byCh = {};
    for (const r of normalSigned) {
      const k = slaKeyFn(r);
      if (!byCh[k]) byCh[k] = { leads: [], refLeads: [] };
      byCh[k].leads.push(Math.round(dayDiff(r.signTime, r.shipDate))); // 总时效=签收-仓库出货
      if (r.refLead > 0) byCh[k].refLeads.push(r.refLead);
    }
    const MIN_SAMPLE = 5;
    const arr = Object.entries(byCh).map(([k, v]) => {
      const n = v.leads.length;
      const avgLead = n ? v.leads.reduce((a, b) => a + b, 0) / n : 0;
      const p95 = percentile(v.leads, 0.95);               // P95 = 建议时效（95%正常单可达成）
      const p90 = percentile(v.leads, 0.9);                // P90 用于参考对比
      const suggested = n >= MIN_SAMPLE ? p95 : avgLead;   // 建议时效 = P95（直观：95%的正常单在此时间内签收）
      const refAvg = v.refLeads.length ? median(v.refLeads) : 0;
      const over = v.leads.filter(l => l > suggested).length;
      const rate = n ? (n - over) / n * 100 : 0;            // SLA达成率（以建议时效为承诺基线）
      return { key: k, n, avgLead, p90, p95, suggested, refAvg, over, rate, small: n < MIN_SAMPLE };
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
        { name: '参考时效(Z列)', type: 'bar', data: arr.map(x => +x.refAvg), itemStyle: { color: '#e08e0b' } },
        { name: 'SLA达成率', type: 'line', yAxisIndex: 1, data: arr.map(x => +x.rate.toFixed(1)), itemStyle: { color: '#d53f8c' }, symbolSize: 7 }
      ]
    });

    // 表格
    const dimNames = { channel: '渠道大类', agent: '代理', customer: '事业部', logistic: '素芸渠道', transport: '运输方式', cat: '国家+类型', month: '月份' };
    const dimName = dimNames[slaDim] || '维度';
    const tb = document.getElementById('sla-table');
    if (tb) {
      const dimLabel = slaPeriod === 'all' ? '全量' : { halfmonth: '近半月', month: '近一月', twomonth: '近两月' }[slaPeriod];
      tb.innerHTML = `<tr><th>${dimName}</th><th>正常已签收</th><th>平均实际时效</th><th>建议时效*</th><th>参考时效(Z)</th><th>超时单数</th><th>SLA达成率</th></tr>` +
        arr.map(x => {
          const rc = x.rate >= 90 ? 'rate-good' : (x.rate >= 80 ? 'rate-mid' : 'rate-bad');
          return `<tr><td>${x.key}</td><td>${x.n}${x.small ? '<span class="pct">样本少</span>' : ''}</td><td>${x.avgLead.toFixed(1)}</td><td><b>${x.suggested.toFixed(1)}</b></td><td>${x.refAvg ? x.refAvg : '-'}</td><td class="rate-bad">${x.over}</td><td class="${rc}">${x.rate.toFixed(1)}%</td></tr>`;
        }).join('') +
        `<tr style="font-weight:700;background:#f3f6fa"><td>合计</td><td>${normalSigned.length}</td><td colspan="5"></td></tr>` +
        `<tr><td colspan="7" style="color:#888;font-size:11px">*建议时效 = P95（95%分位值，即95%的正常已签收单在此天数内完成，已剔除"查验相关"订单：状态查验中/备注含查验/查验时间有值；开查中、索赔赔付保留）。<br>含义：若将此渠道的参考时效设为"建议时效(P95)"，则约 ${arr.length > 0 ? Math.round(arr.filter(x=>x.rate>=95).length/arr.length*100) : 0}% 的渠道可达95%+达成率。<br>数据范围：${dimLabel} | 时效=实际签收时间−仓库出货日期</td></tr>`;
    }
  }

  // ---------------- 筛选状态 ----------------
  let abnCustomerFilter = '全部'; // 异常/预警 客户筛选
  let tmCustomerFilter = '全部';   // 预警 客户筛选

  function setAbnCustomer(c) {
    if (typeof _shareMode !== 'undefined' && _shareMode) { if (typeof showToast === 'function') showToast('分享视图为只读快照，客户筛选不可用，请用原系统查看', 'info'); return; }
    abnCustomerFilter = c;
    document.querySelectorAll('.abn-cust-btn').forEach(b => b.classList.toggle('active', b.dataset.cust === c));
    renderAbnormal();
  }
  function setTmCustomer(c) {
    if (typeof _shareMode !== 'undefined' && _shareMode) { if (typeof showToast === 'function') showToast('分享视图为只读快照，客户筛选不可用，请用原系统查看', 'info'); return; }
    tmCustomerFilter = c;
    document.querySelectorAll('.tm-cust-btn').forEach(b => b.classList.toggle('active', b.dataset.cust === c));
    renderTomorrow();
  }

  // ============================================================
  // ④ 异常监控
  // ============================================================
  // 异常定义对齐：
  //   isInspecting = 货物状态含"查验中"（海关查验进行中，不含快递"开查中"）
  //   isPureAbnormal = 货物状态含"索赔中"或"赔付中"（已进入索赔/赔付流程，不含查验中）
  //   isAbnormal(宽口径) = isInspecting ∪ isPureAbnormal ∪ 开查中 = 含 查验中/开查中/索赔中/赔付中
  //   KPI 卡片展示：查验进行中(isInspecting) + 纯异常单(isPureAbnormal) —— 不重复计数
  function renderAbnormal() {
    if (!todayRecs) { noData('ab-cards'); noData('ab-table'); return; }
    const inspecting = todayRecs.filter(r => r.isInspecting);
    // 开查中 = 快递开查（非海关查验、非索赔赔付）
    const kaicha = todayRecs.filter(r =>
      r.goodsStatus.includes('开查中') &&
      !r.isInspecting &&
      !(r.goodsStatus.includes('索赔中') || r.goodsStatus.includes('赔付中')));
    // 纯异常 = 仅索赔中/赔付中（不含查验中、不含快递"开查中"，避免与查验进行中重复）
    const pureAbnormal = todayRecs.filter(r =>
      !r.isInspecting &&
      !r.goodsStatus.includes('开查中') &&
      (r.goodsStatus.includes('索赔中') || r.goodsStatus.includes('赔付中')));
    const newInspect = todayRecs.filter(r =>
      (r.domInspectDate && withinLastNDays(r.domInspectDate, TODAY, 1)) ||
      (r.destInspectDate && withinLastNDays(r.destInspectDate, TODAY, 1)));

    let deltaHtml = '';
    if (yesterdayRecs) {
      const t = dailyMetrics(todayRecs, TODAY), y = dailyMetrics(yesterdayRecs, yesterdayDate || TODAY);
      const mk = (cur, prev, label) => {
        const d = cur - prev;
        return `<div class="ab-delta-item">${label}<b class="${d >= 0 ? 'up' : 'down'}">${d >= 0 ? '+' : ''}${d}</b><span>较昨日</span></div>`;
      };
      // 去掉"查验进行中 较昨日"——与"近两日新增查验"信息重复
      deltaHtml = `<div class="ab-delta">${mk(t.abnormal - t.inspecting, y.abnormal - y.inspecting, '异常单(索赔/赔付)')}${mk(t.newInspect, y.newInspect, '近两日新增查验')}</div>`;
    } else {
      deltaHtml = '<div class="ov-note" style="color:#888">单表快照模式：按需求已取消双日对比，日环比暂不显示。</div>';
    }

    document.getElementById('ab-cards').innerHTML =
      '<div style="display:flex;gap:10px;flex-wrap:wrap">' +
      `<div class="kpi-card ov-abn"><div class="kpi-num">${inspecting.length}</div><div class="kpi-label">查验进行中</div></div>` +
      `<div class="kpi-card" style="--card-color:#7c3aed"><div class="kpi-num" style="color:#7c3aed">${kaicha.length}</div><div class="kpi-label">开查中</div></div>` +
      `<div class="kpi-card ov-overdue"><div class="kpi-num">${pureAbnormal.length}</div><div class="kpi-label">索赔/赔付</div></div>` +
      `<div class="kpi-card ov-new"><div class="kpi-num">${newInspect.length}</div><div class="kpi-label">近两日新增查验(${fmtMD(new Date(Date.UTC(TODAY.getUTCFullYear(), TODAY.getUTCMonth(), TODAY.getUTCDate() - 1)))}-${fmtMD(TODAY)})</div></div>` +
      '</div>' + deltaHtml;

    // 明细表（异常单）— 支持客户筛选 + 查验持续天数
    const tb = document.getElementById('ab-table');
    if (tb) {
      // 构建客户筛选选项（横排 flex-wrap）
      const allCustomers = [...new Set(todayRecs.filter(r => r.isAbnormal).map(r => r.customer).filter(Boolean))].sort();
      let filterBar = '<div class="cust-filter-bar">' +
        '<span class="cust-filter-label">客户筛选：</span>' +
        `<span class="cust-btn abn-cust-btn${abnCustomerFilter === '全部' ? ' active' : ''}" data-cust="全部" onclick="Daily.setAbnCustomer('全部')">全部</span>`;
      allCustomers.slice(0, 20).forEach(cu => {
        const active = abnCustomerFilter === cu;
        filterBar += `<span class="cust-btn abn-cust-btn${active ? ' active' : ''}" data-cust="${cu}" onclick="Daily.setAbnCustomer('${cu.replace(/'/g, "\\'")}')">${cu}</span>`;
      });
      filterBar += '</div>';

      // 筛选后的行：宽口径异常(isAbnormal) + 客户过滤
      let rows = todayRecs.filter(r => r.isAbnormal &&
        (abnCustomerFilter === '全部' || r.customer === abnCustomerFilter))
        .slice(0, 300).map(r => {
          const t = r.isInspecting ? '查验中'
            : (r.goodsStatus.includes('开查中') ? '开查中'
            : (r.goodsStatus.match(/索赔中|赔付中/) ? r.goodsStatus.match(/索赔中|赔付中/)[0] : '异常'));
          const ticketDisplay = r.tickets.length > 1 ? r.tickets.slice(0, 3).join('<br>') + (r.tickets.length > 3 ? `<br><span style="color:#888;font-size:10px">+${r.tickets.length - 3}更多</span>` : '') : (r.tickets[0] || r.mainTicket);
          // 查验持续天数：仅查验中订单计算 = TODAY - min(国内查验时间, 目的地查验时间)
          // 合理性校验：>365天视为异常数据（日期解析错误等），显示 "-"
          let inspectDays = '';
          if (r.isInspecting && (r.domInspectDate || r.destInspectDate)) {
            const inspectStart = r.domInspectDate && r.destInspectDate
              ? (r.domInspectDate < r.destInspectDate ? r.domInspectDate : r.destInspectDate)
              : (r.domInspectDate || r.destInspectDate);
            const dd = dayDiff(TODAY, inspectStart);
            inspectDays = (dd > 0 && dd <= 365) ? `<td class="rate-bad">${dd}天</td>` : '<td>-</td>';
          } else {
            inspectDays = '<td>-</td>';
          }
          return `<tr>${inspectDays}<td>${ticketDisplay}</td><td>${r.logisticChannel}</td><td>${r.agent}</td><td>${r.customer}</td><td>${r.country}</td><td class="status-inspected">${t}</td><td>${r.goodsStatus}</td></tr>`;
        }).join('');
      tb.innerHTML = filterBar +
        `<table class="data-table"><thead><tr><th>查验天数</th><th>分出仓单号</th><th>渠道</th><th>代理</th><th>客户</th><th>国家</th><th>状态</th><th>货物状态</th></tr></thead>` +
        `<tbody>${rows || '<tr><td colspan="8" style="text-align:center;color:#999">无匹配记录</td></tr>'}</tbody></table>`;
    }

    // —— ⑥ 新增：异常结果呈现（理赔/开查/索赔占比）+ 代理维度 ——
    const abDom = inspecting.length;                       // 查验中
    const abKai = kaicha.length;                           // 开查中
    const abSuo = todayRecs.filter(r => r.isSuoPei).length; // 索赔中
    const abLi = todayRecs.filter(r => r.isLiPei).length;   // 赔付/理赔
    setOpt('ab-resultChart', {
      tooltip: { trigger: 'item', formatter: '{b}: {c} ({d}%)' },
      legend: { type: 'scroll', bottom: 0 },
      series: [{ type: 'pie', radius: ['40%', '66%'], center: ['50%', '44%'], data: [
        { name: '查验中', value: abDom, itemStyle: { color: '#2563eb' } },
        { name: '开查中', value: abKai, itemStyle: { color: '#8b5cf6' } },
        { name: '索赔中', value: abSuo, itemStyle: { color: '#ef4444' } },
        { name: '赔付/理赔', value: abLi, itemStyle: { color: '#f59e0b' } }
      ], label: { formatter: '{b}\n{c}' } }]
    });
    // 代理维度：异常票(宽口径) / 总票 / 异常率 / 其中赔付索赔开查
    const agMap = {};
    todayRecs.forEach(r => {
      const a = r.agent || '未知';
      if (!agMap[a]) agMap[a] = { total: 0, abn: 0, loss: 0 };
      agMap[a].total++;
      if (r.isAbnormal) agMap[a].abn++;
      if (r.lossClaim) agMap[a].loss++;
    });
    const agArr = Object.entries(agMap).map(([k, v]) => ({ k, ...v })).sort((a, b) => b.abn - a.abn).slice(0, 15);
    const agTb = document.getElementById('ab-agentTable');
    if (agTb) {
      agTb.innerHTML = `<table class="data-table"><thead><tr><th>代理</th><th>总票</th><th>异常票</th><th>异常率</th><th>其中赔付/索赔/开查</th></tr></thead><tbody>` +
        agArr.map(x => `<tr><td>${x.k}</td><td>${x.total}</td><td>${x.abn}</td><td class="${x.total ? (x.abn / x.total * 100 >= 10 ? 'rate-bad' : 'rate-mid') : ''}">${x.total ? (x.abn / x.total * 100).toFixed(1) + '%' : '—'}</td><td>${x.loss}</td></tr>`).join('') +
        `</tbody></table>`;
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

    // 构建客户筛选选项（从在途超期单中提取）
    const allTmCustomers = [...new Set(todayRecs.filter(r => r.inTransit && r.latestDeliver).map(r => r.customer).filter(Boolean))].sort();

    // Part A：在途超期（按梯度）— 每个梯度独立展示，Tab式分离 + 客户筛选
    const tiers = [
      { key: '≥10天', test: d => d >= 10, cls: 'tier-10', desc: '严重超期，需立即跟进' },
      { key: '≥7天', test: d => d >= 7 && d < 10, cls: 'tier-7', desc: '明显超期，关注处理' },
      { key: '>5天', test: d => d > 5 && d < 7, cls: 'tier-5', desc: '轻度超期，持续观察' }
    ];
    const list = todayRecs.filter(r => r.inTransit && r.latestDeliver);
    // 应用客户筛选
    const filteredList = tmCustomerFilter === '全部' ? list : list.filter(r => r.customer === tmCustomerFilter);
    const overdueRows = filteredList.map(r => ({ r, d: dayDiff(TODAY, r.latestDeliver) })).filter(x => x.d > 0);

    // 客户筛选栏（横排 flex-wrap）
    let custFilterHtml = '<div class="cust-filter-bar">' +
      '<span class="cust-filter-label">客户筛选：</span>' +
      `<span class="cust-btn tm-cust-btn${tmCustomerFilter === '全部' ? ' active' : ''}" data-cust="全部" onclick="Daily.setTmCustomer('全部')">全部</span>`;
    allTmCustomers.slice(0, 20).forEach(cu => {
      const active = tmCustomerFilter === cu;
      custFilterHtml += `<span class="cust-btn tm-cust-btn${active ? ' active' : ''}" data-cust="${cu}" onclick="Daily.setTmCustomer('${cu.replace(/'/g, "\\'")}')">${cu}</span>`;
    });
    custFilterHtml += '</div>';

    // 构建每个梯度的独立卡片+表格
    let oh = '<div class="tm-tier-group">' + custFilterHtml;
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

  // 分享链接：导出当前所有图表为图片 dataURL（canvas 内容无法随 innerHTML 序列化）
  function getShareCharts() {
    const out = {};
    for (const [id, c] of Object.entries(chartMap)) {
      if (c && typeof c.getDataURL === 'function' && !c.isDisposed()) {
        try { out[id] = c.getDataURL({ type: 'jpeg', pixelRatio: 1, backgroundColor: '#fff' }); } catch (e) {}
      }
    }
    return out;
  }
  function resizeAllCharts() {
    Object.values(chartMap).forEach(c => { if (c && !c.isDisposed()) { try { c.resize(); } catch (e) {} } });
  }

  // ============================================================
  // ⑦ 延误分析（沿用事业部口径）
  // ============================================================
  function renderDelayAnalysis() {
    if (!todayRecs) { ['dl-kpis', 'dl-hist', 'dl-chan', 'dl-blank', 'dl-impactChart', 'dl-impactTable'].forEach(noData); return; }
    const over = todayRecs.filter(r => r.overRefLead);   // 超参考时效未签收
    const overDays = over.map(r => (r.shipDate ? dayDiff(TODAY, r.shipDate) : 0) - (r.refLead || 0)).filter(x => x > 0);
    const avgOver = overDays.length ? overDays.reduce((a, b) => a + b, 0) / overDays.length : 0;
    setKpiRow('dl-kpis', [
      { num: over.length, label: '超参考时效未签收', sub: '在途且已超承诺时效', cls: 'ov-overdue' },
      { num: over.length, label: '待填延误类型', sub: '监控无"延误类型"列，全部列为待跟进', cls: 'ov-new' },
      { num: avgOver.toFixed(1) + '天', label: '平均超期天数', sub: '在途天数−参考时效', cls: 'ov-intransit' },
      { num: todayRecs.filter(r => r.inTransit).length, label: '在途总票数', sub: '含未超期', cls: 'ov-abn' }
    ]);
    // 超期天数分布
    const buckets = ['1-5', '6-10', '11-20', '21-30', '30+']; const cnt = [0, 0, 0, 0, 0];
    overDays.forEach(d => { if (d <= 5) cnt[0]++; else if (d <= 10) cnt[1]++; else if (d <= 20) cnt[2]++; else if (d <= 30) cnt[3]++; else cnt[4]++; });
    setOpt('dl-hist', { tooltip: { trigger: 'axis' }, grid: { left: 50, right: 20, top: 20, bottom: 40 }, xAxis: { type: 'category', data: buckets }, yAxis: { type: 'value', name: '票' }, series: [{ type: 'bar', data: cnt, itemStyle: { color: '#ef4444' } }] });
    // 按渠道大类
    const byCh = {}; over.forEach(r => { const k = r.channelCategory || r.logisticChannel || '未知'; byCh[k] = (byCh[k] || 0) + 1; });
    const chArr = Object.entries(byCh).sort((a, b) => b[1] - a[1]).slice(0, 12);
    setOpt('dl-chan', { tooltip: { trigger: 'axis' }, grid: { left: 50, right: 20, top: 20, bottom: 60 }, xAxis: { type: 'category', data: chArr.map(x => x[0]), axisLabel: { interval: 0, rotate: 30, fontSize: 10 } }, yAxis: { type: 'value', name: '票' }, series: [{ type: 'bar', data: chArr.map(x => x[1]), itemStyle: { color: '#f59e0b' } }] });
    // 明细清单（Top 80 by 超期天数）
    const tb = document.getElementById('dl-blank');
    if (tb) {
      const list = over.map(r => ({ r, over: (r.shipDate ? dayDiff(TODAY, r.shipDate) : 0) - (r.refLead || 0) })).filter(x => x.over > 0).sort((a, b) => b.over - a.over).slice(0, 80);
      tb.innerHTML = `<table class="data-table"><thead><tr><th>分出仓单号</th><th>客户(事业部)</th><th>渠道</th><th>代理</th><th>国家</th><th>出货日期</th><th>在途天数</th><th>参考时效</th><th>超期天数</th></tr></thead><tbody>` +
        (list.length ? list.map(({ r, over }) => `<tr><td>${r.tickets.length > 1 ? r.tickets.slice(0, 2).join('<br>') : (r.tickets[0] || r.mainTicket)}</td><td>${r.customer}</td><td>${r.logisticChannel}</td><td>${r.agent}</td><td>${r.country}</td><td>${r.shipDate ? fmtDate(r.shipDate) : '—'}</td><td>${r.shipDate ? dayDiff(TODAY, r.shipDate) : '—'}</td><td>${r.refLead || '—'}</td><td class="rate-bad">${over}</td></tr>`).join('') : '<tr><td colspan="9" style="text-align:center;color:#999">无超时效在途单</td></tr>') +
        `</tbody></table>`;
    }
    // 查验对延误的影响（仅已签收+含参考时效）
    const signed = todayRecs.filter(r => !r.inTransit && r.refLead > 0 && r.transitDays > 0);
    const dimOf = r => { if (r.domInsp && r.ovsInsp) return '国内+国外'; if (r.domInsp) return '仅起运港'; if (r.ovsInsp) return '仅目的港'; return '无查验'; };
    const dims = ['无查验', '仅起运港', '仅目的港', '国内+国外'];
    const stats = dims.map(dm => { const rs = signed.filter(r => dimOf(r) === dm); const n = rs.length; const at = n ? rs.reduce((a, b) => a + b.transitDays, 0) / n : 0; const ar = n ? rs.reduce((a, b) => a + b.refLead, 0) / n : 0; return { dm, n, at, ar, ex: at - ar }; });
    const base = stats[0].ex;
    const tb2 = document.getElementById('dl-impactTable');
    if (tb2) tb2.innerHTML = `<table class="data-table"><thead><tr><th>查验维度</th><th>票数</th><th>平均全程时效</th><th>平均参考时效</th><th>平均超期</th><th>查验净拖累</th></tr></thead><tbody>` +
      stats.map(s => `<tr><td>${s.dm}</td><td>${s.n}</td><td>${s.at.toFixed(1)}天</td><td>${s.ar.toFixed(1)}天</td><td class="${s.ex > 0 ? 'rate-bad' : 'rate-good'}">${s.ex >= 0 ? '+' : ''}${s.ex.toFixed(1)}天</td><td class="${(s.ex - base) > 0 ? 'rate-bad' : 'rate-good'};font-weight:600">${s.dm === '无查验' ? '基准' : ((s.ex - base) >= 0 ? '+' : '') + (s.ex - base).toFixed(1) + '天'}</td></tr>`).join('') +
      `</tbody></table>`;
    setOpt('dl-impactChart', { tooltip: { trigger: 'axis' }, legend: { data: ['平均全程时效(天)', '平均参考时效(天)'], bottom: 0 }, grid: { left: 50, right: 20, top: 20, bottom: 40 }, xAxis: { type: 'category', data: dims }, yAxis: { type: 'value', name: '天' }, series: [{ name: '平均全程时效(天)', type: 'bar', data: stats.map(s => +s.at.toFixed(1)), itemStyle: { color: '#ef4444' } }, { name: '平均参考时效(天)', type: 'bar', data: stats.map(s => +s.ar.toFixed(1)), itemStyle: { color: '#cbd5e1' } }] });
  }

  // ============================================================
  // ⑧ 按查验发生日期口径的查验率统计
  // ============================================================
  function renderInspByDate() {
    if (!todayRecs) { ['id-kpis', 'id-trend', 'id-table'].forEach(noData); return; }
    const rows = todayRecs;
    // 国内外查验以状态备注文本判定（对齐 chayan.js 真实口径，不依赖损坏的时长列）
    const isDom = r => r.domInsp;   // 备注含"国内查验"
    const isFor = r => r.ovsInsp;   // 备注含"国外查验"
    const months = new Set();
    let domTotal = 0, forTotal = 0, arrTotal = 0, shipTotal = 0;
    rows.forEach(r => {
      if (r.bizMonth && r.bizMonth !== '未知') { months.add(r.bizMonth); shipTotal++; if (isDom(r)) domTotal++; }
      if (r.arrMonth) { months.add(r.arrMonth); arrTotal++; if (isFor(r)) forTotal++; }
    });
    const ms = [...months].sort();
    // 口径：起运港查验率分母=仓库出货日期当月；目的港查验率分母=到港日期当月
    const agg = ms.map(m => {
      const shipped = rows.filter(r => r.bizMonth === m).length;
      const arr = rows.filter(r => r.arrMonth === m).length;
      const dom = rows.filter(r => r.domInsp && r.bizMonth === m).length;
      const dest = rows.filter(r => r.ovsInsp && r.arrMonth === m).length;
      return { m, shipped, arr, dom, dest, domRate: shipped ? +(dom / shipped * 100).toFixed(1) : null, ovsRate: arr ? +(dest / arr * 100).toFixed(1) : null };
    });
    setKpiRow('id-kpis', [
      { num: domTotal, label: '国内(起运港)查验', sub: '备注含"国内查验"累计', cls: 'ov-intransit' },
      { num: forTotal, label: '国外(目的港)查验', sub: '备注含"国外查验"累计', cls: 'ov-abn' },
      { num: ms.length, label: '覆盖月份', sub: ms.length ? ms[0] + ' ~ ' + ms[ms.length - 1] : '—', cls: 'ov-new' },
      { num: shipTotal, label: '有出货日期票', sub: '起运港分母口径', cls: 'ov-overdue' }
    ]);
    setOpt('id-trend', {
      tooltip: { trigger: 'axis' }, legend: { data: ['起运港查验率%', '目的港查验率%'], bottom: 0 },
      grid: { left: 50, right: 20, top: 20, bottom: 60 }, xAxis: { type: 'category', data: ms, axisLabel: { rotate: 30, fontSize: 10 } },
      yAxis: { type: 'value', name: '%', max: 100 },
      series: [
        { name: '起运港查验率%', type: 'line', data: agg.map(a => a.domRate), smooth: true, connectNulls: true, itemStyle: { color: '#2563eb' } },
        { name: '目的港查验率%', type: 'line', data: agg.map(a => a.ovsRate), smooth: true, connectNulls: true, itemStyle: { color: '#f59e0b' } }
      ]
    });
    const tb = document.getElementById('id-table');
    if (tb) tb.innerHTML = `<table class="data-table"><thead><tr><th>月份</th><th>出货票</th><th>国内查验</th><th>起运港率</th><th>到港票</th><th>国外查验</th><th>目的港率</th></tr></thead><tbody>` +
      agg.map(a => `<tr><td>${a.m}</td><td>${a.shipped}</td><td>${a.dom}</td><td class="${a.domRate >= 5 ? 'rate-bad' : a.domRate >= 3 ? 'rate-mid' : 'rate-good'}">${a.domRate != null ? a.domRate + '%' : '—'}</td><td>${a.arr}</td><td>${a.dest}</td><td class="${a.ovsRate >= 5 ? 'rate-bad' : a.ovsRate >= 3 ? 'rate-mid' : 'rate-good'}">${a.ovsRate != null ? a.ovsRate + '%' : '—'}</td></tr>`).join('') +
      `<tr style="font-weight:700;background:#f3f6fa"><td>合计</td><td>${shipTotal}</td><td>${domTotal}</td><td>${shipTotal ? +(domTotal / shipTotal * 100).toFixed(1) + '%' : '—'}</td><td>${arrTotal}</td><td>${forTotal}</td><td>${arrTotal ? +(forTotal / arrTotal * 100).toFixed(1) + '%' : '—'}</td></tr></tbody></table>`;
  }

  // ============================================================
  // ⑨ 美线海运时效对比
  // ============================================================
  const SHIPS = ['美森正班', '美森加班', '快船', '普船'];
  const SHIP_TIER = { '美森正班': 0, '美森加班': 1, '快船': 2, '普船': 3 };
  const SHIP_COLORS = ['#2563eb', '#f59e0b', '#8b5cf6', '#06b6d4'];
  const DELIVS = ['卡派', '快递派'];
  function renderUsOcean() {
    if (!todayRecs) { ['us-kpis', 'us-trend', 'us-bar', 'us-table', 'us-tips'].forEach(noData); return; }
    const base = todayRecs.filter(r => r.country === '美国' && r.shipType && r.deliveryType);
    const signed = r => !r.inTransit && r.transitDays > 0;
    const months = [...new Set(base.map(r => r.bizMonth))].filter(m => m && m !== '未知').sort();
    const shipsPresent = SHIPS.filter(s => base.some(r => r.shipType === s));
    setKpiRow('us-kpis', [{ num: base.length, label: '美线海运票', sub: '已识别船型+派送', cls: 'ov-intransit' }]
      .concat(shipsPresent.map(s => ({ num: base.filter(r => r.shipType === s).length, label: s, cls: ['ov-intransit', 'ov-new', 'ov-abn', 'ov-overdue'][SHIP_TIER[s]] }))));
    // 月度趋势（4 船型）
    const dsTrend = shipsPresent.map(s => {
      const col = SHIP_COLORS[SHIP_TIER[s]];
      const data = months.map(m => { const rs = base.filter(r => r.shipType === s && r.bizMonth === m).filter(signed); return rs.length >= 5 ? +mean(transitOf(rs)).toFixed(1) : null; });
      return { label: s, data, borderColor: col, backgroundColor: col + '22', smooth: true, connectNulls: true };
    });
    setOpt('us-trend', { tooltip: { trigger: 'axis' }, legend: { data: shipsPresent, bottom: 0 }, grid: { left: 50, right: 20, top: 20, bottom: 50 }, xAxis: { type: 'category', data: months, axisLabel: { rotate: 30, fontSize: 10 } }, yAxis: { type: 'value', name: '平均时效(天)' }, series: dsTrend });
    // 船型 × 派送类型 分组柱
    const dsBar = DELIVS.map((d, i) => ({ name: d, type: 'bar', data: shipsPresent.map(s => { const rs = base.filter(r => r.shipType === s && r.deliveryType === d).filter(signed); return rs.length >= 5 ? +mean(transitOf(rs)).toFixed(1) : null; }), itemStyle: { color: SHIP_COLORS[i] } }));
    setOpt('us-bar', { tooltip: { trigger: 'axis' }, legend: { data: DELIVS, bottom: 0 }, grid: { left: 50, right: 20, top: 20, bottom: 40 }, xAxis: { type: 'category', data: shipsPresent }, yAxis: { type: 'value', name: '平均时效(天)' }, series: dsBar });
    // 全景表：船型 × 派送类型
    let rowsArr = [];
    shipsPresent.forEach(s => DELIVS.forEach(d => {
      const rs = base.filter(r => r.shipType === s && r.deliveryType === d);
      if (!rs.length) return;
      const tr = rs.filter(signed);
      const va = rs.filter(r => r.refLead > 0);
      const y = va.filter(r => r.onTime === '是').length;
      rowsArr.push({ s, d, n: rs.length, avg: tr.length ? mean(transitOf(tr)) : 0, med: median(transitOf(tr)), rate: va.length ? (y / va.length * 100) : null });
    }));
    rowsArr.sort((a, b) => SHIP_TIER[a.s] - SHIP_TIER[b.s] || DELIVS.indexOf(a.d) - DELIVS.indexOf(b.d));
    const tb = document.getElementById('us-table');
    if (tb) tb.innerHTML = `<table class="data-table"><thead><tr><th>船型</th><th>派送类型</th><th>票数</th><th>平均时效(天)</th><th>中位时效(天)</th><th>时效达标率</th></tr></thead><tbody>` +
      (rowsArr.length ? rowsArr.map(r => `<tr><td>${r.s}</td><td>${r.d}</td><td>${r.n}</td><td>${r.avg ? r.avg.toFixed(1) : '—'}</td><td>${r.med ? r.med.toFixed(1) : '—'}</td><td class="${r.rate >= 90 ? 'rate-good' : r.rate >= 80 ? 'rate-mid' : 'rate-bad'}">${r.rate != null ? r.rate.toFixed(1) + '%' : '—'}</td></tr>`).join('') : '<tr><td colspan="6" style="text-align:center;color:#999">无匹配（需国家=美国且渠道含海卡/海派+船型关键字）</td></tr>') +
      `</tbody></table>`;
    // 降档优化机会（时效差≤3天）
    const tips = [];
    DELIVS.forEach(d => {
      const present = SHIPS.filter(s => { const rs = base.filter(r => r.shipType === s && r.deliveryType === d).filter(signed); return rs.length >= 30; });
      present.forEach(sA => present.forEach(sB => {
        if (SHIP_TIER[sA] < SHIP_TIER[sB]) {
          const a = mean(transitOf(base.filter(r => r.shipType === sA && r.deliveryType === d).filter(signed)));
          const b = mean(transitOf(base.filter(r => r.shipType === sB && r.deliveryType === d).filter(signed)));
          const gap = Math.abs(a - b);
          if (gap <= 3) tips.push(`【${d}】<b>${sA}</b> 均时效 ${a.toFixed(1)}天 ≈ <b>${sB}</b> ${b.toFixed(1)}天（仅差 ${gap.toFixed(1)}天）→ 可评估改走更便宜的 ${sB} 降运费`);
        }
      }));
    });
    document.getElementById('us-tips').innerHTML = tips.length ? tips.map(t => '• ' + t).join('<br>') : '当前筛选下各船型时效差均 &gt;3 天，暂无明显可降档空间（或样本不足）。';
  }

  // ---------------- 对外接口 ----------------
  return {
    setData, setYesterday, init,
    setCostDim, setCostMetric, setSlaPeriod, setSlaDim, resetSlaDim, setSlaFilter, resetSlaFilters, buildSlaFilterOptions,
    setAbnCustomer, setTmCustomer,
    renderOverview, renderIntransit, renderSLA, renderAbnormal, renderCost, renderTomorrow,
    renderDelayAnalysis, renderInspByDate, renderUsOcean,
    getShareCharts, resizeAllCharts
  };
})();
