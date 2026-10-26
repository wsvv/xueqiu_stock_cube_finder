// ==UserScript==
// @name         雪球 · 按股票找持仓组合
// @namespace    https://xueqiu.com/
// @version      1.7.4
// @updateURL    https://raw.githubusercontent.com/wsvv/xueqiu_stock_cube_finder/main/xueqiu_stock_cube_finder.user.js
// @downloadURL  https://raw.githubusercontent.com/wsvv/xueqiu_stock_cube_finder/main/xueqiu_stock_cube_finder.user.js
// @description  按股票反查组合，支持全量排序、股票结果24小时缓存、永久关停缓存；单条/全部同步更新调仓与核验真实持仓、独立缓存；适中宽度的紧凑表格、右侧操作入口始终可见，并导出 XLSX。
// @author       ChatGPT
// @match        https://xueqiu.com/*
// @match        https://www.xueqiu.com/*
// @run-at       document-idle
// @noframes
// @grant        GM_xmlhttpRequest
// @grant        GM_registerMenuCommand
// @grant        GM_getValue
// @grant        GM_setValue
// @connect      xueqiu.com
// @connect      www.xueqiu.com
// ==/UserScript==

(function () {
  'use strict';

  if (document.getElementById('xq-cube-finder-host')) return;

  const API = '/query/v1/search/cube/stock.json';
  const CANONICAL = 'https://xueqiu.com';
  const PAGE_SIZE = 20;
  const MAX_PAGES = 80; // 安全上限；不是全站遍历，只有当前股票的搜索分页。
  const REQUEST_INTERVAL = 650;
  const FILE_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const QUOTE_BATCH_SIZE = 30;
  const QUOTE_REQUEST_INTERVAL = 500;
  const RB_CONCURRENCY = 1; // 雪球频率限制：批量历史查询改为串行。
  const RB_REQUEST_INTERVAL = 3000;
  const RB_COOLDOWN_MS = 20 * 60 * 1000; // 仅本地保守冷却，不代表雪球官方限流周期。
  const RB_COOLDOWN_KEY = 'xq_stock_cube_finder_rb_cooldown_v174';
  const RB_ISSUE_TTL = 24 * 60 * 60 * 1000;
  const RB_RETRY_COUNT = 2;
  const RB_RETRY_BASE_MS = 1400;
  const LEGACY_CACHE_KEY = 'xq_stock_cube_finder_cache_v11';
  const TEMP_CACHE_KEY = 'xq_stock_cube_finder_recent_v13';
  const PREVIOUS_TEMP_CACHE_KEY = 'xq_stock_cube_finder_recent_v12';
  const CLOSED_CACHE_KEY = 'xq_stock_cube_finder_closed_v12';
  const LEGACY_IGNORE_KEY = 'xq_stock_cube_finder_ignore_v11';
  const QUOTE_TTL = 6 * 60 * 60 * 1000;
  const REBALANCE_TTL = 24 * 60 * 60 * 1000; // 调仓历史过期仍保留值并标记旧缓存。
  const HOLDING_TTL = 60 * 60 * 1000; // 当前持仓缓存 1 小时，过期不作为当前持仓结论。
  const SEARCH_CACHE_KEY = 'xq_stock_cube_finder_stock_search_v15';
  const SEARCH_TTL_KEY = 'xq_stock_cube_finder_search_ttl_v15';
  const SEARCH_TTL_CHOICES = new Set([0, 1, 6, 12, 24, 72, 168]);
  const SEARCH_CACHE_MAX_ENTRIES = 30;
  const SEARCH_CACHE_MAX_CHARS = 3000000; // 避免长时间使用后占用过多脚本存储。

  const state = {
    opened: false,
    busy: false,
    stock: '',
    page: 0,
    totalPages: null,
    data: new Map(),
    controller: null,
    runId: 0,
    completed: false,
    truncated: false,
    lastError: '',
    searchSource: '',
    searchSavedAt: null,
    auxBusy: false,
    auxController: null,
    quoteStates: new Map(),
    rebalancingStates: new Map(),
    holdingStates: new Map(), // 每个组合一次保存完整当前持仓，可以跨股票重复核对
    rebalanceErrors: new Map() // {kind, message, at}，分别标记失败、限流和未执行
  };

  const host = document.createElement('div');
  host.id = 'xq-cube-finder-host';
  document.documentElement.appendChild(host);
  const root = host.attachShadow({ mode: 'open' });
  root.innerHTML = `
<style>
  :host { all: initial; color-scheme: light; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Microsoft YaHei", sans-serif; }
  *, *::before, *::after { box-sizing: border-box; }
  button, input, select { font: inherit; }
  button { cursor: pointer; }
  .fab { position: fixed; bottom: 24px; right: 25px; z-index: 2147483646; padding: 13px 17px; border: none;
    color: white; background: #1366cf; border-radius: 28px; box-shadow: 0 8px 25px #0d3c8260; font-size: 14px; font-weight: 700; }
  .fab:hover { background: #0e55b2; }
  .panel { position: fixed; left: 50%; transform: translateX(-50%); bottom: 72px; z-index: 2147483646;
    width: min(1120px, calc(100vw - 56px)); max-width: calc(100vw - 56px);
    height: min(700px, calc(100vh - 124px)); display: flex; flex-direction: column; background: #fff;
    border: 1px solid #dfe7f0; border-radius: 15px; box-shadow: 0 20px 65px #142d5270; overflow: hidden;
    color: #192a40; font-size: 13px; }
  .hidden { display: none !important; }
  .header { padding: 16px 20px 13px; background: linear-gradient(110deg, #eef5ff, #fafcff); border-bottom: 1px solid #e6eef7; }
  .headrow { display: flex; align-items: center; justify-content: space-between; gap: 10px; }
  h1 { margin: 0; font-size: 18px; line-height: 1.35; font-weight: 750; }
  .subtitle { color: #65768e; font-size: 12px; margin: 5px 0 0; }
  .close { border: 0; background: transparent; color: #65768e; font-size: 24px; padding: 2px 5px; line-height: 1; }
  .body { display: flex; flex-direction: column; min-height: 0; flex: 1; padding: 14px 17px 12px; gap: 10px; }
  .query { display: flex; flex-wrap: wrap; gap: 8px; align-items: center; }
  .query input { flex: 1 1 230px; min-width: 140px; height: 36px; padding: 0 11px; border: 1px solid #c9d6e5; border-radius: 8px; outline: none; font-size: 14px; background: #fff; color: #1e2b3c; }
  .query input:focus, .tools input:focus, .tools select:focus { border-color: #3389de; box-shadow: 0 0 0 3px #dcebff; }
  .btn { height: 35px; padding: 0 12px; border-radius: 8px; border: 1px solid #d2dcea; background: #fff; color: #26415c; font-weight: 650; }
  .btn.primary { color: #fff; background: #146bcc; border-color: #146bcc; }
  .btn.primary:hover { background: #0757b0; }
  .btn.danger { color: #b12735; background: #fff6f6; border-color: #ebc4c8; }
  .btn:disabled { opacity: .45; cursor: not-allowed; }
  .hint { color: #78879b; font-size: 11px; }
  .summary { display: flex; flex-wrap: wrap; gap: 8px; align-items: center; }
  .pill { display: inline-flex; align-items: center; padding: 5px 9px; border-radius: 6px; background: #f0f5fb; color: #36506e; font-size: 12px; }
  .pill.strong { color: #0756a6; background: #e7f2ff; font-weight: 700; }
  .status { min-height: 18px; font-size: 12px; color: #617189; }
  .status.error { color: #b32633; }
  .tools { display: flex; align-items: center; flex-wrap: wrap; gap: 7px; }
  .tools label { font-size: 12px; color: #6b7a8c; }
  .tools select, .tools input { height: 31px; border: 1px solid #d2dbe8; border-radius: 7px; background: #fff; padding: 0 7px; color: #25354a; outline: none; }
  .tools input { width: 160px; }
  .tools .spacer { flex: 1; }
  .tools .mini { width: 125px; }
  .actions { display: flex; flex-wrap: wrap; gap: 8px; align-items: center; }
  .actions label { font-size: 12px; color: #6b7a8c; }
  .actions select { height: 35px; border: 1px solid #d2dbe8; border-radius: 7px; background: #fff; padding: 0 6px; color: #25354a; }
  .rowRefresh { border: 1px solid #bed4ee; background: #eff6ff; color: #1260a8; border-radius: 6px; padding: 5px 8px; font-size: 11px; white-space: nowrap; }
  .rowRefresh:hover { background: #dbefff; }
  .rowRefresh:disabled { cursor: not-allowed; opacity: .45; }
  .stale { color: #b37e2c; }
  .cell-status { font-size: 10px; color: #a17e43; margin-left: 5px; }
  .state.closed { color: #b32633; font-weight: 700; }
  .state.notclosed { color: #087c5b; }
  .state.unknown { color: #8090a2; }
  .state.absent { color: #bf3535; font-weight: 700; }
  .state.held { color: #087c5b; font-weight: 700; }
  .state.stale { color: #b37e2c; }
  .state.pending { color: #8090a2; }
  /* 固定表格布局，避免长列把操作按钮挤出窗口。 */
  .tablebox { flex: 1; min-width: 0; min-height: 0; overflow-y: auto; overflow-x: hidden; border: 1px solid #e1e7ef; border-radius: 9px; background: #fff; }
  table { width: 100%; table-layout: fixed; border-spacing: 0; border-collapse: separate; text-align: left; font-size: 12px; min-width: 0; }
  th { position: sticky; top: 0; z-index: 1; background: #eff4f9; color: #51657d; white-space: nowrap; font-weight: 700;
    padding: 10px 4px; border-bottom: 1px solid #dce5ef; overflow: hidden; text-overflow: ellipsis; }
  td { border-bottom: 1px solid #eef2f6; padding: 9px 4px; vertical-align: middle; white-space: nowrap;
    overflow: hidden; text-overflow: ellipsis; }
  td:first-child, td:nth-child(2) { overflow-wrap: normal; white-space: nowrap; }
  td:first-child a { display: block; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  td:nth-child(11) { font-size: 10.8px; letter-spacing: -.2px; }
  th:last-child { position: sticky; right: 0; z-index: 3; box-shadow: -6px 0 10px -8px #65768e; }
  td:last-child { position: sticky; right: 0; z-index: 2; background: #fff; box-shadow: -6px 0 10px -8px #65768e; }
  .rowRefresh { width: 100%; padding: 5px 2px; font-size: 11px; }
  tbody tr:hover { background: #f7faff; }
  tbody tr:hover td:last-child { background: #f7faff; }
  td .sub { display: block; margin-top: 3px; font-size: 11px; color: #8291a1; }
  a { color: #1266b7; text-decoration: none; font-weight: 650; }
  a:hover { text-decoration: underline; }
  .pos { color: #d33c3c; }
  .neg { color: #159169; }
  .empty { padding: 42px 20px; text-align: center; color: #8492a5; }
  .footer { display: flex; justify-content: space-between; gap: 8px; padding-top: 4px; align-items: center; font-size: 11px; color: #8390a1; }
  .footer .note { max-width: 90%; }
  /* 保留传统表格；视口偏窄时依次收起次要列，绝不变成卡片。 */
  @media (max-width: 1190px) {
    .panel { width: calc(100vw - 40px); max-width: calc(100vw - 40px); }
    .col-monthly, .col-drawdown, thead th:nth-child(8), thead th:nth-child(9), tbody td:nth-child(8), tbody td:nth-child(9) { display: none; }
  }
  @media (max-width: 890px) {
    .col-annual, thead th:nth-child(7), tbody td:nth-child(7) { display: none; }
    .body { padding: 10px; gap: 7px; }
    .header { padding: 12px 14px; }
    .footer .note { max-width: none; }
  }
  @media (max-width: 680px) {
    .fab { right: 14px; bottom: 16px; }
    .panel { bottom: 62px; width: calc(100vw - 16px); max-width: calc(100vw - 16px); height: calc(100vh - 78px); }
    .body { padding: 8px; }
    .col-owner, .col-gain, .col-closed, thead th:nth-child(2), thead th:nth-child(6), thead th:nth-child(10),
    tbody td:nth-child(2), tbody td:nth-child(6), tbody td:nth-child(10) { display: none; }
    table { font-size: 10.5px; }
    th, td { padding-left: 3px; padding-right: 3px; }
    .rowRefresh { font-size: 10px; }
    h1 { font-size: 16px; }
  }
</style>
<button class="fab" type="button" title="按股票名称查询持仓组合">📊 查持仓组合</button>
<section class="panel hidden" aria-label="雪球股票持仓组合查询">
  <header class="header">
    <div class="headrow"><h1>按股票找持仓组合</h1><button type="button" class="close" title="关闭">×</button></div>
    <p class="subtitle">雪球原生接口 · 完整搜索结果可缓存 · 无需建立组合数据库</p>
  </header>
  <div class="body">
    <div class="query">
      <input class="stock" type="text" placeholder="输入中文股票名称，如：震裕科技、贵州茅台" autocomplete="off" maxlength="64"/>
      <button type="button" class="btn primary search">查询全部组合</button>
      <button type="button" class="btn forceSearch" title="跳过本地股票列表缓存，立即重新搜索全部分页">强制刷新列表</button>
      <button type="button" class="btn danger stop" disabled>停止</button>
      <button type="button" class="btn export" disabled>导出 XLSX</button>
    </div>
    <div class="hint">搜索接口按需自动翻页，合并后全量排序；完整搜索结果按设置缓存，命中则不用重新翻页。点击单条或“更新全部”可同时刷新调仓时间和核验当前真实持仓；搜索仓位与核验仓位分开显示。组合并发2，调仓缓存24小时、持仓缓存1小时。</div>
    <div class="summary">
      <span class="pill strong stat-count">已获取 0 个组合</span>
      <span class="pill stat-pages">接口已读 0 页</span>
      <span class="pill stat-stock">等待查询</span>
      <span class="pill stat-cache">尚未查询</span>
    </div>
    <div class="status" role="status" aria-live="polite">输入中文股票名称后点击“查询全部组合”。</div>
    <div class="tools">
      <label for="xqSort">排序</label>
      <select class="sort" id="xqSort">
        <option value="weight-desc">搜索仓位高 → 低</option>
        <option value="weight-asc">搜索仓位低 → 高</option>
        <option value="gain-desc">累计收益高 → 低</option>
        <option value="annual-desc">年化收益高 → 低</option>
        <option value="monthly-desc">近30日收益高 → 低</option>
        <option value="drawdown-desc">回撤较小优先</option>
        <option value="rebalance-desc">最近实际调仓优先（缺失排后）</option>
      </select>
      <label for="xqStatus">状态</label>
      <select class="statusFilter mini" id="xqStatus">
        <option value="all">全部状态</option>
        <option value="notclosed">未见关停</option>
        <option value="closed">已关停</option>
        <option value="unknown">待核实</option>
      </select>
      <label for="xqHolding">持仓核验</label>
      <select class="holdingFilter mini" id="xqHolding">
        <option value="all">全部</option>
        <option value="held">确认持有</option>
        <option value="absent">确认未持有</option>
        <option value="unknown">未核验/过期</option>
      </select>
      <label for="xqFilter">筛选</label>
      <input class="filter" id="xqFilter" type="search" placeholder="组合名 / 主理人"/>
      <span class="spacer"></span>
      <span class="filter-count"></span>
    </div>
    <div class="actions">
      <label for="xqSearchTtl">股票列表缓存</label>
      <select class="searchTtl" id="xqSearchTtl" title="仅影响股票持仓组合搜索结果；关停和调仓缓存独立">
        <option value="0">不缓存</option>
        <option value="1">1小时</option>
        <option value="6">6小时</option>
        <option value="12">12小时</option>
        <option value="24">24小时（默认）</option>
        <option value="72">3天</option>
        <option value="168">7天</option>
      </select>
      <button type="button" class="btn quoteRefresh" disabled>刷新未关停状态（批量）</button>
      <button type="button" class="btn loadRebalance" disabled>更新全部调仓及真实持仓（未关停）</button>
      <button type="button" class="btn danger clearCache">清理缓存</button>
      <span class="hint cacheInfo">关停状态永久缓存；已核验持仓按组合复用（1小时）。</span>
    </div>
    <div class="tablebox">
      <div class="empty">查询后将在这里展示组合名称、主理人和仓位比例。</div>
      <table class="hidden">
        <colgroup>
          <col style="width:14%"><col class="col-owner" style="width:8%"><col style="width:7%"><col style="width:7%">
          <col style="width:8%"><col class="col-gain" style="width:7%"><col class="col-annual" style="width:7%">
          <col class="col-monthly" style="width:6%"><col class="col-drawdown" style="width:6%"><col class="col-closed" style="width:8%">
          <col style="width:13%"><col style="width:9%">
        </colgroup>
        <thead><tr>
          <th>组合名称</th><th>主理人</th><th>搜索仓位</th><th>当前仓位</th><th>持仓核验</th><th>累计收益</th><th>年化收益</th><th>近30日</th><th>回撤</th><th>关停状态</th><th>最近实际调仓</th><th>操作</th>
        </tr></thead>
        <tbody></tbody>
      </table>
    </div>
    <div class="footer">
      <div class="note">只统计 list，不混入 recommend。完整列表在有效期内优先读取本地缓存；点击“强制刷新列表”绕过缓存。关停状态永久缓存；当前持仓缓存1小时，过期不作为确认结论；调仓缓存24小时；不存储登录凭证。</div>
      <span>v1.7.4</span>
    </div>
  </div>
</section>`;

  const $ = selector => root.querySelector(selector);
  const ui = {
    fab: $('.fab'), panel: $('.panel'), close: $('.close'), stock: $('.stock'), search: $('.search'), forceSearch: $('.forceSearch'), searchTtl: $('.searchTtl'), statCache: $('.stat-cache'),
    stop: $('.stop'), export: $('.export'), sort: $('.sort'), filter: $('.filter'),
    statCount: $('.stat-count'), statPages: $('.stat-pages'), statStock: $('.stat-stock'),
    status: $('.status'), filterCount: $('.filter-count'), table: $('table'),
    tableBody: $('tbody'), empty: $('.empty'), statusFilter: $('.statusFilter'),
    quoteRefresh: $('.quoteRefresh'), loadRebalance: $('.loadRebalance'), holdingFilter: $('.holdingFilter'),
    clearCache: $('.clearCache'), cacheInfo: $('.cacheInfo')
  };

  function toggle(force) {
    state.opened = typeof force === 'boolean' ? force : !state.opened;
    ui.panel.classList.toggle('hidden', !state.opened);
    ui.fab.textContent = state.opened ? '📊 收起查询' : '📊 查持仓组合';
    if (state.opened) ui.stock.focus();
  }
  ui.fab.addEventListener('click', () => toggle());
  ui.close.addEventListener('click', () => toggle(false));
  if (typeof GM_registerMenuCommand === 'function') {
    GM_registerMenuCommand('打开「按股票找持仓组合」', () => toggle(true));
  }

  function setStatus(message, error = false) {
    ui.status.textContent = message;
    ui.status.classList.toggle('error', error);
  }

  function updateStats() {
    ui.statCount.textContent = `已获取 ${state.data.size} 个组合`;
    ui.statPages.textContent = `接口已读 ${state.page}${state.totalPages ? ` / ${state.totalPages}` : ''} 页`;
    ui.statStock.textContent = state.stock || '等待查询';
    ui.statCache.textContent = state.searchSource === 'cache'
      ? `列表缓存 · ${formatDate(state.searchSavedAt)} 获取`
      : state.searchSource === 'live' ? '实时查询' : state.searchSource === 'memory' ? '当前列表（已清缓存）' : '尚未查询';
    refreshControls();
  }

  function refreshControls() {
    const working = state.busy || state.auxBusy;
    ui.search.disabled = working;
    ui.forceSearch.disabled = working;
    ui.stop.disabled = !working;
    ui.stock.disabled = working;
    ui.export.disabled = state.data.size === 0;
    ui.quoteRefresh.disabled = working || !state.data.size;
    ui.loadRebalance.disabled = working || !state.data.size;
  }

  function setBusy(busy) { state.busy = busy; refreshControls(); }
  function setAuxBusy(busy) { state.auxBusy = busy; refreshControls(); }

  function num(value) {
    if (value === null || value === undefined || value === '') return null;
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : null;
  }
  function formatPercent(v) {
    const n = num(v);
    return n == null ? '—' : `${n.toFixed(2)}%`;
  }
  function gainClass(v) {
    const n = num(v);
    return n == null ? '' : n > 0 ? 'pos' : n < 0 ? 'neg' : '';
  }
  function readJson(key, fallback = {}) {
    try {
      const text = localStorage.getItem(key);
      return text ? JSON.parse(text) : fallback;
    } catch (_) { return fallback; }
  }
  // 独立缓存“某股票的全部组合搜索结果”，不改变关停/调仓缓存。
  // 只有所有搜索分页正常结束才会落盘；中途停止、错误或达到安全页数上限不会覆盖旧缓存。
  function readSetting(key, fallback) {
    try {
      if (typeof GM_getValue === 'function') return GM_getValue(key, fallback);
    } catch (_) { /* fallback */ }
    try { const v = localStorage.getItem(key); return v == null ? fallback : JSON.parse(v); }
    catch (_) { return fallback; }
  }
  function writeSetting(key, value) {
    try {
      if (typeof GM_setValue === 'function') { GM_setValue(key, value); return; }
      localStorage.setItem(key, JSON.stringify(value));
    } catch (_) { /* 用户禁用本地存储不影响查询 */ }
  }
  function loadSearchCache() {
    const data = readSetting(SEARCH_CACHE_KEY, null);
    return data && data.version === 1 && data.entries && typeof data.entries === 'object' &&
      !Array.isArray(data.entries) ? data.entries : {};
  }
  let searchCache = loadSearchCache();
  function ttlHours() {
    const value = Number(ui.searchTtl.value);
    return SEARCH_TTL_CHOICES.has(value) ? value : 24;
  }
  function getCachedSearch(query) {
    const hours = ttlHours();
    const entry = searchCache[query];
    if (!hours || !entry || !Array.isArray(entry.rows)) return null;
    const age = Date.now() - entry.savedAt;
    if (!Number.isFinite(age) || age < 0 || age >= hours * 3600000 ||
        entry.completed !== true || entry.truncated === true) return null;
    // 数据校验：无效缓存绝不用于恢复列表。
    if (entry.rows.some(row => !row || typeof row.symbol !== 'string' || row.weight == null)) return null;
    entry.accessedAt = Date.now();
    persistSearchCache();
    return entry;
  }
  function persistSearchCache() {
    try {
      const entries = Object.entries(searchCache)
        .filter(([, e]) => e?.savedAt && Array.isArray(e.rows) && Date.now() - e.savedAt < 30 * 86400000)
        .sort((a, b) => (b[1].accessedAt || b[1].savedAt) - (a[1].accessedAt || a[1].savedAt))
        .slice(0, SEARCH_CACHE_MAX_ENTRIES);
      const data = {};
      for (const [key, value] of entries) {
        data[key] = value;
        if (JSON.stringify({version: 1, entries: data}).length > SEARCH_CACHE_MAX_CHARS) delete data[key];
      }
      searchCache = data;
      writeSetting(SEARCH_CACHE_KEY, {version: 1, entries: searchCache});
    } catch (_) { /* 本地存储失败不影响实时查询 */ }
  }
  function storeCompletedSearch(query) {
    if (!ttlHours() || !state.completed || state.truncated) return;
    const now = Date.now();
    searchCache[query] = {
      savedAt: now, accessedAt: now, completed: true, truncated: false,
      page: state.page, totalPages: state.totalPages,
      rows: [...state.data.values()]
    };
    persistSearchCache();
    state.searchSavedAt = now;
  }
  function initializeSearchSettings() {
    const selected = Number(readSetting(SEARCH_TTL_KEY, 24));
    ui.searchTtl.value = String(SEARCH_TTL_CHOICES.has(selected) ? selected : 24);
  }
  function readClosedCache() {
    try {
      if (typeof GM_getValue === 'function') {
        const value = GM_getValue(CLOSED_CACHE_KEY, '');
        if (value) return typeof value === 'string' ? JSON.parse(value) : value;
      }
    } catch (_) { /* 回退到当前站点的 localStorage */ }
    return readJson(CLOSED_CACHE_KEY);
  }
  function writeClosedCache() {
    const closed = Object.fromEntries([...state.quoteStates].filter(([, v]) =>
      v?.status === 'closed' && num(v.closedAt) > 0));
    const json = JSON.stringify(closed);
    // Tampermonkey 存储可在 xueqiu.com 和 www.xueqiu.com 共享，不因域名变化丢失关停记录。
    try {
      if (typeof GM_setValue === 'function') { GM_setValue(CLOSED_CACHE_KEY, json); return; }
    } catch (_) { /* 回退 localStorage */ }
    try { localStorage.setItem(CLOSED_CACHE_KEY, json); }
    catch (_) { /* 浏览器存储额度不足，不影响实时查询 */ }
  }
  function readRecentCache() {
    try {
      if (typeof GM_getValue === 'function') {
        const value = GM_getValue(TEMP_CACHE_KEY, '');
        if (value) return typeof value === 'string' ? JSON.parse(value) : value;
      }
    } catch (_) { /* 回退到 localStorage */ }
    return readJson(TEMP_CACHE_KEY);
  }
  function cacheLoad() {
    const now = Date.now();
    let ignoreLegacy = false;
    try { ignoreLegacy = typeof GM_getValue === 'function' ? !!GM_getValue(LEGACY_IGNORE_KEY, false) : !!readJson(LEGACY_IGNORE_KEY, false); }
    catch (_) { /* 忽略 */ }
    const old = ignoreLegacy ? {} : readJson(LEGACY_CACHE_KEY);
    const previous = (() => {
      try {
        if (typeof GM_getValue === 'function') {
          const data = GM_getValue(PREVIOUS_TEMP_CACHE_KEY, '');
          if (data) return typeof data === 'string' ? JSON.parse(data) : data;
        }
      } catch (_) { /* ignore */ }
      return readJson(PREVIOUS_TEMP_CACHE_KEY);
    })();
    const recent = readRecentCache();
    // 继承关停状态，但不能继承 v1.2 用 current.json 取得的错误调仓时间。
    let migrated = false;
    for (const saved of [old, previous, recent]) {
      for (const [symbol, v] of Object.entries(saved.quote || {})) {
        if (!v || !v.savedAt) continue;
        if (v.status === 'closed' && num(v.closedAt) > 0) {
          state.quoteStates.set(symbol, v); migrated = true;
        } else if (v.status === 'notclosed' && now - v.savedAt < QUOTE_TTL) {
          state.quoteStates.set(symbol, v);
        }
      }
    }
    // 保留近一天的批量失败记录和限流未执行状态，避免刷新页面变成“待加载”。
    for (const [symbol, issue] of Object.entries(recent.rbIssues || {})) {
      if (issue && ['error', 'throttled', 'skipped'].includes(issue.kind) &&
          typeof issue.message === 'string' && Number.isFinite(issue.at) &&
          now >= issue.at && now - issue.at < RB_ISSUE_TTL) state.rebalanceErrors.set(symbol, issue);
    }
    // 仅从 v1.3 新缓存读取历史接口确认的记录；历史解析规则改变时必须更新版本键。
    for (const [symbol, v] of Object.entries(recent.rebalance || {})) {
      if (v?.source === 'history' && v?.savedAt && v.savedAt <= now) {
        // 老记录保留并显示“旧缓存”，手动刷新可重新核验。
        state.rebalancingStates.set(symbol, v);
      }
    }
    // v1.6 新增完整持仓缓存；旧版没有此缓存，避免将搜索 weight 伪装为实盘仓位。
    for (const [symbol, v] of Object.entries(recent.holdings || {})) {
      if (v?.source === 'current' && Array.isArray(v.items) &&
          Number.isFinite(v.savedAt) && v.savedAt <= now && v.items.length <= 200) {
        state.holdingStates.set(symbol, v);
      }
    }
    // 永久关停状态的优先级高于临时缓存，防止已关停组合被误覆盖。
    for (const [symbol, v] of Object.entries(readClosedCache() || {})) {
      if (v?.status === 'closed' && num(v.closedAt) > 0) state.quoteStates.set(symbol, v);
    }
    if (migrated) cacheSave(true);
  }
  function cacheSave(saveClosed = false) {
    try {
      // 只有“未见关停”状态和最近调仓会过期/淘汰。已关停记录另行永久持久化。
      const newest = m => Object.fromEntries([...m.entries()]
        .sort((a, b) => (b[1].savedAt || 0) - (a[1].savedAt || 0)).slice(0, 3000));
      const json = JSON.stringify({
        quote: newest(new Map([...state.quoteStates].filter(([, v]) => v?.status !== 'closed'))),
        rebalance: newest(state.rebalancingStates),
        holdings: newest(state.holdingStates),
        rbIssues: Object.fromEntries([...state.rebalanceErrors].filter(([, v]) => v?.at && Date.now() - v.at < RB_ISSUE_TTL).slice(-3000))
      });
      if (typeof GM_setValue === 'function') GM_setValue(TEMP_CACHE_KEY, json);
      else localStorage.setItem(TEMP_CACHE_KEY, json);
    } catch (_) { /* 临时缓存失败不影响查询 */ }
    if (saveClosed) writeClosedCache();
  }
  function getQuote(symbol) {
    const value = state.quoteStates.get(symbol);
    if (!value) return null;
    if (value.status === 'closed' && num(value.closedAt) > 0) return value; // 永久有效
    return Date.now() - value.savedAt < QUOTE_TTL ? value : null;
  }
  function clearCache() {
    if (state.busy || state.auxBusy) return;
    const closedCount = [...state.quoteStates.values()].filter(v => v?.status === 'closed').length;
    const searchCount = Object.keys(searchCache).length;
    if (!window.confirm(`清空查询助手本地缓存？\n包括 ${searchCount} 条股票搜索缓存、${closedCount} 条永久关停记录、所有调仓/持仓核验/状态临时记录。\n不会删除当前已显示的组合列表，也不会改变选择的缓存时长。`)) return;
    try {
      if (typeof GM_setValue === 'function') {
        GM_setValue(SEARCH_CACHE_KEY, {version: 1, entries: {}});
        GM_setValue(CLOSED_CACHE_KEY, '{}');
        GM_setValue(TEMP_CACHE_KEY, '{}');
        GM_setValue(PREVIOUS_TEMP_CACHE_KEY, '{}');
        GM_setValue(LEGACY_IGNORE_KEY, true); // 避免从另一子域名的旧版缓存重新导入已清理记录
      } else localStorage.setItem(LEGACY_IGNORE_KEY, 'true');
    } catch (_) { /* 清理本地备用存储 */ }
    for (const key of [SEARCH_CACHE_KEY, CLOSED_CACHE_KEY, TEMP_CACHE_KEY, PREVIOUS_TEMP_CACHE_KEY, LEGACY_CACHE_KEY]) {
      try { localStorage.removeItem(key); } catch (_) { /* 忽略不可写存储 */ }
    }
    state.quoteStates.clear();
    state.rebalancingStates.clear();
    state.rebalanceErrors.clear();
    writeSetting(RB_COOLDOWN_KEY, 0);
    state.holdingStates.clear();
    searchCache = {};
    if (state.searchSource === 'cache') state.searchSource = 'memory';
    state.searchSavedAt = null;
    updateStats();
    render();
    setStatus('股票列表缓存、永久关停和调仓缓存均已清理；当前显示列表保留，可手动重新获取。');
  }
  function isRateLimitedError(error) {
    const message = String(error?.message || error || '');
    return error?.rateLimited === true || /110017|HTTP\s+(403|429)\b|操作过于频繁|访问过于频繁|限流|请求频繁/.test(message);
  }
  function cooldownUntil() {
    const until = Number(readSetting(RB_COOLDOWN_KEY, 0));
    return Number.isFinite(until) && until > Date.now() ? until : 0;
  }
  function startCooldown() {
    const until = Date.now() + RB_COOLDOWN_MS;
    writeSetting(RB_COOLDOWN_KEY, until);
    return until;
  }
  function getRb(symbol) {
    // 所有历史调仓记录持久展示；超过有效期只标记旧缓存，不悄悄遗忘已查过的值。
    return state.rebalancingStates.get(symbol) || null;
  }
  function isRbFresh(symbol) {
    const value = getRb(symbol);
    return !!value && Number.isFinite(value.savedAt) &&
      Date.now() >= value.savedAt && Date.now() - value.savedAt < REBALANCE_TTL;
  }
  function getHolding(symbol) {
    return state.holdingStates.get(symbol) || null;
  }
  function isHoldingFresh(symbol) {
    const value = getHolding(symbol);
    return !!value && Number.isFinite(value.savedAt) &&
      Date.now() >= value.savedAt && Date.now() - value.savedAt < HOLDING_TTL;
  }
  function normalizeStockName(name) {
    return String(name ?? '').normalize('NFKC').replace(/[\s\u200b\u00a0]/g, '').toLocaleLowerCase('zh-CN');
  }
  // 只有核验快照仍在有效期内才能得出“当前持有/未持有”的结论。
  function holdingCheck(symbol) {
    const snapshot = getHolding(symbol);
    if (!snapshot) return {status: 'pending', weight: null};
    if (!isHoldingFresh(symbol)) return {status: 'stale', weight: null, savedAt: snapshot.savedAt};
    const wanted = normalizeStockName(state.stock);
    const matches = snapshot.items.filter(item => normalizeStockName(item.name) === wanted);
    if (matches.length) {
      const weight = matches.reduce((sum, item) => sum + (num(item.weight) ?? 0), 0);
      return weight > 0 ? {status: 'held', weight, savedAt: snapshot.savedAt} :
        {status: 'absent', weight: 0, savedAt: snapshot.savedAt};
    }
    // 快照有代码无名称时，不能由中文名称直接推断“未持有”。
    if (snapshot.items.some(item => !normalizeStockName(item.name) && num(item.weight) > 0)) {
      return {status: 'unknown', weight: null, savedAt: snapshot.savedAt};
    }
    return {status: 'absent', weight: 0, savedAt: snapshot.savedAt};
  }
  // 确认已成交完整快照后才会保存；接口失败、结构不完整绝不能标记“未持有”。
  function parseCurrentHoldings(payload) {
    const node = payload?.data && typeof payload.data === 'object' ? payload.data : payload;
    const rb = node?.last_success_rb;
    if (!rb || rb.status !== 'success' || !Array.isArray(rb.holdings)) {
      throw new Error('current.json 缺少已成交的完整 holdings，不能确认真实仓位');
    }
    if (rb.holdings.length > 200) throw new Error('当前持仓返回异常：股票数量超过安全上限');
    const items = rb.holdings.map(h => {
      if (!h || typeof h !== 'object' || !Number.isFinite(num(h.weight))) {
        throw new Error('持仓快照字段不完整，跳过核验');
      }
      return {symbol: String(h.stock_symbol ?? ''), name: String(h.stock_name ?? ''), weight: num(h.weight)};
    });
    return {items, source: 'current', savedAt: Date.now()};
  }
  function formatDate(timestamp) {
    const n = num(timestamp);
    if (n === null || n <= 0) return '—';
    const ms = n < 1e11 ? n * 1000 : n;
    const date = new Date(ms);
    return Number.isNaN(date.getTime()) ? '—' : date.toLocaleString('zh-CN', {
      year: 'numeric', month: '2-digit', day: '2-digit',
      hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false, timeZone: 'Asia/Shanghai'
    });
  }
  function statusText(symbol) {
    const status = getQuote(symbol)?.status || 'unknown';
    return status === 'closed' ? '已关停' : status === 'notclosed' ? '未见关停' : '待核实';
  }
  function sortableRows() {
    const keyword = ui.filter.value.trim().toLowerCase();
    const requestedStatus = ui.statusFilter.value;
    const holdingFilter = ui.holdingFilter.value;
    const values = [...state.data.values()].filter(r => {
      if (keyword && ![r.name, r.owner_name, r.symbol].some(v =>
        String(v || '').toLowerCase().includes(keyword))) return false;
      if (requestedStatus !== 'all' && (getQuote(r.symbol)?.status || 'unknown') !== requestedStatus) return false;
      const check = holdingCheck(r.symbol);
      if (holdingFilter === 'held' && check.status !== 'held') return false;
      if (holdingFilter === 'absent' && check.status !== 'absent') return false;
      if (holdingFilter === 'unknown' && ['held','absent'].includes(check.status)) return false;
      return true;
    });
    const [field, direction] = ui.sort.value.split('-');
    const fields = {
      weight: 'weight', gain: 'total_gain', annual: 'annualized_gain_rate',
      monthly: 'monthly_gain', drawdown: 'draw_down'
    };
    const sign = direction === 'asc' ? 1 : -1;
    values.sort((a, b) => {
      if (field === 'rebalance') {
        const ac = getQuote(a.symbol)?.status === 'closed';
        const bc = getQuote(b.symbol)?.status === 'closed';
        if (ac !== bc) return ac ? 1 : -1; // 排序时关停组合排后
      }
      const va = field === 'rebalance' ? (num(getRb(a.symbol)?.timestamp) || null) : num(a[fields[field] || 'weight']);
      const vb = field === 'rebalance' ? (num(getRb(b.symbol)?.timestamp) || null) : num(b[fields[field] || 'weight']);
      if (va === null && vb === null) return String(a.symbol).localeCompare(String(b.symbol));
      if (va === null) return 1;
      if (vb === null) return -1;
      return sign * (va - vb) || String(a.symbol).localeCompare(String(b.symbol));
    });
    return values;
  }

  function render() {
    // 表格不再分页：排序、筛选始终作用于当前股票全部已获取组合。
    let rows = sortableRows();
    const eligible = rows.filter(r => getQuote(r.symbol)?.status !== 'closed');
    const rbLoaded = eligible.filter(r => isRbFresh(r.symbol)).length;
    const rbStale = eligible.filter(r => !!getRb(r.symbol) && !isRbFresh(r.symbol)).length;
    const verified = eligible.filter(r => ['held', 'absent'].includes(holdingCheck(r.symbol).status)).length;
    const recentSort = ui.sort.querySelector('option[value="rebalance-desc"]');
    // 即使仅部分组合获取成功，也支持按已知时间排序；无时间记录排最后。
    recentSort.disabled = rows.length === 0;
    const closedCount = [...state.quoteStates.values()].filter(v => v?.status === 'closed').length;
    ui.filterCount.textContent = state.data.size ? `显示 ${rows.length} / ${state.data.size} 条 · 全量排序` : '';
    const failedCount = eligible.filter(r => ['error', 'throttled'].includes(state.rebalanceErrors.get(r.symbol)?.kind)).length;
    const skippedCount = eligible.filter(r => state.rebalanceErrors.get(r.symbol)?.kind === 'skipped').length;
    const until = cooldownUntil();
    ui.cacheInfo.textContent = `列表缓存 ${Object.keys(searchCache).length} 只股票；关停永久缓存 ${closedCount} 个；当前筛选未关停调仓已核验 ${rbLoaded}/${eligible.length}${rbStale ? `，旧缓存 ${rbStale}` : ''}，请求失败/限流 ${failedCount}，限流未执行 ${skippedCount}；实时持仓已核验 ${verified}/${eligible.length}${until ? `；调仓冷却至 ${formatDate(until)}` : ''}。`;
    // 详情更新按钮需要随着请求状态切换为禁用。
    ui.table.classList.toggle('hidden', !rows.length);
    ui.empty.classList.toggle('hidden', !!rows.length);
    if (!rows.length) {
      ui.empty.textContent = state.data.size ? '没有匹配筛选条件的组合。' : '暂无组合结果。';
      ui.tableBody.replaceChildren();      return;
    }
    const frag = document.createDocumentFragment();
    for (const row of rows) {
      const tr = document.createElement('tr');
      const nameCell = document.createElement('td');
      const a = document.createElement('a');
      a.href = `https://xueqiu.com/P/${encodeURIComponent(row.symbol)}`;
      a.target = '_blank';
      a.rel = 'noopener noreferrer';
      a.textContent = row.name || row.symbol;
      nameCell.appendChild(a);
      const subtitle = document.createElement('span');
      subtitle.className = 'sub';
      subtitle.textContent = row.symbol;
      nameCell.appendChild(subtitle);
      tr.appendChild(nameCell);
      const owner = document.createElement('td'); owner.textContent = row.owner_name || '—'; tr.appendChild(owner);
      const searchWeight = document.createElement('td');
      searchWeight.textContent = formatPercent(row.weight);
      searchWeight.title = '雪球反查接口的搜索权重，可能不是当前实际持仓';
      tr.appendChild(searchWeight);
      const check = holdingCheck(row.symbol);
      const actualWeight = document.createElement('td');
      actualWeight.textContent = check.status === 'held' ? formatPercent(check.weight) :
        check.status === 'absent' ? '0.00%' : '—';
      actualWeight.title = check.status === 'absent' ? '已成功核验：当前持仓不包含所查股票' :
        check.status === 'held' ? `实际持仓已核验：${formatDate(check.savedAt)}` : '需要更新真实持仓才能确认';
      tr.appendChild(actualWeight);
      const verifyCell = document.createElement('td');
      verifyCell.textContent = ({held:'当前持有',absent:'当前未持有',stale:'旧缓存需更新',unknown:'无法确认',pending:'待核验'})[check.status];
      verifyCell.className = `state ${check.status}`;
      tr.appendChild(verifyCell);
      for (const [field, isGain] of [
        ['total_gain', true], ['annualized_gain_rate', true],
        ['monthly_gain', true], ['draw_down', true]
      ]) {
        const td = document.createElement('td');
        td.textContent = formatPercent(row[field]);
        if (isGain) td.className = gainClass(row[field]);
        tr.appendChild(td);
      }
      const quote = getQuote(row.symbol);
      const status = document.createElement('td');
      status.textContent = statusText(row.symbol);
      status.className = `state ${quote?.status || 'unknown'}`;
      if (quote?.status === 'closed' && quote.closedAt) {
        status.title = `关停时间：${formatDate(quote.closedAt)}`;
      }
      tr.appendChild(status);
      const rb = getRb(row.symbol);
      const recent = document.createElement('td');
      const fresh = isRbFresh(row.symbol);
      recent.textContent = rb ? formatDate(rb.timestamp) : '待加载';
      if (rb && !rb.timestamp) recent.textContent = '近50条未确认主动调仓';
      if (rb && !fresh) {
        recent.classList.add('stale');
        const mark = document.createElement('span');
        mark.className = 'cell-status'; mark.textContent = '旧缓存'; recent.appendChild(mark);
      }
      if (rb) recent.title = `最后核验：${formatDate(rb.savedAt)}（北京时间） · 来源：历史调仓 · ${rb.category || '未确认'} · ID ${rb.id || '未知'}`;
      const rbError = state.rebalanceErrors.get(row.symbol);
      if (rbError && !isRbFresh(row.symbol)) {
        if (!rb) recent.textContent = rbError.kind === 'skipped' ? '限流后未执行' :
          rbError.kind === 'throttled' ? '接口限流（110017）' : '获取失败（可重试）';
        recent.title = `${rbError.message} · 记录时间 ${formatDate(rbError.at)}；可在限制解除后补齐`;
      }
      tr.appendChild(recent);
      const action = document.createElement('td');
      const refresh = document.createElement('button');
      refresh.type = 'button'; refresh.className = 'rowRefresh';
      refresh.textContent = '↻ 调仓+持仓';
      refresh.title = `同时更新 ${row.symbol} 的最近主动调仓及完整持仓快照（各1次请求）`;
      refresh.disabled = state.busy || state.auxBusy;
      refresh.addEventListener('click', () => refreshRebalancing([row.symbol], 'single'));
      action.appendChild(refresh); tr.appendChild(action);
      // 表格在窄视口隐藏部分收益列；被收起的字段始终保留在 Excel 导出中。
      const columnNames = ['组合名称', '主理人', '搜索仓位', '当前仓位', '持仓核验', '累计收益', '年化收益', '近30日', '回撤', '关停状态', '最近实际调仓', '操作'];
      [...tr.children].forEach((cell, index) => {
        cell.dataset.label = columnNames[index] || '详情';
        if (!cell.title && index !== 11) cell.title = String(cell.textContent || '').trim();
      });
      frag.appendChild(tr);
    }
    ui.tableBody.replaceChildren(frag);
  }

  function pause(ms) { return new Promise(resolve => setTimeout(resolve, ms)); }
  function isAbort(err) { return err?.name === 'AbortError'; }

  async function sameOriginRequest(path, signal) {
    const response = await fetch(`${location.origin}${path}`, {
      method: 'GET', credentials: 'include', signal, cache: 'no-store',
      headers: { Accept: 'application/json' }
    });
    if (!response.ok) throw new Error(`网页请求 HTTP ${response.status}`);
    const raw = await response.text();
    try { return JSON.parse(raw); }
    catch (_) { throw new Error('网页接口没有返回 JSON（可能需要登录）'); }
  }

  function gmRequest(path, signal) {
    return new Promise((resolve, reject) => {
      if (signal.aborted) return reject(new DOMException('已取消', 'AbortError'));
      let request;
      const onAbort = () => request?.abort();
      signal.addEventListener('abort', onAbort, { once: true });
      const done = (fn, value) => {
        signal.removeEventListener('abort', onAbort);
        fn(value);
      };
      try {
        request = GM_xmlhttpRequest({
          method: 'GET', url: `${CANONICAL}${path}`,
          headers: { Accept: 'application/json' },
          timeout: 20000,
          onload: res => {
            if (signal.aborted) return done(reject, new DOMException('已取消', 'AbortError'));
            if (res.status !== 200) return done(reject, new Error(`跨域备用请求 HTTP ${res.status}`));
            try { done(resolve, JSON.parse(res.responseText)); }
            catch (_) { done(reject, new Error('备用接口没有返回 JSON（请检查雪球登录状态）')); }
          },
          onerror: () => done(reject, new Error('备用网络请求失败')),
          ontimeout: () => done(reject, new Error('备用请求超时')),
          onabort: () => done(reject, new DOMException('已取消', 'AbortError'))
        });
      } catch (e) { done(reject, e); }
    });
  }

  async function requestJson(path, signal) {
    try {
      return await sameOriginRequest(path, signal);
    } catch (primaryErr) {
      if (signal.aborted || isAbort(primaryErr)) throw primaryErr;
      // 服务端明确拒绝/限流时不要通过跨域备用通道立刻重试，避免把请求放大两倍。
      if (/HTTP (403|429)\b/.test(primaryErr.message || '')) throw primaryErr;
      try { return await gmRequest(path, signal); }
      catch (fallbackErr) {
        if (signal.aborted) throw new DOMException('已取消', 'AbortError');
        throw new Error(`直连及备用请求均失败：${primaryErr.message}；${fallbackErr.message}`);
      }
    }
  }

  // 行情接口支持以逗号分隔的多个组合代码。closed_at 是判断关停的主要依据。
  function parseQuoteRecord(value) {
    if (!value || typeof value !== 'object') return null;
    const raw = value.closed_at;
    const closedAt = num(raw);
    if (closedAt != null && closedAt > 0) {
      return { status: 'closed', closedAt, savedAt: Date.now() };
    }
    // 只有字段确实存在时，才显示“未见关停”；不能用 hasexist=false 判断关停。
    if (Object.prototype.hasOwnProperty.call(value, 'closed_at')) {
      return { status: 'notclosed', closedAt: null, savedAt: Date.now() };
    }
    return null;
  }

  async function loadQuoteAll(force = false, fromSearch = false) {
    if (state.auxBusy || (state.busy && !fromSearch) || !state.data.size) return;
    const symbols = [...state.data.keys()].filter(symbol => {
      const cached = getQuote(symbol);
      if (cached?.status === 'closed') return false; // 永久关停不再重复请求，清缓存可解除
      return force || !cached;
    });
    if (!symbols.length) {
      if (!fromSearch) setStatus('所有组合的关停状态均已命中缓存；已关停组合不会重新请求。');
      render();
      return;
    }
    const controller = new AbortController();
    const currentRun = state.runId;
    state.auxController = controller;
    setAuxBusy(true);
    let checked = 0, unknown = 0, failed = 0;
    try {
      for (let i = 0; i < symbols.length; i += QUOTE_BATCH_SIZE) {
        if (currentRun !== state.runId || controller.signal.aborted) return;
        const chunk = symbols.slice(i, i + QUOTE_BATCH_SIZE);
        const params = new URLSearchParams({ code: chunk.join(',') });
        try {
          const payload = await requestJson(`/cubes/quote.json?${params}`, controller.signal);
          if (currentRun !== state.runId || controller.signal.aborted) return;
          const source = payload?.data && !chunk.some(x => Object.hasOwn(payload, x)) ? payload.data : payload;
          if (!source || typeof source !== 'object' || Array.isArray(source)) throw new Error('行情接口格式异常');
          // 不对缺失的组合赋予未关停状态。
          for (const symbol of chunk) {
            const parsed = parseQuoteRecord(source[symbol]);
            if (parsed) state.quoteStates.set(symbol, parsed);
            else unknown++;
          }
          checked += chunk.length;
          cacheSave(true);
        } catch (e) {
          if (controller.signal.aborted || isAbort(e)) return;
          failed += chunk.length;
        }
        render();
        setStatus(`批量查关停：已处理 ${Math.min(i + chunk.length, symbols.length)} / ${symbols.length}，未确认 ${unknown}，请求失败覆盖 ${failed} 条。`);
        if (i + chunk.length < symbols.length) await pause(QUOTE_REQUEST_INTERVAL);
      }
      if (currentRun === state.runId) {
        setStatus(`批量关停状态完成：本次处理 ${checked} 条，未知 ${unknown} 条，失败 ${failed} 条；已关停记录永久保存。调仓可按需逐批加载。${failed ? '如遇限流请稍后手动刷新。' : ''}`, !!failed);
      }
    } finally {
      if (state.auxController === controller) { state.auxController = null; setAuxBusy(false); render(); }
    }
  }

  // history.json 是独立的调仓事件列表；current.json 的 last_success_rb.created_at
  // 可能是近期生成的持仓快照/调整记录，不等价于组合页展示的“最新调仓”。
  // 仅认可成功的用户调仓、且交易变动列表中确实包含股票权重变动的记录。
  function isEffectiveRebalance(record) {
    if (!record || typeof record !== 'object') return false;
    if (record.status !== 'success') return false;
    if (record.category && record.category !== 'user_rebalancing') return false;
    const changes = record.rebalancing_histories;
    if (!Array.isArray(changes) || !changes.length) return false;
    return changes.some(change => {
      if (!change || typeof change !== 'object') return false;
      if (change.proactive === true) return true;
      const before = num(change.prev_target_weight ?? change.prev_weight);
      const after = num(change.target_weight ?? change.weight);
      // 新买入与清仓时 before/after 可能为 null，视作 0。
      if (before !== null || after !== null) {
        return Math.abs((after ?? 0) - (before ?? 0)) >= 0.001;
      }
      // 缺少仓位字段时不猜测“有调仓”，避免将系统快照误当交易。
      return false;
    });
  }

  function parseLatestRebalance(payload) {
    const node = payload?.data && typeof payload.data === 'object' ? payload.data : payload;
    const list = Array.isArray(node?.list) ? node.list : (Array.isArray(node?.items) ? node.items : null);
    if (!list) throw new Error('调仓历史接口格式异常：缺少 list');
    const record = list.filter(isEffectiveRebalance)
      .sort((a, b) => (num(b.created_at) ?? 0) - (num(a.created_at) ?? 0))[0];
    if (record) {
      const timestamp = num(record.created_at);
      if (timestamp && timestamp > 0) return {
        timestamp, id: record.id ?? null, category: record.category || 'user_rebalancing',
        source: 'history', savedAt: Date.now()
      };
    }
    // 不回退 current.json 的快照时间，避免显示一个看似精确却与网页不一致的日期。
    return { timestamp: null, id: null, category: null, source: 'history', savedAt: Date.now() };
  }

  // 单条/批量同步刷新调仓历史和当前持仓。两个接口各请求一次，互不影响各自缓存写入。
  // 每个 worker 同一时间仅执行一个请求，因此并发上限仍然是 RB_CONCURRENCY=2。
  async function refreshRebalancing(symbols, mode = 'all') {
    if (state.busy || state.auxBusy || !state.data.size) return;
    const unique = [...new Set(symbols)].filter(sym => state.data.has(sym));
    const until = cooldownUntil();
    if (until) {
      setStatus(`雪球调仓接口此前出现频率限制（110017），本地查询暂停至 ${formatDate(until)}（北京时间）；请稍后再试，已缓存的数据不受影响。`, true);
      return;
    }
    const targets = mode === 'single' ? unique : unique.filter(sym => getQuote(sym)?.status !== 'closed' && (!isRbFresh(sym) || !isHoldingFresh(sym)));
    const countHistory = targets.filter(sym => mode === 'single' || !isRbFresh(sym)).length;
    const countHoldings = targets.filter(sym => mode === 'single' || !isHoldingFresh(sym)).length;
    if (!targets.length) {
      setStatus(mode === 'single' ? '找不到对应组合。' : '未关停组合的调仓和持仓缓存都在有效期内，不需要重复请求。');
      return;
    }
    if (mode !== 'single' && targets.length > 100 &&
      !window.confirm(`将检查 ${targets.length} 个未关停或状态未知的组合，预计 ${countHistory} 次调仓 + ${countHoldings} 次持仓请求（不计临时失败重试），仅补齐缺失或过期的数据。\n批量串行，每两次请求至少间隔约 ${RB_REQUEST_INTERVAL/1000} 秒；已关停不请求。是否继续？`)) return;

    const controller = new AbortController();
    const currentRun = state.runId;
    state.auxController = controller;
    setAuxBusy(true);
    let cursor = 0, completed = 0, rbOk = 0, holdingOk = 0, rbFailed = 0, holdingFailed = 0;
    let throttled = false, pauseUntil = 0;
    const attemptedHistory = new Set();
    const kind = mode === 'single' ? '单条' : '全部未关停';
    setStatus(`正在${kind}更新：0 / ${targets.length}（调仓 ${countHistory}、持仓 ${countHoldings}；串行限速）…`);
    async function checkedRequest(path) {
      const data = await requestJson(path, controller.signal);
      if (Number(data?.error_code) === 110017) {
        const err = new Error(`雪球接口限流 110017：${data.error_description || '操作过于频繁，请稍后再试'}`);
        err.rateLimited = true;
        throw err;
      }
      if (data?.error_code && data.error_code !== 0) throw new Error(data.error_description || `雪球接口错误 ${data.error_code}`);
      if (data?.success === false || (data?.code != null && data.code !== 200))
        throw new Error(data?.message || `雪球接口 code=${data?.code}`);
      return data;
    }
    async function retryHistory(symbol) {
      for (let attempt = 0; ; attempt++) {
        try {
          return await checkedRequest(`/cubes/rebalancing/history.json?cube_symbol=${encodeURIComponent(symbol)}&count=50&page=1`);
        } catch (e) {
          if (controller.signal.aborted || isAbort(e)) throw e;
          const message = String(e.message || e);
          if (isRateLimitedError(e) || /\b401\b|未登录/.test(message) || attempt >= RB_RETRY_COUNT) throw e;
          await pause(RB_RETRY_BASE_MS * (attempt + 1));
        }
      }
    }
    async function worker() {
      while (!throttled) {
        if (currentRun !== state.runId || controller.signal.aborted) return;
        const index = cursor++;
        if (index >= targets.length) return;
        const symbol = targets[index];
        for (const task of ['history', 'holdings']) {
          if (mode !== 'single' && (task === 'history' ? isRbFresh(symbol) : isHoldingFresh(symbol))) continue;
          if (throttled || currentRun !== state.runId || controller.signal.aborted) break;
          if (task === 'holdings' && mode !== 'single') await pause(RB_REQUEST_INTERVAL);
          if (throttled || currentRun !== state.runId || controller.signal.aborted) break;
          try {
            if (task === 'history') {
              attemptedHistory.add(symbol);
              const data = await retryHistory(symbol);
              if (controller.signal.aborted || currentRun !== state.runId) return;
              state.rebalancingStates.set(symbol, parseLatestRebalance(data));
              state.rebalanceErrors.delete(symbol);
              rbOk++;
            } else {
              const data = await checkedRequest(`/cubes/rebalancing/current.json?cube_symbol=${encodeURIComponent(symbol)}`);
              if (controller.signal.aborted || currentRun !== state.runId) return;
              state.holdingStates.set(symbol, parseCurrentHoldings(data));
              holdingOk++;
            }
          } catch (e) {
            if (controller.signal.aborted || isAbort(e)) return;
            if (task === 'history') {
              rbFailed++;
              state.rebalanceErrors.set(symbol, {
                kind: isRateLimitedError(e) ? 'throttled' : 'error',
                message: String(e.message || e).slice(0, 180), at: Date.now()
              });
            } else holdingFailed++;
            if (isRateLimitedError(e)) {
              throttled = true;
              pauseUntil = startCooldown();
              setStatus(`雪球接口限流，已暂停至 ${formatDate(pauseUntil)}；已成功获取的数据和缓存保留。`, true);
            } else if (mode === 'single') {
              setStatus(`组合 ${symbol} 的${task === 'history' ? '调仓历史' : '当前持仓'}获取失败：${e.message || e}；其余项目可正常保存。`, true);
            }
          }
        }
        completed++;
        if (mode === 'single' || completed % 5 === 0 || completed === targets.length || throttled) {
          cacheSave();
          render();
        }
        if (!throttled) setStatus(`${kind}同步更新：完成 ${completed}/${targets.length}；调仓成功 ${rbOk}、失败 ${rbFailed}；持仓成功 ${holdingOk}、失败 ${holdingFailed}。`);
        if (mode !== 'single' && cursor < targets.length) await pause(RB_REQUEST_INTERVAL);
      }
    }
    try {
      await Promise.all(Array.from({length: Math.min(RB_CONCURRENCY, targets.length)}, worker));
      if (currentRun !== state.runId || controller.signal.aborted) return;
      const message = throttled
        ? `雪球接口限流，已暂停：调仓成功 ${rbOk}、持仓成功 ${holdingOk}；已保存结果，其余保留旧缓存。`
        : `${kind}同步更新完成：调仓成功 ${rbOk}/${targets.length}、失败 ${rbFailed}；当前持仓成功 ${holdingOk}/${targets.length}、失败 ${holdingFailed}。当前仓位和核验状态已更新。`;
      setStatus(message, throttled || rbFailed > 0 || holdingFailed > 0);
    } finally {
      if (throttled && mode !== 'single' && !controller.signal.aborted) {
        for (const symbol of targets) {
          if (!attemptedHistory.has(symbol) && !isRbFresh(symbol)) {
            state.rebalanceErrors.set(symbol, {
              kind: 'skipped',
              message: '批量任务遇到雪球接口限流后停止：本组合尚未执行调仓请求',
              at: Date.now()
            });
          }
        }
      }
      cacheSave();
      if (state.auxController === controller) {state.auxController = null; setAuxBusy(false); render();}
    }
  }

  function updateAllRebalancing() {
    return refreshRebalancing([...state.data.keys()], 'all');
  }

  function stopSearch() {
    if (!state.busy && !state.auxBusy) return;
    state.runId++;
    state.controller?.abort();
    state.auxController?.abort();
    setBusy(false);
    setAuxBusy(false);
    render();
    setStatus(`已停止请求。保留前 ${state.page} 页的 ${state.data.size} 个组合和已加载扩展信息。`);
  }

  async function startSearch(force = false) {
    const query = ui.stock.value.trim();
    if (!query) { setStatus('请先输入股票中文名称。', true); ui.stock.focus(); return; }
    if (/^(SH|SZ|BJ|HK|US)?\d{4,6}$/i.test(query)) {
      setStatus('请填写中文股票名称（例如：震裕科技），接口按纯股票代码通常查不到持仓组合。', true);
      return;
    }
    if (state.busy || state.auxBusy) return;
    state.runId++;
    const currentRun = state.runId;
    const controller = new AbortController();
    state.controller = controller;
    state.stock = query;
    state.page = 0;
    state.totalPages = null;
    state.data.clear();
    state.completed = false;
    state.truncated = false;
    state.lastError = '';
    state.searchSource = '';
    state.searchSavedAt = null;
    const cached = !force ? getCachedSearch(query) : null;
    if (cached) {
      for (const row of cached.rows) state.data.set(row.symbol, row);
      state.page = cached.page || 0;
      state.totalPages = cached.totalPages ?? null;
      state.completed = true;
      state.searchSource = 'cache';
      state.searchSavedAt = cached.savedAt;
      updateStats();
      render();
      setStatus(`命中 ${query} 的完整列表缓存：${state.data.size} 个组合，获取于 ${formatDate(cached.savedAt)}；可点击“强制刷新列表”获取最新数据。`);
      if (state.data.size) await loadQuoteAll(false, true); // 只刷新已过期的状态，不重复查询股票列表。
      return;
    }
    state.searchSource = 'live';
    setBusy(true);
    updateStats();
    render();
    setStatus(force ? '强制刷新：正在重新请求第 1 页…' : '正在请求第 1 页…');
    try {
      for (let page = 1; page <= MAX_PAGES; page++) {
        if (currentRun !== state.runId || controller.signal.aborted) return;
        const params = new URLSearchParams({ q: query, count: String(PAGE_SIZE), page: String(page) });
        const payload = await requestJson(`${API}?${params.toString()}`, controller.signal);
        if (currentRun !== state.runId || controller.signal.aborted) return;
        if (payload?.success === false || (payload?.code != null && payload.code !== 200)) {
          throw new Error(payload?.message || `接口返回 code=${payload?.code}`);
        }
        if (!Array.isArray(payload?.data?.list)) throw new Error('返回数据缺少 data.list；雪球可能调整了接口。');
        const list = payload.data.list;
        for (const item of list) {
          if (item?.symbol && item.weight != null) state.data.set(item.symbol, item);
        }
        state.page = page;
        const m = payload.meta || {};
        state.totalPages = Number.isFinite(Number(m.max_page)) && Number(m.max_page) > 0
          ? Number(m.max_page) : null;
        updateStats();
        render();
        const more = m.has_next_page === true || (state.totalPages !== null && page < state.totalPages);
        setStatus(`第 ${page}${state.totalPages ? ` / ${state.totalPages}` : ''} 页完成，当前 ${state.data.size} 个不同的持仓组合。`);
        if (!more || list.length === 0) { state.completed = true; break; }
        if (page === MAX_PAGES) { state.truncated = true; break; }
        await pause(REQUEST_INTERVAL);
      }
      if (currentRun !== state.runId) return;
      if (state.completed && !state.truncated) storeCompletedSearch(query);
      setStatus(state.truncated
        ? `已达到 ${MAX_PAGES} 页安全上限；共 ${state.data.size} 个组合，后续页未读取，不写入缓存。`
        : `查询完成：${query}，${state.data.size} 个持仓组合，共 ${state.page} 页。${ttlHours() ? `列表已保存至本地（有效期 ${ttlHours()} 小时）。` : '列表缓存已关闭。'}`);
      updateStats();
      if (state.data.size && currentRun === state.runId && !controller.signal.aborted) {
        await loadQuoteAll(false, true);
      }
    } catch (e) {
      if (currentRun !== state.runId || controller.signal.aborted) return;
      state.lastError = e.message || String(e);
      setStatus(`查询中断（第 ${state.page + 1} 页）：${state.lastError}。已获取的数据保留，可先导出。`, true);
    } finally {
      if (currentRun === state.runId) { setBusy(false); updateStats(); render(); }
    }
  }

  // 零第三方依赖的 XLSX 导出：创建合法的 Office Open XML 压缩包。
  function xml(s) {
    return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' })[c]);
  }
  function excelCol(n) {
    let out = '';
    for (let x = n + 1; x > 0; x = Math.floor((x - 1) / 26)) out = String.fromCharCode(65 + (x - 1) % 26) + out;
    return out;
  }
  function zipStored(files) {
    const encoder = new TextEncoder();
    const local = [], central = [];
    let offset = 0;
    const table = new Uint32Array(256);
    for (let i = 0; i < 256; i++) {
      let c = i;
      for (let j = 0; j < 8; j++) c = c & 1 ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      table[i] = c >>> 0;
    }
    const crc32 = bytes => {
      let crc = 0xFFFFFFFF;
      for (const b of bytes) crc = table[(crc ^ b) & 255] ^ (crc >>> 8);
      return (crc ^ 0xFFFFFFFF) >>> 0;
    };
    function header(size) {
      const b = new Uint8Array(size);
      const v = new DataView(b.buffer);
      return { b, v };
    }
    const concat = arr => {
      const result = new Uint8Array(arr.reduce((sum, b) => sum + b.length, 0));
      let n = 0;
      for (const bytes of arr) { result.set(bytes, n); n += bytes.length; }
      return result;
    };
    for (const [filename, contents] of files) {
      const name = encoder.encode(filename);
      const body = encoder.encode(contents);
      const crc = crc32(body);
      const lh = header(30);
      lh.v.setUint32(0, 0x04034B50, true); lh.v.setUint16(4, 20, true); lh.v.setUint16(6, 0x0800, true);
      lh.v.setUint16(8, 0, true); lh.v.setUint32(14, crc, true);
      lh.v.setUint32(18, body.length, true); lh.v.setUint32(22, body.length, true); lh.v.setUint16(26, name.length, true);
      local.push(lh.b, name, body);
      const ch = header(46);
      ch.v.setUint32(0, 0x02014B50, true); ch.v.setUint16(4, 20, true); ch.v.setUint16(6, 20, true);
      ch.v.setUint16(8, 0x0800, true); ch.v.setUint16(10, 0, true); ch.v.setUint32(16, crc, true);
      ch.v.setUint32(20, body.length, true); ch.v.setUint32(24, body.length, true);
      ch.v.setUint16(28, name.length, true); ch.v.setUint32(42, offset, true);
      central.push(ch.b, name);
      offset += lh.b.length + name.length + body.length;
    }
    const centralSize = central.reduce((n, bytes) => n + bytes.length, 0);
    const end = header(22);
    end.v.setUint32(0, 0x06054B50, true);
    end.v.setUint16(8, files.length, true); end.v.setUint16(10, files.length, true);
    end.v.setUint32(12, centralSize, true); end.v.setUint32(16, offset, true);
    return concat([...local, ...central, end.b]);
  }
  function makeXlsx(rows) {
    const headers = ['组合名称', '组合代码', '主理人', '搜索权重(%)', '当前持仓(%)', '持仓核验状态', '持仓核验时间', '累计收益(%)', '年化收益(%)', '近30日收益(%)', '最大回撤(%)', '关停状态', '关停时间', '最近实际调仓', '调仓数据状态', '调仓核验时间', '组合链接', '查询股票', '导出时间'];
    const time = new Date().toLocaleString('zh-CN', { hour12: false });
    const data = [headers, ...rows.map(r => [
      r.name, r.symbol, r.owner_name, num(r.weight),
      ['held', 'absent'].includes(holdingCheck(r.symbol).status) ? holdingCheck(r.symbol).weight : null,
      ({held:'当前持有', absent:'当前未持有', stale:'旧缓存需更新', unknown:'无法确认', pending:'待核验'})[holdingCheck(r.symbol).status],
      formatDate(getHolding(r.symbol)?.savedAt),
      num(r.total_gain), num(r.annualized_gain_rate), num(r.monthly_gain), num(r.draw_down),
      statusText(r.symbol), formatDate(getQuote(r.symbol)?.closedAt),
      formatDate(getRb(r.symbol)?.timestamp),
      !getRb(r.symbol) ? '待加载' : isRbFresh(r.symbol) ? '24小时内已核验' : '旧缓存（超24小时）',
      formatDate(getRb(r.symbol)?.savedAt),
      `https://xueqiu.com/P/${r.symbol}`, state.stock, time
    ])];
    const lastCell = `${excelCol(headers.length - 1)}${data.length}`;
    const widths = [26, 17, 22, 15, 15, 17, 23, 15, 15, 17, 15, 16, 23, 23, 21, 23, 44, 20, 23];
    const sheetCols = widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('');
    const lines = data.map((row, ri) => `<row r="${ri + 1}">${row.map((value, ci) => {
      const addr = `${excelCol(ci)}${ri + 1}`;
      const style = ri === 0 ? ' s="1"' : '';
      if (typeof value === 'number' && Number.isFinite(value)) {
        return `<c r="${addr}"${style}><v>${value}</v></c>`;
      }
      const safe = xml(value == null ? '' : value);
      return `<c r="${addr}" t="inlineStr"${style}><is><t xml:space="preserve">${safe}</t></is></c>`;
    }).join('')}</row>`).join('');
    const worksheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:${lastCell}"/>` +
      `<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>` +
      `<cols>${sheetCols}</cols><sheetData>${lines}</sheetData><autoFilter ref="A1:${lastCell}"/></worksheet>`;
    const files = [
      ['[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`],
      ['_rels/.rels', `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`],
      ['xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="持仓组合" sheetId="1" r:id="rId1"/></sheets></workbook>`],
      ['xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`],
      ['xl/styles.xml', `<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="微软雅黑"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="微软雅黑"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1766BA"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs></styleSheet>`],
      ['xl/worksheets/sheet1.xml', worksheet]
    ];
    return new Blob([zipStored(files)], { type: FILE_MIME });
  }
  function exportXlsx() {
    if (!state.data.size) return;
    try {
      const rows = sortableRows();
      if (!rows.length) { setStatus('当前筛选结果为空，无法导出。', true); return; }
      const blob = makeXlsx(rows);
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      const safeName = state.stock.replace(/[\\/:*?"<>|]/g, '_');
      const ymd = new Date().toLocaleDateString('en-CA');
      link.download = `雪球持仓组合_${safeName}_${ymd}.xlsx`;
      document.body.appendChild(link);
      link.click();
      link.remove();
      setTimeout(() => URL.revokeObjectURL(url), 2500);
      setStatus(`已生成 ${rows.length} 条记录的 Excel 文件（当前排序及筛选结果）。`);
    } catch (e) { setStatus(`导出 XLSX 失败：${e.message}`, true); }
  }

  ui.search.addEventListener('click', () => startSearch(false));
  ui.forceSearch.addEventListener('click', () => startSearch(true));
  ui.searchTtl.addEventListener('change', () => {
    writeSetting(SEARCH_TTL_KEY, ttlHours());
    setStatus(ttlHours()
      ? `股票列表缓存时长已设置为 ${ttlHours()} 小时，下一次搜索生效；可点击“强制刷新列表”跳过缓存。`
      : '股票列表缓存已关闭，下一次查询会实时搜索；已有缓存暂时保留，点击“清理缓存”可删除。');
  });
  ui.stop.addEventListener('click', stopSearch);
  ui.export.addEventListener('click', exportXlsx);
  ui.sort.addEventListener('change', render);
  ui.filter.addEventListener('input', render);
  ui.statusFilter.addEventListener('change', render);
  ui.holdingFilter.addEventListener('change', render);
  ui.quoteRefresh.addEventListener('click', () => loadQuoteAll(true));
  ui.loadRebalance.addEventListener('click', updateAllRebalancing);
  ui.clearCache.addEventListener('click', clearCache);
  ui.stock.addEventListener('keydown', e => { if (e.key === 'Enter') startSearch(false); });
  initializeSearchSettings();
  cacheLoad();
  updateStats();
  render();
})();