const BASE_METRICS = [
  { id: 'db_count', label: 'Databases', category: 'Inventory', fmt: 'int', summary: true },
  { id: 'host_count', label: 'Hosts', category: 'Inventory', fmt: 'int', summary: true },
  { id: 'instance_count', label: 'Instances', category: 'Inventory', fmt: 'int', summary: true },
  { id: 'vcpu_total', label: 'vCPU (sum)', category: 'CPU', fmt: 'dec1', summary: true },
  { id: 'memory_gb_total', label: 'Memory (GB)', category: 'Memory', fmt: 'dec1', summary: true },
  { id: 'allocated_storage_gb', label: 'Allocated Storage (GB)', category: 'Storage', fmt: 'dec1', summary: true },
  { id: 'used_storage_gb', label: 'Used Storage (GB)', category: 'Storage', fmt: 'dec1', summary: true },
  { id: 'db_iops_total', label: 'DB IOPS (rollup)', category: 'IO', fmt: 'dec1', summary: true },
  { id: 'db_logons_total', label: 'DB Logons (rollup)', category: 'Activity', fmt: 'dec1', summary: true },
];

const SPECIAL_METRICS = [
  { id: 'topsql', label: 'Top SQL panel', category: 'Advanced', fmt: 'none', summary: false },
  { id: 'segmentio', label: 'Segment IO panel', category: 'Advanced', fmt: 'none', summary: false },
];

const state = {
  raw: null,
  graph: null,
  metricCatalog: [],
  defaultMetricIds: new Set(),
  activeTab: 'web',
  linkSelections: true,
  selection: {
    web: { cohorts: new Set(), dbs: new Set(), instances: new Set(), metrics: new Set() },
    ppt: { cohorts: new Set(), dbs: new Set(), instances: new Set(), metrics: new Set() },
  },
  drill: { cohort: null, db: null, instance: null },
  sort: {
    cohortTable: { key: 'cohort', dir: 'asc' },
    dbStatsTable: { key: 'metric', dir: 'asc' },
    topSqlTable: { key: 'elapsed', dir: 'desc' },
    segmentIoTable: { key: 'physical_io_tot', dir: 'desc' },
  },
  page: { mode: 'global', cohort: null, instance: null },
  cpuScaleMode: 'dynamic',
  vcpuSizingTargetPct: null,
  memorySizingTargetPct: null,
  iopsSizingTargetPct: null,
  storageSizingTargetPct: null,
  cohortTargets: {},
  layoutEditMode: false,
  layoutDefaultOrder: [],
  layoutDraggingId: null,
  chartSizes: {},
  cardSpans: {},
  chartResizeSession: null,
  cardResizeSession: null,
  isResizingChart: false,
  cardHeights: {},
  chartTypes: {},
  pptSlidePlan: [],
  pptSlideSelected: new Set(),
  previewSlideId: null,
  previewChartCache: {},
  previewRenderToken: 0,
  exportInProgress: false,
  templateApiAvailable: null,
  templatePath: '',
  advisorExportProfile: 'compact',
  referenceLibrary: [],
  reportMeta: {
    customerName: '',
    opportunityNumber: '',
    salesRepName: '',
    architectName: '',
    engineerName: '',
  },
  collapsedSections: {},
  sidebarCollapsed: false,
  aiChat: {
    panelOpen: false,
    connected: false,
    connectionSource: 'unknown',
    connectionError: '',
    model: 'gpt-4.1-mini',
    messages: [],
    docs: [],
    question: '',
    prompt: '',
    response: '',
    appPayload: null,
    payloadErrors: [],
    busy: false,
  },
};

const $ = (id) => document.getElementById(id);
const LINE_COLORS = ['#b11f1f', '#0f766e', '#1d4ed8', '#b45309', '#7c3aed', '#be185d', '#0e7490', '#4d7c0f', '#374151'];
const ITEM_COLORS = ['#b11f1f', '#0f766e', '#1d4ed8', '#b45309', '#7c3aed', '#be185d', '#0891b2', '#4d7c0f', '#475569', '#dc2626', '#0284c7', '#16a34a'];
const TEMPLATE_EXPORT_ONLY = true;
const LAYOUT_STORAGE_KEY = 'awr_review_layout_v1';
const CHART_SIZE_STORAGE_KEY = 'awr_review_chart_sizes_v1';
const CARD_SPAN_STORAGE_KEY = 'awr_review_card_spans_v1';
const CARD_HEIGHT_STORAGE_KEY = 'awr_review_card_heights_v1';
const CHART_TYPE_STORAGE_KEY = 'awr_review_chart_types_v1';
const SIDEBAR_COLLAPSED_STORAGE_KEY = 'awr_review_sidebar_collapsed_v1';
const CHATBOT_ENABLED = false;
const INFRA_TIERS = ['Base Database', 'Exascale', 'Exadata Dedicated', 'Exadata Cloud Services Dedicated'];
const AWR_DBA_ADVISOR_URL = 'https://chatgpt.com/g/g-69d9751bd6008191afe58d05e76405b6-awr-dba-advisor';
const TARGET_METRICS = ['vcpu', 'memory', 'iops', 'storage'];

function normalizeReportMeta(meta = {}) {
  const src = meta && typeof meta === 'object' ? meta : {};
  return {
    customerName: String(src.customerName || '').trim(),
    opportunityNumber: String(src.opportunityNumber || '').trim(),
    salesRepName: String(src.salesRepName || '').trim(),
    architectName: String(src.architectName || '').trim(),
    engineerName: String(src.engineerName || '').trim(),
  };
}

function applyReportMetaToInputs() {
  if ($('metaCustomerName')) $('metaCustomerName').value = state.reportMeta.customerName || '';
  if ($('metaOpportunityNumber')) $('metaOpportunityNumber').value = state.reportMeta.opportunityNumber || '';
  if ($('metaSalesRepName')) $('metaSalesRepName').value = state.reportMeta.salesRepName || '';
  if ($('metaArchitectName')) $('metaArchitectName').value = state.reportMeta.architectName || '';
  if ($('metaEngineerName')) $('metaEngineerName').value = state.reportMeta.engineerName || '';
}

function applySidebarCollapsedUi() {
  const layout = $('mainLayout');
  if (layout) layout.classList.toggle('sidebar-collapsed', Boolean(state.sidebarCollapsed));
  const btn = $('sidebarToggleBtn');
  if (btn) {
    btn.textContent = '☰';
    btn.setAttribute('aria-label', state.sidebarCollapsed ? 'Show filters' : 'Hide filters');
    btn.title = state.sidebarCollapsed ? 'Show filters' : 'Hide filters';
  }
}

function saveSidebarCollapsedPref() {
  try {
    localStorage.setItem(SIDEBAR_COLLAPSED_STORAGE_KEY, state.sidebarCollapsed ? '1' : '0');
  } catch {
    // ignore
  }
}

function loadSidebarCollapsedPref() {
  try {
    const raw = localStorage.getItem(SIDEBAR_COLLAPSED_STORAGE_KEY);
    state.sidebarCollapsed = raw === '1';
  } catch {
    state.sidebarCollapsed = false;
  }
  applySidebarCollapsedUi();
}

function fmt(n, kind) {
  if (n == null || Number.isNaN(Number(n))) return 'N/A';
  const value = Number(n);
  if (kind === 'int') return new Intl.NumberFormat().format(Math.round(value));
  if (kind === 'dec1') return new Intl.NumberFormat(undefined, { maximumFractionDigits: 1 }).format(value);
  if (kind === 'dec3') return new Intl.NumberFormat(undefined, { maximumFractionDigits: 3 }).format(value);
  return String(value);
}

function fmtBytes(bytes) {
  const b = Number(bytes) || 0;
  if (b <= 0) return '0 B';
  if (b < 1024) return `${b} B`;
  if (b < 1024 * 1024) return `${(b / 1024).toFixed(1)} KB`;
  if (b < 1024 * 1024 * 1024) return `${(b / (1024 * 1024)).toFixed(1)} MB`;
  return `${(b / (1024 * 1024 * 1024)).toFixed(2)} GB`;
}

function cpuFixedMax() {
  return state.cpuScaleMode === 'fixed100' ? 100 : null;
}

function esc(s) {
  return String(s)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

function encodeKey(value) {
  return encodeURIComponent(String(value ?? ''));
}

function decodeKey(value) {
  try {
    return decodeURIComponent(String(value ?? ''));
  } catch {
    return String(value ?? '');
  }
}

function hashString(input) {
  let h = 2166136261;
  const s = String(input ?? '');
  for (let i = 0; i < s.length; i += 1) {
    h ^= s.charCodeAt(i);
    h = Math.imul(h, 16777619);
  }
  return Math.abs(h >>> 0);
}

function colorForItem(item) {
  const idx = hashString(item) % ITEM_COLORS.length;
  return ITEM_COLORS[idx];
}

function decodeFloat64B64(b64) {
  try {
    const normalized = String(b64)
      .replaceAll('\\u002f', '/')
      .replaceAll('\\u002b', '+')
      .replaceAll('\\u003d', '=');
    const padded = normalized + '='.repeat((4 - (normalized.length % 4)) % 4);
    const bin = atob(padded);
    const buffer = new ArrayBuffer(bin.length);
    const bytes = new Uint8Array(buffer);
    for (let i = 0; i < bin.length; i += 1) bytes[i] = bin.charCodeAt(i);
    const view = new DataView(buffer);
    const out = [];
    for (let i = 0; i < bin.length; i += 8) out.push(view.getFloat64(i, true));
    return out;
  } catch {
    return [];
  }
}

function parseBlackLineSeriesFromPlotHtml(fileContent) {
  const blackRe = /"line":\{"color":"black","width":\d+\},"mode":"lines","x":(\[[\s\S]*?\]),"y":(?:\{"dtype":"f8","bdata":"([^"]+)"\}|(\[[\s\S]*?\]))[\s\S]*?"type":"scatter"/;
  const genericTimeRe = /"x":(\["\d{4}-\d{2}-\d{2}T[\s\S]*?\]),"y":(?:\{"dtype":"f8","bdata":"([^"]+)"\}|(\[[\s\S]*?\]))[\s\S]*?"type":"scatter"/;
  const m = fileContent.match(blackRe) || fileContent.match(genericTimeRe);
  if (!m) return null;
  let x = [];
  let y = [];
  try {
    x = JSON.parse(m[1]);
  } catch {
    x = [];
  }
  if (m[2]) y = decodeFloat64B64(m[2]);
  else if (m[3]) {
    try {
      y = JSON.parse(m[3]);
    } catch {
      y = [];
    }
  }
  if (!x.length || !y.length) return null;
  const n = Math.min(x.length, y.length);
  return { x: x.slice(0, n), y: y.slice(0, n) };
}

function parseCpuTimeSeries(raw) {
  const byCohort = {};
  for (const plot of raw.plots_html || []) {
    const name = plot.file_name || '';
    if (!name.endsWith('_DB_vCPU_unadjusted.html')) continue;
    const cohortKey = name.replace('_DB_vCPU_unadjusted.html', '');
    const parsed = parseBlackLineSeriesFromPlotHtml(plot.file_content || '');
    if (!parsed) continue;
    byCohort[cohortKey] = parsed;
    if (plot.cohort && !byCohort[plot.cohort]) byCohort[plot.cohort] = parsed;
  }
  return byCohort;
}

function resolveCpuSeriesForCohort(cohort) {
  const map = state.graph.cpuSeriesByCohort || {};
  if (map[cohort]) return map[cohort];
  const keys = Object.keys(map);
  const exact = keys.find((k) => k.toUpperCase() === String(cohort).toUpperCase());
  if (exact) return map[exact];
  const fuzzy = keys.find((k) => k.startsWith(cohort) || String(cohort).startsWith(k));
  if (fuzzy) return map[fuzzy];
  return null;
}

function setStatus(text) {
  $('fileStatus').textContent = text;
}

function truncateText(s, max = 4000) {
  const t = String(s || '');
  if (t.length <= max) return t;
  return `${t.slice(0, max)}\n...[truncated ${t.length - max} chars]`;
}

function escapeForPromptJson(v) {
  try {
    return JSON.stringify(v, null, 2);
  } catch {
    return String(v);
  }
}

function isObj(x) {
  return x && typeof x === 'object' && !Array.isArray(x);
}

function validateStringArray(arr, path, errors, maxItems = 20) {
  if (!Array.isArray(arr)) {
    errors.push(`${path} must be an array.`);
    return;
  }
  if (arr.length > maxItems) errors.push(`${path} must have at most ${maxItems} items.`);
  arr.forEach((v, i) => {
    if (typeof v !== 'string' || !v.trim()) errors.push(`${path}[${i}] must be a non-empty string.`);
  });
}

function validateMapOfStringArrays(obj, path, errors, maxItems = 12) {
  if (!isObj(obj)) {
    errors.push(`${path} must be an object.`);
    return;
  }
  Object.entries(obj).forEach(([k, v]) => {
    if (!String(k).trim()) errors.push(`${path} has an empty key.`);
    validateStringArray(v, `${path}.${k}`, errors, maxItems);
  });
}

function validateAppPayload(payload) {
  const errors = [];
  if (!isObj(payload)) {
    return { valid: false, errors: ['Payload must be a JSON object.'] };
  }
  const required = [
    'global_comments',
    'cohort_comments',
    'instance_comments',
    'infrastructure_recommendation',
    'assumptions',
    'data_gaps',
  ];
  required.forEach((k) => {
    if (!(k in payload)) errors.push(`Missing required field: ${k}.`);
  });
  const allowedTop = new Set(required);
  Object.keys(payload).forEach((k) => {
    if (!allowedTop.has(k)) errors.push(`Unexpected top-level field: ${k}.`);
  });

  validateStringArray(payload.global_comments, 'global_comments', errors, 12);
  validateMapOfStringArrays(payload.cohort_comments, 'cohort_comments', errors, 12);
  validateMapOfStringArrays(payload.instance_comments, 'instance_comments', errors, 12);
  validateStringArray(payload.assumptions, 'assumptions', errors, 20);
  validateStringArray(payload.data_gaps, 'data_gaps', errors, 20);

  const infra = payload.infrastructure_recommendation;
  if (!isObj(infra)) {
    errors.push('infrastructure_recommendation must be an object.');
  } else {
    const infraReq = [
      'recommended_tier',
      'scale_position',
      'rationale',
      'deployment_plan',
      'cohort_tier_suggestions',
      'instance_tier_suggestions',
      'confidence',
    ];
    infraReq.forEach((k) => {
      if (!(k in infra)) errors.push(`Missing infrastructure_recommendation.${k}.`);
    });
    const allowedInfra = new Set(infraReq);
    Object.keys(infra).forEach((k) => {
      if (!allowedInfra.has(k)) errors.push(`Unexpected infrastructure_recommendation field: ${k}.`);
    });
    if (!INFRA_TIERS.includes(infra.recommended_tier)) {
      errors.push('infrastructure_recommendation.recommended_tier has invalid value.');
    }
    if (![1, 2, 3].includes(infra.scale_position)) {
      errors.push('infrastructure_recommendation.scale_position must be 1, 2, or 3.');
    }
    const expectedPos =
      infra.recommended_tier === 'Base Database'
        ? 1
        : infra.recommended_tier === 'Exascale'
          ? 2
          : (infra.recommended_tier === 'Exadata Cloud Services Dedicated' || infra.recommended_tier === 'Exadata Dedicated')
            ? 3
            : null;
    if (expectedPos != null && infra.scale_position !== expectedPos) {
      errors.push('infrastructure_recommendation.scale_position does not match recommended_tier.');
    }
    validateStringArray(infra.rationale, 'infrastructure_recommendation.rationale', errors, 12);
    if (!isObj(infra.deployment_plan)) {
      errors.push('infrastructure_recommendation.deployment_plan must be an object.');
    } else {
      if (infra.deployment_plan.base_database !== 'one Base Database per instance') {
        errors.push('deployment_plan.base_database invalid.');
      }
      if (infra.deployment_plan.exascale !== 'one Exascale deployment per cohort') {
        errors.push('deployment_plan.exascale invalid.');
      }
      if (infra.deployment_plan.exadata_dedicated !== 'all cohorts consolidated in Exadata Dedicated') {
        errors.push('deployment_plan.exadata_dedicated invalid.');
      }
    }
    if (!isObj(infra.cohort_tier_suggestions)) {
      errors.push('infrastructure_recommendation.cohort_tier_suggestions must be an object.');
    } else {
      Object.entries(infra.cohort_tier_suggestions).forEach(([k, v]) => {
        if (!INFRA_TIERS.includes(v)) errors.push(`cohort_tier_suggestions.${k} invalid tier.`);
      });
    }
    if (!isObj(infra.instance_tier_suggestions)) {
      errors.push('infrastructure_recommendation.instance_tier_suggestions must be an object.');
    } else {
      Object.entries(infra.instance_tier_suggestions).forEach(([k, v]) => {
        if (!INFRA_TIERS.includes(v)) errors.push(`instance_tier_suggestions.${k} invalid tier.`);
      });
    }
    if (typeof infra.confidence !== 'number' || Number.isNaN(infra.confidence) || infra.confidence < 0 || infra.confidence > 1) {
      errors.push('infrastructure_recommendation.confidence must be a number between 0 and 1.');
    }
  }
  return { valid: errors.length === 0, errors };
}

function stripJsonCodeFence(text) {
  const raw = String(text || '').trim();
  const m = raw.match(/^```(?:json)?\s*([\s\S]*?)\s*```$/i);
  return m ? m[1].trim() : raw;
}

function updateTemplateStatusText() {
  const node = $('templateStatus');
  if (!node) return;
  if (!state.templateApiAvailable) {
    node.textContent = 'Active Template: server unavailable';
    return;
  }
  const p = String(state.templatePath || '').trim();
  if (!p) {
    node.textContent = 'Active Template: connected';
    return;
  }
  const parts = p.split(/[\\/]/);
  const file = parts[parts.length - 1] || p;
  node.textContent = `Active Template: ${file}`;
}

function renderReferenceLibrary() {
  const list = $('referenceBookList');
  const status = $('referenceBookStatus');
  if (!list || !status) return;
  const docs = Array.isArray(state.referenceLibrary) ? state.referenceLibrary : [];
  if (!docs.length) {
    list.innerHTML = '<p class="muted">No reference docs found in <code>docs/</code>.</p>';
    status.textContent = 'Reference library is empty.';
    return;
  }
  status.textContent = `${fmt(docs.length, 'int')} document(s) found in docs/.`;
  list.innerHTML = docs
    .map((d) => {
      const name = d?.name || d?.path || 'Document';
      const url = String(d?.url || '').trim();
      const size = fmtBytes(d?.size_bytes || 0);
      return `<div class="reference-item"><a href="${escAttr(url)}" target="_blank" rel="noopener" title="${escAttr(name)}">${esc(
        name,
      )}</a><span class="reference-meta">${esc(size)}</span></div>`;
    })
    .join('');
}

async function loadReferenceLibrary() {
  const status = $('referenceBookStatus');
  if (status) status.textContent = 'Loading reference library...';
  try {
    const res = await fetch('/api/reference-library', { method: 'GET' });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const data = await res.json().catch(() => ({}));
    state.referenceLibrary = Array.isArray(data?.docs) ? data.docs : [];
  } catch {
    state.referenceLibrary = [];
    if (status) status.textContent = 'Reference library unavailable. Ensure app_server.py is running.';
  }
  renderReferenceLibrary();
}

function selectionToPojo(sel) {
  return {
    cohorts: [...(sel?.cohorts || [])],
    dbs: [...(sel?.dbs || [])],
    instances: [...(sel?.instances || [])],
    metrics: [...(sel?.metrics || [])],
  };
}

function selectionFromPojo(rawSel, fallback) {
  return {
    cohorts: new Set(Array.isArray(rawSel?.cohorts) ? rawSel.cohorts : [...fallback.cohorts]),
    dbs: new Set(Array.isArray(rawSel?.dbs) ? rawSel.dbs : [...fallback.dbs]),
    instances: new Set(Array.isArray(rawSel?.instances) ? rawSel.instances : [...fallback.instances]),
    metrics: new Set(Array.isArray(rawSel?.metrics) ? rawSel.metrics : [...fallback.metrics]),
  };
}

function setDefaultsForLoadedRaw(raw) {
  state.raw = raw;
  state.graph = normalize(raw);
  state.metricCatalog = buildMetricCatalog(raw);

  const defaultIds = [
    ...BASE_METRICS.map((m) => m.id),
    ...state.metricCatalog
      .filter((m) => m.id.startsWith('dbstat:') && ['Allocated Storage (GB)', 'DB IOPS', 'DB vCPU', 'DB Memory (MB)'].includes(m.dbStatName))
      .map((m) => m.id),
    'topsql',
  ];
  state.defaultMetricIds = new Set(defaultIds.filter((id) => state.metricCatalog.some((m) => m.id === id)));

  ['web', 'ppt'].forEach((k) => {
    state.selection[k] = {
      cohorts: new Set(state.graph.cohorts),
      dbs: new Set(state.graph.dbs),
      instances: new Set(state.graph.instances.map((x) => x.instance)),
      metrics: new Set(state.defaultMetricIds),
    };
  });

  state.drill = { cohort: null, db: null, instance: null };
  state.page = { mode: 'global', cohort: null, instance: null };
  state.vcpuSizingTargetPct = 100;
  state.memorySizingTargetPct = 100;
  state.iopsSizingTargetPct = 100;
  state.storageSizingTargetPct = 100;
  state.cohortTargets = {};
  state.reportMeta = normalizeReportMeta({});
  state.collapsedSections = {};
  applyReportMetaToInputs();
}

function enableReportActions() {
  const saveBtn = $('saveReportBtn');
  if (saveBtn) saveBtn.disabled = !state.raw;
  const advisorBtn = $('exportAdvisorJsonBtn');
  if (advisorBtn) advisorBtn.disabled = !state.raw;
  const exportBtn = $('exportBtn');
  if (exportBtn) exportBtn.disabled = !state.raw || !state.pptSlideSelected || state.pptSlideSelected.size === 0;
  const previewExportBtn = $('previewExportBtn');
  if (previewExportBtn) previewExportBtn.disabled = state.exportInProgress || !state.raw || !state.pptSlideSelected || state.pptSlideSelected.size === 0;
}

function applyLocalPrefsToStorage() {
  try {
    localStorage.setItem(LAYOUT_STORAGE_KEY, JSON.stringify(getCurrentLayoutOrder()));
    localStorage.setItem(CHART_SIZE_STORAGE_KEY, JSON.stringify(state.chartSizes || {}));
    localStorage.setItem(CARD_SPAN_STORAGE_KEY, JSON.stringify(state.cardSpans || {}));
    localStorage.setItem(CARD_HEIGHT_STORAGE_KEY, JSON.stringify(state.cardHeights || {}));
    localStorage.setItem(CHART_TYPE_STORAGE_KEY, JSON.stringify(state.chartTypes || {}));
  } catch {
    // ignore
  }
}

function reportStateSnapshot() {
  return {
    schema_version: 1,
    saved_at: new Date().toISOString(),
    raw_data: state.raw,
    report_metadata: normalizeReportMeta(state.reportMeta),
    ui_state: {
      activeTab: state.activeTab,
      linkSelections: state.linkSelections,
      selection: {
        web: selectionToPojo(state.selection.web),
        ppt: selectionToPojo(state.selection.ppt),
      },
      cpuScaleMode: state.cpuScaleMode,
      vcpuSizingTargetPct:
        Number.isFinite(state.vcpuSizingTargetPct) ? Number(state.vcpuSizingTargetPct) : null,
      memorySizingTargetPct:
        Number.isFinite(state.memorySizingTargetPct) ? Number(state.memorySizingTargetPct) : null,
      iopsSizingTargetPct:
        Number.isFinite(state.iopsSizingTargetPct) ? Number(state.iopsSizingTargetPct) : null,
      storageSizingTargetPct:
        Number.isFinite(state.storageSizingTargetPct) ? Number(state.storageSizingTargetPct) : null,
      cohortTargets: state.cohortTargets && typeof state.cohortTargets === 'object' ? state.cohortTargets : {},
      layoutOrder: getCurrentLayoutOrder(),
      cardSpans: state.cardSpans || {},
      cardHeights: state.cardHeights || {},
      chartSizes: state.chartSizes || {},
      chartTypes: state.chartTypes || {},
      collapsedSections: state.collapsedSections || {},
      sidebarCollapsed: Boolean(state.sidebarCollapsed),
      advisorExportProfile: state.advisorExportProfile === 'comprehensive' ? 'comprehensive' : 'compact',
      reportMeta: normalizeReportMeta(state.reportMeta),
      pptSlidesSelected: [...(state.pptSlideSelected || [])],
      aiChat: {
        panelOpen: Boolean(state.aiChat?.panelOpen),
        connected: Boolean(state.aiChat?.connected),
        connectionSource: state.aiChat?.connectionSource || 'unknown',
        model: state.aiChat?.model || 'gpt-4.1-mini',
        messages: (state.aiChat?.messages || []).map((m) => ({
          role: m.role === 'assistant' ? 'assistant' : 'user',
          content: String(m.content || ''),
        })),
        docs: (state.aiChat?.docs || []).map((d) => ({ name: d.name, content: d.content })),
        question: state.aiChat?.question || '',
        prompt: state.aiChat?.prompt || '',
        response: state.aiChat?.response || '',
        appPayload: state.aiChat?.appPayload || null,
        payloadErrors: Array.isArray(state.aiChat?.payloadErrors) ? state.aiChat.payloadErrors : [],
      },
    },
  };
}

function downloadJsonFile(payload, prefix, pretty = true) {
  const stamp = new Date().toISOString().replaceAll(':', '-').slice(0, 19);
  const blob = new Blob([JSON.stringify(payload, null, pretty ? 2 : 0)], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `${prefix}_${stamp}.json`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(a.href);
  return a.download;
}

function groupedArray(values, size) {
  const n = Math.max(1, Number(size) || 1);
  const out = [];
  for (let i = 0; i < values.length; i += n) out.push(values.slice(i, i + n));
  return out;
}

function statsFromRowsForMetric(rows, metric) {
  if (metric === 'memory') {
    const vals = rows
      .map((r) => {
        const total = Number(r.mem_gb) || 0;
        const used = (Number(r.sga_size_gb) || 0) + (Number(r.pga_size_gb) || 0);
        if (total <= 0) return null;
        return (used / total) * 100;
      })
      .filter((v) => Number.isFinite(v));
    return statsFromValues(vals);
  }
  return null;
}

function computeByCohortFromRows(rows) {
  const groups = new Map();
  for (const r of rows) {
    if (!groups.has(r.cohort)) {
      groups.set(r.cohort, { cohort: r.cohort, dbs: new Set(), instances: 0, hosts: new Set(), vcpu: 0, mem: 0 });
    }
    const g = groups.get(r.cohort);
    g.dbs.add(r.db);
    g.hosts.add(r.host);
    g.instances += 1;
    g.vcpu += Number(r.init_cpu_count || r.logical_cpu_count || 0);
    g.mem += Number(r.mem_gb || 0);
  }
  return [...groups.values()]
    .map((g) => ({
      cohort: g.cohort,
      dbs: g.dbs.size,
      instances: g.instances,
      hosts: g.hosts.size,
      vcpu: g.vcpu,
      mem: g.mem,
      allocated: state.graph.cohortRollups[g.cohort]?.allocated || 0,
      used: state.graph.cohortRollups[g.cohort]?.used || 0,
      iops: state.graph.cohortRollups[g.cohort]?.db_iops || 0,
      logons: state.graph.cohortRollups[g.cohort]?.db_logons || 0,
    }))
    .sort((a, b) => a.cohort.localeCompare(b.cohort));
}

function computeScopedCpuStats(rows, cohorts) {
  const values = [];
  cohorts.forEach((cohort) => {
    const base = resolveCpuSeriesForCohort(cohort);
    if (!base) return;
    const cRows = rows.filter((r) => r.cohort === cohort);
    if (!cRows.length) return;
    const total = cRows.reduce((a, r) => a + (Number(r.init_cpu_count) || 0), 0);
    const fallbackShare = 1 / cRows.length;
    cRows.forEach((r) => {
      const share = total > 0 ? (Number(r.init_cpu_count) || 0) / total : fallbackShare;
      base.y.forEach((v) => values.push((Number(v) || 0) * share));
    });
  });
  return statsFromValues(values);
}

function computeScopedIopsStats(cohortSet, dbSet) {
  const rows = (state.graph?.dbStats || []).filter(
    (r) =>
      cohortSet.has(r.cohort) &&
      dbSet.has(r.db) &&
      r.metric === 'DB IOPS' &&
      (!r.summary_of || r.summary_of === 'cdb'),
  );
  if (!rows.length) return null;
  const cap = Math.max(
    ...rows.map((r) => Math.max(Number(r.max) || 0, Number(r.p99) || 0, Number(r.p95) || 0)),
    0,
  );
  if (!Number.isFinite(cap) || cap <= 0) return null;
  const values = rows
    .flatMap((r) => [r.min, r.p30, r.p50, r.p70, r.p95, r.p99, r.max])
    .map((v) => ((Number(v) || 0) / cap) * 100)
    .filter((v) => Number.isFinite(v));
  return statsFromValues(values);
}

function computeScopedStorageStats(cohortSet, dbSet) {
  const allocByDb = new Map();
  const usedByDb = new Map();
  (state.graph?.dbStats || []).forEach((r) => {
    if (!cohortSet.has(r.cohort)) return;
    if (!dbSet.has(r.db)) return;
    if (r.summary_of && r.summary_of !== 'cdb') return;
    if (r.metric === 'Allocated Storage (GB)') allocByDb.set(r.db, Number(r.p95 || r.max || 0));
    if (r.metric === 'Used Storage (GB)') usedByDb.set(r.db, Number(r.p95 || r.max || 0));
  });
  const vals = [...allocByDb.keys()]
    .map((db) => {
      const a = allocByDb.get(db) || 0;
      const u = usedByDb.get(db) || 0;
      if (a <= 0) return null;
      return (u / a) * 100;
    })
    .filter((v) => Number.isFinite(v));
  return statsFromValues(vals);
}

function filterRawDataForScope(raw, cohortSet, dbSet, instanceSet) {
  const out = {};
  Object.entries(raw || {}).forEach(([k, v]) => {
    if (!Array.isArray(v)) out[k] = v;
  });
  out.instances = (raw.instances || []).filter((r) => {
    const c = r.cdb_cohort || r.db_cohort || 'UNASSIGNED';
    const db = r.cdb || r.WHICH_DB || 'UNKNOWN_DB';
    const inst = r.db || `${db}$${r.host || 'UNKNOWN_HOST'}`;
    return cohortSet.has(c) && dbSet.has(db) && instanceSet.has(inst);
  });
  out.database_statistics = (raw.database_statistics || []).filter((r) => {
    const c = r.cdb_cohort || 'UNASSIGNED';
    const db = r.cdb || 'UNKNOWN_DB';
    return cohortSet.has(c) && dbSet.has(db);
  });
  out.top_sql = (raw.top_sql || []).filter((r) => dbSet.has(r.cdb || 'UNKNOWN_DB'));
  out.segment_io = (raw.segment_io || []).filter((r) => dbSet.has(r.cdb || 'UNKNOWN_DB'));
  out.cohort_rollups = (raw.cohort_rollups || []).filter((r) => cohortSet.has(r.Name || ''));
  out.databases = (raw.databases || []).filter((d) => dbSet.has(d.cdb || d.WHICH_DB || 'UNKNOWN_DB'));
  out.properties_database = (raw.properties_database || []).filter((d) => dbSet.has(d.WHICH_DB || d.cdb || 'UNKNOWN_DB'));
  out.plots_html = (raw.plots_html || []).filter((p) => {
    const cohortHint = String(p?.cohort || p?.name || p?.title || p || '').toUpperCase();
    return [...cohortSet].some((c) => cohortHint.includes(String(c).toUpperCase()));
  });
  return out;
}

function buildCompactDbStatsSnapshot(scopedRaw) {
  const keepMetrics = new Set([
    'DB IOPS',
    'DB CPU',
    'DB Memory (MB)',
    'Allocated Storage (GB)',
    'Used Storage (GB)',
    'PGA in use',
    'SGA in use',
  ]);
  const rows = (scopedRaw?.database_statistics || [])
    .filter((r) => keepMetrics.has(String(r.metric || '')))
    .map((r) => ({
      cohort: r.cdb_cohort || 'UNASSIGNED',
      db: r.cdb || 'UNKNOWN_DB',
      metric: String(r.metric || ''),
      p95: Number(Number(r.p95 || 0).toFixed(4)),
      max: Number(Number(r.max || 0).toFixed(4)),
      summary_of: r.summary_of || '',
    }));

  return rows.slice(0, 1200);
}

function buildComprehensiveDbStatsSnapshot(scopedRaw) {
  const keepMetrics = new Set([
    'DB IOPS',
    'DB CPU',
    'DB Memory (MB)',
    'Allocated Storage (GB)',
    'Used Storage (GB)',
    'PGA in use',
    'SGA in use',
    'CPU Utilization (%)',
    'Read IO MBPS',
    'Write IO MBPS',
  ]);
  const rows = (scopedRaw?.database_statistics || [])
    .filter((r) => keepMetrics.has(String(r.metric || '')))
    .map((r) => ({
      cohort: r.cdb_cohort || 'UNASSIGNED',
      db: r.cdb || 'UNKNOWN_DB',
      metric: String(r.metric || ''),
      min: Number(Number(r.min || 0).toFixed(4)),
      p30: Number(Number(r.p30 || 0).toFixed(4)),
      p50: Number(Number(r.p50 || 0).toFixed(4)),
      p70: Number(Number(r.p70 || 0).toFixed(4)),
      p95: Number(Number(r.p95 || 0).toFixed(4)),
      p99: Number(Number(r.p99 || 0).toFixed(4)),
      max: Number(Number(r.max || 0).toFixed(4)),
      summary_of: r.summary_of || '',
    }));
  return rows.slice(0, 3500);
}

function buildDbMetricSummary(cohortSet, dbSet) {
  const keep = new Set([
    'DB CPU',
    'DB IOPS',
    'DB Memory (MB)',
    'Allocated Storage (GB)',
    'Used Storage (GB)',
    'PGA in use',
    'SGA in use',
  ]);
  const byDb = new Map();
  (state.graph?.dbStats || []).forEach((r) => {
    if (!cohortSet.has(r.cohort)) return;
    if (!dbSet.has(r.db)) return;
    if (!keep.has(String(r.metric || ''))) return;
    if (r.summary_of && r.summary_of !== 'cdb') return;
    if (!byDb.has(r.db)) byDb.set(r.db, {});
    byDb.get(r.db)[r.metric] = {
      p95: Number(Number(r.p95 || 0).toFixed(4)),
      p99: Number(Number(r.p99 || 0).toFixed(4)),
      max: Number(Number(r.max || 0).toFixed(4)),
    };
  });
  return byDb;
}

function buildAdvisorTechnicalSummary(rows, cohortSet, dbSet) {
  const dbMetricSummary = buildDbMetricSummary(cohortSet, dbSet);
  const byDbRows = new Map();
  rows.forEach((r) => {
    if (!byDbRows.has(r.db)) byDbRows.set(r.db, []);
    byDbRows.get(r.db).push(r);
  });

  const dbs = [...byDbRows.keys()].sort((a, b) => a.localeCompare(b));
  const db_summary = dbs.map((db) => {
    const rws = byDbRows.get(db) || [];
    const vcpu = rws.reduce((a, x) => a + (Number(x.init_cpu_count || x.logical_cpu_count || 0) || 0), 0);
    const mem = rws.reduce((a, x) => a + (Number(x.mem_gb || 0) || 0), 0);
    const sga = rws.reduce((a, x) => a + (Number(x.sga_size_gb || 0) || 0), 0);
    const pga = rws.reduce((a, x) => a + (Number(x.pga_size_gb || 0) || 0), 0);
    const m = dbMetricSummary.get(db) || {};
    return {
      db,
      db_name: displayDbName(db),
      cohort: rws[0]?.cohort || 'UNASSIGNED',
      version: state.graph?.dbVersionByDb?.[db] || 'Unknown',
      instances: rws.length,
      hosts: [...new Set(rws.map((x) => x.host))].length,
      vcpu_total: Number(vcpu.toFixed(3)),
      memory_gb_total: Number(mem.toFixed(3)),
      sga_gb_total: Number(sga.toFixed(3)),
      pga_gb_total: Number(pga.toFixed(3)),
      db_cpu: m['DB CPU'] || null,
      db_iops: m['DB IOPS'] || null,
      db_memory_mb: m['DB Memory (MB)'] || null,
      allocated_storage_gb: m['Allocated Storage (GB)'] || null,
      used_storage_gb: m['Used Storage (GB)'] || null,
      sga_in_use: m['SGA in use'] || null,
      pga_in_use: m['PGA in use'] || null,
    };
  });

  const top_sql = (state.graph?.topSql || [])
    .filter((r) => dbSet.has(r.db))
    .sort((a, b) => (Number(b.elapsed) || 0) - (Number(a.elapsed) || 0))
    .slice(0, 25)
    .map((r) => ({
      db: r.db,
      db_name: displayDbName(r.db),
      sql_id: r.sql_id,
      elapsed: Number(Number(r.elapsed || 0).toFixed(3)),
      execs: Number(Number(r.execs || 0).toFixed(3)),
      log_reads: Number(Number(r.log_reads || 0).toFixed(3)),
      phy_read_gb: Number(Number(r.phy_read_gb || 0).toFixed(3)),
    }));

  const segment_io = (state.graph?.segmentIo || [])
    .filter((r) => dbSet.has(r.db))
    .sort((a, b) => (Number(b.physical_io_tot) || 0) - (Number(a.physical_io_tot) || 0))
    .slice(0, 25)
    .map((r) => ({
      db: r.db,
      db_name: displayDbName(r.db),
      owner: r.owner,
      object_name: r.object_name,
      object_type: r.object_type,
      physical_io_tot: Number(Number(r.physical_io_tot || 0).toFixed(3)),
    }));

  const timeline = [...cohortSet]
    .sort((a, b) => a.localeCompare(b))
    .map((cohort) => {
      const ts = resolveCpuSeriesForCohort(cohort);
      if (!ts || !Array.isArray(ts.x) || !ts.x.length) {
        return { cohort, points: 0, start: null, end: null, cpu_pct: null };
      }
      const vals = (ts.y || []).map((v) => Number(v)).filter((v) => Number.isFinite(v));
      const st = statsFromValues(vals);
      return {
        cohort,
        points: ts.x.length,
        start: ts.x[0] || null,
        end: ts.x[ts.x.length - 1] || null,
        cpu_pct: st
          ? {
              p50: Number(Number(st.p50 || 0).toFixed(4)),
              p95: Number(Number(st.p95 || 0).toFixed(4)),
              max: Number(Number(st.max || 0).toFixed(4)),
            }
          : null,
      };
    });

  return {
    cohort_timeline: timeline,
    db_summary,
    top_sql,
    segment_io,
  };
}

function buildAdvisorExportPayloadForScope(kind, rows, chunkCohorts, allRows, profile = 'compact') {
  if (!state.raw || !state.graph) return null;
  const sel = state.selection[kind];
  const byCohort = computeByCohortFromRows(rows);
  ensureCohortTargetsInitialized(kind);
  const globalTargets = {
    vcpu_pct: getGlobalTarget('vcpu'),
    memory_pct: getGlobalTarget('memory'),
    iops_pct: getGlobalTarget('iops'),
    storage_pct: getGlobalTarget('storage'),
  };
  const allByCohort = computeByCohortFromRows(allRows);
  const metricWeightTotals = {
    vcpu: allByCohort.reduce((a, c) => a + (Number(c.vcpu) || 0), 0),
    memory: allByCohort.reduce((a, c) => a + (Number(c.mem) || 0), 0),
    iops: allByCohort.reduce((a, c) => a + (Number(c.iops) || 0), 0),
    storage: allByCohort.reduce((a, c) => a + (Number(c.allocated) || 0), 0),
  };
  const cohortSet = new Set(chunkCohorts);
  const dbSet = new Set(rows.map((r) => r.db));
  const instSet = new Set(rows.map((r) => r.instance));
  const scopedRaw = filterRawDataForScope(state.raw, cohortSet, dbSet, instSet);
  const technicalSummary = buildAdvisorTechnicalSummary(rows, cohortSet, dbSet);
  const cohorts = byCohort.map((c) => {
    const target = ensureCohortTargetContainer(c.cohort);
    const wVcpu = Number(c.vcpu) || 0;
    const wMem = Number(c.mem) || 0;
    const wIops = Number(c.iops) || 0;
    const wStorage = Number(c.allocated) || 0;
    return {
      cohort: c.cohort,
      summary: c,
      targets_pct: {
        vcpu_pct: parseTargetPctOrDefault(target.vcpu, globalTargets.vcpu_pct),
        memory_pct: parseTargetPctOrDefault(target.memory, globalTargets.memory_pct),
        iops_pct: parseTargetPctOrDefault(target.iops, globalTargets.iops_pct),
        storage_pct: parseTargetPctOrDefault(target.storage, globalTargets.storage_pct),
      },
      weighted_share_pct: {
        vcpu: metricWeightTotals.vcpu > 0 ? (wVcpu / metricWeightTotals.vcpu) * 100 : 0,
        memory: metricWeightTotals.memory > 0 ? (wMem / metricWeightTotals.memory) * 100 : 0,
        iops: metricWeightTotals.iops > 0 ? (wIops / metricWeightTotals.iops) * 100 : 0,
        storage: metricWeightTotals.storage > 0 ? (wStorage / metricWeightTotals.storage) * 100 : 0,
      },
      stats: {
        cpu_pct: computeScopedCpuStats(rows.filter((r) => r.cohort === c.cohort), [c.cohort]),
        memory_pct: statsFromRowsForMetric(rows.filter((r) => r.cohort === c.cohort), 'memory'),
        iops_pct: computeScopedIopsStats(new Set([c.cohort]), new Set(rows.filter((r) => r.cohort === c.cohort).map((r) => r.db))),
        storage_pct: computeScopedStorageStats(new Set([c.cohort]), new Set(rows.filter((r) => r.cohort === c.cohort).map((r) => r.db))),
      },
    };
  });

  return {
    schema_version: 1,
    exported_at: new Date().toISOString(),
    purpose: 'ChatGPT AWR DBA Advisor context',
    export_profile: profile === 'comprehensive' ? 'comprehensive_v1' : 'compact_v1',
    report_metadata: normalizeReportMeta(state.reportMeta),
    scope: {
      mode: kind,
      selected_cohorts: [...chunkCohorts],
      selected_databases: [...dbSet],
      selected_instances: [...instSet],
      selected_metrics: [...sel.metrics],
      selected_instances_count: rows.length,
    },
    targets: {
      global_pct: globalTargets,
      cohorts,
    },
    summary: {
      global: computeSummaryFromRows(rows),
      by_cohort: byCohort,
      cpu_pct: computeScopedCpuStats(rows, chunkCohorts),
      memory_pct: statsFromRowsForMetric(rows, 'memory'),
      iops_pct: computeScopedIopsStats(cohortSet, dbSet),
      storage_pct: computeScopedStorageStats(cohortSet, dbSet),
    },
    selected_instances: rows.map((r) => ({
      cohort: r.cohort,
      db: r.db,
      db_name: displayDbName(r.db),
      instance: r.instance,
      instance_name: displayInstanceName(r.instance),
      host: r.host,
      vcpu: Number(r.init_cpu_count || r.logical_cpu_count || 0),
      memory_gb: Number(r.mem_gb || 0),
      sga_gb: Number(r.sga_size_gb || 0),
      pga_gb: Number(r.pga_size_gb || 0),
      db_version: state.graph?.dbVersionByDb?.[r.db] || 'Unknown',
      rac: Boolean(r.is_rac || r.rac || r.rac_instance || r.clustered),
      pdb_count: Number(r.pdb_count || 0),
    })),
    compact_data: {
      cohort_rollups: (scopedRaw.cohort_rollups || []).map((r) => ({
        cohort: r.Name || '',
        allocated_storage_gb: Number(Number(r['Allocated Storage (GB)'] || 0).toFixed(3)),
        used_storage_gb: Number(Number(r['Used Storage (GB)'] || 0).toFixed(3)),
        db_iops: Number(Number(r['DB IOPS'] || 0).toFixed(3)),
        db_vcpu: Number(Number(r['DB vCPU'] || 0).toFixed(3)),
      })),
      technical_summary: technicalSummary,
    },
    ...(profile === 'comprehensive'
      ? {
          comprehensive_data: {
            database_statistics: buildComprehensiveDbStatsSnapshot(scopedRaw),
            top_sql: (scopedRaw.top_sql || []).slice(0, 120),
            segment_io: (scopedRaw.segment_io || []).slice(0, 120),
          },
        }
      : {}),
  };
}

function buildAdvisorExportPayloads(kind = 'web', cohortsPerFile = 1, profile = 'compact') {
  if (!state.raw || !state.graph) return [];
  const allRows = selectedInstances(kind);
  const cohorts = [...new Set(allRows.map((r) => r.cohort))].sort((a, b) => a.localeCompare(b));
  const chunks = groupedArray(cohorts, cohortsPerFile);
  return chunks
    .map((chunkCohorts, idx) => {
      const chunkSet = new Set(chunkCohorts);
      const rows = allRows.filter((r) => chunkSet.has(r.cohort));
      const payload = buildAdvisorExportPayloadForScope(kind, rows, chunkCohorts, allRows, profile);
      if (!payload) return null;
      payload.chunk = {
        index: idx + 1,
        total: chunks.length,
        cohorts_per_file: cohortsPerFile,
      };
      return payload;
    })
    .filter(Boolean);
}

function applySavedUiState(uiState = {}, reportMeta = null) {
  state.activeTab = 'web';
  state.linkSelections = true;

  const fallbackWeb = {
    cohorts: new Set(state.graph.cohorts),
    dbs: new Set(state.graph.dbs),
    instances: new Set(state.graph.instances.map((x) => x.instance)),
    metrics: new Set(state.defaultMetricIds),
  };
  const fallbackPpt = {
    cohorts: new Set(fallbackWeb.cohorts),
    dbs: new Set(fallbackWeb.dbs),
    instances: new Set(fallbackWeb.instances),
    metrics: new Set(fallbackWeb.metrics),
  };

  state.selection.web = selectionFromPojo(uiState.selection?.web, fallbackWeb);
  state.selection.ppt = selectionFromPojo(uiState.selection?.ppt, fallbackPpt);

  state.cpuScaleMode = uiState.cpuScaleMode === 'fixed100' ? 'fixed100' : 'dynamic';
  state.vcpuSizingTargetPct = parseTargetPctOrDefault(uiState.vcpuSizingTargetPct, 100);
  state.memorySizingTargetPct = parseTargetPctOrDefault(uiState.memorySizingTargetPct, 100);
  state.iopsSizingTargetPct = parseTargetPctOrDefault(uiState.iopsSizingTargetPct, 100);
  state.storageSizingTargetPct = parseTargetPctOrDefault(uiState.storageSizingTargetPct, 100);
  state.cohortTargets = uiState.cohortTargets && typeof uiState.cohortTargets === 'object' ? uiState.cohortTargets : {};
  state.chartSizes = uiState.chartSizes && typeof uiState.chartSizes === 'object' ? uiState.chartSizes : {};
  state.cardSpans = uiState.cardSpans && typeof uiState.cardSpans === 'object' ? uiState.cardSpans : {};
  state.cardHeights = uiState.cardHeights && typeof uiState.cardHeights === 'object' ? uiState.cardHeights : {};
  state.chartTypes = uiState.chartTypes && typeof uiState.chartTypes === 'object' ? uiState.chartTypes : {};
  state.collapsedSections = uiState.collapsedSections && typeof uiState.collapsedSections === 'object' ? uiState.collapsedSections : {};
  if (Object.prototype.hasOwnProperty.call(uiState || {}, 'sidebarCollapsed')) {
    state.sidebarCollapsed = Boolean(uiState.sidebarCollapsed);
  }
  state.advisorExportProfile = uiState.advisorExportProfile === 'comprehensive' ? 'comprehensive' : 'compact';
  saveSidebarCollapsedPref();
  applySidebarCollapsedUi();
  const advisorProfileSelect = $('advisorExportProfile');
  if (advisorProfileSelect) advisorProfileSelect.value = state.advisorExportProfile;
  state.reportMeta = normalizeReportMeta(reportMeta || uiState.reportMeta || {});

  if (!Number.isFinite(state.vcpuSizingTargetPct)) {
    state.vcpuSizingTargetPct = 100;
  }
  if (!Number.isFinite(state.memorySizingTargetPct)) {
    state.memorySizingTargetPct = 100;
  }
  if (!Number.isFinite(state.iopsSizingTargetPct)) {
    state.iopsSizingTargetPct = 100;
  }
  if (!Number.isFinite(state.storageSizingTargetPct)) {
    state.storageSizingTargetPct = 100;
  }
  applyReportMetaToInputs();
  state.pptSlideSelected = new Set(Array.isArray(uiState.pptSlidesSelected) ? uiState.pptSlidesSelected : []);
  const ai = uiState.aiChat && typeof uiState.aiChat === 'object' ? uiState.aiChat : {};
  state.aiChat = {
    panelOpen: Boolean(ai.panelOpen),
    connected: Boolean(ai.connected),
    connectionSource: String(ai.connectionSource || 'unknown'),
    model: String(ai.model || 'gpt-4.1-mini'),
    messages: Array.isArray(ai.messages)
      ? ai.messages
          .filter((m) => m && typeof m.content === 'string')
          .map((m) => ({ role: m.role === 'assistant' ? 'assistant' : 'user', content: String(m.content || '') }))
      : [],
    docs: Array.isArray(ai.docs)
      ? ai.docs
          .filter((d) => d && typeof d.name === 'string')
          .map((d) => ({ name: d.name, content: String(d.content || '') }))
      : [],
    question: String(ai.question || ''),
    prompt: String(ai.prompt || ''),
    response: String(ai.response || ''),
    appPayload: ai.appPayload && typeof ai.appPayload === 'object' ? ai.appPayload : null,
    payloadErrors: Array.isArray(ai.payloadErrors) ? ai.payloadErrors.map((x) => String(x || '')) : [],
    busy: false,
  };

  applySavedCardSpans();
  applySavedCardHeights();
  applyLayoutOrder(Array.isArray(uiState.layoutOrder) ? uiState.layoutOrder : state.layoutDefaultOrder);
  applyLocalPrefsToStorage();

}

function defaultCardCollapsed(cardId) {
  if (cardId === 'globalSummaryCard') return false;
  return cardId !== 'globalSummaryCard';
}

function setCardCollapsedUi(card, collapsed) {
  if (!card?.id) return;
  const next = Boolean(collapsed);
  state.collapsedSections[card.id] = next;
  card.classList.add('section-collapsible');
  card.classList.toggle('section-collapsed', next);
  const btn = card.querySelector('.section-toggle-btn');
  if (btn) {
    btn.textContent = next ? '>' : '⌄';
    btn.setAttribute('aria-expanded', String(!next));
    btn.setAttribute('aria-label', next ? 'Expand section' : 'Collapse section');
    btn.title = next ? 'Expand section' : 'Collapse section';
  }
}

function applyCollapsedSections() {
  document.querySelectorAll('.content > .card[data-layout-item]').forEach((card) => {
    if (card.id === 'globalSummaryCard') {
      setCardCollapsedUi(card, false);
      return;
    }
    const explicit = Object.prototype.hasOwnProperty.call(state.collapsedSections || {}, card.id);
    const collapsed = explicit ? Boolean(state.collapsedSections[card.id]) : defaultCardCollapsed(card.id);
    setCardCollapsedUi(card, collapsed);
  });
}

function initCollapsibleSections() {
  const cards = [...document.querySelectorAll('.content > .card[data-layout-item]')];
  cards.forEach((card) => {
    if (!card?.id) return;
    if (card.id === 'globalSummaryCard') {
      card.classList.remove('section-collapsible', 'section-collapsed');
      card.querySelector('.section-toggle-btn')?.remove();
      return;
    }
    let head = card.querySelector(':scope > .section-head');
    if (!head) {
      const h2 = card.querySelector(':scope > h2');
      if (h2) {
        head = document.createElement('div');
        head.className = 'section-head';
        card.insertBefore(head, h2);
        head.appendChild(h2);
      }
    }
    if (!head || head.querySelector('.section-toggle-btn')) return;
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'btn tiny secondary section-toggle-btn';
    btn.addEventListener('click', () => {
      const current = card.classList.contains('section-collapsed');
      setCardCollapsedUi(card, !current);
      if (state.page.mode === 'global') {
        renderAll();
      }
    });
    head.appendChild(btn);
  });
  applyCollapsedSections();
}

function findChartTitle(chartBox) {
  const h3 = chartBox.querySelector('h3');
  if (h3?.textContent) return h3.textContent.trim();
  const h2 = chartBox.querySelector('h2');
  if (h2?.textContent) return h2.textContent.trim();
  const section = chartBox.closest('.card');
  const sectionTitle = section?.querySelector('h2');
  return sectionTitle?.textContent?.trim() || 'Chart';
}

function openChartModalFromBox(chartBox) {
  const modal = $('chartModal');
  const body = $('chartModalBody');
  const title = $('chartModalTitle');
  if (!modal || !body || !title) return;
  const svg = chartBox.querySelector('svg');
  if (!svg) return;

  title.textContent = findChartTitle(chartBox);
  const legend = chartBox.querySelector('.line-legend');
  body.innerHTML = svg.outerHTML + (legend ? legend.outerHTML : '');
  modal.classList.add('open');
  modal.setAttribute('aria-hidden', 'false');
}

function closeChartModal() {
  const modal = $('chartModal');
  if (!modal) return;
  modal.classList.remove('open');
  modal.setAttribute('aria-hidden', 'true');
}

function showChartTooltip(text, x, y) {
  const t = $('chartTooltip');
  if (!t) return;
  t.textContent = text;
  t.style.display = 'block';
  const pad = 12;
  const maxX = window.innerWidth - t.offsetWidth - 8;
  const maxY = window.innerHeight - t.offsetHeight - 8;
  t.style.left = `${Math.max(8, Math.min(x + pad, maxX))}px`;
  t.style.top = `${Math.max(8, Math.min(y + pad, maxY))}px`;
}

function hideChartTooltip() {
  const t = $('chartTooltip');
  if (!t) return;
  t.style.display = 'none';
}

function setupChartInteractions() {
  if (document.body.dataset.chartInteractionsInit) return;
  document.body.dataset.chartInteractionsInit = '1';

  document.body.addEventListener('click', (e) => {
    if (state.layoutEditMode) return;
    if (state.isResizingChart) {
      state.isResizingChart = false;
      return;
    }
    if (e.target.closest('.chart-type-controls')) return;
    if (e.target.closest('.chart-resize-handle')) return;
    if (e.target.closest('.global-vcpu-sizer-summary')) return;
    if (e.target.closest('input, textarea, select')) return;
    if (e.target.closest('[data-close-chart-modal="1"]') || e.target.id === 'chartModalClose') {
      closeChartModal();
      return;
    }
    const chartBox = e.target.closest('.chart-box');
    if (!chartBox) return;
    if (!chartBox.querySelector('svg')) return;
    openChartModalFromBox(chartBox);
  });

  document.body.addEventListener('mousemove', (e) => {
    const tipNode = e.target.closest('[data-tip]');
    if (!tipNode) {
      hideChartTooltip();
      return;
    }
    const txt = tipNode.getAttribute('data-tip');
    if (!txt) {
      hideChartTooltip();
      return;
    }
    showChartTooltip(txt, e.clientX, e.clientY);
  });

  document.body.addEventListener('mouseleave', hideChartTooltip);

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') closeChartModal();
  });
}

function chartBoxKey(chartBox) {
  const nodeWithId = chartBox.querySelector('[id]');
  if (nodeWithId?.id) return `id:${nodeWithId.id}`;
  const title = findChartTitle(chartBox);
  return `title:${title}`;
}

function loadChartSizes() {
  try {
    const raw = localStorage.getItem(CHART_SIZE_STORAGE_KEY);
    const parsed = raw ? JSON.parse(raw) : {};
    state.chartSizes = parsed && typeof parsed === 'object' ? parsed : {};
  } catch {
    state.chartSizes = {};
  }
}

function saveChartSizes() {
  try {
    localStorage.setItem(CHART_SIZE_STORAGE_KEY, JSON.stringify(state.chartSizes));
  } catch {
    // Ignore storage write failures.
  }
}

function normalizeChartType(value) {
  const v = String(value || '').toLowerCase();
  if (v === 'text' || v === 'column' || v === 'bar' || v === 'pie' || v === 'line' || v === 'box') return v;
  return 'bar';
}

function loadChartTypes() {
  try {
    const raw = localStorage.getItem(CHART_TYPE_STORAGE_KEY);
    const parsed = raw ? JSON.parse(raw) : {};
    state.chartTypes = parsed && typeof parsed === 'object' ? parsed : {};
  } catch {
    state.chartTypes = {};
  }
}

function saveChartTypes() {
  try {
    localStorage.setItem(CHART_TYPE_STORAGE_KEY, JSON.stringify(state.chartTypes || {}));
  } catch {
    // Ignore storage write failures.
  }
}

function getChartType(containerId, defaultType = 'bar') {
  const t = state.chartTypes?.[containerId];
  return t ? normalizeChartType(t) : normalizeChartType(defaultType);
}

function setChartType(containerId, type) {
  if (!containerId) return;
  state.chartTypes[containerId] = normalizeChartType(type);
  saveChartTypes();
}

function applyChartSizeFromState(chartBox) {
  const key = chartBoxKey(chartBox);
  const h = Number(state.chartSizes[key]);
  if (!h || Number.isNaN(h)) return;
  chartBox.style.setProperty('--chart-height', `${h}px`);
  chartBox.dataset.chartHeight = String(h);
}

function syncResizableCharts() {
  document.querySelectorAll('.chart-box').forEach((box) => {
    applyChartSizeFromState(box);
    if (box.querySelector('.chart-resize-handle')) return;
    const handle = document.createElement('button');
    handle.type = 'button';
    handle.className = 'chart-resize-handle';
    handle.setAttribute('aria-label', 'Resize chart');
    handle.title = 'Resize chart';
    box.appendChild(handle);
  });
}

function chartContainerIdFromBox(box) {
  const preferred = box.querySelector('div[id]');
  return preferred?.id || '';
}

function isFixedBoxChartId(id) {
  return (
    id === 'cpuBoxChart' ||
    id === 'memBoxChart' ||
    id === 'cohortCpuBoxChart' ||
    id === 'cohortMemBoxChart' ||
    id === 'instanceCpuBoxChart' ||
    id === 'instanceMemBoxChart'
  );
}

function supportsLineType(id) {
  return !isFixedBoxChartId(id) && /cpu/i.test(String(id || ''));
}

function defaultChartTypeForId(id) {
  return /^cpuLine/i.test(String(id || '')) ? 'line' : 'bar';
}

function updateChartTypeButtons(box) {
  const id = chartContainerIdFromBox(box);
  if (!id) return;
  const current = getChartType(id, defaultChartTypeForId(id));
  box.querySelectorAll('.chart-type-btn').forEach((btn) => {
    btn.classList.toggle('active', btn.dataset.chartType === current);
  });
}

function syncChartTypeControls() {
  document.querySelectorAll('.chart-box').forEach((box) => {
    const id = chartContainerIdFromBox(box);
    if (!id) return;
    if (isFixedBoxChartId(id)) {
      const existing = box.querySelector('.chart-type-controls');
      if (existing) existing.remove();
      return;
    }
    if (!box.querySelector('.chart-type-controls')) {
      const lineBtn = supportsLineType(id)
        ? `<button type="button" class="btn tiny secondary chart-type-btn" data-chart-id="${esc(id)}" data-chart-type="line">Line</button>`
        : '';
      const controls = document.createElement('div');
      controls.className = 'chart-type-controls';
      controls.innerHTML = `
        <span class="chart-type-label">Chart Type:</span>
        <button type="button" class="btn tiny secondary chart-type-btn" data-chart-id="${esc(id)}" data-chart-type="text">Text</button>
        <button type="button" class="btn tiny secondary chart-type-btn" data-chart-id="${esc(id)}" data-chart-type="column">Column</button>
        <button type="button" class="btn tiny secondary chart-type-btn" data-chart-id="${esc(id)}" data-chart-type="bar">Bar</button>
        <button type="button" class="btn tiny secondary chart-type-btn" data-chart-id="${esc(id)}" data-chart-type="pie">Pie</button>
        ${lineBtn}
      `;
      box.insertBefore(controls, box.firstChild);
    }
    updateChartTypeButtons(box);
  });
}

function setupChartResize() {
  if (document.body.dataset.chartResizeInit) return;
  document.body.dataset.chartResizeInit = '1';

  document.body.addEventListener('mousedown', (e) => {
    const handle = e.target.closest('.chart-resize-handle');
    if (!handle) return;
    const box = handle.closest('.chart-box');
    if (!box) return;
    e.preventDefault();
    const startHeight = Number(box.dataset.chartHeight) || box.getBoundingClientRect().height || 280;
    state.chartResizeSession = {
      box,
      startY: e.clientY,
      startHeight,
    };
    document.body.classList.add('resizing-charts');
  });

  document.addEventListener('mousemove', (e) => {
    if (!state.chartResizeSession) return;
    const { box, startY, startHeight } = state.chartResizeSession;
    const delta = e.clientY - startY;
    const nextHeight = Math.max(160, Math.min(760, Math.round(startHeight + delta)));
    box.style.setProperty('--chart-height', `${nextHeight}px`);
    box.dataset.chartHeight = String(nextHeight);
    state.isResizingChart = true;
  });

  document.addEventListener('mouseup', () => {
    if (!state.chartResizeSession) return;
    const box = state.chartResizeSession.box;
    const key = chartBoxKey(box);
    const h = Number(box.dataset.chartHeight) || Math.round(box.getBoundingClientRect().height);
    state.chartSizes[key] = h;
    saveChartSizes();
    state.chartResizeSession = null;
    document.body.classList.remove('resizing-charts');
  });
}

function normalizeCardSpan(value) {
  const n = Number(value);
  if (n === 4 || n === 8 || n === 12) return n;
  return 12;
}

function normalizeCardHeightSize(value) {
  const v = String(value || '').toLowerCase();
  if (v === 's' || v === 'm' || v === 'l') return v;
  return 'm';
}

function loadCardSpans() {
  try {
    const raw = localStorage.getItem(CARD_SPAN_STORAGE_KEY);
    const parsed = raw ? JSON.parse(raw) : {};
    state.cardSpans = parsed && typeof parsed === 'object' ? parsed : {};
  } catch {
    state.cardSpans = {};
  }
}

function saveCardSpans() {
  try {
    localStorage.setItem(CARD_SPAN_STORAGE_KEY, JSON.stringify(state.cardSpans || {}));
  } catch {
    // Ignore storage write failures.
  }
}

function setCardSpan(card, span) {
  if (!card?.id) return;
  const normalized = normalizeCardSpan(span);
  card.dataset.cardSpan = String(normalized);
  state.cardSpans[card.id] = normalized;
  saveCardSpans();
}

function updateCardSpanButtons(card) {
  if (!card) return;
  const current = normalizeCardSpan(card.dataset.cardSpan || 12);
  card.querySelectorAll('.card-size-btn').forEach((btn) => {
    btn.classList.toggle('active', Number(btn.dataset.span) === current);
  });
}

function updateCardHeightButtons(card) {
  if (!card) return;
  const current = normalizeCardHeightSize(card.dataset.cardHeight || 'm');
  card.querySelectorAll('.card-height-btn').forEach((btn) => {
    btn.classList.toggle('active', String(btn.dataset.height) === current);
  });
}

function setCardHeightSize(card, size) {
  if (!card?.id) return;
  const normalized = normalizeCardHeightSize(size);
  card.dataset.cardHeight = normalized;
  state.cardHeights[card.id] = normalized;
  saveCardHeights();
}

function moveCardByOffset(card, offset) {
  if (!card || !offset) return;
  const cards = movableCards();
  const idx = cards.findIndex((c) => c === card);
  if (idx < 0) return;
  const nextIdx = Math.max(0, Math.min(cards.length - 1, idx + offset));
  if (nextIdx === idx) return;
  const target = cards[nextIdx];
  if (!target || !target.parentElement) return;
  const parent = target.parentElement;
  if (offset > 0) parent.insertBefore(target, card);
  else parent.insertBefore(card, target);
}

function applySavedCardSpans() {
  movableCards().forEach((card) => {
    const saved = state.cardSpans?.[card.id];
    card.dataset.cardSpan = String(normalizeCardSpan(saved || 12));
    updateCardSpanButtons(card);
  });
}

function syncCardSizeControls() {
  movableCards().forEach((card) => {
    if (card.querySelector('.card-size-controls')) {
      updateCardSpanButtons(card);
      updateCardHeightButtons(card);
      return;
    }

    const controls = document.createElement('div');
    controls.className = 'card-size-controls';
    controls.innerHTML = `
      <span class="card-size-label">Move:</span>
      <button type="button" class="btn tiny secondary card-move-btn" data-card-id="${esc(card.id)}" data-move="-1">Up</button>
      <button type="button" class="btn tiny secondary card-move-btn" data-card-id="${esc(card.id)}" data-move="1">Down</button>
      <span class="card-size-sep" aria-hidden="true"></span>
      <span class="card-size-label">Width:</span>
      <button type="button" class="btn tiny secondary card-size-btn" data-card-id="${esc(card.id)}" data-span="4">1/3</button>
      <button type="button" class="btn tiny secondary card-size-btn" data-card-id="${esc(card.id)}" data-span="8">2/3</button>
      <button type="button" class="btn tiny secondary card-size-btn" data-card-id="${esc(card.id)}" data-span="12">3/3</button>
      <span class="card-size-sep" aria-hidden="true"></span>
      <span class="card-size-label">Height:</span>
      <button type="button" class="btn tiny secondary card-height-btn" data-card-id="${esc(card.id)}" data-height="s">S</button>
      <button type="button" class="btn tiny secondary card-height-btn" data-card-id="${esc(card.id)}" data-height="m">M</button>
      <button type="button" class="btn tiny secondary card-height-btn" data-card-id="${esc(card.id)}" data-height="l">L</button>
    `;
    card.insertBefore(controls, card.firstChild);
    updateCardSpanButtons(card);
    updateCardHeightButtons(card);
  });
}

function loadCardHeights() {
  try {
    const raw = localStorage.getItem(CARD_HEIGHT_STORAGE_KEY);
    const parsed = raw ? JSON.parse(raw) : {};
    state.cardHeights = parsed && typeof parsed === 'object' ? parsed : {};
  } catch {
    state.cardHeights = {};
  }
}

function saveCardHeights() {
  try {
    localStorage.setItem(CARD_HEIGHT_STORAGE_KEY, JSON.stringify(state.cardHeights || {}));
  } catch {
    // Ignore storage write failures.
  }
}

function applySavedCardHeights() {
  movableCards().forEach((card) => {
    const savedRaw = state.cardHeights?.[card.id];
    let size = 'm';
    if (typeof savedRaw === 'number') {
      if (savedRaw <= 380) size = 's';
      else if (savedRaw >= 640) size = 'l';
      else size = 'm';
    } else {
      size = normalizeCardHeightSize(savedRaw);
    }
    card.style.height = '';
    card.classList.remove('card-height-custom');
    card.dataset.cardHeight = size;
    updateCardHeightButtons(card);
  });
}

function syncCardResizeHandles() {
  // Simplified UX: section height is controlled by S/M/L buttons.
}

function movableCards() {
  return [...document.querySelectorAll('.content > .card[data-layout-item]')];
}

function getCurrentLayoutOrder() {
  return movableCards().map((card) => card.id).filter(Boolean);
}

function setLayoutStatus(text) {
  const node = $('layoutStatus');
  if (!node) return;
  node.textContent = text || '';
}

function applyLayoutOrder(order) {
  const content = document.querySelector('.content');
  if (!content || !Array.isArray(order) || !order.length) return;
  const fixedAnchor = $('cohortPage') || $('instancePage');
  const cards = movableCards();
  const byId = new Map(cards.map((card) => [card.id, card]));
  const placed = new Set();

  order.forEach((id) => {
    const card = byId.get(id);
    if (!card) return;
    if (fixedAnchor) content.insertBefore(card, fixedAnchor);
    else content.appendChild(card);
    placed.add(id);
  });

  cards.forEach((card) => {
    if (placed.has(card.id)) return;
    if (fixedAnchor) content.insertBefore(card, fixedAnchor);
    else content.appendChild(card);
  });
}

function loadSavedLayout() {
  try {
    const raw = localStorage.getItem(LAYOUT_STORAGE_KEY);
    if (!raw) return false;
    const saved = JSON.parse(raw);
    if (!Array.isArray(saved) || !saved.length) return false;
    applyLayoutOrder(saved);
    setLayoutStatus('Saved view loaded.');
    return true;
  } catch {
    return false;
  }
}

function saveCurrentLayout() {
  try {
    const order = getCurrentLayoutOrder();
    localStorage.setItem(LAYOUT_STORAGE_KEY, JSON.stringify(order));
    setLayoutStatus('View saved on this browser.');
  } catch {
    setLayoutStatus('Unable to save view in this browser.');
  }
}

function resetLayout() {
  applyLayoutOrder(state.layoutDefaultOrder);
  try {
    localStorage.removeItem(LAYOUT_STORAGE_KEY);
  } catch {
    // Ignore storage failures and still keep in-memory reset behavior.
  }
  setLayoutStatus('View reset to default.');
}

function clearDropTargets() {
  movableCards().forEach((card) => card.classList.remove('layout-drop-target'));
}

function setLayoutEditMode(enabled) {
  state.layoutEditMode = Boolean(enabled);
  const editBtn = $('layoutEditBtn');
  const saveBtn = $('layoutSaveBtn');
  if (editBtn) editBtn.textContent = state.layoutEditMode ? 'Customize: On' : 'Customize: Off';
  if (saveBtn) saveBtn.disabled = !state.layoutEditMode;
  document.body.classList.toggle('layout-editing', state.layoutEditMode);

  movableCards().forEach((card) => {
    card.draggable = false;
    card.classList.remove('layout-draggable');
  });

  if (!state.layoutEditMode) {
    state.layoutDraggingId = null;
    clearDropTargets();
    movableCards().forEach((card) => card.classList.remove('layout-dragging'));
  }
}

function setupLayoutDnD() {
  // Simplified UX: order is controlled by Up/Down buttons.
}

function setupCardResize() {
  // Simplified UX: section height is controlled by S/M/L buttons.
}

function computeSummaryFromRows(rows) {
  const dbs = new Set(rows.map((r) => r.db));
  const hosts = new Set(rows.map((r) => r.host));
  const cohorts = new Set(rows.map((r) => r.cohort));
  let allocated = 0;
  let used = 0;
  let iops = 0;
  let logons = 0;
  cohorts.forEach((c) => {
    const item = state.graph.cohortRollups[c];
    allocated += item?.allocated || 0;
    used += item?.used || 0;
    iops += item?.db_iops || 0;
    logons += item?.db_logons || 0;
  });
  return {
    db_count: dbs.size,
    host_count: hosts.size,
    instance_count: rows.length,
    vcpu_total: rows.reduce((a, r) => a + (r.init_cpu_count || r.logical_cpu_count || 0), 0),
    memory_gb_total: rows.reduce((a, r) => a + r.mem_gb, 0),
    allocated_storage_gb: allocated,
    used_storage_gb: used,
    db_iops_total: iops,
    db_logons_total: logons,
  };
}

function compareValues(a, b) {
  const aNum = Number(a);
  const bNum = Number(b);
  const aIsNum = Number.isFinite(aNum);
  const bIsNum = Number.isFinite(bNum);
  if (aIsNum && bIsNum) return aNum - bNum;
  return String(a ?? '').localeCompare(String(b ?? ''));
}

function sortRows(tableId, rows) {
  const cfg = state.sort[tableId];
  if (!cfg) return rows;
  const sorted = [...rows].sort((x, y) => compareValues(x[cfg.key], y[cfg.key]));
  return cfg.dir === 'asc' ? sorted : sorted.reverse();
}

function updateSortHeaderIndicators(tableId) {
  const table = $(tableId);
  if (!table) return;
  const cfg = state.sort[tableId];
  table.querySelectorAll('th[data-sort-key]').forEach((th) => {
    const key = th.dataset.sortKey;
    const base = th.dataset.baseLabel || th.textContent.replace(/\s+\[(asc|desc)\]$/, '');
    th.dataset.baseLabel = base;
    th.classList.toggle('sort-active', key === cfg.key);
    th.textContent = key === cfg.key ? `${base} [${cfg.dir}]` : base;
  });
}

function setupSortableHeaders() {
  Object.keys(state.sort).forEach((tableId) => {
    const table = $(tableId);
    if (!table || table.dataset.sortInit) return;
    table.dataset.sortInit = '1';
    table.querySelectorAll('th[data-sort-key]').forEach((th) => {
      th.addEventListener('click', () => {
        const key = th.dataset.sortKey;
        const cfg = state.sort[tableId];
        if (cfg.key === key) cfg.dir = cfg.dir === 'asc' ? 'desc' : 'asc';
        else {
          cfg.key = key;
          cfg.dir = 'asc';
        }
        renderAll();
      });
    });
    updateSortHeaderIndicators(tableId);
  });
}

function normalize(raw) {
  const dbDisplayByDb = {};
  (raw.properties_database || []).forEach((r) => {
    const key = r.WHICH_DB || r.cdb;
    const nm = r.DB_NAME || r.db_name || r.dbname;
    if (key && nm) dbDisplayByDb[key] = String(nm);
  });

  const inferDbDisplay = (dbKey) => {
    const key = String(dbKey || 'UNKNOWN_DB');
    if (dbDisplayByDb[key]) return dbDisplayByDb[key];
    let t = key;
    if (t.startsWith('db_')) t = t.slice(3);
    t = t.split('_')[0] || t;
    return t || 'UNKNOWN_DB';
  };
  const inferInstanceDisplay = (instKey) => {
    const rawInst = String(instKey || 'UNKNOWN_INSTANCE');
    const m = rawInst.match(/^db_([^_$]+)/i);
    if (m?.[1]) return m[1];
    if (rawInst.includes('$')) {
      const parts = rawInst.split('$');
      if (parts.length >= 3) return `${parts[parts.length - 2]}$${parts[parts.length - 1]}`;
      return parts.slice(1).join('$') || rawInst;
    }
    return rawInst;
  };
  const instanceDisplayByInstance = {};

  const instances = (raw.instances || []).map((r) => {
    const cohort = r.cdb_cohort || r.db_cohort || 'UNASSIGNED';
    const db = r.cdb || r.WHICH_DB || 'UNKNOWN_DB';
    const instance = r.db || `${db}$${r.host || 'UNKNOWN_HOST'}`;
    const db_display = inferDbDisplay(db);
    const instance_display = inferInstanceDisplay(instance);
    instanceDisplayByInstance[instance] = instance_display;
    return {
      cohort,
      db,
      db_display,
      instance,
      instance_display,
      host: r.host || 'UNKNOWN_HOST',
      init_cpu_count: Number(r.init_cpu_count) || 0,
      logical_cpu_count: Number(r.logical_cpu_count) || 0,
      mem_gb: Number(r.mem_gb) || 0,
      sga_size_gb: Number(r.sga_size_gb) || 0,
      pga_size_gb: Number(r.pga_size_gb) || 0,
    };
  });

  const cohorts = [...new Set(instances.map((x) => x.cohort))].sort();
  const dbs = [...new Set(instances.map((x) => x.db))].sort();

  const cohortRollups = {};
  for (const r of raw.cohort_rollups || []) {
    const c = r.Name;
    if (!c) continue;
    cohortRollups[c] = {
      allocated: Number(r['Allocated Storage (GB)']) || 0,
      used: Number(r['Used Storage (GB)']) || 0,
      db_vcpu: Number(r['DB vCPU']) || 0,
      db_iops: Number(r['DB IOPS']) || 0,
      db_logons: Number(r['DB Logons']) || 0,
    };
  }

  const dbStats = (raw.database_statistics || []).map((r) => ({
    cohort: r.cdb_cohort || 'UNASSIGNED',
    db: r.cdb || 'UNKNOWN_DB',
    metric: r.metric || 'UNKNOWN_METRIC',
    min: Number(r.min) || 0,
    p30: Number(r.p30) || 0,
    p50: Number(r.p50) || 0,
    p70: Number(r.p70) || 0,
    p95: Number(r.p95) || 0,
    p99: Number(r.p99) || 0,
    max: Number(r.max) || 0,
    summary_of: r.summary_of || '',
  }));

  const topSql = (raw.top_sql || []).map((r) => ({
    db: r.cdb || 'UNKNOWN_DB',
    sql_id: r.SQL_ID || '',
    elapsed: Number(r.ELAP) || 0,
    execs: Number(r.EXECS) || 0,
    log_reads: Number(r.LOG_READS) || 0,
    phy_read_gb: Number(r.PHY_READ_GB) || 0,
  }));

  const segmentIo = (raw.segment_io || []).map((r) => ({
    db: r.cdb || 'UNKNOWN_DB',
    owner: r.OWNER || '',
    object_name: r.OBJECT_NAME || '',
    object_type: r.OBJECT_TYPE || '',
    physical_io_tot: Number(r.PHYSICAL_IO_TOT || r.SEG_PHYSICAL_IO_TOT || 0),
  }));

  const cpuSeriesByCohort = parseCpuTimeSeries(raw);

  const dbVersionByDb = {};
  (raw.databases || []).forEach((d) => {
    const db = d.cdb || d.WHICH_DB;
    if (!db) return;
    dbVersionByDb[db] = d.cdb_version || d.db_version || 'Unknown';
  });

  return { instances, cohorts, dbs, cohortRollups, dbStats, topSql, segmentIo, cpuSeriesByCohort, dbVersionByDb, dbDisplayByDb, instanceDisplayByInstance };
}

function displayDbName(db) {
  const key = String(db || 'UNKNOWN_DB');
  const mapped = state.graph?.dbDisplayByDb?.[key];
  if (mapped) return mapped;
  let t = key;
  if (t.startsWith('db_')) t = t.slice(3);
  t = t.split('_')[0] || t;
  return t || key;
}

function displayInstanceName(instance) {
  const key = String(instance || 'UNKNOWN_INSTANCE');
  const mapped = state.graph?.instanceDisplayByInstance?.[key];
  if (mapped) return mapped;
  const m = key.match(/^db_([^_$]+)/i);
  if (m?.[1]) return m[1];
  if (key.includes('$')) {
    const parts = key.split('$');
    if (parts.length >= 3) return `${parts[parts.length - 2]}$${parts[parts.length - 1]}`;
    return parts.slice(1).join('$') || key;
  }
  return key;
}

function buildMetricCatalog(raw) {
  const list = [...BASE_METRICS];
  const names = [...new Set((raw.database_statistics || []).map((r) => r.metric).filter(Boolean))].sort();
  names.forEach((name) => {
    list.push({
      id: `dbstat:${name}`,
      label: `DB Stat: ${name}`,
      category: 'Database Statistics',
      fmt: 'dec1',
      summary: false,
      dbStatName: name,
    });
  });
  list.push(...SPECIAL_METRICS);
  return list;
}

function currentMetricCatalogById() {
  return new Map(state.metricCatalog.map((m) => [m.id, m]));
}

function escAttr(s) {
  return esc(s).replaceAll('"', '&quot;');
}

function ensureValidSelection(kind) {
  const sel = state.selection[kind];
  const g = state.graph;
  if (!g) return;

  if (sel.cohorts.size === 0) g.cohorts.forEach((c) => sel.cohorts.add(c));

  const allowedDbs = new Set(g.instances.filter((x) => sel.cohorts.has(x.cohort)).map((x) => x.db));
  sel.dbs = new Set([...sel.dbs].filter((d) => allowedDbs.has(d)));
  if (sel.dbs.size === 0) allowedDbs.forEach((d) => sel.dbs.add(d));

  const allowedInstances = new Set(
    g.instances.filter((x) => sel.cohorts.has(x.cohort) && sel.dbs.has(x.db)).map((x) => x.instance),
  );
  sel.instances = new Set([...sel.instances].filter((i) => allowedInstances.has(i)));
  if (sel.instances.size === 0) allowedInstances.forEach((i) => sel.instances.add(i));

  const allowedMetrics = new Set(state.metricCatalog.map((m) => m.id));
  sel.metrics = new Set([...sel.metrics].filter((id) => allowedMetrics.has(id)));
  if (sel.metrics.size === 0) {
    state.defaultMetricIds.forEach((id) => sel.metrics.add(id));
  }
}

function activeSelection() {
  return state.selection.web;
}

function selectedInstances(kind) {
  const sel = state.selection[kind];
  return state.graph.instances.filter(
    (x) => sel.cohorts.has(x.cohort) && sel.dbs.has(x.db) && sel.instances.has(x.instance),
  );
}

function slideBadge(slide) {
  if (slide.type === 'executive-summary') return 'EXEC';
  if (slide.type?.startsWith('deployment-')) return 'DEPLOY';
  if (slide.type === 'title') return 'TITLE';
  if (slide.type === 'summary') return 'SUMMARY';
  if (slide.type === 'separator') return 'SEPARATOR';
  if (slide.type === 'cohort') return 'COHORT';
  return 'INSTANCE';
}

function buildPptSlidePlan() {
  const rows = selectedInstances('ppt');
  const byCohort = new Map();
  rows.forEach((r) => {
    if (!byCohort.has(r.cohort)) byCohort.set(r.cohort, []);
    byCohort.get(r.cohort).push(r);
  });

  const cohorts = [...byCohort.keys()].sort((a, b) => a.localeCompare(b));
  const plan = [];
  plan.push({
    id: 'executive-summary',
    type: 'executive-summary',
    title: 'Executive Summary',
    subtitle: 'DBA analytical overview and deployment recommendation',
  });
  plan.push({
    id: 'deployment-base',
    type: 'deployment-base',
    title: 'Service Fit Metrics - Base Database',
    subtitle: 'Informative slide: fit parameters, formulas, thresholds, and references',
  });
  plan.push({
    id: 'deployment-exascale',
    type: 'deployment-exascale',
    title: 'Service Fit Metrics - Exascale',
    subtitle: 'Informative slide: fit parameters, formulas, thresholds, and references',
  });
  plan.push({
    id: 'deployment-exadata',
    type: 'deployment-exadata',
    title: 'Service Fit Metrics - Exadata Dedicated',
    subtitle: 'Informative slide: fit parameters, formulas, thresholds, and references',
  });
  plan.push({ id: 'title', type: 'title', title: 'AWR Analysis', subtitle: 'Generated report overview' });
  plan.push({ id: 'summary-global', type: 'summary', title: 'Global Summary', subtitle: 'All selected databases and cohorts' });
  plan.push({ id: 'sep-cohorts', type: 'separator', title: 'Cohort Analysis', subtitle: 'Cohort-level drill down' });
  cohorts.forEach((cohort) => {
    plan.push({
      id: `cohort:${cohort}`,
      type: 'cohort',
      cohort,
      title: `Cohort: ${cohort}`,
      subtitle: `Summary and metrics for cohort ${cohort}`,
    });
  });
  plan.push({ id: 'sep-instances', type: 'separator', title: 'Instance Analysis', subtitle: 'Instance-level drill down' });

  cohorts.forEach((cohort) => {
    plan.push({
      id: `sep-instance-cohort:${cohort}`,
      type: 'separator',
      cohort,
      title: `Instances - ${cohort}`,
      subtitle: `Instance details for cohort ${cohort}`,
    });
    const instances = [...byCohort.get(cohort)].sort((a, b) => a.instance.localeCompare(b.instance));
    instances.forEach((r) => {
      plan.push({
        id: `instance:${cohort}:${r.instance}`,
        type: 'instance',
        cohort,
        instance: r.instance,
        db: r.db,
        title: `Instance: ${displayInstanceName(r.instance)}`,
        subtitle: `${displayDbName(r.db)} on ${r.host}`,
      });
    });
  });
  return plan;
}

function syncPptSlideSelection(plan) {
  const next = new Set();
  const existing = state.pptSlideSelected || new Set();
  plan.forEach((s) => {
    if (existing.has(s.id) || existing.size === 0) next.add(s.id);
  });
  state.pptSlideSelected = next;
}

function renderPptStoryboard() {
  const list = $('pptSlidesList');
  const status = $('pptSlidesStatus');
  if (!list || !status) return;
  if (!state.graph) {
    list.innerHTML = '';
    status.textContent = 'Load JSON to generate slides.';
    return;
  }

  const plan = buildPptSlidePlan();
  state.pptSlidePlan = plan;
  syncPptSlideSelection(plan);
  const selectedCount = plan.filter((s) => state.pptSlideSelected.has(s.id)).length;
  status.textContent = `${selectedCount}/${plan.length} slides selected for export.`;
  list.innerHTML = plan
    .map(
      (s, idx) =>
        `<div class="slide-row">
          <input class="slide-check" aria-label="Include slide ${idx + 1}" type="checkbox" data-ppt-slide-id="${escAttr(s.id)}" ${state.pptSlideSelected.has(s.id) ? 'checked' : ''} />
          <span class="slide-badge">${slideBadge(s)}</span>
          <span class="slide-title">${idx + 1}. ${esc(s.title)}</span>
        </div>`,
    )
    .join('');
  enableReportActions();
}

function findSlideById(id) {
  return (state.pptSlidePlan || []).find((s) => s.id === id);
}

const PPT_STYLE = {
  // Oracle FY26 template-aligned (Light theme / white slides)
  bg: 'FFFFFF',
  text: '2A2F2F',
  muted: '6B747A',
  primary: 'C74634',
  primaryDark: '2A2F2F',
  border: 'D9DEDE',
  panel: 'F8FAFC',
  accentSoft: 'FBF2F0',
  font: 'Oracle Sans Tab',
  titleFont: 'Oracle Sans Tab',
  accent1: '04536F',
  accent2: '6C3F49',
  accent3: 'C74634',
  accent4: 'F0CC71',
  accent5: '89B2B0',
  accent6: '86B596',
};

const PPT_CANVAS = { w: 13.333, h: 7.5 };
const ORACLE_TEMPLATE_GRID = {
  title: { x: 0.8409, y: 0.1995, w: 11.67, h: 0.9 },
  subtitle: { x: 0.8385, y: 1.1044, w: 11.67, h: 0.3615 },
  slideNo: { x: 0.8333, y: 7.025, w: 0.4, h: 0.4 },
  footer: { x: 1.2333, y: 7.0253, w: 6.2832, h: 0.3993 },
  date: { x: 7.5166, y: 7.0271, w: 3.0, h: 0.3975 },
  logo: { x: 12.4733, y: 7.07, w: 0.43, h: 0.43 },
};

const PPT_LAYOUT_MAP = {
  headerLineY: 1.52,
  kpiStartY: 1.62,
  kpiCardH: 0.74,
  kpiGapY: 0.1,
  tableY: 3.36,
  tableH: 0.88,
  chartPrimary: { x: 0.5, y: 4.28, w: 6.05, h: 2.62 },
  chartSecondary: { x: 6.78, y: 4.28, w: 6.05, h: 2.62 },
  coverBox: { x: 0.7, y: 1.75, w: 12.0, h: 2.95 },
};

function geomToPct(g) {
  const p = (n, total) => `${((n / total) * 100).toFixed(2)}%`;
  return {
    left: p(g.x, PPT_CANVAS.w),
    top: p(g.y, PPT_CANVAS.h),
    width: p(g.w, PPT_CANVAS.w),
    height: p(g.h, PPT_CANVAS.h),
  };
}

function slideDataForSelection() {
  const rows = selectedInstances('ppt');
  const byCohort = new Map();
  rows.forEach((r) => {
    if (!byCohort.has(r.cohort)) byCohort.set(r.cohort, []);
    byCohort.get(r.cohort).push(r);
  });
  return { rows, byCohort };
}

function addPptFrame(slide, pageNo, totalPages) {
  slide.background = { color: PPT_STYLE.bg };
  // Keep a subtle top guide line while preserving white template look.
  slide.addShape('line', { x: 0.84, y: 0.19, w: 11.67, h: 0, line: { color: PPT_STYLE.border, pt: 1 } });
  slide.addShape('line', {
    x: 0.5,
    y: 7.13,
    w: 12.3,
    h: 0,
    line: { color: PPT_STYLE.border, pt: 1 },
  });
  slide.addText(`${pageNo}`, {
    ...ORACLE_TEMPLATE_GRID.slideNo,
    fontSize: 8,
    color: PPT_STYLE.muted,
    align: 'left',
    fontFace: PPT_STYLE.font,
  });
  slide.addText('Copyright © 2026, Oracle and/or its affiliates |  Confidential: Internal/Restricted/Highly Restricted', {
    ...ORACLE_TEMPLATE_GRID.footer,
    fontSize: 8,
    color: PPT_STYLE.muted,
    align: 'left',
    fontFace: PPT_STYLE.font,
  });
  const d = new Date();
  const dateTxt = d.toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: '2-digit' });
  slide.addText(dateTxt, {
    ...ORACLE_TEMPLATE_GRID.date,
    fontSize: 8,
    color: PPT_STYLE.muted,
    align: 'left',
    fontFace: PPT_STYLE.font,
  });
  slide.addText('Oracle', {
    ...ORACLE_TEMPLATE_GRID.logo,
    fontSize: 8,
    bold: true,
    color: PPT_STYLE.primary,
    align: 'right',
    fontFace: PPT_STYLE.font,
  });
}

function addPptHeader(slide, title, subtitle = '', pageNo = 1, totalPages = 1) {
  addPptFrame(slide, pageNo, totalPages);
  slide.addText(title, {
    ...ORACLE_TEMPLATE_GRID.title,
    bold: true,
    fontSize: 24,
    color: PPT_STYLE.primaryDark,
    fontFace: PPT_STYLE.titleFont,
  });
  if (subtitle) {
    slide.addText(subtitle, {
      ...ORACLE_TEMPLATE_GRID.subtitle,
      fontSize: 12,
      color: PPT_STYLE.muted,
      fontFace: PPT_STYLE.font,
    });
  }
  slide.addShape('line', { x: 0.84, y: PPT_LAYOUT_MAP.headerLineY, w: 11.67, h: 0, line: { color: PPT_STYLE.border, pt: 1 } });
}

function addPptKpiGrid(slide, metrics, startY = PPT_LAYOUT_MAP.kpiStartY) {
  const cols = 4;
  const cardW = 2.95;
  const cardH = PPT_LAYOUT_MAP.kpiCardH;
  const gapX = 0.15;
  const gapY = PPT_LAYOUT_MAP.kpiGapY;
  metrics.forEach((m, idx) => {
    const col = idx % cols;
    const row = Math.floor(idx / cols);
    const x = 0.5 + col * (cardW + gapX);
    const y = startY + row * (cardH + gapY);
    slide.addShape('roundRect', {
      x,
      y,
      w: cardW,
      h: cardH,
      radius: 0.06,
      line: { color: PPT_STYLE.border, pt: 1 },
      fill: { color: PPT_STYLE.panel },
    });
    slide.addText(m.label, { x: x + 0.1, y: y + 0.06, w: cardW - 0.2, h: 0.22, fontSize: 9, color: PPT_STYLE.muted, fontFace: PPT_STYLE.font });
    slide.addText(m.value, { x: x + 0.1, y: y + 0.27, w: cardW - 0.2, h: 0.34, bold: true, fontSize: 14, color: PPT_STYLE.text, fontFace: PPT_STYLE.font });
  });
}

function addPptSummaryTable(slide, rows, startY = PPT_LAYOUT_MAP.tableY, maxRows = 8) {
  const header = [
    { text: 'Cohort', options: { bold: true, color: '0F172A' } },
    { text: 'DBs', options: { bold: true, color: '0F172A' } },
    { text: 'Instances', options: { bold: true, color: '0F172A' } },
    { text: 'Hosts', options: { bold: true, color: '0F172A' } },
    { text: 'vCPU', options: { bold: true, color: '0F172A' } },
    { text: 'Memory (GB)', options: { bold: true, color: '0F172A' } },
  ];
  const body = rows.slice(0, maxRows).map((r) => [r.cohort, fmt(r.dbs, 'int'), fmt(r.instances, 'int'), fmt(r.hosts, 'int'), fmt(r.vcpu, 'dec1'), fmt(r.mem, 'dec1')]);
  slide.addTable([header, ...body], {
    x: 0.5,
    y: startY,
    w: 12.3,
    h: PPT_LAYOUT_MAP.tableH,
    fontSize: 10,
    color: PPT_STYLE.text,
    border: { pt: 1, color: PPT_STYLE.border },
    fill: 'FFFFFF',
    valign: 'mid',
  });
}

function ensureSvgXmlNs(svgText) {
  const txt = String(svgText || '');
  if (txt.includes('xmlns=')) return txt;
  return txt.replace('<svg', '<svg xmlns="http://www.w3.org/2000/svg"');
}

async function svgElementToPngDataUrl(svgEl, outW = 1400, outH = 800) {
  if (!svgEl) return null;
  const svgText = ensureSvgXmlNs(new XMLSerializer().serializeToString(svgEl));
  const blob = new Blob([svgText], { type: 'image/svg+xml;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  try {
    const img = await new Promise((resolve, reject) => {
      const im = new Image();
      im.onload = () => resolve(im);
      im.onerror = reject;
      im.src = url;
    });
    const canvas = document.createElement('canvas');
    canvas.width = outW;
    canvas.height = outH;
    const ctx = canvas.getContext('2d');
    if (!ctx) return null;
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, outW, outH);
    ctx.drawImage(img, 0, 0, outW, outH);
    return canvas.toDataURL('image/png');
  } finally {
    URL.revokeObjectURL(url);
  }
}

function chartPixelsFromGeom(geom, base = 420) {
  const pxW = Math.max(1000, Math.round(geom.w * base));
  const pxH = Math.max(460, Math.round(geom.h * base));
  return { outW: pxW, outH: pxH };
}

async function renderTempChartPng(targetId, renderFn, outW, outH) {
  const host = document.createElement('div');
  host.style.position = 'fixed';
  host.style.left = '-10000px';
  host.style.top = '0';
  host.style.width = `${Math.max(620, Math.round(outW / 2))}px`;
  host.style.background = '#fff';
  host.innerHTML = `<div class="chart-box" style="--chart-height:${Math.max(220, Math.round(outH / 2))}px;"><div id="${targetId}"></div></div>`;
  document.body.appendChild(host);
  try {
    renderFn(targetId);
    const svg = host.querySelector('svg');
    if (!svg) return null;
    return await svgElementToPngDataUrl(svg, outW, outH);
  } catch {
    return null;
  } finally {
    host.remove();
  }
}

async function buildSlideChartImages(slideDef) {
  if (!slideDef) return {};
  const primaryPx = chartPixelsFromGeom(PPT_LAYOUT_MAP.chartPrimary);
  const secondaryPx = chartPixelsFromGeom(PPT_LAYOUT_MAP.chartSecondary);
  if (slideDef.type === 'summary') {
    const series = [...state.selection.ppt.cohorts]
      .map((cohort) => {
        const ts = resolveCpuSeriesForCohort(cohort);
        if (!ts) return null;
        return { name: cohort, x: ts.x, y: ts.y };
      })
      .filter(Boolean);
    const img1 = await renderTempChartPng(
      'cpuLinePptTmpSummary',
      (id) => renderMultiLineChart(id, series, cpuFixedMax()),
      primaryPx.outW,
      primaryPx.outH,
    );
    const versionRows = computeVersionRows('ppt');
    const img2 = await renderTempChartPng(
      'versionPptTmpSummary',
      (id) =>
        renderCategoricalChart(
          id,
          versionRows.map((r) => ({ label: r.version, value: r.count })),
          { fmtKind: 'int', defaultType: 'pie' },
        ),
      secondaryPx.outW,
      secondaryPx.outH,
    );
    return { primary: img1, secondary: img2 };
  }
  if (slideDef.type === 'cohort') {
    const series = buildInstanceCpuTimeSeries('ppt', slideDef.cohort);
    const img1 = await renderTempChartPng(
      'cpuLinePptTmpCohort',
      (id) => renderMultiLineChart(id, series, cpuFixedMax()),
      primaryPx.outW,
      primaryPx.outH,
    );
    const memRows = aggregateMetricByDb('ppt', 'DB Memory (MB)', (v) => v / 1024, slideDef.cohort);
    const img2 = await renderTempChartPng(
      'memPptTmpCohort',
      (id) =>
        renderCategoricalChart(
          id,
          memRows.map((r) => ({ label: displayDbName(r.db), value: r.p95 || r.max || 0 })),
          { fmtKind: 'dec1', defaultType: 'bar' },
        ),
      secondaryPx.outW,
      secondaryPx.outH,
    );
    return { primary: img1, secondary: img2 };
  }
  if (slideDef.type === 'instance') {
    const series = buildInstanceCpuTimeSeries('ppt', slideDef.cohort).filter((s) => s.key === slideDef.instance);
    const img1 = await renderTempChartPng(
      'cpuLinePptTmpInst',
      (id) => renderMultiLineChart(id, series, cpuFixedMax()),
      primaryPx.outW,
      primaryPx.outH,
    );
    const dbScope = new Set([slideDef.db]);
    const memRows = aggregateMetricByDb('ppt', 'DB Memory (MB)', (v) => v / 1024, slideDef.cohort, dbScope);
    const img2 = await renderTempChartPng(
      'memPptTmpInst',
      (id) =>
        renderCategoricalChart(
          id,
          memRows.map((r) => ({ label: displayDbName(r.db), value: r.p95 || r.max || 0 })),
          { fmtKind: 'dec1', defaultType: 'bar' },
        ),
      secondaryPx.outW,
      secondaryPx.outH,
    );
    return { primary: img1, secondary: img2 };
  }
  return {};
}

function normalizePptTimeLabel(raw) {
  const s = String(raw || '');
  if (!s) return '';
  if (s.includes('T')) return s.replace('T', ' ').slice(5, 16);
  return s.slice(0, 12);
}

function downsampleLineSeries(seriesRaw, maxPoints = 36) {
  const clean = (seriesRaw || []).filter((s) => Array.isArray(s?.x) && Array.isArray(s?.y) && s.x.length && s.y.length);
  if (!clean.length) return { labels: [], series: [] };
  const base = clean.reduce((a, b) => (a.x.length >= b.x.length ? a : b));
  const xLen = base.x.length;
  const idx = [];
  if (xLen <= maxPoints) {
    for (let i = 0; i < xLen; i += 1) idx.push(i);
  } else {
    const last = xLen - 1;
    for (let i = 0; i < maxPoints; i += 1) idx.push(Math.round((i * last) / (maxPoints - 1)));
  }
  const uniqIdx = [...new Set(idx)].sort((a, b) => a - b);
  const labels = uniqIdx.map((i) => normalizePptTimeLabel(base.x[i]));
  const series = clean.map((s) => {
    const vals = uniqIdx.map((i) => Number(s.y[Math.min(i, s.y.length - 1)]) || 0);
    return { name: s.name, values: vals };
  });
  return { labels, series };
}

function resolvePptChartType(pptx, key, fallback) {
  return pptx?.ChartType?.[key] || fallback;
}

function pptColorForItem(item) {
  const palette = [PPT_STYLE.accent3, PPT_STYLE.accent1, PPT_STYLE.accent2, PPT_STYLE.accent5, PPT_STYLE.accent6, PPT_STYLE.accent4];
  const idx = hashString(item) % palette.length;
  return palette[idx];
}

function addNativeLineChart(slide, pptx, geom, seriesRaw, opts = {}) {
  const ds = downsampleLineSeries(seriesRaw, 42);
  if (!ds.labels.length || !ds.series.length) return false;
  const topSeries = [...ds.series]
    .sort((a, b) => Math.max(...b.values) - Math.max(...a.values))
    .slice(0, 8);
  if (!topSeries.length) return false;
  const data = topSeries.map((s) => ({
    name: shortLabel(s.name, 28),
    labels: ds.labels,
    values: s.values,
  }));
  const chartColors = topSeries.map((s) => pptColorForItem(s.name));
  const chartOpts = {
    x: geom.x,
    y: geom.y,
    w: geom.w,
    h: geom.h,
    showLegend: true,
    legendPos: 'b',
    catAxisLabelSize: 8,
    valAxisLabelSize: 8,
    valAxisMinVal: 0,
    lineSize: 2,
    chartColors,
  };
  if (opts.fixedMax != null) chartOpts.valAxisMaxVal = Number(opts.fixedMax);
  slide.addChart(resolvePptChartType(pptx, 'line', 'line'), data, chartOpts);
  return true;
}

function addNativeBarChart(slide, pptx, geom, items, seriesName = 'Value') {
  const top = (items || []).filter((x) => Number(x?.value) > 0).slice(0, 10);
  if (!top.length) return false;
  const data = [
    {
      name: seriesName,
      labels: top.map((x) => shortLabel(x.label, 22)),
      values: top.map((x) => Number(x.value) || 0),
    },
  ];
  slide.addChart(resolvePptChartType(pptx, 'bar', 'bar'), data, {
    x: geom.x,
    y: geom.y,
    w: geom.w,
    h: geom.h,
    barDir: 'col',
    barGrouping: 'clustered',
    showLegend: false,
    catAxisLabelSize: 8,
    valAxisLabelSize: 8,
    valAxisMinVal: 0,
    chartColors: [PPT_STYLE.primary],
  });
  return true;
}

function addNativePieChart(slide, pptx, geom, items, seriesName = 'Distribution') {
  const top = (items || [])
    .filter((x) => Number(x?.value) > 0)
    .sort((a, b) => (Number(b.value) || 0) - (Number(a.value) || 0))
    .slice(0, 8);
  if (!top.length) return false;
  const data = [
    {
      name: seriesName,
      labels: top.map((x) => shortLabel(x.label, 24)),
      values: top.map((x) => Number(x.value) || 0),
    },
  ];
  slide.addChart(resolvePptChartType(pptx, 'pie', 'pie'), data, {
    x: geom.x,
    y: geom.y,
    w: geom.w,
    h: geom.h,
    showLegend: true,
    legendPos: 'r',
    chartColors: top.map((x) => pptColorForItem(x.label)),
  });
  return true;
}

function addNativeChartsForSlide(slide, pptx, slideDef) {
  const out = { primary: false, secondary: false };
  if (!slideDef) return out;
  try {
    if (slideDef.type === 'summary') {
      const cpuSeries = [...state.selection.ppt.cohorts]
        .map((cohort) => {
          const ts = resolveCpuSeriesForCohort(cohort);
          if (!ts) return null;
          return { name: cohort, x: ts.x, y: ts.y };
        })
        .filter(Boolean);
      out.primary = addNativeLineChart(slide, pptx, PPT_LAYOUT_MAP.chartPrimary, cpuSeries, {
        fixedMax: cpuFixedMax(),
      });
      const versionRows = computeVersionRows('ppt').map((r) => ({ label: r.version, value: r.count }));
      out.secondary = addNativePieChart(slide, pptx, PPT_LAYOUT_MAP.chartSecondary, versionRows, 'DB Versions');
      return out;
    }
    if (slideDef.type === 'cohort') {
      const cpuSeries = buildInstanceCpuTimeSeries('ppt', slideDef.cohort);
      out.primary = addNativeLineChart(slide, pptx, PPT_LAYOUT_MAP.chartPrimary, cpuSeries, {
        fixedMax: cpuFixedMax(),
      });
      const memRows = aggregateMetricByDb('ppt', 'DB Memory (MB)', (v) => v / 1024, slideDef.cohort)
        .sort((a, b) => (b.p95 || b.max || 0) - (a.p95 || a.max || 0))
        .map((r) => ({ label: displayDbName(r.db), value: r.p95 || r.max || 0 }));
      out.secondary = addNativeBarChart(slide, pptx, PPT_LAYOUT_MAP.chartSecondary, memRows, 'P95 Memory (GB)');
      return out;
    }
    if (slideDef.type === 'instance') {
      const cpuSeries = buildInstanceCpuTimeSeries('ppt', slideDef.cohort).filter((s) => s.key === slideDef.instance);
      out.primary = addNativeLineChart(slide, pptx, PPT_LAYOUT_MAP.chartPrimary, cpuSeries, {
        fixedMax: cpuFixedMax(),
      });
      const dbScope = new Set([slideDef.db]);
      const m = aggregateMetricByDb('ppt', 'DB Memory (MB)', (v) => v / 1024, slideDef.cohort, dbScope)[0];
      const memItems = m
        ? [
            { label: 'Min', value: m.min || 0 },
            { label: 'P50', value: m.p50 || 0 },
            { label: 'P95', value: m.p95 || 0 },
            { label: 'Max', value: m.max || 0 },
          ]
        : [];
      out.secondary = addNativeBarChart(slide, pptx, PPT_LAYOUT_MAP.chartSecondary, memItems, 'Memory Stats (GB)');
      return out;
    }
  } catch {
    return out;
  }
  return out;
}

async function exportPptFromStoryboard() {
  if (!state.graph || !state.raw) {
    setStatus('No data loaded. Load JSON first.');
    return;
  }
  if (!window.PptxGenJS || !window.JSZip) {
    setStatus('PPT libraries not loaded (PptxGenJS/JSZip). Hard refresh and try again.');
    return;
  }
  const selectedIds = (state.pptSlidePlan || []).filter((s) => state.pptSlideSelected.has(s.id)).map((s) => s.id);
  if (!selectedIds.length) {
    setStatus('No PPT slides selected.');
    return;
  }

  const previewBtn = $('previewExportBtn');
  const previewStatus = $('pptPreviewStatus');
  state.exportInProgress = true;
  enableReportActions();
  setStatus(`Starting PPT export (${selectedIds.length} slides)...`);
  if (previewBtn) previewBtn.textContent = 'Exporting...';
  if (previewStatus) previewStatus.textContent = `Preparing export (0/${selectedIds.length})...`;

  const pptx = new window.PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.author = 'AWR Review App';
  pptx.subject = 'AWR Analysis';
  pptx.title = 'AWR Analysis Report';
  pptx.theme = {
    lang: 'en-US',
    headFontFace: PPT_STYLE.font,
    bodyFontFace: PPT_STYLE.font,
  };

  const { rows: pptRows } = slideDataForSelection();
  const byCohortRows = computeByCohort('ppt');
  const summary = computeGlobal('ppt');

  const totalSlides = selectedIds.length;
  try {
    for (let idx = 0; idx < selectedIds.length; idx += 1) {
      const id = selectedIds[idx];
      const def = findSlideById(id);
      if (!def) continue;
      const slide = pptx.addSlide();
      const pageNo = idx + 1;
      if (previewStatus) previewStatus.textContent = `Generating slide ${pageNo}/${selectedIds.length}: ${def.title}`;

      if (def.type === 'title') {
        const customer = String(state.reportMeta?.customerName || '').trim();
        const subtitle = customer ? `AWR Analysis for: ${customer}` : 'AWR Analysis';
        addPptHeader(slide, 'AWR Analysis', subtitle, pageNo, totalSlides);
        slide.addShape('roundRect', {
          x: PPT_LAYOUT_MAP.coverBox.x,
          y: PPT_LAYOUT_MAP.coverBox.y,
          w: PPT_LAYOUT_MAP.coverBox.w,
          h: PPT_LAYOUT_MAP.coverBox.h,
          radius: 0.08,
          line: { color: PPT_STYLE.border, pt: 1 },
          fill: { color: PPT_STYLE.panel },
        });
        const roleLines = [
          ['Sales Rep', String(state.reportMeta?.salesRepName || '').trim()],
          ['Architect', String(state.reportMeta?.architectName || '').trim()],
          ['Engineer', String(state.reportMeta?.engineerName || '').trim()],
        ]
          .filter(([, value]) => value)
          .map(([role, value]) => `${role}: ${value}`);
        if (roleLines.length) {
          slide.addText(roleLines.join('\n'), {
            x: 0.866, y: 5.339, w: 6.22, h: 1.26, fontSize: 16, color: PPT_STYLE.text, align: 'left', fontFace: PPT_STYLE.font,
          });
        }
        slide.addText(`Generated on ${new Date().toLocaleString()}`, {
          x: 0.866, y: 7.087, w: 4.13, h: 0.31, fontSize: 12, color: PPT_STYLE.muted, align: 'left', fontFace: PPT_STYLE.font,
        });
        continue;
      }

    if (def.type === 'separator') {
      addPptHeader(slide, def.title, def.subtitle || '', pageNo, totalSlides);
      slide.addShape('roundRect', {
        x: 0.8, y: 2.2, w: 11.7, h: 2.2, radius: 0.08, fill: { color: PPT_STYLE.panel }, line: { color: PPT_STYLE.border, pt: 1 },
      });
      slide.addText(def.title, { x: 1.1, y: 2.8, w: 11.1, h: 0.6, bold: true, fontSize: 32, color: PPT_STYLE.primaryDark, align: 'center', fontFace: PPT_STYLE.font });
        continue;
      }

    if (def.type === 'summary') {
      addPptHeader(slide, 'Global Summary', 'Selected scope from current PPT filters', pageNo, totalSlides);
      addPptKpiGrid(slide, [
        { label: 'Databases', value: fmt(summary.db_count, 'int') },
        { label: 'Hosts', value: fmt(summary.host_count, 'int') },
        { label: 'Instances', value: fmt(summary.instance_count, 'int') },
        { label: 'vCPU', value: fmt(summary.vcpu_total, 'dec1') },
        { label: 'Memory (GB)', value: fmt(summary.memory_gb_total, 'dec1') },
        { label: 'Allocated (GB)', value: fmt(summary.allocated_storage_gb, 'dec1') },
        { label: 'Used (GB)', value: fmt(summary.used_storage_gb, 'dec1') },
        { label: 'DB IOPS', value: fmt(summary.db_iops_total, 'dec1') },
      ], PPT_LAYOUT_MAP.kpiStartY);
      addPptSummaryTable(slide, byCohortRows, PPT_LAYOUT_MAP.tableY, 3);
      const native = addNativeChartsForSlide(slide, pptx, def);
      if (!native.primary || !native.secondary) {
        const imgs = await buildSlideChartImages(def);
        if (!native.primary && imgs.primary) slide.addImage({ data: imgs.primary, ...PPT_LAYOUT_MAP.chartPrimary });
        if (!native.secondary && imgs.secondary) slide.addImage({ data: imgs.secondary, ...PPT_LAYOUT_MAP.chartSecondary });
      }
        continue;
      }

    if (def.type === 'cohort') {
      const cohortRows = pptRows.filter((r) => r.cohort === def.cohort);
      const cSummary = computeSummaryFromRows(cohortRows);
      addPptHeader(slide, `Cohort: ${def.cohort}`, `Databases: ${cSummary.db_count} | Instances: ${cSummary.instance_count}`, pageNo, totalSlides);
      addPptKpiGrid(slide, [
        { label: 'Databases', value: fmt(cSummary.db_count, 'int') },
        { label: 'Hosts', value: fmt(cSummary.host_count, 'int') },
        { label: 'Instances', value: fmt(cSummary.instance_count, 'int') },
        { label: 'vCPU', value: fmt(cSummary.vcpu_total, 'dec1') },
        { label: 'Memory (GB)', value: fmt(cSummary.memory_gb_total, 'dec1') },
        { label: 'Allocated (GB)', value: fmt(cSummary.allocated_storage_gb, 'dec1') },
      ], PPT_LAYOUT_MAP.kpiStartY);
      const dbCounts = new Map();
      cohortRows.forEach((r) => {
        dbCounts.set(r.db, (dbCounts.get(r.db) || 0) + 1);
      });
      const dbRows = [...dbCounts.entries()]
        .sort((a, b) => b[1] - a[1] || String(a[0]).localeCompare(String(b[0])))
        .slice(0, 5)
        .map(([db, n]) => ({ db, n }));
      slide.addTable(
        [
          [
            { text: 'Database', options: { bold: true } },
            { text: 'Instances', options: { bold: true } },
          ],
          ...dbRows.map((r) => [shortLabel(displayDbName(r.db), 34), String(r.n)]),
        ],
        { x: 0.5, y: PPT_LAYOUT_MAP.tableY, w: 12.3, h: PPT_LAYOUT_MAP.tableH, fontSize: 9, border: { pt: 1, color: PPT_STYLE.border }, color: PPT_STYLE.text },
      );
      const native = addNativeChartsForSlide(slide, pptx, def);
      if (!native.primary || !native.secondary) {
        const imgs = await buildSlideChartImages(def);
        if (!native.primary && imgs.primary) slide.addImage({ data: imgs.primary, ...PPT_LAYOUT_MAP.chartPrimary });
        if (!native.secondary && imgs.secondary) slide.addImage({ data: imgs.secondary, ...PPT_LAYOUT_MAP.chartSecondary });
      }
        continue;
      }

    if (def.type === 'instance') {
      const row = pptRows.find((r) => r.cohort === def.cohort && r.instance === def.instance);
      addPptHeader(slide, `Instance: ${displayInstanceName(def.instance)}`, `${def.cohort} | ${displayDbName(def.db)}`, pageNo, totalSlides);
      addPptKpiGrid(slide, [
        { label: 'Cohort', value: def.cohort || 'N/A' },
        { label: 'Database', value: displayDbName(row?.db || def.db || 'N/A') },
        { label: 'Host', value: row?.host || 'N/A' },
        { label: 'vCPU', value: fmt(row?.init_cpu_count || row?.logical_cpu_count || 0, 'dec1') },
        { label: 'Memory (GB)', value: fmt(row?.mem_gb || 0, 'dec1') },
        { label: 'SGA (GB)', value: fmt(row?.sga_size_gb || 0, 'dec1') },
        { label: 'PGA (GB)', value: fmt(row?.pga_size_gb || 0, 'dec1') },
      ], PPT_LAYOUT_MAP.kpiStartY);
      slide.addText('This slide follows your current selection and uses native editable PowerPoint charts.', {
        x: 0.6, y: PPT_LAYOUT_MAP.tableY, w: 12.0, h: 0.35, fontSize: 10, color: PPT_STYLE.muted, fontFace: PPT_STYLE.font,
      });
      const native = addNativeChartsForSlide(slide, pptx, def);
      if (!native.primary || !native.secondary) {
        const imgs = await buildSlideChartImages(def);
        if (!native.primary && imgs.primary) slide.addImage({ data: imgs.primary, ...PPT_LAYOUT_MAP.chartPrimary });
        if (!native.secondary && imgs.secondary) slide.addImage({ data: imgs.secondary, ...PPT_LAYOUT_MAP.chartSecondary });
      }
      }
    }

    if (previewStatus) previewStatus.textContent = 'Writing .pptx file...';
    const stamp = new Date().toISOString().replaceAll(':', '-').slice(0, 19);
    await pptx.writeFile({ fileName: `AWR_Analysis_${stamp}.pptx` });
    setStatus(`PPT exported with ${selectedIds.length} slides.`);
    if (previewStatus) previewStatus.textContent = `Export completed: ${selectedIds.length} slides.`;
  } finally {
    state.exportInProgress = false;
    enableReportActions();
    if (previewBtn) previewBtn.textContent = 'Export PPT';
  }
}

function fileNameFromDisposition(headerVal) {
  const raw = String(headerVal || '');
  const m = raw.match(/filename\*=UTF-8''([^;]+)|filename=\"?([^\";]+)\"?/i);
  const enc = m?.[1];
  const plain = m?.[2];
  if (enc) {
    try {
      return decodeURIComponent(enc);
    } catch {
      return enc;
    }
  }
  if (plain) return plain;
  return '';
}

function toBase64(bytes) {
  let binary = '';
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    const sub = bytes.subarray(i, i + chunk);
    binary += String.fromCharCode(...sub);
  }
  return btoa(binary);
}

async function probeTemplateApi() {
  // Keep quick path only when already confirmed reachable.
  // If it was unreachable before, retry because server/template state can change.
  if (state.templateApiAvailable === true) {
    updateTemplateStatusText();
    return state.templateApiAvailable;
  }
  try {
    const res = await fetch('/api/health', { method: 'GET' });
    if (!res.ok) {
      state.templateApiAvailable = false;
      state.templatePath = '';
      updateTemplateStatusText();
      return false;
    }
    const data = await res.json().catch(() => ({}));
    state.templateApiAvailable = Boolean(data?.ok);
    state.templatePath = data?.template_path || '';
    updateTemplateStatusText();
    return state.templateApiAvailable;
  } catch {
    state.templateApiAvailable = false;
    state.templatePath = '';
    updateTemplateStatusText();
    return false;
  }
}

async function probeAiConnection() {
  try {
    const res = await fetch('/api/chat-status', { method: 'GET' });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const data = await res.json().catch(() => ({}));
    state.aiChat.connected = Boolean(data?.connected);
    state.aiChat.connectionSource = data?.source || 'unknown';
    state.aiChat.model = data?.model || state.aiChat.model || 'gpt-4.1-mini';
    state.aiChat.connectionError = data?.error || '';
    renderAiChatAssistant();
    return state.aiChat.connected;
  } catch {
    state.aiChat.connected = false;
    state.aiChat.connectionSource = 'unavailable';
    state.aiChat.connectionError = 'Connection check failed.';
    renderAiChatAssistant();
    return false;
  }
}

async function sendAiMessage() {
  const question = (state.aiChat.question || '').trim();
  if (!question) {
    setStatus('Type a question first.');
    return;
  }
  const connected = await probeAiConnection();
  if (!connected) {
    await sendAiViaWeb();
    return;
  }
  const prompt = buildAiPromptText();
  state.aiChat.prompt = prompt;
  state.aiChat.busy = true;
  state.aiChat.messages = [...(state.aiChat.messages || []), { role: 'user', content: question }];
  renderAiChatAssistant();

  try {
    const res = await fetch('/api/chat-assistant', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        model: state.aiChat.model || 'gpt-4.1-mini',
        messages: [
          { role: 'system', content: 'You are an Oracle AWR analyst. Respond with factual and quantitative findings first.' },
          { role: 'user', content: prompt },
        ],
      }),
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok || !data?.ok) {
      throw new Error(data?.error || `HTTP ${res.status}`);
    }
    const answer = String(data.answer || '').trim() || 'No response content.';
    state.aiChat.response = answer;
    try {
      const parsed = JSON.parse(stripJsonCodeFence(answer));
      const check = validateAppPayload(parsed);
      if (check.valid) {
        state.aiChat.appPayload = parsed;
        state.aiChat.payloadErrors = [];
      } else {
        state.aiChat.appPayload = null;
        state.aiChat.payloadErrors = check.errors;
      }
    } catch {
      // non-JSON assistant responses are allowed; user can paste APP_PAYLOAD manually
    }
    state.aiChat.messages = [...(state.aiChat.messages || []), { role: 'assistant', content: answer }];
    state.aiChat.question = '';
    setStatus('AI response received.');
  } catch (err) {
    setStatus(`AI request failed: ${err?.message || err}`);
  } finally {
    state.aiChat.busy = false;
    renderAiChatAssistant();
  }
}

async function sendAiViaWeb() {
  const basePrompt = state.aiChat.prompt || buildAiPromptText();
  const prompt = `${basePrompt}\n\nIMPORTANT: Return APP_PAYLOAD JSON only (no markdown), valid for app_payload.schema.json.`;
  state.aiChat.prompt = prompt;
  try {
    if (navigator?.clipboard?.writeText) {
      await navigator.clipboard.writeText(prompt);
    } else {
      const ta = document.createElement('textarea');
      ta.value = prompt;
      document.body.appendChild(ta);
      ta.select();
      document.execCommand('copy');
      ta.remove();
    }
    window.open(AWR_DBA_ADVISOR_URL, '_blank', 'noopener,noreferrer');
    setStatus('Prompt copied. AWR DBA Advisor opened. Paste and send there.');
  } catch {
    window.open(AWR_DBA_ADVISOR_URL, '_blank', 'noopener,noreferrer');
    setStatus('AWR DBA Advisor opened. Copy prompt manually from the panel and paste there.');
  } finally {
    renderAiChatAssistant();
  }
}

async function uploadTemplateToApi(file) {
  if (!file) return false;
  const ok = await probeTemplateApi();
  if (!ok) {
    throw new Error('Template server unavailable. Start app_server.py first.');
  }
  const ab = await file.arrayBuffer();
  const b64 = toBase64(new Uint8Array(ab));
  const res = await fetch('/api/template', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      filename: file.name,
      content_b64: b64,
    }),
  });
  if (!res.ok) {
    let msg = `HTTP ${res.status}`;
    try {
      const j = await res.json();
      if (j?.error) msg = j.error;
    } catch {
      // ignore
    }
    throw new Error(msg);
  }
  const data = await res.json().catch(() => ({}));
  state.templateApiAvailable = true;
  state.templatePath = data?.template_path || '';
  updateTemplateStatusText();
  setStatus(`Template loaded: ${file.name}`);
  const previewStatus = $('pptPreviewStatus');
  if (previewStatus && data?.template_path) {
    previewStatus.textContent = `Template active: ${data.template_path}`;
  }
  return true;
}

async function exportPptViaTemplateApi() {
  if (!state.graph || !state.raw) {
    setStatus('No data loaded. Load JSON first.');
    return false;
  }
  const selectedIds = (state.pptSlidePlan || []).filter((s) => state.pptSlideSelected.has(s.id)).map((s) => s.id);
  if (!selectedIds.length) {
    setStatus('No PPT slides selected.');
    return false;
  }

  const previewBtn = $('previewExportBtn');
  const previewStatus = $('pptPreviewStatus');
  state.exportInProgress = true;
  enableReportActions();
  if (previewBtn) previewBtn.textContent = 'Exporting...';
  if (previewStatus) previewStatus.textContent = 'Generating template-native PPT...';
  setStatus(`Starting template-native PPT export (${selectedIds.length} slides)...`);

  try {
    const payload = { report: reportStateSnapshot() };
    const res = await fetch('/api/export-template', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });
    if (!res.ok) {
      let msg = `HTTP ${res.status}`;
      try {
        const j = await res.json();
        if (j?.error) msg = j.error;
      } catch {
        // ignore parse errors
      }
      throw new Error(msg);
    }
    const blob = await res.blob();
    const nameFromHeader = fileNameFromDisposition(res.headers.get('Content-Disposition'));
    const stamp = new Date().toISOString().replaceAll(':', '-').slice(0, 19);
    const fileName = nameFromHeader || `AWR_Analysis_Template_${stamp}.pptx`;
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    setStatus(`Template-native PPT exported (${fileName}).`);
    if (previewStatus) previewStatus.textContent = 'Template-native export completed.';
    return true;
  } finally {
    state.exportInProgress = false;
    enableReportActions();
    if (previewBtn) previewBtn.textContent = 'Export PPT';
  }
}

function computeGlobal(kind) {
  const rows = selectedInstances(kind);
  const dbs = new Set(rows.map((r) => r.db));
  const hosts = new Set(rows.map((r) => r.host));
  const cohorts = new Set(rows.map((r) => r.cohort));

  let allocated = 0;
  let used = 0;
  let iops = 0;
  let logons = 0;
  cohorts.forEach((c) => {
    const item = state.graph.cohortRollups[c];
    allocated += item?.allocated || 0;
    used += item?.used || 0;
    iops += item?.db_iops || 0;
    logons += item?.db_logons || 0;
  });

  return {
    db_count: dbs.size,
    host_count: hosts.size,
    instance_count: rows.length,
    vcpu_total: rows.reduce((a, r) => a + (r.init_cpu_count || r.logical_cpu_count || 0), 0),
    memory_gb_total: rows.reduce((a, r) => a + r.mem_gb, 0),
    allocated_storage_gb: allocated,
    used_storage_gb: used,
    db_iops_total: iops,
    db_logons_total: logons,
  };
}

function computeByCohort(kind) {
  const rows = selectedInstances(kind);
  const groups = new Map();
  for (const r of rows) {
    if (!groups.has(r.cohort)) {
      groups.set(r.cohort, { cohort: r.cohort, dbs: new Set(), instances: 0, hosts: new Set(), vcpu: 0, mem: 0 });
    }
    const g = groups.get(r.cohort);
    g.dbs.add(r.db);
    g.hosts.add(r.host);
    g.instances += 1;
    g.vcpu += r.init_cpu_count || r.logical_cpu_count || 0;
    g.mem += r.mem_gb;
  }
  return [...groups.values()]
    .map((g) => ({
      cohort: g.cohort,
      dbs: g.dbs.size,
      instances: g.instances,
      hosts: g.hosts.size,
      vcpu: g.vcpu,
      mem: g.mem,
      allocated: state.graph.cohortRollups[g.cohort]?.allocated || 0,
      used: state.graph.cohortRollups[g.cohort]?.used || 0,
      iops: state.graph.cohortRollups[g.cohort]?.db_iops || 0,
      logons: state.graph.cohortRollups[g.cohort]?.db_logons || 0,
    }))
    .sort((a, b) => a.cohort.localeCompare(b.cohort));
}

function metricGroupsForRender(sel) {
  const groups = new Map();
  state.metricCatalog.forEach((m) => {
    if (!groups.has(m.category)) groups.set(m.category, []);
    groups.get(m.category).push(m);
  });
  return [...groups.entries()].map(([category, metrics]) => ({
    category,
    metrics: metrics.sort((a, b) => a.label.localeCompare(b.label)).map((m) => ({ ...m, checked: sel.metrics.has(m.id) })),
  }));
}

function renderSelectors() {
  const g = state.graph;
  const sel = activeSelection();

  $('cohortSelect').innerHTML = g.cohorts
    .map(
      (c) => `<label class="scope-item"><input type="checkbox" data-scope-kind="cohort" value="${escAttr(c)}" ${
        sel.cohorts.has(c) ? 'checked' : ''
      } /><span title="${escAttr(c)}">${esc(c)}</span></label>`,
    )
    .join('');

  const dbOptions = [...new Set(g.instances.filter((x) => sel.cohorts.has(x.cohort)).map((x) => x.db))].sort();
  $('dbSelect').innerHTML = dbOptions
    .map(
      (d) => `<label class="scope-item"><input type="checkbox" data-scope-kind="db" value="${escAttr(d)}" ${
        sel.dbs.has(d) ? 'checked' : ''
      } /><span title="${escAttr(displayDbName(d))}">${esc(displayDbName(d))}</span></label>`,
    )
    .join('');

  const instanceOptions = [
    ...new Set(g.instances.filter((x) => sel.cohorts.has(x.cohort) && sel.dbs.has(x.db)).map((x) => x.instance)),
  ].sort();
  $('instanceSelect').innerHTML = instanceOptions
    .map(
      (i) => `<label class="scope-item"><input type="checkbox" data-scope-kind="instance" value="${escAttr(i)}" ${
        sel.instances.has(i) ? 'checked' : ''
      } /><span title="${escAttr(displayInstanceName(i))}">${esc(displayInstanceName(i))}</span></label>`,
    )
    .join('');

  $('metricsList').innerHTML = metricGroupsForRender(sel)
    .map(
      (group) => `
      <div class="metric-group">
        <h3>${esc(group.category)}</h3>
        ${group.metrics
          .map(
            (m) => `<label class="metric-item"><input type="checkbox" data-metric="${esc(m.id)}" ${
              m.checked ? 'checked' : ''
            } /> ${esc(m.label)}</label>`,
          )
          .join('')}
      </div>
    `,
    )
    .join('');

}

function renderBars(containerId, rows, valueKey) {
  renderCategoricalChart(
    containerId,
    rows.map((r) => ({ label: r.cohort, value: Number(r[valueKey]) || 0 })),
    { fmtKind: 'dec1', defaultType: 'bar' },
  );
}

function renderCohortMetricCharts(rows) {
  const grid = $('cohortBarsGrid');
  if (!grid) return;
  const charts = [
    { key: 'instances', label: 'Instances by Cohort' },
    { key: 'vcpu', label: 'vCPU by Cohort' },
    { key: 'mem', label: 'Memory (GB) by Cohort' },
    { key: 'allocated', label: 'Allocated Storage (GB) by Cohort' },
    { key: 'used', label: 'Used Storage (GB) by Cohort' },
    { key: 'iops', label: 'DB IOPS by Cohort' },
    { key: 'logons', label: 'DB Logons by Cohort' },
  ];

  grid.innerHTML = charts
    .map((c, idx) => `<div class="chart-box mini-chart"><h3>${esc(c.label)}</h3><div id="cohortBars_${idx}"></div></div>`)
    .join('');

  charts.forEach((c, idx) => {
    renderBars(`cohortBars_${idx}`, rows, c.key);
  });
}

function computeVersionRows(kind) {
  const rows = selectedInstances(kind);
  const uniqueDbs = [...new Set(rows.map((r) => r.db))];
  const grouped = new Map();
  uniqueDbs.forEach((db) => {
    const version = state.graph.dbVersionByDb[db] || 'Unknown';
    if (!grouped.has(version)) grouped.set(version, { version, count: 0 });
    grouped.get(version).count += 1;
  });
  return [...grouped.values()].sort((a, b) => b.count - a.count || a.version.localeCompare(b.version));
}

function renderVersionChart(kind) {
  const versionRows = computeVersionRows(kind);
  renderCategoricalChart(
    'dbVersionChart',
    versionRows.map((r) => ({ label: r.version, value: r.count })),
    { fmtKind: 'int', defaultType: 'bar' },
  );
}

function shortLabel(text, max = 18) {
  if (text.length <= max) return text;
  return `${text.slice(0, max - 1)}...`;
}

function pieArcPath(cx, cy, r, startAngle, endAngle) {
  const x1 = cx + r * Math.cos(startAngle);
  const y1 = cy + r * Math.sin(startAngle);
  const x2 = cx + r * Math.cos(endAngle);
  const y2 = cy + r * Math.sin(endAngle);
  const largeArc = endAngle - startAngle > Math.PI ? 1 : 0;
  return `M ${cx} ${cy} L ${x1} ${y1} A ${r} ${r} 0 ${largeArc} 1 ${x2} ${y2} Z`;
}

function renderCategoricalChart(containerId, rawItems, opts = {}) {
  const container = $(containerId);
  if (!container) return;
  const fmtKind = opts.fmtKind || 'dec1';
  const defaultType = opts.defaultType || 'bar';
  const chartType = getChartType(containerId, defaultType);

  const items = (rawItems || [])
    .map((x) => ({ label: String(x.label), value: Number(x.value) || 0, color: x.color || colorForItem(x.label) }))
    .filter((x) => x.label);

  if (!items.length) {
    container.innerHTML = '<p class="muted">No data in current scope.</p>';
    return;
  }

  const topItems = [...items].sort((a, b) => b.value - a.value).slice(0, 10);
  const total = topItems.reduce((a, b) => a + b.value, 0);

  if (chartType === 'text') {
    container.innerHTML = `
      <div class="table-wrap">
        <table>
          <thead><tr><th>Item</th><th>Value</th><th>%</th></tr></thead>
          <tbody>
            ${topItems
              .map(
                (it) =>
                  `<tr><td>${esc(it.label)}</td><td>${fmt(it.value, fmtKind)}</td><td>${fmt(total ? (it.value / total) * 100 : 0, 'dec1')}%</td></tr>`,
              )
              .join('')}
          </tbody>
        </table>
      </div>
    `;
    return;
  }

  if (chartType === 'pie') {
    const w = 620;
    const h = 320;
    const cx = 190;
    const cy = 160;
    const r = 110;
    let a = -Math.PI / 2;
    let slices = '';
    topItems.forEach((it) => {
      const ratio = total ? it.value / total : 0;
      const delta = ratio * Math.PI * 2;
      const next = a + delta;
      const path = pieArcPath(cx, cy, r, a, next);
      slices += `<path d="${path}" fill="${it.color}" stroke="#ffffff" stroke-width="1.2" data-tip="${esc(
        `${it.label}: ${fmt(it.value, fmtKind)} (${fmt(ratio * 100, 'dec1')}%)`,
      )}"></path>`;
      a = next;
    });
    const legend = topItems
      .map(
        (it, i) =>
          `<text x="340" y="${28 + i * 22}" fill="#1f2a37" font-size="12">${esc(shortLabel(it.label, 28))}: ${fmt(it.value, fmtKind)}</text>` +
          `<rect x="320" y="${20 + i * 22}" width="12" height="12" fill="${it.color}" />`,
      )
      .join('');
    container.innerHTML = `<svg viewBox="0 0 ${w} ${h}" role="img">${slices}${legend}</svg>`;
    return;
  }

  if (chartType === 'line') {
    const w = 700;
    const h = 300;
    const left = 26;
    const right = 12;
    const top = 16;
    const bottom = 62;
    const plotW = w - left - right;
    const plotH = h - top - bottom;
    const maxValue = Math.max(...topItems.map((r) => r.value), 1);
    const n = Math.max(topItems.length, 2);
    const x = (i) => left + (i / (n - 1)) * plotW;
    const y = (v) => top + (1 - v / maxValue) * plotH;

    const pathD = topItems.map((it, i) => `${i === 0 ? 'M' : 'L'} ${x(i)} ${y(it.value)}`).join(' ');
    const points = topItems
      .map((it, i) => {
        const px = x(i);
        const py = y(it.value);
        return (
          `<circle cx="${px}" cy="${py}" r="3" fill="${it.color}" stroke="#fff" stroke-width="0.8" data-tip="${esc(
            `${it.label}: ${fmt(it.value, fmtKind)}`,
          )}"></circle>` +
          `<text x="${px}" y="${h - 42}" text-anchor="middle" font-size="10" fill="#596579">${esc(shortLabel(it.label, 11))}</text>`
        );
      })
      .join('');
    container.innerHTML = `<svg viewBox="0 0 ${w} ${h}" role="img"><path d="${pathD}" fill="none" stroke="#334155" stroke-width="2"></path>${points}</svg>`;
    return;
  }

  if (chartType === 'column') {
    const w = 700;
    const h = 300;
    const left = 24;
    const right = 12;
    const top = 18;
    const bottom = 86;
    const plotW = w - left - right;
    const plotH = h - top - bottom;
    const maxValue = Math.max(...topItems.map((r) => r.value), 1);
    const colW = plotW / Math.max(topItems.length, 1);
    let out = '';
    topItems.forEach((it, i) => {
      const height = (it.value / maxValue) * plotH;
      const x = left + i * colW + colW * 0.14;
      const y = top + plotH - height;
      const rw = colW * 0.72;
      out += `<rect x="${x}" y="${y}" width="${rw}" height="${height}" rx="4" fill="${it.color}" data-tip="${esc(`${it.label}: ${fmt(
        it.value,
        fmtKind,
      )}`)}"></rect>`;
      out += `<text x="${x + rw / 2}" y="${h - 52}" text-anchor="middle" font-size="10" fill="#596579">${esc(
        shortLabel(it.label, 10),
      )}</text>`;
      out += `<text x="${x + rw / 2}" y="${Math.max(y - 4, 12)}" text-anchor="middle" font-size="10" fill="#334155">${fmt(
        it.value,
        fmtKind,
      )}</text>`;
    });
    container.innerHTML = `<svg viewBox="0 0 ${w} ${h}" role="img">${out}</svg>`;
    return;
  }

  const maxValue = Math.max(...topItems.map((r) => r.value), 1);
  const barArea = 300;
  const barHeight = 18;
  const gap = 12;
  const left = 140;
  const top = 16;
  const h = top + topItems.length * (barHeight + gap) + 20;
  const w = 500;
  let bars = '';
  topItems.forEach((it, i) => {
    const y = top + i * (barHeight + gap);
    const width = Math.round((it.value / maxValue) * barArea);
    bars += `<text x="0" y="${y + 13}" fill="#1f2a37" font-size="11">${esc(shortLabel(it.label))}</text>`;
    bars += `<rect x="${left}" y="${y}" width="${barArea}" height="${barHeight}" rx="3" fill="#eef2f7"></rect>`;
    bars += `<rect x="${left}" y="${y}" width="${width}" height="${barHeight}" rx="3" fill="${it.color}" opacity="0.92" data-tip="${esc(
      `${it.label}: ${fmt(it.value, fmtKind)}`,
    )}"></rect>`;
    bars += `<text x="${left + width + 6}" y="${y + 13}" fill="#374151" font-size="11">${fmt(it.value, fmtKind)}</text>`;
  });
  container.innerHTML = `<svg viewBox="0 0 ${w} ${h}" role="img">${bars}</svg>`;
}

function aggregateMetricByCohort(kind, metricName, transformValue = (v) => v) {
  const sel = state.selection[kind];
  const rows = state.graph.dbStats.filter(
    (r) =>
      sel.cohorts.has(r.cohort) &&
      sel.dbs.has(r.db) &&
      r.metric === metricName &&
      (!r.summary_of || r.summary_of === 'cdb'),
  );

  const grouped = new Map();
  rows.forEach((r) => {
    if (!grouped.has(r.cohort)) {
      grouped.set(r.cohort, { cohort: r.cohort, count: 0, min: 0, p30: 0, p50: 0, p70: 0, p95: 0, p99: 0, max: 0 });
    }
    const g = grouped.get(r.cohort);
    g.count += 1;
    g.min += r.min || 0;
    g.p30 += r.p30 || 0;
    g.p50 += r.p50 || 0;
    g.p70 += r.p70 || 0;
    g.p95 += r.p95 || 0;
    g.p99 += r.p99 || 0;
    g.max += r.max || 0;
  });

  return [...grouped.values()].map((g) => ({
    cohort: g.cohort,
    min: transformValue(g.min / g.count),
    p30: transformValue(g.p30 / g.count),
    p50: transformValue(g.p50 / g.count),
    p70: transformValue(g.p70 / g.count),
    p95: transformValue(g.p95 / g.count),
    p99: transformValue(g.p99 / g.count),
    max: transformValue(g.max / g.count),
  }));
}

function aggregateMetricByDb(kind, metricName, transformValue = (v) => v, cohortFilter = null, dbFilter = null) {
  const sel = state.selection[kind];
  const rows = state.graph.dbStats.filter(
    (r) =>
      sel.cohorts.has(r.cohort) &&
      sel.dbs.has(r.db) &&
      r.metric === metricName &&
      (!r.summary_of || r.summary_of === 'cdb') &&
      (cohortFilter ? r.cohort === cohortFilter : true) &&
      (dbFilter ? dbFilter.has(r.db) : true),
  );

  const grouped = new Map();
  rows.forEach((r) => {
    if (!grouped.has(r.db)) {
      grouped.set(r.db, { db: r.db, count: 0, min: 0, p30: 0, p50: 0, p70: 0, p95: 0, p99: 0, max: 0 });
    }
    const g = grouped.get(r.db);
    g.count += 1;
    g.min += r.min || 0;
    g.p30 += r.p30 || 0;
    g.p50 += r.p50 || 0;
    g.p70 += r.p70 || 0;
    g.p95 += r.p95 || 0;
    g.p99 += r.p99 || 0;
    g.max += r.max || 0;
  });

  return [...grouped.values()].map((g) => ({
    db: g.db,
    min: transformValue(g.min / g.count),
    p30: transformValue(g.p30 / g.count),
    p50: transformValue(g.p50 / g.count),
    p70: transformValue(g.p70 / g.count),
    p95: transformValue(g.p95 / g.count),
    p99: transformValue(g.p99 / g.count),
    max: transformValue(g.max / g.count),
  }));
}

function renderRankBarsByGroup(containerId, rows, valueKey, fmtKind, labelKey, fixedMax = null) {
  const topRows = [...rows].sort((a, b) => b[valueKey] - a[valueKey]).slice(0, 10);
  const items = topRows.map((r) => ({ label: r[labelKey], value: Number(r[valueKey]) || 0 }));
  renderCategoricalChart(containerId, items, { fmtKind, defaultType: 'bar' });
}

function renderBoxPlotByGroup(containerId, rows, fmtKind, labelKey, fixedMax = null) {
  const container = $(containerId);
  if (!rows.length) {
    container.innerHTML = '<p class="muted">No data in current scope.</p>';
    return;
  }

  const topRows = [...rows].sort((a, b) => b.max - a.max).slice(0, 8);
  const minScale = fixedMax != null ? 0 : Math.min(...topRows.map((r) => r.min), 0);
  const maxScale = fixedMax != null ? Math.max(fixedMax, 1) : Math.max(...topRows.map((r) => r.max), 1);
  const span = Math.max(maxScale - minScale, 1);
  const scaleX = (v) => 140 + ((v - minScale) / span) * 300;

  const rowHeight = 26;
  const top = 16;
  const h = top + rowHeight * topRows.length + 20;
  const w = 500;

  let shapes = '';
  topRows.forEach((r, i) => {
    const y = top + i * rowHeight;
    const yMid = y + 10;
    const xMin = scaleX(r.min);
    const xQ1 = scaleX(r.p30);
    const xMed = scaleX(r.p50);
    const xQ3 = scaleX(r.p70);
    const xMax = scaleX(r.max);

    const itemKey = r[labelKey];
    const color = colorForItem(itemKey);
    shapes += `<text x="0" y="${yMid + 4}" fill="#1f2a37" font-size="11">${esc(shortLabel(r[labelKey]))}</text>`;
    shapes += `<line x1="${xMin}" y1="${yMid}" x2="${xMax}" y2="${yMid}" stroke="#7d8593" stroke-width="1.2" data-tip="${esc(
      `${r[labelKey]} min/max: ${fmt(r.min, fmtKind)} / ${fmt(r.max, fmtKind)}`,
    )}"><title>${esc(`${r[labelKey]} min/max: ${fmt(r.min, fmtKind)} / ${fmt(r.max, fmtKind)}`)}</title></line>`;
    shapes += `<rect x="${Math.min(xQ1, xQ3)}" y="${yMid - 6}" width="${Math.max(xQ3 - xQ1, 2)}" height="12" fill="${color}22" stroke="${color}" data-tip="${esc(
      `${r[labelKey]} p30-p70: ${fmt(r.p30, fmtKind)} - ${fmt(r.p70, fmtKind)}`,
    )}"><title>${esc(`${r[labelKey]} p30-p70: ${fmt(r.p30, fmtKind)} - ${fmt(r.p70, fmtKind)}`)}</title></rect>`;
    shapes += `<line x1="${xMed}" y1="${yMid - 7}" x2="${xMed}" y2="${yMid + 7}" stroke="${color}" stroke-width="1.5" data-tip="${esc(
      `${r[labelKey]} p50: ${fmt(r.p50, fmtKind)}`,
    )}"><title>${esc(`${r[labelKey]} p50: ${fmt(r.p50, fmtKind)}`)}</title></line>`;
    shapes += `<circle cx="${xMin}" cy="${yMid}" r="1.7" fill="#7d8593"></circle>`;
    shapes += `<circle cx="${xMax}" cy="${yMid}" r="1.9" fill="${color}"></circle>`;
  });

  shapes += `<text x="140" y="${h - 4}" font-size="10" fill="#596579">${fmt(minScale, fmtKind)}</text>`;
  shapes += `<text x="430" y="${h - 4}" font-size="10" fill="#596579">${fmt(maxScale, fmtKind)}</text>`;
  container.innerHTML = `<svg viewBox="0 0 ${w} ${h}" role="img">${shapes}</svg>`;
}

function renderCpuMemorySections(kind) {
  const sel = state.selection[kind];
  const cpuSeries = [...sel.cohorts]
    .map((cohort) => {
      const ts = resolveCpuSeriesForCohort(cohort);
      if (!ts) return null;
      return { name: cohort, x: ts.x, y: ts.y };
    })
    .filter(Boolean);
  const cpuRows = statsRowsFromSeries(cpuSeries, 'cohort');
  renderCpuProfileLineGlobal(kind);
  const cpuMax = cpuFixedMax();
  renderRankBarsByGroup('cpuTopChart', cpuRows, 'max', 'dec1', 'cohort', cpuMax);
  renderRankBarsByGroup('cpuP95Chart', cpuRows, 'p95', 'dec3', 'cohort', cpuMax);
  renderBoxPlotByGroup('cpuBoxChart', cpuRows, 'dec1', 'cohort', cpuMax);

  const memRows = aggregateMetricByCohort(kind, 'DB Memory (MB)', (v) => v / 1024);
  renderRankBarsByGroup('memTopChart', memRows, 'max', 'dec1', 'cohort');
  renderRankBarsByGroup('memP95Chart', memRows, 'p95', 'dec1', 'cohort');
  renderBoxPlotByGroup('memBoxChart', memRows, 'dec1', 'cohort');
}

function clampPct(v) {
  const n = Number(v);
  if (!Number.isFinite(n)) return 0;
  return Math.max(0, Math.min(100, n));
}

function nonNegativePct(v) {
  const n = Number(v);
  if (!Number.isFinite(n)) return 0;
  return Math.max(0, n);
}

function sizerAxisMax(values, fallback = 100) {
  const clean = (values || []).map((v) => Number(v)).filter((v) => Number.isFinite(v) && v >= 0);
  const rawMax = clean.length ? Math.max(...clean) : fallback;
  const base = Math.max(fallback, rawMax);
  if (base <= 10) return 10;
  if (base <= 25) return 25;
  if (base <= 50) return 50;
  if (base <= 100) return 100;
  if (base <= 150) return 150;
  if (base <= 200) return 200;
  return Math.ceil(base / 50) * 50;
}

function clampTargetPct(v) {
  const n = Number(v);
  if (!Number.isFinite(n)) return 0;
  return Math.max(0, Math.min(200, n));
}

function parseTargetPctOrDefault(raw, fallback = 100) {
  if (raw === null || raw === undefined) return fallback;
  const s = String(raw).trim();
  if (!s) return fallback;
  const n = Number(s);
  if (!Number.isFinite(n)) return fallback;
  return clampTargetPct(n);
}

function globalTargetField(metric) {
  if (metric === 'vcpu') return 'vcpuSizingTargetPct';
  if (metric === 'memory') return 'memorySizingTargetPct';
  if (metric === 'iops') return 'iopsSizingTargetPct';
  return 'storageSizingTargetPct';
}

function getGlobalTarget(metric) {
  return parseTargetPctOrDefault(state[globalTargetField(metric)], 100);
}

function setGlobalTarget(metric, value) {
  state[globalTargetField(metric)] = parseTargetPctOrDefault(value, 100);
}

function ensureCohortTargetContainer(cohort) {
  if (!state.cohortTargets || typeof state.cohortTargets !== 'object') state.cohortTargets = {};
  if (!state.cohortTargets[cohort] || typeof state.cohortTargets[cohort] !== 'object') state.cohortTargets[cohort] = {};
  return state.cohortTargets[cohort];
}

function ensureCohortTargetsInitialized(kind) {
  const rows = computeByCohort(kind);
  rows.forEach((r) => {
    const item = ensureCohortTargetContainer(r.cohort);
    TARGET_METRICS.forEach((m) => {
      if (!Number.isFinite(Number(item[m]))) item[m] = getGlobalTarget(m);
      item[m] = clampTargetPct(item[m]);
    });
  });
}

function cascadeGlobalTargetsToCohorts(kind, metric) {
  ensureCohortTargetsInitialized(kind);
  const rows = computeByCohort(kind);
  const v = getGlobalTarget(metric);
  rows.forEach((r) => {
    const item = ensureCohortTargetContainer(r.cohort);
    item[metric] = v;
  });
}

function cohortWeight(metric, row) {
  if (metric === 'vcpu') return Number(row.vcpu) || 0;
  if (metric === 'memory') return Number(row.mem) || 0;
  if (metric === 'iops') return Number(row.iops) || 0;
  return Number(row.allocated) || 0;
}

function recomputeGlobalTargetFromCohorts(kind, metric) {
  ensureCohortTargetsInitialized(kind);
  const rows = computeByCohort(kind);
  if (!rows.length) return;
  const sumW = rows.reduce((a, r) => a + cohortWeight(metric, r), 0);
  const fallbackW = rows.length > 0 ? 1 / rows.length : 0;
  let acc = 0;
  rows.forEach((r) => {
    const w = sumW > 0 ? cohortWeight(metric, r) / sumW : fallbackW;
    const item = ensureCohortTargetContainer(r.cohort);
    const v = parseTargetPctOrDefault(item[metric], getGlobalTarget(metric));
    acc += w * v;
  });
  setGlobalTarget(metric, acc);
}

function computeGlobalInstanceCpuStats(kind) {
  const sel = state.selection[kind];
  const vals = [...sel.cohorts]
    .map((cohort) => resolveCpuSeriesForCohort(cohort))
    .filter(Boolean)
    .flatMap((series) => series.y || [])
    .map((v) => Number(v))
    .filter((v) => Number.isFinite(v));
  if (!vals.length) return null;
  const sorted = [...vals].sort((a, b) => a - b);
  return {
    count: sorted.length,
    min: quantileFromSorted(sorted, 0),
    p30: quantileFromSorted(sorted, 0.3),
    p50: quantileFromSorted(sorted, 0.5),
    p70: quantileFromSorted(sorted, 0.7),
    p95: quantileFromSorted(sorted, 0.95),
    p99: quantileFromSorted(sorted, 0.99),
    max: quantileFromSorted(sorted, 1),
  };
}

function computeGlobalInstanceMemoryStats(kind) {
  const rows = selectedInstances(kind);
  const vals = rows
    .map((r) => {
      const total = Number(r.mem_gb) || 0;
      const used = (Number(r.sga_size_gb) || 0) + (Number(r.pga_size_gb) || 0);
      if (total <= 0) return null;
      return (used / total) * 100;
    })
    .filter((v) => Number.isFinite(v));
  if (!vals.length) return null;
  const sorted = [...vals].sort((a, b) => a - b);
  return {
    count: sorted.length,
    min: quantileFromSorted(sorted, 0),
    p30: quantileFromSorted(sorted, 0.3),
    p50: quantileFromSorted(sorted, 0.5),
    p70: quantileFromSorted(sorted, 0.7),
    p95: quantileFromSorted(sorted, 0.95),
    p99: quantileFromSorted(sorted, 0.99),
    max: quantileFromSorted(sorted, 1),
  };
}

function computeGlobalIopsStats(kind) {
  const sel = state.selection[kind];
  const rows = (state.graph?.dbStats || []).filter(
    (r) =>
      sel.cohorts.has(r.cohort) &&
      sel.dbs.has(r.db) &&
      r.metric === 'DB IOPS' &&
      (!r.summary_of || r.summary_of === 'cdb'),
  );
  if (!rows.length) return null;
  const cap = Math.max(
    ...rows.map((r) => Math.max(Number(r.max) || 0, Number(r.p99) || 0, Number(r.p95) || 0)),
    0,
  );
  if (!Number.isFinite(cap) || cap <= 0) return null;
  const values = rows
    .flatMap((r) => [r.min, r.p30, r.p50, r.p70, r.p95, r.p99, r.max])
    .map((v) => ((Number(v) || 0) / cap) * 100)
    .filter((v) => Number.isFinite(v));
  if (!values.length) return null;
  const sorted = [...values].sort((a, b) => a - b);
  return {
    count: sorted.length,
    min: quantileFromSorted(sorted, 0),
    p30: quantileFromSorted(sorted, 0.3),
    p50: quantileFromSorted(sorted, 0.5),
    p70: quantileFromSorted(sorted, 0.7),
    p95: quantileFromSorted(sorted, 0.95),
    p99: quantileFromSorted(sorted, 0.99),
    max: quantileFromSorted(sorted, 1),
  };
}

function computeGlobalStorageStats(kind) {
  const rows = computeByCohort(kind);
  const values = rows
    .map((r) => {
      const allocated = Number(r.allocated) || 0;
      const used = Number(r.used) || 0;
      if (allocated <= 0) return null;
      return (used / allocated) * 100;
    })
    .filter((v) => Number.isFinite(v));
  if (!values.length) return null;
  const sorted = [...values].sort((a, b) => a - b);
  return {
    count: sorted.length,
    min: quantileFromSorted(sorted, 0),
    p30: quantileFromSorted(sorted, 0.3),
    p50: quantileFromSorted(sorted, 0.5),
    p70: quantileFromSorted(sorted, 0.7),
    p95: quantileFromSorted(sorted, 0.95),
    p99: quantileFromSorted(sorted, 0.99),
    max: quantileFromSorted(sorted, 1),
  };
}

function statsFromValues(values) {
  const clean = (values || []).map((v) => Number(v)).filter((v) => Number.isFinite(v));
  if (!clean.length) return null;
  const sorted = [...clean].sort((a, b) => a - b);
  return {
    count: sorted.length,
    min: quantileFromSorted(sorted, 0),
    p30: quantileFromSorted(sorted, 0.3),
    p50: quantileFromSorted(sorted, 0.5),
    p70: quantileFromSorted(sorted, 0.7),
    p95: quantileFromSorted(sorted, 0.95),
    p99: quantileFromSorted(sorted, 0.99),
    max: quantileFromSorted(sorted, 1),
  };
}

function computeCohortCpuStats(kind, cohort) {
  const sel = state.selection[kind];
  if (!sel?.cohorts?.has(cohort)) return null;
  const series = resolveCpuSeriesForCohort(cohort);
  if (!series) return null;
  return statsFromValues(series.y || []);
}

function computeCohortMemoryStats(kind, cohort) {
  const rows = selectedInstances(kind).filter((r) => r.cohort === cohort);
  const vals = rows
    .map((r) => {
      const total = Number(r.mem_gb) || 0;
      const used = (Number(r.sga_size_gb) || 0) + (Number(r.pga_size_gb) || 0);
      if (total <= 0) return null;
      return (used / total) * 100;
    })
    .filter((v) => Number.isFinite(v));
  return statsFromValues(vals);
}

function computeCohortIopsStats(kind, cohort) {
  const sel = state.selection[kind];
  const rows = (state.graph?.dbStats || []).filter(
    (r) =>
      r.cohort === cohort &&
      sel.dbs.has(r.db) &&
      r.metric === 'DB IOPS' &&
      (!r.summary_of || r.summary_of === 'cdb'),
  );
  if (!rows.length) return null;
  const cap = Math.max(...rows.map((r) => Math.max(Number(r.max) || 0, Number(r.p99) || 0, Number(r.p95) || 0)), 0);
  if (!Number.isFinite(cap) || cap <= 0) return null;
  const vals = rows
    .flatMap((r) => [r.min, r.p30, r.p50, r.p70, r.p95, r.p99, r.max])
    .map((v) => ((Number(v) || 0) / cap) * 100)
    .filter((v) => Number.isFinite(v));
  return statsFromValues(vals);
}

function computeCohortStorageStats(kind, cohort) {
  const sel = state.selection[kind];
  const allocByDb = new Map();
  const usedByDb = new Map();
  (state.graph?.dbStats || []).forEach((r) => {
    if (r.cohort !== cohort) return;
    if (!sel.dbs.has(r.db)) return;
    if (r.summary_of && r.summary_of !== 'cdb') return;
    if (r.metric === 'Allocated Storage (GB)') allocByDb.set(r.db, Number(r.p95 || r.max || 0));
    if (r.metric === 'Used Storage (GB)') usedByDb.set(r.db, Number(r.p95 || r.max || 0));
  });
  const vals = [...allocByDb.keys()]
    .map((db) => {
      const a = allocByDb.get(db) || 0;
      const u = usedByDb.get(db) || 0;
      if (a <= 0) return null;
      return (u / a) * 100;
    })
    .filter((v) => Number.isFinite(v));
  return statsFromValues(vals);
}

function computeGlobalCohortMetricRows(kind, metric) {
  return computeByCohort(kind)
    .map((row) => {
      const stats =
        metric === 'vcpu'
          ? computeCohortCpuStats(kind, row.cohort)
          : metric === 'memory'
            ? computeCohortMemoryStats(kind, row.cohort)
            : metric === 'iops'
              ? computeCohortIopsStats(kind, row.cohort)
              : computeCohortStorageStats(kind, row.cohort);
      if (!stats) return null;
      return { cohort: row.cohort, ...stats };
    })
    .filter(Boolean)
    .sort((a, b) => a.cohort.localeCompare(b.cohort));
}

function cohortProvisionLines(kind, cohort, metric) {
  const BASE = { cpu: 256, mem: 512, iops: 8000, storage_tb: 80 };
  const XS = { ecpu: 200, storage_tb: 100 };
  const EXA = { ecpu: 1520, mem: 973, iops: 89000000, storage_tb: 5000 };
  const rows = selectedInstances(kind).filter((r) => r.cohort === cohort);
  const c = computeByCohort(kind).find((x) => x.cohort === cohort);
  if (!rows.length || !c) return [];

  const maxInstCpu = Math.max(...rows.map((r) => Number(r.init_cpu_count || r.logical_cpu_count || 0)), 0);
  const maxInstMem = Math.max(...rows.map((r) => Number(r.mem_gb || 0)), 0);
  const maxInstCount = Math.max(rows.length, 1);

  const iopsByDb = aggregateMetricByDb(kind, 'DB IOPS', (v) => v, cohort);
  const maxInstIops = Math.max(...iopsByDb.map((r) => Number(r.p95 || r.max || 0)), 0);
  const cohortIops = Math.max(...iopsByDb.map((r) => Number(r.p95 || r.max || 0)), 0);

  const storageByDb = aggregateMetricByDb(kind, 'Allocated Storage (GB)', (v) => v, cohort);
  const maxInstStorageTb = Math.max(...storageByDb.map((r) => Number(r.p95 || r.max || 0) / 1024), 0);
  const cohortStorageTb = Number(c.allocated || 0) / 1024;

  const cohortCpuEcpu = Number(c.vcpu || 0) * 2;

  const utilBase =
    metric === 'vcpu'
      ? (maxInstCpu / BASE.cpu) * 100
      : metric === 'memory'
        ? (maxInstMem / BASE.mem) * 100
        : metric === 'iops'
          ? (maxInstIops / BASE.iops) * 100
          : (maxInstStorageTb / BASE.storage_tb) * 100;

  const utilXs =
    metric === 'vcpu'
      ? (cohortCpuEcpu / XS.ecpu) * 100
      : metric === 'memory'
        ? (Number(c.mem || 0) / EXA.mem) * 100
        : metric === 'iops'
          ? (cohortIops / EXA.iops) * 100
          : (cohortStorageTb / XS.storage_tb) * 100;

  const utilExa =
    metric === 'vcpu'
      ? (cohortCpuEcpu / EXA.ecpu) * 100
      : metric === 'memory'
        ? (Number(c.mem || 0) / EXA.mem) * 100
        : metric === 'iops'
          ? (cohortIops / EXA.iops) * 100
          : (cohortStorageTb / EXA.storage_tb) * 100;

  const perInstScale = 1 / maxInstCount; // keeps Base visually comparable by cohort scope
  return [
    { name: 'Base Provisioned', value: utilBase * perInstScale, color: '#b45309' },
    { name: 'Exascale Provisioned', value: utilXs, color: '#0f766e' },
    { name: 'Exadata Provisioned', value: utilExa, color: '#1d4ed8' },
  ];
}

function renderVerticalSizerChart(chartNode, stats, selectedPct, titleText, colorMain, provisionLines = []) {
  const w = 340;
  const h = 420;
  const left = 44;
  const right = 126;
  const top = 18;
  const bottom = 28;
  const plotW = w - left - right;
  const plotH = h - top - bottom;
  const xMid = left + plotW * 0.46;
  const boxHalf = Math.max(16, plotW * 0.24);
  const labelX = xMid + boxHalf + 22;

  const p = {
    min: nonNegativePct(stats.min),
    p30: nonNegativePct(stats.p30),
    p50: nonNegativePct(stats.p50),
    p70: nonNegativePct(stats.p70),
    p95: nonNegativePct(stats.p95),
    p99: nonNegativePct(stats.p99),
    max: nonNegativePct(stats.max),
  };
  const axisMax = sizerAxisMax([p.max, p.p99, p.p95, selectedPct, ...(provisionLines || []).map((ln) => ln?.value)], 100);
  const y = (pct) => top + ((axisMax - Math.min(nonNegativePct(pct), axisMax)) / axisMax) * plotH;

  const yMin = y(p.min);
  const yQ1 = y(p.p30);
  const yMed = y(p.p50);
  const yQ3 = y(p.p70);
  const yMax = y(p.max);
  const yP95 = y(p.p95);
  const yP99 = y(p.p99);
  const ySelected = y(selectedPct);
  const hasUpperOutlier = p.max > p.p99 + 0.0001;
  const yUpperWhisker = hasUpperOutlier ? yP99 : yMax;
  const outlierDiamond =
    hasUpperOutlier
      ? `<polygon points="${xMid},${yMax - 5} ${xMid + 5},${yMax} ${xMid},${yMax + 5} ${xMid - 5},${yMax}" fill="#7c3aed" stroke="#ffffff" stroke-width="0.8"></polygon>`
      : '';

  const ticks = [];
  const tickStep = axisMax <= 100 ? 10 : axisMax <= 200 ? 20 : 50;
  for (let v = 0; v <= axisMax; v += tickStep) {
    const yy = y(v);
    ticks.push(
      `<line x1="${left - 5}" y1="${yy}" x2="${left + plotW}" y2="${yy}" stroke="${v % 20 === 0 ? '#e2e8f0' : '#edf2f7'}" />` +
        `<text x="${left - 10}" y="${yy + 3}" text-anchor="end" font-size="9" fill="#64748b">${v}%</text>`,
    );
  }

  const marker = (name, val, yy, color = '#334155') =>
    `<line x1="${xMid + boxHalf}" y1="${yy}" x2="${labelX - 6}" y2="${yy}" stroke="${color}" stroke-width="1" />` +
    `<circle cx="${labelX - 8}" cy="${yy}" r="2.1" fill="${color}" />` +
    `<text x="${labelX}" y="${yy + 3}" font-size="9" fill="#334155">${name} ${fmt(val, 'dec1')}%</text>`;

  const provisionSvg = (provisionLines || [])
    .map((ln, i) => {
      const yy = y(ln.value);
      const lx = left + 6 + (i % 2) * 4;
      return (
        `<line x1="${left}" y1="${yy}" x2="${left + plotW}" y2="${yy}" stroke="${ln.color || '#334155'}" stroke-width="1.4" stroke-dasharray="3 3"></line>` +
        `<text x="${lx}" y="${Math.max(10, yy - 2)}" font-size="8.5" fill="${ln.color || '#334155'}">${esc(shortLabel(`${ln.name} ${fmt(ln.value, 'dec1')}%`, 28))}</text>`
      );
    })
    .join('');

  chartNode.innerHTML = `
    <svg viewBox="0 0 ${w} ${h}" role="img" aria-label="${escAttr(titleText)}">
      ${ticks.join('')}
      <text x="${left + plotW / 2}" y="${h - 6}" text-anchor="middle" font-size="10" fill="#64748b">${esc(titleText)}</text>
      <text x="${left + plotW - 2}" y="${top + 10}" text-anchor="end" font-size="9" fill="#64748b">Max ${fmt(axisMax, 'dec1')}%</text>
      <line x1="${xMid}" y1="${yUpperWhisker}" x2="${xMid}" y2="${yMin}" stroke="#64748b" stroke-width="1.5"></line>
      <rect x="${xMid - boxHalf}" y="${Math.min(yQ3, yQ1)}" width="${boxHalf * 2}" height="${Math.max(2, Math.abs(yQ1 - yQ3))}" fill="${colorMain}22" stroke="${colorMain}" stroke-width="1.2"></rect>
      <line x1="${xMid - boxHalf}" y1="${yMed}" x2="${xMid + boxHalf}" y2="${yMed}" stroke="${colorMain}" stroke-width="2"></line>
      <line x1="${xMid - 10}" y1="${yMin}" x2="${xMid + 10}" y2="${yMin}" stroke="#64748b" stroke-width="1.2"></line>
      <line x1="${xMid - 10}" y1="${yUpperWhisker}" x2="${xMid + 10}" y2="${yUpperWhisker}" stroke="#64748b" stroke-width="1.2"></line>
      ${outlierDiamond}
      ${provisionSvg}
      <line class="sizer-target-line" x1="${left}" y1="${ySelected}" x2="${left + plotW}" y2="${ySelected}" stroke="#b11f1f" stroke-width="2" stroke-dasharray="5 3"></line>
      <text class="sizer-target-label" x="${left + plotW + 6}" y="${Math.max(12, ySelected - 4)}" font-size="9" fill="#b11f1f">${fmt(clampTargetPct(
        selectedPct,
      ), 'dec1')}%</text>
      ${marker('P30', p.p30, yQ1)}
      ${marker('P50', p.p50, yMed)}
      ${marker('P70', p.p70, yQ3)}
      ${marker('P95', p.p95, yP95, '#b45309')}
      ${marker('P99', p.p99, yP99, '#7c3aed')}
      ${hasUpperOutlier ? marker('MAX(outlier)', p.max, yMax, '#7c3aed') : ''}
    </svg>
  `;

  return { y, p, axisMax, targetLine: chartNode.querySelector('.sizer-target-line'), targetLabel: chartNode.querySelector('.sizer-target-label') };
}

function renderMultiVerticalSizerChart(chartNode, rows, globalTargetPct, titleText, cohortTargetMap = {}, colorMain = '#0f766e') {
  if (!chartNode) return null;
  if (!rows.length) {
    chartNode.innerHTML = '<p class="muted">No cohort utilization data in current scope.</p>';
    return null;
  }

  const count = rows.length;
  const w = Math.max(360, 88 + count * 64);
  const h = 420;
  const left = 40;
  const right = 24;
  const top = 18;
  const bottom = 54;
  const plotW = w - left - right;
  const plotH = h - top - bottom;
  const step = plotW / Math.max(count, 1);
  const boxHalf = Math.max(10, Math.min(18, step * 0.22));
  const axisMax = sizerAxisMax(
    [
      ...rows.flatMap((row) => [row.min, row.p30, row.p50, row.p70, row.p95, row.p99, row.max]),
      globalTargetPct,
      ...rows.map((row) => cohortTargetMap?.[row.cohort]),
    ],
    100,
  );
  const y = (pct) => top + ((axisMax - Math.min(nonNegativePct(pct), axisMax)) / axisMax) * plotH;
  const globalY = y(globalTargetPct);

  const ticks = [];
  const tickStep = axisMax <= 100 ? 10 : axisMax <= 200 ? 20 : 50;
  for (let v = 0; v <= axisMax; v += tickStep) {
    const yy = y(v);
    ticks.push(
      `<line x1="${left - 4}" y1="${yy}" x2="${w - right}" y2="${yy}" stroke="${v % 20 === 0 ? '#e2e8f0' : '#edf2f7'}" />` +
        `<text x="${left - 8}" y="${yy + 3}" text-anchor="end" font-size="9" fill="#64748b">${v}%</text>`,
    );
  }

  const cohortTargetLines = [];
  const cohortLabelEls = [];
  const boxEls = [];
  rows.forEach((row, idx) => {
    const xMid = left + step * idx + step / 2;
    const p = {
      min: nonNegativePct(row.min),
      p30: nonNegativePct(row.p30),
      p50: nonNegativePct(row.p50),
      p70: nonNegativePct(row.p70),
      p95: nonNegativePct(row.p95),
      p99: nonNegativePct(row.p99),
      max: nonNegativePct(row.max),
    };
    const yMin = y(p.min);
    const yQ1 = y(p.p30);
    const yMed = y(p.p50);
    const yQ3 = y(p.p70);
    const yP99 = y(p.p99);
    const yMax = y(p.max);
    const hasUpperOutlier = p.max > p.p99 + 0.0001;
    const yUpperWhisker = hasUpperOutlier ? yP99 : yMax;
    const itemColor = colorForItem(row.cohort);
    const cohortTarget = parseTargetPctOrDefault(cohortTargetMap?.[row.cohort], globalTargetPct);
    const yCohortTarget = y(cohortTarget);

    boxEls.push(`
      <line x1="${xMid}" y1="${yUpperWhisker}" x2="${xMid}" y2="${yMin}" stroke="#64748b" stroke-width="1.2"></line>
      <line x1="${xMid - 8}" y1="${yMin}" x2="${xMid + 8}" y2="${yMin}" stroke="#64748b" stroke-width="1.1"></line>
      <line x1="${xMid - 8}" y1="${yUpperWhisker}" x2="${xMid + 8}" y2="${yUpperWhisker}" stroke="#64748b" stroke-width="1.1"></line>
      <rect x="${xMid - boxHalf}" y="${Math.min(yQ3, yQ1)}" width="${boxHalf * 2}" height="${Math.max(2, Math.abs(yQ1 - yQ3))}" fill="${itemColor}22" stroke="${itemColor}" stroke-width="1.2"></rect>
      <line x1="${xMid - boxHalf}" y1="${yMed}" x2="${xMid + boxHalf}" y2="${yMed}" stroke="${itemColor}" stroke-width="2"></line>
      ${
        hasUpperOutlier
          ? `<polygon points="${xMid},${yMax - 5} ${xMid + 5},${yMax} ${xMid},${yMax + 5} ${xMid - 5},${yMax}" fill="#7c3aed" stroke="#ffffff" stroke-width="0.8"></polygon>`
          : ''
      }
    `);
    cohortTargetLines.push(
      `<line class="cohort-target-line" data-cohort-key="${escAttr(row.cohort)}" x1="${xMid - boxHalf - 2}" y1="${yCohortTarget}" x2="${xMid + boxHalf + 2}" y2="${yCohortTarget}" stroke="${itemColor}" stroke-width="2" stroke-dasharray="3 2"></line>`,
    );
    cohortLabelEls.push(
      `<text x="${xMid}" y="${h - 22}" text-anchor="middle" font-size="9" fill="#475569">${esc(shortLabel(row.cohort, 12))}</text>`,
    );
  });

  chartNode.innerHTML = `
    <svg viewBox="0 0 ${w} ${h}" role="img" aria-label="${escAttr(titleText)}">
      ${ticks.join('')}
      <line class="sizer-target-line" x1="${left}" y1="${globalY}" x2="${w - right}" y2="${globalY}" stroke="#b11f1f" stroke-width="2" stroke-dasharray="5 3"></line>
      <text class="sizer-target-label" x="${w - right - 2}" y="${Math.max(12, globalY - 4)}" text-anchor="end" font-size="9" fill="#b11f1f">${fmt(
        clampTargetPct(globalTargetPct),
        'dec1',
      )}%</text>
      ${boxEls.join('')}
      ${cohortTargetLines.join('')}
      ${cohortLabelEls.join('')}
      <text x="${left + plotW / 2}" y="${h - 6}" text-anchor="middle" font-size="10" fill="#64748b">${esc(titleText)}</text>
      <text x="${w - right - 2}" y="${top + 10}" text-anchor="end" font-size="9" fill="#64748b">Max ${fmt(axisMax, 'dec1')}%</text>
      <text x="${left + 2}" y="${top + 10}" font-size="9" fill="${colorMain}">One boxplot per cohort</text>
    </svg>
  `;

  const cohortTargets = new Map();
  chartNode.querySelectorAll('.cohort-target-line').forEach((el) => {
    cohortTargets.set(decodeKey(el.getAttribute('data-cohort-key')), el);
  });
  return {
    y,
    axisMax,
    targetLine: chartNode.querySelector('.sizer-target-line'),
    targetLabel: chartNode.querySelector('.sizer-target-label'),
    cohortTargets,
  };
}

function renderGlobalVcpuSizer(kind) {
  const chartNode = $('globalVcpuSizerChart');
  const memChartNode = $('globalMemSizerChart');
  const iopsChartNode = $('globalIopsSizerChart');
  const storageChartNode = $('globalStorageSizerChart');
  const controlsNode = $('globalSizerControls');
  if (!chartNode || !memChartNode || !iopsChartNode || !storageChartNode || !controlsNode) return;

  ensureCohortTargetsInitialized(kind);
  const stats = computeGlobalInstanceCpuStats(kind);
  const memStats = computeGlobalInstanceMemoryStats(kind);
  const iopsStats = computeGlobalIopsStats(kind);
  const storageStats = computeGlobalStorageStats(kind);
  const vcpuRows = computeGlobalCohortMetricRows(kind, 'vcpu');
  const memRows = computeGlobalCohortMetricRows(kind, 'memory');
  const iopsRows = computeGlobalCohortMetricRows(kind, 'iops');
  const storageRows = computeGlobalCohortMetricRows(kind, 'storage');
  const globalTotals = computeGlobal(kind);
  const cohortTargetsByMetric = {
    vcpu: Object.fromEntries(computeByCohort(kind).map((row) => [row.cohort, parseTargetPctOrDefault(ensureCohortTargetContainer(row.cohort).vcpu, getGlobalTarget('vcpu'))])),
    memory: Object.fromEntries(computeByCohort(kind).map((row) => [row.cohort, parseTargetPctOrDefault(ensureCohortTargetContainer(row.cohort).memory, getGlobalTarget('memory'))])),
    iops: Object.fromEntries(computeByCohort(kind).map((row) => [row.cohort, parseTargetPctOrDefault(ensureCohortTargetContainer(row.cohort).iops, getGlobalTarget('iops'))])),
    storage: Object.fromEntries(computeByCohort(kind).map((row) => [row.cohort, parseTargetPctOrDefault(ensureCohortTargetContainer(row.cohort).storage, getGlobalTarget('storage'))])),
  };
  const metricCfg = {
    vcpu: { total: Number(globalTotals?.vcpu_total) || 0, unit: 'vCPU', fmt: 'dec1' },
    memory: { total: Number(globalTotals?.memory_gb_total) || 0, unit: 'GB', fmt: 'dec1' },
    iops: { total: Number(globalTotals?.db_iops_total) || 0, unit: 'IOPS', fmt: 'dec1' },
    storage: { total: Number(globalTotals?.allocated_storage_gb) || 0, unit: 'GB', fmt: 'dec1' },
  };
  if (!stats && !memStats && !iopsStats && !storageStats) {
    chartNode.innerHTML = '<p class="muted">No CPU series available in the current scope.</p>';
    memChartNode.innerHTML = '<p class="muted">No memory utilization series available in the current scope.</p>';
    iopsChartNode.innerHTML = '<p class="muted">No IOPS utilization series available in the current scope.</p>';
    storageChartNode.innerHTML = '<p class="muted">No storage utilization series available in the current scope.</p>';
    controlsNode.innerHTML = '';
    return;
  }

  if (stats && !Number.isFinite(state.vcpuSizingTargetPct)) {
    state.vcpuSizingTargetPct = 100;
  }
  if (memStats && !Number.isFinite(state.memorySizingTargetPct)) {
    state.memorySizingTargetPct = 100;
  }
  if (iopsStats && !Number.isFinite(state.iopsSizingTargetPct)) {
    state.iopsSizingTargetPct = 100;
  }
  if (storageStats && !Number.isFinite(state.storageSizingTargetPct)) {
    state.storageSizingTargetPct = 100;
  }

  const vcpuModel = vcpuRows.length
    ? renderMultiVerticalSizerChart(chartNode, vcpuRows, state.vcpuSizingTargetPct, 'VCPU Utilization', cohortTargetsByMetric.vcpu, '#0f766e')
    : null;
  const memModel = memRows.length
    ? renderMultiVerticalSizerChart(memChartNode, memRows, state.memorySizingTargetPct, 'Memory Utilization', cohortTargetsByMetric.memory, '#0e7490')
    : null;
  if (!iopsStats) iopsChartNode.innerHTML = '<p class="muted">No IOPS utilization data in current scope.</p>';
  const iopsModel = iopsRows.length
    ? renderMultiVerticalSizerChart(iopsChartNode, iopsRows, state.iopsSizingTargetPct, 'IOPS Utilization', cohortTargetsByMetric.iops, '#6d28d9')
    : null;
  if (!storageStats) storageChartNode.innerHTML = '<p class="muted">No storage utilization data in current scope.</p>';
  const storageModel = storageRows.length
    ? renderMultiVerticalSizerChart(storageChartNode, storageRows, state.storageSizingTargetPct, 'Storage Utilization', cohortTargetsByMetric.storage, '#b45309')
    : null;

  controlsNode.innerHTML = `
    <div class="global-sizer-controls-grid">
      <div class="global-sizer-control-card">
        <label for="globalVcpuTargetInput" style="margin:0 0 4px;color:#475569;font-size:12px;">VCPU target input (0-200%)</label>
        <input id="globalVcpuTargetInput" type="number" min="0" max="200" step="0.1" value="${fmt(clampTargetPct(state.vcpuSizingTargetPct), 'dec1')}" style="width:100%;padding:6px 8px;border:1px solid #cfd6df;border-radius:8px;" />
        <div id="globalVcpuInputError" style="min-height:16px;color:#b91c1c;font-size:11px;margin-top:4px;"></div>
        <div id="globalVcpuSummary" class="line" style="font-size:11px;color:#475569;"></div>
        <div id="globalVcpuQuant" class="line" style="font-size:11px;color:#334155;"></div>
      </div>
      <div class="global-sizer-control-card">
        <label for="globalMemTargetInput" style="margin:0 0 4px;color:#475569;font-size:12px;">Memory target input (0-200%)</label>
        <input id="globalMemTargetInput" type="number" min="0" max="200" step="0.1" value="${fmt(clampTargetPct(state.memorySizingTargetPct), 'dec1')}" style="width:100%;padding:6px 8px;border:1px solid #cfd6df;border-radius:8px;" />
        <div id="globalMemInputError" style="min-height:16px;color:#b91c1c;font-size:11px;margin-top:4px;"></div>
        <div id="globalMemSummary" class="line" style="font-size:11px;color:#475569;"></div>
        <div id="globalMemQuant" class="line" style="font-size:11px;color:#334155;"></div>
      </div>
      <div class="global-sizer-control-card">
        <label for="globalIopsTargetInput" style="margin:0 0 4px;color:#475569;font-size:12px;">IOPS target input (0-200%)</label>
        <input id="globalIopsTargetInput" type="number" min="0" max="200" step="0.1" value="${fmt(clampTargetPct(state.iopsSizingTargetPct), 'dec1')}" style="width:100%;padding:6px 8px;border:1px solid #cfd6df;border-radius:8px;" />
        <div id="globalIopsInputError" style="min-height:16px;color:#b91c1c;font-size:11px;margin-top:4px;"></div>
        <div id="globalIopsSummary" class="line" style="font-size:11px;color:#475569;"></div>
        <div id="globalIopsQuant" class="line" style="font-size:11px;color:#334155;"></div>
      </div>
      <div class="global-sizer-control-card">
        <label for="globalStorageTargetInput" style="margin:0 0 4px;color:#475569;font-size:12px;">Storage target input (0-200%)</label>
        <input id="globalStorageTargetInput" type="number" min="0" max="200" step="0.1" value="${fmt(clampTargetPct(state.storageSizingTargetPct), 'dec1')}" style="width:100%;padding:6px 8px;border:1px solid #cfd6df;border-radius:8px;" />
        <div id="globalStorageInputError" style="min-height:16px;color:#b91c1c;font-size:11px;margin-top:4px;"></div>
        <div id="globalStorageSummary" class="line" style="font-size:11px;color:#475569;"></div>
        <div id="globalStorageQuant" class="line" style="font-size:11px;color:#334155;"></div>
      </div>
    </div>
  `;
  const vcpuErrNode = controlsNode.querySelector('#globalVcpuInputError');
  const memErrNode = controlsNode.querySelector('#globalMemInputError');
  const iopsErrNode = controlsNode.querySelector('#globalIopsInputError');
  const storageErrNode = controlsNode.querySelector('#globalStorageInputError');
  const vcpuInput = controlsNode.querySelector('#globalVcpuTargetInput');
  const memInput = controlsNode.querySelector('#globalMemTargetInput');
  const iopsInput = controlsNode.querySelector('#globalIopsTargetInput');
  const storageInput = controlsNode.querySelector('#globalStorageTargetInput');
  const vcpuSummary = controlsNode.querySelector('#globalVcpuSummary');
  const memSummary = controlsNode.querySelector('#globalMemSummary');
  const iopsSummary = controlsNode.querySelector('#globalIopsSummary');
  const storageSummary = controlsNode.querySelector('#globalStorageSummary');
  const vcpuQuant = controlsNode.querySelector('#globalVcpuQuant');
  const memQuant = controlsNode.querySelector('#globalMemQuant');
  const iopsQuant = controlsNode.querySelector('#globalIopsQuant');
  const storageQuant = controlsNode.querySelector('#globalStorageQuant');

  const validateInput = (raw) => {
    if (raw == null || String(raw).trim() === '') return { ok: false, msg: 'Enter a value between 0 and 200.' };
    const n = Number(raw);
    if (!Number.isFinite(n)) return { ok: false, msg: 'Invalid number.' };
    if (n < 0) return { ok: false, msg: 'Minimum is 0.' };
    if (n > 200) return { ok: false, msg: 'Max is 200.' };
    return { ok: true, value: n };
  };

  const updateOneUi = (which) => {
    const modelByType = { vcpu: vcpuModel, memory: memModel, iops: iopsModel, storage: storageModel };
    const statsByType = { vcpu: stats, memory: memStats, iops: iopsStats, storage: storageStats };
    const summaryByType = { vcpu: vcpuSummary, memory: memSummary, iops: iopsSummary, storage: storageSummary };
    const quantByType = { vcpu: vcpuQuant, memory: memQuant, iops: iopsQuant, storage: storageQuant };
    const selectedByType = {
      vcpu: state.vcpuSizingTargetPct,
      memory: state.memorySizingTargetPct,
      iops: state.iopsSizingTargetPct,
      storage: state.storageSizingTargetPct,
    };
    const selected = clampTargetPct(selectedByType[which]);
    const model = modelByType[which];
    const baseline = statsByType[which];
    const summaryNode = summaryByType[which];
    const quantNode = quantByType[which];
    const cfg = metricCfg[which];
    if (!model || !baseline) return;
    const p95v = nonNegativePct(baseline.p95);
    const relPct = p95v > 0 ? ((selected - p95v) / p95v) * 100 : 0;
    const yy = model.y(selected);
    if (model.targetLine) {
      model.targetLine.setAttribute('y1', String(yy));
      model.targetLine.setAttribute('y2', String(yy));
    }
    if (model.targetLabel) {
      model.targetLabel.setAttribute('y', String(Math.max(12, yy - 4)));
      model.targetLabel.textContent = `${fmt(selected, 'dec1')}%${selected > 100 ? ' (off-scale)' : ''}`;
    }
    if (model.cohortTargets instanceof Map) {
      model.cohortTargets.forEach((el, cohort) => {
        const cohortSelected = parseTargetPctOrDefault(ensureCohortTargetContainer(cohort)[which], selected);
        const cohortY = model.y(cohortSelected);
        el.setAttribute('y1', String(cohortY));
        el.setAttribute('y2', String(cohortY));
      });
    }
    if (summaryNode) {
      const cohortVals = computeByCohort(kind).map((row) => parseTargetPctOrDefault(ensureCohortTargetContainer(row.cohort)[which], selected));
      const cohortMin = cohortVals.length ? Math.min(...cohortVals) : selected;
      const cohortMax = cohortVals.length ? Math.max(...cohortVals) : selected;
      summaryNode.textContent = `P95 ${fmt(p95v, 'dec1')}% | Global target ${fmt(selected, 'dec1')}% | Cohort targets ${fmt(cohortMin, 'dec1')}%-${fmt(cohortMax, 'dec1')}% | Δ ${fmt(
        selected - p95v,
        'dec1',
      )} pts (${fmt(relPct, 'dec1')}%)`;
    }
    if (quantNode && cfg) {
      const maxQty = cfg.total;
      const targetQty = (maxQty * selected) / 100;
      quantNode.textContent = `100% = ${fmt(maxQty, cfg.fmt)} ${cfg.unit} | ${fmt(selected, 'dec1')}% = ${fmt(targetQty, cfg.fmt)} ${cfg.unit}`;
    }
  };

  if (vcpuInput) {
    vcpuInput.addEventListener('input', () => {
      const v = validateInput(vcpuInput.value);
      if (!v.ok) {
        if (vcpuErrNode) vcpuErrNode.textContent = v.msg;
        return;
      }
      if (vcpuErrNode) vcpuErrNode.textContent = '';
      setGlobalTarget('vcpu', v.value);
      cascadeGlobalTargetsToCohorts(kind, 'vcpu');
      updateOneUi('vcpu');
    });
    vcpuInput.addEventListener('blur', () => {
      const v = validateInput(vcpuInput.value);
      if (!v.ok) {
        vcpuInput.value = String(fmt(clampTargetPct(state.vcpuSizingTargetPct), 'dec1'));
        if (vcpuErrNode) vcpuErrNode.textContent = '';
        return;
      }
      setGlobalTarget('vcpu', v.value);
      cascadeGlobalTargetsToCohorts(kind, 'vcpu');
      vcpuInput.value = String(fmt(getGlobalTarget('vcpu'), 'dec1'));
      if (vcpuErrNode) vcpuErrNode.textContent = '';
      updateOneUi('vcpu');
    });
  }

  if (memInput) {
    memInput.addEventListener('input', () => {
      const v = validateInput(memInput.value);
      if (!v.ok) {
        if (memErrNode) memErrNode.textContent = v.msg;
        return;
      }
      if (memErrNode) memErrNode.textContent = '';
      setGlobalTarget('memory', v.value);
      cascadeGlobalTargetsToCohorts(kind, 'memory');
      updateOneUi('memory');
    });
    memInput.addEventListener('blur', () => {
      const v = validateInput(memInput.value);
      if (!v.ok) {
        memInput.value = String(fmt(clampTargetPct(state.memorySizingTargetPct), 'dec1'));
        if (memErrNode) memErrNode.textContent = '';
        return;
      }
      setGlobalTarget('memory', v.value);
      cascadeGlobalTargetsToCohorts(kind, 'memory');
      memInput.value = String(fmt(getGlobalTarget('memory'), 'dec1'));
      if (memErrNode) memErrNode.textContent = '';
      updateOneUi('memory');
    });
  }

  if (iopsInput) {
    iopsInput.addEventListener('input', () => {
      const v = validateInput(iopsInput.value);
      if (!v.ok) {
        if (iopsErrNode) iopsErrNode.textContent = v.msg;
        return;
      }
      if (iopsErrNode) iopsErrNode.textContent = '';
      setGlobalTarget('iops', v.value);
      cascadeGlobalTargetsToCohorts(kind, 'iops');
      updateOneUi('iops');
    });
    iopsInput.addEventListener('blur', () => {
      const v = validateInput(iopsInput.value);
      if (!v.ok) {
        iopsInput.value = String(fmt(clampTargetPct(state.iopsSizingTargetPct), 'dec1'));
        if (iopsErrNode) iopsErrNode.textContent = '';
        return;
      }
      setGlobalTarget('iops', v.value);
      cascadeGlobalTargetsToCohorts(kind, 'iops');
      iopsInput.value = String(fmt(getGlobalTarget('iops'), 'dec1'));
      if (iopsErrNode) iopsErrNode.textContent = '';
      updateOneUi('iops');
    });
  }

  if (storageInput) {
    storageInput.addEventListener('input', () => {
      const v = validateInput(storageInput.value);
      if (!v.ok) {
        if (storageErrNode) storageErrNode.textContent = v.msg;
        return;
      }
      if (storageErrNode) storageErrNode.textContent = '';
      setGlobalTarget('storage', v.value);
      cascadeGlobalTargetsToCohorts(kind, 'storage');
      updateOneUi('storage');
    });
    storageInput.addEventListener('blur', () => {
      const v = validateInput(storageInput.value);
      if (!v.ok) {
        storageInput.value = String(fmt(clampTargetPct(state.storageSizingTargetPct), 'dec1'));
        if (storageErrNode) storageErrNode.textContent = '';
        return;
      }
      setGlobalTarget('storage', v.value);
      cascadeGlobalTargetsToCohorts(kind, 'storage');
      storageInput.value = String(fmt(getGlobalTarget('storage'), 'dec1'));
      if (storageErrNode) storageErrNode.textContent = '';
      updateOneUi('storage');
    });
  }

  updateOneUi('vcpu');
  updateOneUi('memory');
  updateOneUi('iops');
  updateOneUi('storage');
}

function renderMultiLineChart(containerId, seriesRaw, fixedMax = null) {
  const container = $(containerId);
  if (!container) return;
  if (!seriesRaw.length) {
    container.innerHTML = '<p class="muted">No CPU profile data in current scope.</p>';
    return;
  }

  const chartType = getChartType(containerId, 'line');
  if (chartType !== 'column' && chartType !== 'bar' && chartType !== 'pie' && chartType !== 'text') {
    // no-op, fallback handled below by normalize
  }

  if (chartType !== 'line') {
    const items = seriesRaw.map((s) => {
      const avg = s.y?.length ? s.y.reduce((a, b) => a + b, 0) / s.y.length : 0;
      return { label: s.name, value: avg };
    });
    renderCategoricalChart(containerId, items, { fmtKind: 'dec1', defaultType: 'column' });
    return;
  }

  const series = [...seriesRaw]
    .sort((a, b) => Math.max(...b.y) - Math.max(...a.y))
    .slice(0, 10)
    .map((s, i) => ({
      ...s,
      color: colorForItem(s.name),
      dash: ['0', '8 4', '3 3', '10 4 2 4'][i % 4],
      marker: ['circle', 'square', 'diamond', 'triangle'][i % 4],
    }));

  const maxY = fixedMax != null ? Math.max(fixedMax, 1) : Math.max(...series.flatMap((s) => s.y), 1);
  const w = 760;
  const h = 280;
  const left = 50;
  const right = 16;
  const top = 10;
  const bottom = 32;
  const plotW = w - left - right;
  const plotH = h - top - bottom;
  const xLen = Math.max(...series.map((s) => s.x.length), 2);

  const x = (i) => left + (i / (xLen - 1)) * plotW;
  const y = (v) => top + (1 - v / maxY) * plotH;

  const baseX = series[0].x;
  const tickIdx = [0, Math.floor((xLen - 1) * 0.25), Math.floor((xLen - 1) * 0.5), Math.floor((xLen - 1) * 0.75), xLen - 1];
  const uniqTickIdx = [...new Set(tickIdx)];
  const xTicks = uniqTickIdx
    .map((i) => {
      const labelRaw = baseX[i] || '';
      const label = String(labelRaw).replace('T', ' ').slice(0, 16);
      return `<text x="${x(i)}" y="${h - 10}" text-anchor="middle" font-size="10" fill="#596579">${esc(label)}</text>`;
    })
    .join('');
  const grid = [0, 0.25, 0.5, 0.75, 1]
    .map((f) => {
      const val = maxY * f;
      const yy = y(val);
      return `<line x1="${left}" y1="${yy}" x2="${w - right}" y2="${yy}" stroke="#e5e7eb" />` +
        `<text x="4" y="${yy + 3}" font-size="10" fill="#6b7280">${fmt(val, 'dec1')}</text>`;
    })
    .join('');

  const paths = series
    .map((s) => {
      const d = s.y.map((v, i) => `${i === 0 ? 'M' : 'L'} ${x(i)} ${y(v)}`).join(' ');
      const areaD = `${d} L ${x(s.y.length - 1)} ${top + plotH} L ${x(0)} ${top + plotH} Z`;
      const points = s.y
        .map((v, i) => {
          const tip = esc(`${s.name} | ${(s.x[i] || '').replace('T', ' ')} | ${fmt(v, 'dec1')}`);
          const px = x(i);
          const py = y(v);
          if (s.marker === 'square') {
            return `<rect x="${px - 2.4}" y="${py - 2.4}" width="4.8" height="4.8" fill="${s.color}" stroke="#ffffff" stroke-width="0.8" data-tip="${tip}"></rect>`;
          }
          if (s.marker === 'diamond') {
            return `<polygon points="${px},${py - 3} ${px + 3},${py} ${px},${py + 3} ${px - 3},${py}" fill="${s.color}" stroke="#ffffff" stroke-width="0.8" data-tip="${tip}"></polygon>`;
          }
          if (s.marker === 'triangle') {
            return `<polygon points="${px},${py - 3.4} ${px + 3.2},${py + 2.8} ${px - 3.2},${py + 2.8}" fill="${s.color}" stroke="#ffffff" stroke-width="0.8" data-tip="${tip}"></polygon>`;
          }
          return `<circle cx="${px}" cy="${py}" r="2.4" fill="${s.color}" stroke="#ffffff" stroke-width="0.8" data-tip="${tip}"></circle>`;
        })
        .join('');
      const dashAttr = s.dash === '0' ? '' : `stroke-dasharray="${s.dash}"`;
      return `<path d="${areaD}" fill="${s.color}22"></path><path d="${d}" fill="none" stroke="${s.color}" stroke-width="2.4" ${dashAttr}><title>${esc(
        `${s.name} time series`,
      )}</title></path>${points}`;
    })
    .join('');

  const legend = series
    .map((s) => `<span style="--legend-color:${s.color}">${esc(shortLabel(s.name, 26))}</span>`)
    .join('');

  container.innerHTML = `<svg viewBox="0 0 ${w} ${h}" role="img">${grid}${paths}${xTicks}</svg><div class="line-legend">${legend}</div>`;
}

function renderCpuProfileLineGlobal(kind) {
  const sel = state.selection[kind];
  const series = [...sel.cohorts]
    .map((cohort) => {
      const ts = resolveCpuSeriesForCohort(cohort);
      if (!ts) return null;
      return { name: cohort, x: ts.x, y: ts.y };
    })
    .filter(Boolean);
  renderMultiLineChart('cpuLineGlobal', series, cpuFixedMax());
}

function quantileFromSorted(arr, q) {
  if (!arr.length) return 0;
  if (arr.length === 1) return arr[0];
  const pos = (arr.length - 1) * q;
  const base = Math.floor(pos);
  const rest = pos - base;
  if (arr[base + 1] !== undefined) return arr[base] + rest * (arr[base + 1] - arr[base]);
  return arr[base];
}

function statsRowsFromSeries(series, labelKey) {
  return series
    .filter((s) => Array.isArray(s.y) && s.y.length)
    .map((s) => {
      const sorted = [...s.y].sort((a, b) => a - b);
      return {
        [labelKey]: s.name,
        min: quantileFromSorted(sorted, 0),
        p30: quantileFromSorted(sorted, 0.3),
        p50: quantileFromSorted(sorted, 0.5),
        p70: quantileFromSorted(sorted, 0.7),
        p95: quantileFromSorted(sorted, 0.95),
        p99: quantileFromSorted(sorted, 0.99),
        max: quantileFromSorted(sorted, 1),
      };
    });
}

function buildInstanceCpuTimeSeries(kind, cohort) {
  const base = resolveCpuSeriesForCohort(cohort);
  if (!base) return [];
  const rows = selectedInstances(kind).filter((r) => r.cohort === cohort);
  if (!rows.length) return [];
  const total = rows.reduce((a, r) => a + (r.init_cpu_count || 0), 0);
  const fallbackShare = 1 / rows.length;
  return rows.map((r) => {
    const share = total > 0 ? (r.init_cpu_count || 0) / total : fallbackShare;
    return {
      key: r.instance,
      name: displayInstanceName(r.instance),
      x: base.x,
      y: base.y.map((v) => v * share),
    };
  });
}

function renderCohortTable(rows) {
  const sortedRows = sortRows('cohortTable', rows);
  const tbody = $('cohortTable').querySelector('tbody');
  tbody.innerHTML = sortedRows
    .map(
      (r) =>
        `<tr class="clickable" data-cohort-key="${encodeKey(r.cohort)}"><td>${esc(r.cohort)}</td><td>${fmt(r.dbs, 'int')}</td><td>${fmt(
          r.instances,
          'int',
        )}</td><td>${fmt(r.hosts, 'int')}</td><td>${fmt(r.vcpu, 'dec1')}</td><td>${fmt(r.mem, 'dec1')}</td><td>${fmt(
          r.allocated,
          'dec1',
        )}</td><td>${fmt(r.used, 'dec1')}</td></tr>`,
    )
    .join('');

  if (!tbody.dataset.clickInit) {
    tbody.dataset.clickInit = '1';
    tbody.addEventListener('click', (e) => {
      const tr = e.target.closest('tr[data-cohort-key]');
      if (!tr) return;
      goToCohort(decodeKey(tr.dataset.cohortKey));
    });
  }
  updateSortHeaderIndicators('cohortTable');
}

function renderGlobal() {
  const kind = state.activeTab;
  const sel = state.selection[kind];
  const metricById = currentMetricCatalogById();
  const summary = computeGlobal(kind);
  const cohortRows = computeByCohort(kind);

  const kpiMetrics = [...sel.metrics]
    .map((id) => metricById.get(id))
    .filter((m) => m && m.summary);

  $('kpiGrid').innerHTML = kpiMetrics
    .map((m) => `<div class="kpi"><div class="name">${esc(m.label)}</div><div class="value">${fmt(summary[m.id], m.fmt)}</div></div>`)
    .join('');
  renderGlobalVcpuSizer(kind);

  renderCohortTable(cohortRows);
  renderCohortMetricCharts(cohortRows);
  renderVersionChart(kind);
  renderCpuMemorySections(kind);

  renderDbStats(kind);
  renderTopSql(kind, sel.metrics.has('topsql'));
  renderSegmentIo(kind, sel.metrics.has('segmentio'));
}

function goToGlobal() {
  state.page = { mode: 'global', cohort: null, instance: null };
  renderAll();
}

function goToCohort(cohort) {
  state.page = { mode: 'cohort', cohort, instance: null };
  renderAll();
}

function goToInstance(cohort, instance) {
  state.page = { mode: 'instance', cohort, instance };
  renderAll();
}

function goToPptPreview() {
  state.page = { mode: 'pptPreview', cohort: null, instance: null };
  state.previewChartCache = state.previewChartCache || {};
  renderAll();
}

function selectedPreviewSlides() {
  return (state.pptSlidePlan || []).filter((s) => state.pptSlideSelected.has(s.id));
}

function previewDetailForSlide(slide, rows, byCohort) {
  let detail = slide.subtitle || '';
  if (slide.type === 'summary' || slide.type === 'executive-summary' || slide.type?.startsWith('deployment-')) {
    const g = computeGlobal('ppt');
    detail = `${fmt(g.db_count, 'int')} DBs | ${fmt(g.instance_count, 'int')} instances | ${fmt(g.host_count, 'int')} hosts`;
  }
  if (slide.type === 'cohort') {
    const cRows = byCohort.get(slide.cohort) || [];
    const dbs = new Set(cRows.map((r) => r.db)).size;
    detail = `${dbs} DBs | ${cRows.length} instances`;
  }
  if (slide.type === 'instance') {
    const row = rows.find((r) => r.cohort === slide.cohort && r.instance === slide.instance);
    detail = `${displayDbName(row?.db || slide.db || '')} | ${row?.host || 'Host N/A'}`;
  }
  return detail;
}

function previewKpisForSlide(slide, rows, byCohort) {
  if (slide.type === 'summary' || slide.type === 'executive-summary' || slide.type?.startsWith('deployment-')) {
    const g = computeGlobal('ppt');
    return [
      { k: 'DBs', v: fmt(g.db_count, 'int') },
      { k: 'Hosts', v: fmt(g.host_count, 'int') },
      { k: 'Instances', v: fmt(g.instance_count, 'int') },
      { k: 'vCPU', v: fmt(g.vcpu_total, 'dec1') },
      { k: 'Memory(GB)', v: fmt(g.memory_gb_total, 'dec1') },
      { k: 'Allocated', v: fmt(g.allocated_storage_gb, 'dec1') },
    ];
  }
  if (slide.type === 'cohort') {
    const cRows = byCohort.get(slide.cohort) || [];
    const s = computeSummaryFromRows(cRows);
    return [
      { k: 'DBs', v: fmt(s.db_count, 'int') },
      { k: 'Hosts', v: fmt(s.host_count, 'int') },
      { k: 'Instances', v: fmt(s.instance_count, 'int') },
      { k: 'vCPU', v: fmt(s.vcpu_total, 'dec1') },
      { k: 'Memory(GB)', v: fmt(s.memory_gb_total, 'dec1') },
      { k: 'Allocated', v: fmt(s.allocated_storage_gb, 'dec1') },
    ];
  }
  if (slide.type === 'instance') {
    const row = rows.find((r) => r.cohort === slide.cohort && r.instance === slide.instance);
    return [
      { k: 'DB', v: displayDbName(row?.db || slide.db || 'N/A') },
      { k: 'Host', v: row?.host || 'N/A' },
      { k: 'vCPU', v: fmt(row?.init_cpu_count || row?.logical_cpu_count || 0, 'dec1') },
      { k: 'Memory', v: fmt(row?.mem_gb || 0, 'dec1') },
      { k: 'SGA', v: fmt(row?.sga_size_gb || 0, 'dec1') },
      { k: 'PGA', v: fmt(row?.pga_size_gb || 0, 'dec1') },
    ];
  }
  return [];
}

function previewRowsForSlide(slide, rows, byCohort) {
  if (slide.type === 'summary' || slide.type === 'executive-summary' || slide.type?.startsWith('deployment-')) {
    return computeByCohort('ppt')
      .slice(0, 5)
      .map((r) => [r.cohort, fmt(r.dbs, 'int'), fmt(r.instances, 'int'), fmt(r.vcpu, 'dec1')]);
  }
  if (slide.type === 'cohort') {
    const cRows = byCohort.get(slide.cohort) || [];
    const byDb = new Map();
    cRows.forEach((r) => {
      if (!byDb.has(r.db)) byDb.set(r.db, 0);
      byDb.set(r.db, byDb.get(r.db) + 1);
    });
    return [...byDb.entries()]
      .sort((a, b) => b[1] - a[1])
      .slice(0, 5)
      .map(([db, n]) => [displayDbName(db), String(n)]);
  }
  if (slide.type === 'instance') {
    const db = rows.find((r) => r.cohort === slide.cohort && r.instance === slide.instance)?.db || slide.db;
    return (state.graph.topSql || [])
      .filter((r) => r.db === db)
      .sort((a, b) => b.elapsed - a.elapsed)
      .slice(0, 4)
      .map((r) => [r.sql_id || '-', fmt(r.elapsed, 'dec1')]);
  }
  return [];
}

function buildSlideMockHtml(slide, detail, compact = false, rows = [], byCohort = new Map(), chartImgs = null) {
  const chips = [];
  if (slide.type === 'executive-summary') chips.push('Executive Analysis', 'Infrastructure Fit');
  if (slide.type === 'deployment-base') chips.push('Base Database', 'Detailed Sizing');
  if (slide.type === 'deployment-exascale') chips.push('Exascale', 'Detailed Sizing');
  if (slide.type === 'deployment-exadata') chips.push('Exadata Dedicated', 'Detailed Sizing');
  if (slide.type === 'title') chips.push('Cover Slide');
  if (slide.type === 'summary') chips.push('Global KPIs', 'Cohort Table');
  if (slide.type === 'separator') chips.push('Section Divider');
  if (slide.type === 'cohort') chips.push('Cohort KPIs', 'DB Breakdown');
  if (slide.type === 'instance') chips.push('Instance KPIs', 'Top SQL');
  const kpis = previewKpisForSlide(slide, rows, byCohort);
  const dataRows = previewRowsForSlide(slide, rows, byCohort);
  const sectionClass = slide.type === 'separator' ? 'separator' : slide.type === 'title' ? 'cover' : '';

  const kpiHtml =
    compact || !kpis.length
      ? ''
      : `<div class="mock-kpis">${kpis
          .map((x) => `<div class="mock-kpi"><span>${esc(x.k)}</span><strong>${esc(x.v)}</strong></div>`)
          .join('')}</div>`;

  const chartA = geomToPct(PPT_LAYOUT_MAP.chartPrimary);
  const chartB = geomToPct(PPT_LAYOUT_MAP.chartSecondary);
  const tableGeom = geomToPct({ x: 0.5, y: PPT_LAYOUT_MAP.tableY, w: 12.3, h: PPT_LAYOUT_MAP.tableH });
  const chartZones =
    compact || slide.type === 'separator' || slide.type === 'title'
      ? ''
      : `
      <div class="mock-layout-zone table" style="left:${tableGeom.left};top:${tableGeom.top};width:${tableGeom.width};height:${tableGeom.height};">Table Zone</div>
      <div class="mock-layout-zone chart image" style="left:${chartA.left};top:${chartA.top};width:${chartA.width};height:${chartA.height};">${
        chartImgs?.primary
          ? `<img src="${chartImgs.primary}" alt="Chart preview A" />`
          : '<span>Chart A</span>'
      }</div>
      <div class="mock-layout-zone chart image" style="left:${chartB.left};top:${chartB.top};width:${chartB.width};height:${chartB.height};">${
        chartImgs?.secondary
          ? `<img src="${chartImgs.secondary}" alt="Chart preview B" />`
          : '<span>Chart B</span>'
      }</div>
    `;

  let tableHtml = '';
  if (!compact && dataRows.length) {
    const hdr =
      slide.type === 'summary' || slide.type === 'executive-summary'
        ? ['Cohort', 'DBs', 'Inst', 'vCPU']
        : slide.type === 'cohort'
          ? ['Database', 'Instances']
          : ['SQL ID', 'Elapsed'];
    tableHtml = `
      <div class="mock-table">
        <div class="mock-row head">${hdr.map((h) => `<span>${esc(h)}</span>`).join('')}</div>
        ${dataRows.map((r) => `<div class="mock-row">${r.map((c) => `<span>${esc(c)}</span>`).join('')}</div>`).join('')}
      </div>
    `;
  }

  return `
    <div class="ppt-slide-mock ${compact ? 'compact' : ''} ${sectionClass}">
      <div class="top-band"></div>
      <div class="left-band"></div>
      <div class="content">
        <div class="title">${esc(shortLabel(slide.title || 'Slide', compact ? 36 : 96))}</div>
        <div class="subtitle">${esc(detail || slide.subtitle || '')}</div>
        ${kpiHtml}
        ${tableHtml}
        <div class="chips">${chips.map((c) => `<span class="chip">${esc(c)}</span>`).join('')}</div>
        <div class="note">${esc(slideBadge(slide))}</div>
        ${chartZones}
      </div>
    </div>
  `;
}

function movePreviewSelection(offset) {
  const slides = selectedPreviewSlides();
  if (!slides.length) return;
  const idx = Math.max(0, slides.findIndex((s) => s.id === state.previewSlideId));
  const nextIdx = Math.max(0, Math.min(slides.length - 1, idx + offset));
  state.previewSlideId = slides[nextIdx].id;
  renderPptPreviewPage();
}

async function fillPreviewChartsForSlide(slide) {
  if (!slide || slide.type === 'separator' || slide.type === 'title') return {};
  if (state.previewChartCache[slide.id]) return state.previewChartCache[slide.id];
  const imgs = await buildSlideChartImages(slide);
  state.previewChartCache[slide.id] = imgs || {};
  return state.previewChartCache[slide.id];
}

function renderPptPreviewPage() {
  const token = ++state.previewRenderToken;
  void renderPptPreviewPageAsync(token);
}

async function renderPptPreviewPageAsync(token) {
  const status = $('pptPreviewStatus');
  const thumbs = $('pptPreviewThumbs');
  const body = $('pptPreviewBody');
  const exportBtn = $('previewExportBtn');
  const prevBtn = $('previewPrevBtn');
  const nextBtn = $('previewNextBtn');
  if (!status || !body || !thumbs) return;
  if (!state.graph) {
    status.textContent = 'Load JSON to preview slides.';
    thumbs.innerHTML = '';
    body.innerHTML = '';
    if (exportBtn) exportBtn.disabled = true;
    return;
  }

  const selectedSlides = selectedPreviewSlides();
  if (!state.previewSlideId || !selectedSlides.some((s) => s.id === state.previewSlideId)) {
    state.previewSlideId = selectedSlides[0]?.id || null;
  }
  const activeSlide = selectedSlides.find((s) => s.id === state.previewSlideId) || null;
  const activeIndex = Math.max(0, selectedSlides.findIndex((s) => s.id === state.previewSlideId));
  status.textContent = selectedSlides.length
    ? `${activeIndex + 1}/${selectedSlides.length} selected slides. Review content and sequence before export.`
    : 'No slides selected.';
  if (exportBtn) exportBtn.disabled = selectedSlides.length === 0;
  if (prevBtn) prevBtn.disabled = activeIndex <= 0 || selectedSlides.length === 0;
  if (nextBtn) nextBtn.disabled = activeIndex >= selectedSlides.length - 1 || selectedSlides.length === 0;

  const rows = selectedInstances('ppt');
  const byCohort = new Map();
  rows.forEach((r) => {
    if (!byCohort.has(r.cohort)) byCohort.set(r.cohort, []);
    byCohort.get(r.cohort).push(r);
  });

  thumbs.innerHTML = selectedSlides
    .map((s, idx) => {
      const detail = previewDetailForSlide(s, rows, byCohort);
      const imgs = state.previewChartCache[s.id] || null;
      return `
        <div class="ppt-thumb ${s.id === state.previewSlideId ? 'active' : ''}" data-preview-slide-id="${encodeKey(s.id)}">
          ${buildSlideMockHtml(s, detail, true, rows, byCohort, imgs)}
          <div class="ppt-thumb-title">${idx + 1}. ${esc(shortLabel(s.title, 52))}</div>
        </div>
      `;
    })
    .join('');

  if (!activeSlide) {
    body.innerHTML = '<p class="muted">No slides selected for preview.</p>';
    return;
  }

  const detail = previewDetailForSlide(activeSlide, rows, byCohort);
  const activeImgs = state.previewChartCache[activeSlide.id] || null;
  body.innerHTML = `
    <div class="ppt-stage-meta">
      <span class="slide-badge">${slideBadge(activeSlide)}</span>
      <span><strong>${activeIndex + 1}. ${esc(activeSlide.title)}</strong></span>
      <span>${esc(detail)}</span>
    </div>
    ${buildSlideMockHtml(activeSlide, detail, false, rows, byCohort, activeImgs)}
  `;

  thumbs.querySelectorAll('[data-preview-slide-id]').forEach((el) => {
    el.addEventListener('click', () => {
      state.previewSlideId = decodeKey(el.getAttribute('data-preview-slide-id'));
      renderPptPreviewPage();
    });
  });

  const needThumb = selectedSlides.filter((s) => !state.previewChartCache[s.id] && s.type !== 'separator' && s.type !== 'title').slice(0, 4);
  if (needThumb.length) {
    for (const s of needThumb) {
      await fillPreviewChartsForSlide(s);
      if (token !== state.previewRenderToken) return;
    }
    renderPptPreviewPage();
    return;
  }

  if (!state.previewChartCache[activeSlide.id] && activeSlide.type !== 'separator' && activeSlide.type !== 'title') {
    await fillPreviewChartsForSlide(activeSlide);
    if (token !== state.previewRenderToken) return;
    renderPptPreviewPage();
  }
}

function renderCohortPage(kind) {
  const rows = selectedInstances(kind).filter((r) => r.cohort === state.page.cohort);
  const summary = computeSummaryFromRows(rows);
  ensureCohortTargetsInitialized(kind);
  const cohortCfg = ensureCohortTargetContainer(state.page.cohort);
  const cohortTargets = {
    vcpu: parseTargetPctOrDefault(cohortCfg.vcpu, getGlobalTarget('vcpu')),
    memory: parseTargetPctOrDefault(cohortCfg.memory, getGlobalTarget('memory')),
    iops: parseTargetPctOrDefault(cohortCfg.iops, getGlobalTarget('iops')),
    storage: parseTargetPctOrDefault(cohortCfg.storage, getGlobalTarget('storage')),
  };

  $('cohortPageTitle').textContent = `Cohort: ${state.page.cohort}`;
  $('cohortPageBody').innerHTML = `
    <div class="kpi-grid">
      <div class="kpi"><div class="name">Databases</div><div class="value">${fmt(summary.db_count, 'int')}</div></div>
      <div class="kpi"><div class="name">Hosts</div><div class="value">${fmt(summary.host_count, 'int')}</div></div>
      <div class="kpi"><div class="name">Instances</div><div class="value">${fmt(summary.instance_count, 'int')}</div></div>
      <div class="kpi"><div class="name">vCPU</div><div class="value">${fmt(summary.vcpu_total, 'dec1')}</div></div>
      <div class="kpi"><div class="name">Memory (GB)</div><div class="value">${fmt(summary.memory_gb_total, 'dec1')}</div></div>
      <div class="kpi"><div class="name">Allocated (GB)</div><div class="value">${fmt(summary.allocated_storage_gb, 'dec1')}</div></div>
    </div>
    <div class="chart-box" style="margin-top:10px;">
      <h3>Cohort Sizing Overrides</h3>
      <p class="muted">Adjust this cohort targets. Global target is recalculated as weighted average.</p>
      <div class="global-sizer-controls-grid">
        <div class="global-sizer-control-card">
          <label for="cohortTargetVcpuInput" style="margin:0 0 4px;color:#475569;font-size:12px;">VCPU target (0-200%)</label>
          <input id="cohortTargetVcpuInput" type="number" min="0" max="200" step="0.1" value="${fmt(cohortTargets.vcpu, 'dec1')}" style="width:100%;padding:6px 8px;border:1px solid #cfd6df;border-radius:8px;" />
          <div id="cohortTargetVcpuError" style="min-height:16px;color:#b91c1c;font-size:11px;margin-top:4px;"></div>
          <div id="cohortTargetVcpuSummary" class="line" style="font-size:11px;color:#475569;"></div>
        </div>
        <div class="global-sizer-control-card">
          <label for="cohortTargetMemInput" style="margin:0 0 4px;color:#475569;font-size:12px;">Memory target (0-200%)</label>
          <input id="cohortTargetMemInput" type="number" min="0" max="200" step="0.1" value="${fmt(cohortTargets.memory, 'dec1')}" style="width:100%;padding:6px 8px;border:1px solid #cfd6df;border-radius:8px;" />
          <div id="cohortTargetMemError" style="min-height:16px;color:#b91c1c;font-size:11px;margin-top:4px;"></div>
          <div id="cohortTargetMemSummary" class="line" style="font-size:11px;color:#475569;"></div>
        </div>
        <div class="global-sizer-control-card">
          <label for="cohortTargetIopsInput" style="margin:0 0 4px;color:#475569;font-size:12px;">IOPS target (0-200%)</label>
          <input id="cohortTargetIopsInput" type="number" min="0" max="200" step="0.1" value="${fmt(cohortTargets.iops, 'dec1')}" style="width:100%;padding:6px 8px;border:1px solid #cfd6df;border-radius:8px;" />
          <div id="cohortTargetIopsError" style="min-height:16px;color:#b91c1c;font-size:11px;margin-top:4px;"></div>
          <div id="cohortTargetIopsSummary" class="line" style="font-size:11px;color:#475569;"></div>
        </div>
        <div class="global-sizer-control-card">
          <label for="cohortTargetStorageInput" style="margin:0 0 4px;color:#475569;font-size:12px;">Storage target (0-200%)</label>
          <input id="cohortTargetStorageInput" type="number" min="0" max="200" step="0.1" value="${fmt(cohortTargets.storage, 'dec1')}" style="width:100%;padding:6px 8px;border:1px solid #cfd6df;border-radius:8px;" />
          <div id="cohortTargetStorageError" style="min-height:16px;color:#b91c1c;font-size:11px;margin-top:4px;"></div>
          <div id="cohortTargetStorageSummary" class="line" style="font-size:11px;color:#475569;"></div>
        </div>
      </div>
    </div>
    <div class="chart-box" style="margin-top:10px;">
      <h3>Cohort Effective Sizing by Deployment</h3>
      <p class="muted">Provisioned lines shown for Base, Exascale, and Exadata per metric.</p>
      <div class="global-vcpu-sizer-wrap">
        <div id="cohortSizerChartVcpu" class="global-vcpu-sizer-chart"></div>
        <div id="cohortSizerChartMem" class="global-vcpu-sizer-chart"></div>
        <div id="cohortSizerChartIops" class="global-vcpu-sizer-chart"></div>
        <div id="cohortSizerChartStorage" class="global-vcpu-sizer-chart"></div>
      </div>
    </div>
    <p class="muted" style="margin-top:10px;">Select an instance to continue.</p>
    <div class="chart-box" style="margin-bottom:10px;">
      <h3>CPU Profile Line Chart by Instance</h3>
      <div id="cpuLineCohortPage"></div>
    </div>
    <div class="triple-grid" style="margin-bottom:10px;">
      <div class="chart-box mini-chart">
        <h3>Top CPU by Instance (max)</h3>
        <div id="cohortCpuTopChart"></div>
      </div>
      <div class="chart-box mini-chart">
        <h3>CPU by Instance p95</h3>
        <div id="cohortCpuP95Chart"></div>
      </div>
      <div class="chart-box mini-chart">
        <h3>CPU Box Plot by Instance</h3>
        <div id="cohortCpuBoxChart"></div>
      </div>
    </div>
    <div class="triple-grid" style="margin-bottom:10px;">
      <div class="chart-box mini-chart">
        <h3>Top Memory by Database (max)</h3>
        <div id="cohortMemTopChart"></div>
      </div>
      <div class="chart-box mini-chart">
        <h3>Memory by Database p95</h3>
        <div id="cohortMemP95Chart"></div>
      </div>
      <div class="chart-box mini-chart">
        <h3>Memory Box Plot by Database</h3>
        <div id="cohortMemBoxChart"></div>
      </div>
    </div>
    <div class="table-wrap">
      <table>
        <thead><tr><th>Instance</th><th>Database</th><th>Host</th><th>vCPU</th><th>Memory (GB)</th><th>SGA (GB)</th><th>PGA (GB)</th></tr></thead>
        <tbody>
          ${rows
            .sort((a, b) => a.instance.localeCompare(b.instance))
            .map(
              (r) =>
                `<tr class="clickable" data-page-instance-key="${encodeKey(r.instance)}"><td>${esc(displayInstanceName(r.instance))}</td><td>${esc(
                  shortLabel(displayDbName(r.db), 36),
                )}</td><td>${esc(r.host)}</td><td>${fmt(r.init_cpu_count, 'dec1')}</td><td>${fmt(r.mem_gb, 'dec1')}</td><td>${fmt(
                  r.sga_size_gb,
                  'dec1',
                )}</td><td>${fmt(r.pga_size_gb, 'dec1')}</td></tr>`,
            )
            .join('')}
        </tbody>
      </table>
    </div>
  `;

  $('cohortPageBody').querySelectorAll('tr[data-page-instance-key]').forEach((tr) => {
    tr.addEventListener('click', () => {
      goToInstance(state.page.cohort, decodeKey(tr.dataset.pageInstanceKey));
    });
  });

  const validateTargetInput = (raw) => {
    if (raw == null || String(raw).trim() === '') return { ok: false, msg: 'Enter a value between 0 and 200.' };
    const n = Number(raw);
    if (!Number.isFinite(n)) return { ok: false, msg: 'Invalid number.' };
    if (n < 0) return { ok: false, msg: 'Minimum is 0.' };
    if (n > 200) return { ok: false, msg: 'Max is 200.' };
    return { ok: true, value: n };
  };

  const hookCohortTarget = (metric, inputId, errId, summaryId) => {
    const input = $('cohortPageBody').querySelector(`#${inputId}`);
    const errNode = $('cohortPageBody').querySelector(`#${errId}`);
    const summaryNode = $('cohortPageBody').querySelector(`#${summaryId}`);
    if (!input) return;
    const refreshSummary = () => {
      const cohortVal = parseTargetPctOrDefault(ensureCohortTargetContainer(state.page.cohort)[metric], 100);
      const globalVal = getGlobalTarget(metric);
      if (summaryNode) summaryNode.textContent = `Cohort: ${fmt(cohortVal, 'dec1')}% | Weighted global: ${fmt(globalVal, 'dec1')}%`;
    };
    input.addEventListener('input', () => {
      const v = validateTargetInput(input.value);
      if (!v.ok) {
        if (errNode) errNode.textContent = v.msg;
        return;
      }
      if (errNode) errNode.textContent = '';
      ensureCohortTargetContainer(state.page.cohort)[metric] = v.value;
      recomputeGlobalTargetFromCohorts(kind, metric);
      refreshSummary();
    });
    input.addEventListener('blur', () => {
      const v = validateTargetInput(input.value);
      if (!v.ok) {
        input.value = String(fmt(parseTargetPctOrDefault(ensureCohortTargetContainer(state.page.cohort)[metric], 100), 'dec1'));
        if (errNode) errNode.textContent = '';
        return;
      }
      ensureCohortTargetContainer(state.page.cohort)[metric] = clampTargetPct(v.value);
      recomputeGlobalTargetFromCohorts(kind, metric);
      input.value = String(fmt(parseTargetPctOrDefault(ensureCohortTargetContainer(state.page.cohort)[metric], 100), 'dec1'));
      if (errNode) errNode.textContent = '';
      refreshSummary();
    });
    refreshSummary();
  };

  hookCohortTarget('vcpu', 'cohortTargetVcpuInput', 'cohortTargetVcpuError', 'cohortTargetVcpuSummary');
  hookCohortTarget('memory', 'cohortTargetMemInput', 'cohortTargetMemError', 'cohortTargetMemSummary');
  hookCohortTarget('iops', 'cohortTargetIopsInput', 'cohortTargetIopsError', 'cohortTargetIopsSummary');
  hookCohortTarget('storage', 'cohortTargetStorageInput', 'cohortTargetStorageError', 'cohortTargetStorageSummary');

  const vcpuStats = computeCohortCpuStats(kind, state.page.cohort);
  const memStats = computeCohortMemoryStats(kind, state.page.cohort);
  const iopsStats = computeCohortIopsStats(kind, state.page.cohort);
  const storageStats = computeCohortStorageStats(kind, state.page.cohort);

  if (vcpuStats) {
    renderVerticalSizerChart(
      $('cohortPageBody').querySelector('#cohortSizerChartVcpu'),
      vcpuStats,
      cohortTargets.vcpu,
      'VCPU Utilization',
      '#0f766e',
      cohortProvisionLines(kind, state.page.cohort, 'vcpu'),
    );
  }
  if (memStats) {
    renderVerticalSizerChart(
      $('cohortPageBody').querySelector('#cohortSizerChartMem'),
      memStats,
      cohortTargets.memory,
      'Memory Utilization',
      '#0e7490',
      cohortProvisionLines(kind, state.page.cohort, 'memory'),
    );
  }
  if (iopsStats) {
    renderVerticalSizerChart(
      $('cohortPageBody').querySelector('#cohortSizerChartIops'),
      iopsStats,
      cohortTargets.iops,
      'IOPS Utilization',
      '#6d28d9',
      cohortProvisionLines(kind, state.page.cohort, 'iops'),
    );
  }
  if (storageStats) {
    renderVerticalSizerChart(
      $('cohortPageBody').querySelector('#cohortSizerChartStorage'),
      storageStats,
      cohortTargets.storage,
      'Storage Utilization',
      '#b45309',
      cohortProvisionLines(kind, state.page.cohort, 'storage'),
    );
  }

  const instanceSeries = buildInstanceCpuTimeSeries(kind, state.page.cohort);
  const cpuMax = cpuFixedMax();
  renderMultiLineChart('cpuLineCohortPage', instanceSeries, cpuMax);

  const cpuRows = statsRowsFromSeries(instanceSeries, 'instance');
  renderRankBarsByGroup('cohortCpuTopChart', cpuRows, 'max', 'dec1', 'instance', cpuMax);
  renderRankBarsByGroup('cohortCpuP95Chart', cpuRows, 'p95', 'dec3', 'instance', cpuMax);
  renderBoxPlotByGroup('cohortCpuBoxChart', cpuRows, 'dec1', 'instance', cpuMax);

  const memRows = aggregateMetricByDb(kind, 'DB Memory (MB)', (v) => v / 1024, state.page.cohort);
  renderRankBarsByGroup('cohortMemTopChart', memRows, 'max', 'dec1', 'db');
  renderRankBarsByGroup('cohortMemP95Chart', memRows, 'p95', 'dec1', 'db');
  renderBoxPlotByGroup('cohortMemBoxChart', memRows, 'dec1', 'db');
}

function renderInstancePage(kind) {
  const rows = selectedInstances(kind).filter((r) => r.cohort === state.page.cohort && r.instance === state.page.instance);
  const item = rows[0];
  const db = item?.db || '';
  const sqlRows = state.graph.topSql
    .filter((r) => r.db === db)
    .sort((a, b) => b.elapsed - a.elapsed)
    .slice(0, 12);

  $('instancePageTitle').textContent = `Instance: ${displayInstanceName(state.page.instance)}`;
  $('instancePageBody').innerHTML = `
    <div class="kpi-grid">
      <div class="kpi"><div class="name">Cohort</div><div class="value" style="font-size:16px;">${esc(item?.cohort || '')}</div></div>
      <div class="kpi"><div class="name">Database</div><div class="value" style="font-size:16px;">${esc(shortLabel(displayDbName(db), 26))}</div></div>
      <div class="kpi"><div class="name">Host</div><div class="value" style="font-size:16px;">${esc(item?.host || '')}</div></div>
      <div class="kpi"><div class="name">vCPU</div><div class="value">${fmt(item?.init_cpu_count || 0, 'dec1')}</div></div>
      <div class="kpi"><div class="name">Memory (GB)</div><div class="value">${fmt(item?.mem_gb || 0, 'dec1')}</div></div>
      <div class="kpi"><div class="name">SGA/PGA (GB)</div><div class="value">${fmt(item?.sga_size_gb || 0, 'dec1')} / ${fmt(
        item?.pga_size_gb || 0,
        'dec1',
      )}</div></div>
    </div>
    <div class="chart-box" style="margin:10px 0;">
      <h3>CPU Profile Line Chart (Instance)</h3>
      <div id="cpuLineInstancePage"></div>
    </div>
    <div class="triple-grid" style="margin:10px 0;">
      <div class="chart-box mini-chart">
        <h3>Top CPU (max)</h3>
        <div id="instanceCpuTopChart"></div>
      </div>
      <div class="chart-box mini-chart">
        <h3>CPU p95</h3>
        <div id="instanceCpuP95Chart"></div>
      </div>
      <div class="chart-box mini-chart">
        <h3>CPU Box Plot</h3>
        <div id="instanceCpuBoxChart"></div>
      </div>
    </div>
    <div class="triple-grid" style="margin-bottom:10px;">
      <div class="chart-box mini-chart">
        <h3>Top Memory (max)</h3>
        <div id="instanceMemTopChart"></div>
      </div>
      <div class="chart-box mini-chart">
        <h3>Memory p95</h3>
        <div id="instanceMemP95Chart"></div>
      </div>
      <div class="chart-box mini-chart">
        <h3>Memory Box Plot</h3>
        <div id="instanceMemBoxChart"></div>
      </div>
    </div>
    <p class="muted" style="margin-top:10px;">Top SQL for this instance's database</p>
    <div class="table-wrap">
      <table>
        <thead><tr><th>SQL ID</th><th>Elapsed</th><th>Execs</th><th>Logical Reads</th><th>Physical Read GB</th></tr></thead>
        <tbody>
          ${sqlRows
            .map(
              (r) =>
                `<tr><td>${esc(r.sql_id)}</td><td>${fmt(r.elapsed, 'dec1')}</td><td>${fmt(r.execs, 'dec1')}</td><td>${fmt(
                  r.log_reads,
                  'dec1',
                )}</td><td>${fmt(r.phy_read_gb, 'dec1')}</td></tr>`,
            )
            .join('')}
        </tbody>
      </table>
    </div>
  `;

  const instanceSeries = buildInstanceCpuTimeSeries(kind, state.page.cohort).filter((s) => s.key === state.page.instance);
  const cpuMax = cpuFixedMax();
  renderMultiLineChart('cpuLineInstancePage', instanceSeries, cpuMax);

  const dbScope = new Set([db]);
  const cpuRows = statsRowsFromSeries(instanceSeries, 'instance');
  renderRankBarsByGroup('instanceCpuTopChart', cpuRows, 'max', 'dec1', 'instance', cpuMax);
  renderRankBarsByGroup('instanceCpuP95Chart', cpuRows, 'p95', 'dec3', 'instance', cpuMax);
  renderBoxPlotByGroup('instanceCpuBoxChart', cpuRows, 'dec1', 'instance', cpuMax);

  const memRows = aggregateMetricByDb(kind, 'DB Memory (MB)', (v) => v / 1024, state.page.cohort, dbScope);
  renderRankBarsByGroup('instanceMemTopChart', memRows, 'max', 'dec1', 'db');
  renderRankBarsByGroup('instanceMemP95Chart', memRows, 'p95', 'dec1', 'db');
  renderBoxPlotByGroup('instanceMemBoxChart', memRows, 'dec1', 'db');
}

function renderPageMode() {
  const nav = $('pageNav');
  const cohortPage = $('cohortPage');
  const instancePage = $('instancePage');
  const pptPreviewPage = $('pptPreviewPage');
  const sections = Array.from(document.querySelectorAll('.content > .card'));
  const globalCards = sections.filter((s) => !['navCard', 'cohortPage', 'instancePage', 'pptPreviewPage'].includes(s.id));
  const kind = state.activeTab;

  if (state.page.mode === 'global') {
    globalCards.forEach((c) => (c.style.display = ''));
    applyCollapsedSections();
    cohortPage.style.display = 'none';
    instancePage.style.display = 'none';
    if (pptPreviewPage) pptPreviewPage.style.display = 'none';
    nav.textContent = 'Global';
    return;
  }

  globalCards.forEach((c) => (c.style.display = 'none'));

  if (state.page.mode === 'pptPreview') {
    nav.innerHTML = `<span class="page-link" id="navGlobal">Global</span> / PPT Preview`;
    cohortPage.style.display = 'none';
    instancePage.style.display = 'none';
    if (pptPreviewPage) pptPreviewPage.style.display = '';
    renderPptPreviewPage();
    $('navGlobal')?.addEventListener('click', goToGlobal);
    return;
  }

  if (state.page.mode === 'cohort') {
    nav.innerHTML = `<span class="page-link" id="navGlobal">Global</span> / ${esc(state.page.cohort)}`;
    cohortPage.style.display = '';
    instancePage.style.display = 'none';
    if (pptPreviewPage) pptPreviewPage.style.display = 'none';
    renderCohortPage(kind);
  } else {
    nav.innerHTML = `<span class="page-link" id="navGlobal">Global</span> / <span class="page-link" id="navCohort">${esc(
      state.page.cohort,
    )}</span> / ${esc(shortLabel(displayInstanceName(state.page.instance), 28))}`;
    cohortPage.style.display = 'none';
    instancePage.style.display = '';
    if (pptPreviewPage) pptPreviewPage.style.display = 'none';
    renderInstancePage(kind);
  }

  $('navGlobal')?.addEventListener('click', goToGlobal);
  $('navCohort')?.addEventListener('click', () => goToCohort(state.page.cohort));
}

function renderCpuScaleToggleLabel() {
  const btn = $('cpuScaleToggle');
  if (!btn) return;
  btn.textContent = state.cpuScaleMode === 'fixed100' ? 'Scale: 100%' : 'Scale: Dynamic';
}

function renderDbStats(kind) {
  const sel = state.selection[kind];
  const selectedDbStatNames = [...sel.metrics]
    .filter((id) => id.startsWith('dbstat:'))
    .map((id) => id.slice('dbstat:'.length));

  const tbody = $('dbStatsTable').querySelector('tbody');
  if (!selectedDbStatNames.length) {
    tbody.innerHTML = '<tr><td colspan="5">No database statistic metric selected.</td></tr>';
    updateSortHeaderIndicators('dbStatsTable');
    return;
  }

  const rows = state.graph.dbStats.filter((r) => sel.cohorts.has(r.cohort) && sel.dbs.has(r.db) && selectedDbStatNames.includes(r.metric));
  if (!rows.length) {
    tbody.innerHTML = '<tr><td colspan="5">No rows available for current scope and metric selection.</td></tr>';
    updateSortHeaderIndicators('dbStatsTable');
    return;
  }

  const grouped = new Map();
  rows.forEach((r) => {
    if (!grouped.has(r.metric)) grouped.set(r.metric, []);
    grouped.get(r.metric).push(r);
  });

  const out = [...grouped.entries()].map(([metric, items]) => {
    const avg = (field) => items.reduce((a, r) => a + (Number(r[field]) || 0), 0) / Math.max(items.length, 1);
    const top = items.reduce((best, r) => ((r.max || 0) > (best.max || 0) ? r : best), items[0]);
    return {
      metric,
      p95: avg('p95'),
      p99: avg('p99'),
      max: avg('max'),
      topDb: top.db,
      topDbDisplay: displayDbName(top.db),
      topMax: top.max,
    };
  });

  const sortedOut = sortRows('dbStatsTable', out);
  tbody.innerHTML = sortedOut
    .map(
      (r) =>
        `<tr><td>${esc(r.metric)}</td><td>${fmt(r.p95, 'dec1')}</td><td>${fmt(r.p99, 'dec1')}</td><td>${fmt(
          r.max,
          'dec1',
        )}</td><td>${esc(r.topDbDisplay)} (${fmt(r.topMax, 'dec1')})</td></tr>`,
    )
    .join('');
  updateSortHeaderIndicators('dbStatsTable');
}

function renderTopSql(kind, enabled) {
  const sel = state.selection[kind];
  const tbody = $('topSqlTable').querySelector('tbody');
  if (!enabled) {
    tbody.innerHTML = '<tr><td colspan="6">Select metric "Top SQL panel" to display this table.</td></tr>';
    updateSortHeaderIndicators('topSqlTable');
    return;
  }

  const rows = state.graph.topSql.filter((r) => sel.dbs.has(r.db));
  const sortedRows = sortRows('topSqlTable', rows).slice(0, 20);

  if (!sortedRows.length) {
    tbody.innerHTML = '<tr><td colspan="6">No top SQL rows for current scope.</td></tr>';
    updateSortHeaderIndicators('topSqlTable');
    return;
  }

  tbody.innerHTML = sortedRows
    .map(
      (r) =>
        `<tr><td>${esc(displayDbName(r.db))}</td><td>${esc(r.sql_id)}</td><td>${fmt(r.elapsed, 'dec1')}</td><td>${fmt(r.execs, 'dec1')}</td><td>${fmt(
          r.log_reads,
          'dec1',
        )}</td><td>${fmt(r.phy_read_gb, 'dec1')}</td></tr>`,
    )
    .join('');
  updateSortHeaderIndicators('topSqlTable');
}

function renderSegmentIo(kind, enabled) {
  const sel = state.selection[kind];
  const tbody = $('segmentIoTable').querySelector('tbody');
  if (!enabled) {
    tbody.innerHTML = '<tr><td colspan="5">Select metric "Segment IO panel" to display this table.</td></tr>';
    updateSortHeaderIndicators('segmentIoTable');
    return;
  }

  const rows = state.graph.segmentIo.filter((r) => sel.dbs.has(r.db));
  const sortedRows = sortRows('segmentIoTable', rows).slice(0, 20);

  if (!sortedRows.length) {
    tbody.innerHTML = '<tr><td colspan="5">No segment IO rows for current scope.</td></tr>';
    updateSortHeaderIndicators('segmentIoTable');
    return;
  }

  tbody.innerHTML = sortedRows
    .map(
      (r) =>
        `<tr><td>${esc(displayDbName(r.db))}</td><td>${esc(r.owner)}</td><td>${esc(r.object_name)}</td><td>${esc(r.object_type)}</td><td>${fmt(
          r.physical_io_tot,
          'dec1',
        )}</td></tr>`,
    )
    .join('');
  updateSortHeaderIndicators('segmentIoTable');
}

function renderDrilldown() {
  const panel = $('drillPanel');
  const kind = state.activeTab;
  const rows = selectedInstances(kind);
  const cohortSet = new Set(rows.map((r) => r.cohort));
  if (state.drill.cohort && !cohortSet.has(state.drill.cohort)) {
    state.drill = { cohort: null, db: null, instance: null };
  }

  const cohortRows = state.drill.cohort ? rows.filter((r) => r.cohort === state.drill.cohort) : rows;
  const dbSet = new Set(cohortRows.map((r) => r.db));
  if (state.drill.db && !dbSet.has(state.drill.db)) {
    state.drill.db = null;
    state.drill.instance = null;
  }

  const dbRows = state.drill.db ? cohortRows.filter((r) => r.db === state.drill.db) : cohortRows;
  const instanceSet = new Set(dbRows.map((r) => r.instance));
  if (state.drill.instance && !instanceSet.has(state.drill.instance)) {
    state.drill.instance = null;
  }

  const crumbs = [
    `<button class="btn tiny secondary" data-drill-level="global">Global</button>`,
    state.drill.cohort ? `<button class="btn tiny secondary" data-drill-level="cohort">${esc(state.drill.cohort)}</button>` : '',
    state.drill.db ? `<button class="btn tiny secondary" data-drill-level="db">${esc(shortLabel(displayDbName(state.drill.db), 28))}</button>` : '',
    state.drill.instance ? `<button class="btn tiny secondary" data-drill-level="instance">${esc(shortLabel(displayInstanceName(state.drill.instance), 24))}</button>` : '',
  ]
    .filter(Boolean)
    .join(' <span class="muted">/</span> ');

  $('drillTitle').textContent = 'Drilldown Navigator';

  if (!state.drill.cohort) {
    const byCohort = new Map();
    rows.forEach((r) => {
      if (!byCohort.has(r.cohort)) byCohort.set(r.cohort, { cohort: r.cohort, dbs: new Set(), instances: 0, hosts: new Set() });
      const g = byCohort.get(r.cohort);
      g.dbs.add(r.db);
      g.instances += 1;
      g.hosts.add(r.host);
    });

    panel.innerHTML = `
      <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:10px;">${crumbs}</div>
      <p class="muted">Start by selecting a cohort.</p>
      <div class="table-wrap">
        <table>
          <thead><tr><th>Cohort</th><th>Databases</th><th>Instances</th><th>Hosts</th></tr></thead>
          <tbody>
            ${[...byCohort.values()]
              .sort((a, b) => a.cohort.localeCompare(b.cohort))
              .map(
                (x) =>
                  `<tr class="clickable" data-cohort-key="${encodeKey(x.cohort)}"><td>${esc(x.cohort)}</td><td>${fmt(x.dbs.size, 'int')}</td><td>${fmt(
                    x.instances,
                    'int',
                  )}</td><td>${fmt(x.hosts.size, 'int')}</td></tr>`,
              )
              .join('')}
          </tbody>
        </table>
      </div>
    `;

    panel.querySelectorAll('tr[data-cohort-key]').forEach((tr) => {
      tr.addEventListener('click', () => {
        state.drill.cohort = decodeKey(tr.dataset.cohortKey);
        state.drill.db = null;
        state.drill.instance = null;
        renderDrilldown();
      });
    });
  } else if (!state.drill.db) {
    const byDb = new Map();
    cohortRows.forEach((r) => {
      if (!byDb.has(r.db)) byDb.set(r.db, { db: r.db, instances: 0, hosts: new Set(), mem: 0, vcpu: 0, sga: 0, pga: 0 });
      const g = byDb.get(r.db);
      g.instances += 1;
      g.hosts.add(r.host);
      g.mem += r.mem_gb;
      g.vcpu += r.init_cpu_count || r.logical_cpu_count || 0;
      g.sga += r.sga_size_gb;
      g.pga += r.pga_size_gb;
    });

    panel.innerHTML = `
      <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:10px;">${crumbs}</div>
      <p class="muted">Now select a database in cohort ${esc(state.drill.cohort)}.</p>
      <div class="table-wrap">
      <table>
        <thead><tr><th>Database</th><th>Instances</th><th>Hosts</th><th>vCPU</th><th>Memory (GB)</th><th>SGA (GB)</th><th>PGA (GB)</th></tr></thead>
        <tbody>
          ${[...byDb.values()]
            .sort((a, b) => a.db.localeCompare(b.db))
            .map(
                (x) =>
                `<tr class="clickable" data-db-key="${encodeKey(x.db)}"><td>${esc(displayDbName(x.db))}</td><td>${fmt(x.instances, 'int')}</td><td>${fmt(
                  x.hosts.size,
                  'int',
                )}</td><td>${fmt(x.vcpu, 'dec1')}</td><td>${fmt(x.mem, 'dec1')}</td><td>${fmt(x.sga, 'dec1')}</td><td>${fmt(
                  x.pga,
                  'dec1',
                )}</td></tr>`,
            )
            .join('')}
        </tbody>
      </table>
      </div>
    `;

    panel.querySelectorAll('tr[data-db-key]').forEach((tr) => {
      tr.addEventListener('click', () => {
        state.drill.db = decodeKey(tr.dataset.dbKey);
        state.drill.instance = null;
        renderDrilldown();
      });
    });
  } else if (!state.drill.instance) {
    panel.innerHTML = `
      <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:10px;">${crumbs}</div>
      <p class="muted">Select an instance in database ${esc(displayDbName(state.drill.db))}.</p>
      <div class="table-wrap">
        <table>
          <thead><tr><th>Instance</th><th>Host</th><th>vCPU</th><th>Logical CPU</th><th>Memory (GB)</th><th>SGA (GB)</th><th>PGA (GB)</th></tr></thead>
          <tbody>
            ${dbRows
              .sort((a, b) => a.instance.localeCompare(b.instance))
              .map(
                (x) =>
                  `<tr class="clickable" data-instance-key="${encodeKey(x.instance)}"><td>${esc(displayInstanceName(x.instance))}</td><td>${esc(x.host)}</td><td>${fmt(
                    x.init_cpu_count,
                    'dec1',
                  )}</td><td>${fmt(x.logical_cpu_count, 'int')}</td><td>${fmt(x.mem_gb, 'dec1')}</td><td>${fmt(
                    x.sga_size_gb,
                    'dec1',
                  )}</td><td>${fmt(x.pga_size_gb, 'dec1')}</td></tr>`,
              )
              .join('')}
          </tbody>
        </table>
      </div>
    `;

    panel.querySelectorAll('tr[data-instance-key]').forEach((tr) => {
      tr.addEventListener('click', () => {
        state.drill.instance = decodeKey(tr.dataset.instanceKey);
        renderDrilldown();
      });
    });
  } else {
    const item = dbRows.find((r) => r.instance === state.drill.instance);
    const dbSql = state.graph.topSql
      .filter((r) => r.db === state.drill.db)
      .sort((a, b) => b.elapsed - a.elapsed)
      .slice(0, 8);

    panel.innerHTML = `
      <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:10px;">${crumbs}</div>
      <p class="muted">Instance details.</p>
      <div class="kpi-grid" style="margin-bottom:10px;">
        <div class="kpi"><div class="name">Instance</div><div class="value" style="font-size:16px;">${esc(displayInstanceName(item?.instance || ''))}</div></div>
        <div class="kpi"><div class="name">Host</div><div class="value" style="font-size:16px;">${esc(item?.host || '')}</div></div>
        <div class="kpi"><div class="name">vCPU</div><div class="value">${fmt(item?.init_cpu_count || 0, 'dec1')}</div></div>
        <div class="kpi"><div class="name">Memory (GB)</div><div class="value">${fmt(item?.mem_gb || 0, 'dec1')}</div></div>
        <div class="kpi"><div class="name">SGA (GB)</div><div class="value">${fmt(item?.sga_size_gb || 0, 'dec1')}</div></div>
        <div class="kpi"><div class="name">PGA (GB)</div><div class="value">${fmt(item?.pga_size_gb || 0, 'dec1')}</div></div>
      </div>
      <p class="muted">Top SQL for database ${esc(displayDbName(state.drill.db))}</p>
      <div class="table-wrap">
        <table>
          <thead><tr><th>SQL ID</th><th>Elapsed</th><th>Execs</th><th>Logical Reads</th><th>Physical Read GB</th></tr></thead>
          <tbody>
            ${dbSql
              .map(
                (r) =>
                  `<tr><td>${esc(r.sql_id)}</td><td>${fmt(r.elapsed, 'dec1')}</td><td>${fmt(r.execs, 'dec1')}</td><td>${fmt(
                    r.log_reads,
                    'dec1',
                  )}</td><td>${fmt(r.phy_read_gb, 'dec1')}</td></tr>`,
              )
              .join('')}
          </tbody>
        </table>
      </div>
    `;
  }

  panel.querySelectorAll('button[data-drill-level]').forEach((btn) => {
    btn.addEventListener('click', () => {
      const level = btn.dataset.drillLevel;
      if (level === 'global') state.drill = { cohort: null, db: null, instance: null };
      if (level === 'cohort') {
        state.drill.db = null;
        state.drill.instance = null;
      }
      if (level === 'db') {
        state.drill.instance = null;
      }
      renderDrilldown();
    });
  });
}

function updateFromMultiSelect(selectEl, setRef) {
  setRef.clear();
  [...selectEl.selectedOptions].forEach((o) => setRef.add(o.value));
}

function applyScopeCheckboxChange(kind, value, checked) {
  if (!state.graph) return;
  const sel = activeSelection();
  const setRef = kind === 'cohort' ? sel.cohorts : kind === 'db' ? sel.dbs : sel.instances;
  if (!setRef) return;
  if (checked) setRef.add(value);
  else setRef.delete(value);
  ensureValidSelection('web');
  mirrorIfLinked('web');
  renderAll();
}

function mirrorIfLinked(originKind) {
  const targetKind = originKind === 'web' ? 'ppt' : 'web';
  state.selection[targetKind] = {
    cohorts: new Set(state.selection[originKind].cohorts),
    dbs: new Set(state.selection[originKind].dbs),
    instances: new Set(state.selection[originKind].instances),
    metrics: new Set(state.selection[originKind].metrics),
  };
}

function renderAll() {
  if (!state.graph) {
    renderAiChatAssistant();
    return;
  }
  state.activeTab = 'web';
  state.linkSelections = true;
  mirrorIfLinked('web');
  ensureValidSelection('web');
  ensureValidSelection('ppt');
  renderSelectors();
  if (state.page.mode === 'global') {
    renderGlobal();
    renderDrilldown();
  }
  renderPageMode();
  renderCpuScaleToggleLabel();
  renderPptStoryboard();
  renderAiChatAssistant();
  syncResizableCharts();
  syncChartTypeControls();
}

function buildAiPromptText() {
  const docs = (state.aiChat?.docs || []).map((d) => `### ${d.name}\n${truncateText(d.content, 5000)}`).join('\n\n');
  const question = (state.aiChat?.question || '').trim();

  if (!state.graph) {
    return [
      'You are assisting with Oracle AWR analysis.',
      'No AWR JSON is loaded yet. Base your answer only on the user question and additional docs.',
      '',
      '## User Question',
      question || 'Provide a concise analysis based on the loaded documentation.',
      '',
      '## Additional Documentation',
      docs || 'No extra documents loaded.',
      '',
      '## Response Format',
      '1) Key findings (quantitative where possible)',
      '2) Risks / anomalies',
      '3) Suggested next checks',
    ].join('\n');
  }

  const kind = state.activeTab === 'ppt' ? 'ppt' : 'web';
  const sel = state.selection[kind];
  const rows = selectedInstances(kind);
  const global = computeGlobal(kind);
  const byCohort = computeByCohort(kind).map((r) => ({
    cohort: r.cohort,
    dbs: r.dbs,
    instances: r.instances,
    hosts: r.hosts,
    vcpu: Number((r.vcpu || 0).toFixed(3)),
    memory_gb: Number((r.mem || 0).toFixed(3)),
  }));

  const context = {
    scope: {
      mode: kind,
      selected_cohorts: [...sel.cohorts],
      selected_dbs: [...sel.dbs],
      selected_instances: [...sel.instances],
      selected_metrics: [...sel.metrics],
      selected_rows_count: rows.length,
    },
    summary: global,
    cohorts: byCohort,
  };

  return [
    'You are assisting with Oracle AWR analysis.',
    'Use the context below and answer with factual, quantitative points first.',
    '',
    '## User Question',
    question || 'Provide a concise analysis of the selected scope.',
    '',
    '## AWR Context (from current app selection)',
    '```json',
    escapeForPromptJson(context),
    '```',
    '',
    '## Additional Documentation',
    docs || 'No extra documents loaded.',
    '',
    '## Response Format',
    '1) Key findings (quantitative)',
    '2) Risks / anomalies',
    '3) Suggested next checks',
  ].join('\n');
}

function renderAiChatAssistant() {
  const panel = $('aiPanel');
  const floatBtn = $('aiFloatBtn');
  const statusNode = $('aiConnectionStatus');
  const payloadStatusNode = $('aiPayloadStatus');
  const sendBtn = $('aiSendBtn');
  const msgNode = $('aiMessages');
  const docsNode = $('aiDocsList');
  const qNode = $('aiUserQuestion');
  const pNode = $('aiPromptOutput');
  const rNode = $('aiChatResponse');
  if (!docsNode || !qNode || !pNode || !rNode || !panel || !floatBtn || !statusNode || !msgNode) return;
  if (!CHATBOT_ENABLED) {
    panel.classList.remove('open');
    panel.setAttribute('aria-hidden', 'true');
    panel.style.display = 'none';
    floatBtn.style.display = 'none';
    return;
  }

  panel.classList.toggle('open', Boolean(state.aiChat.panelOpen));
  panel.setAttribute('aria-hidden', state.aiChat.panelOpen ? 'false' : 'true');
  floatBtn.textContent = state.aiChat.panelOpen ? 'Close AI' : 'AI Assistant';
  if (state.aiChat.connected) {
    statusNode.textContent = `Connected (${state.aiChat.connectionSource}) • Model: ${state.aiChat.model}`;
  } else if (state.aiChat.connectionError) {
    statusNode.textContent = `Not connected: ${state.aiChat.connectionError} Use "Send to AWR DBA Advisor".`;
  } else if (state.aiChat.connectionSource === 'unavailable') {
    statusNode.textContent = 'Connection check failed. Ensure app_server.py is running, or use "Send to AWR DBA Advisor".';
  } else {
    statusNode.textContent = 'Not connected. Will try loading OPENAI_API_KEY from ~/.codex/auth.json or .env, or use "Send to AWR DBA Advisor".';
  }
  if (sendBtn) sendBtn.disabled = Boolean(state.aiChat.busy);

  docsNode.innerHTML = (state.aiChat.docs || []).length
    ? state.aiChat.docs
        .map(
          (d) =>
            `<div class="ai-doc-item"><span>${esc(d.name)}</span><span>${fmt((d.content || '').length, 'int')} chars</span></div>`,
        )
        .join('')
    : '<p class="muted">No extra docs loaded.</p>';

  msgNode.innerHTML = (state.aiChat.messages || []).length
    ? state.aiChat.messages
        .slice(-12)
        .map((m) => `<div class="ai-msg ${m.role === 'assistant' ? 'assistant' : 'user'}"><strong>${m.role}:</strong> ${esc(m.content)}</div>`)
        .join('')
    : '<p class="muted">Start a conversation about the loaded AWR data.</p>';

  qNode.value = state.aiChat.question || '';
  pNode.value = state.aiChat.prompt || '';
  rNode.value = state.aiChat.response || '';
  if (payloadStatusNode) {
    if (state.aiChat.appPayload) {
      const infra = state.aiChat.appPayload?.infrastructure_recommendation || {};
      payloadStatusNode.textContent = `APP_PAYLOAD validated. Tier: ${infra.recommended_tier || 'N/A'} | Scale: ${infra.scale_position || 'N/A'}`;
    } else if (state.aiChat.payloadErrors?.length) {
      payloadStatusNode.textContent = `APP_PAYLOAD invalid: ${state.aiChat.payloadErrors[0]}`;
    } else {
      payloadStatusNode.textContent = 'No APP_PAYLOAD validated yet.';
    }
  }
}

function initEvents() {
  setupSortableHeaders();
  setupChartInteractions();
  setupChartResize();
  setupCardResize();
  setupLayoutDnD();
  loadChartSizes();
  loadChartTypes();
  loadCardSpans();
  loadCardHeights();
  loadSidebarCollapsedPref();
  syncCardSizeControls();
  syncCardResizeHandles();
  applySavedCardSpans();
  applySavedCardHeights();
  syncResizableCharts();
  syncChartTypeControls();

  state.layoutDefaultOrder = getCurrentLayoutOrder();
  loadSavedLayout();
  setLayoutEditMode(false);
  enableReportActions();

  $('sidebarToggleBtn')?.addEventListener('click', () => {
    state.sidebarCollapsed = !state.sidebarCollapsed;
    applySidebarCollapsedUi();
    saveSidebarCollapsedPref();
  });

  $('layoutEditBtn')?.addEventListener('click', () => {
    setLayoutEditMode(!state.layoutEditMode);
    if (state.layoutEditMode) setLayoutStatus('Use Up/Down, Width, and Height controls, then click "Save View".');
  });

  $('layoutSaveBtn')?.addEventListener('click', () => {
    saveCurrentLayout();
    setLayoutEditMode(false);
  });

  $('layoutResetBtn')?.addEventListener('click', () => {
    resetLayout();
    state.cardSpans = {};
    saveCardSpans();
    applySavedCardSpans();
    state.cardHeights = {};
    saveCardHeights();
    applySavedCardHeights();
    setLayoutEditMode(false);
  });

  document.body.addEventListener('click', (e) => {
    const typeBtn = e.target.closest('.chart-type-btn');
    if (typeBtn && state.layoutEditMode) {
      const chartId = typeBtn.dataset.chartId;
      const chartType = typeBtn.dataset.chartType;
      if (!chartId || !chartType) return;
      setChartType(chartId, chartType);
      renderAll();
      setLayoutStatus('Chart type updated.');
      return;
    }

    const moveBtn = e.target.closest('.card-move-btn');
    if (moveBtn && state.layoutEditMode) {
      const cardId = moveBtn.dataset.cardId;
      const card = cardId ? document.getElementById(cardId) : null;
      if (!card) return;
      const offset = Number(moveBtn.dataset.move) || 0;
      moveCardByOffset(card, offset);
      setLayoutStatus('Section order updated.');
      return;
    }

    const btn = e.target.closest('.card-size-btn');
    if (!btn || !state.layoutEditMode) return;
    const cardId = btn.dataset.cardId;
    const card = cardId ? document.getElementById(cardId) : null;
    if (!card) return;
    setCardSpan(card, btn.dataset.span);
    updateCardSpanButtons(card);
    setLayoutStatus('Section width updated.');
  });

  document.body.addEventListener('click', (e) => {
    const btn = e.target.closest('.card-height-btn');
    if (!btn || !state.layoutEditMode) return;
    const cardId = btn.dataset.cardId;
    const card = cardId ? document.getElementById(cardId) : null;
    if (!card) return;
    setCardHeightSize(card, btn.dataset.height);
    updateCardHeightButtons(card);
    setLayoutStatus('Section height updated.');
  });

  document.body.addEventListener('mousedown', (e) => {
    if (
      !e.target.closest('.card-size-btn') &&
      !e.target.closest('.card-height-btn') &&
      !e.target.closest('.card-move-btn') &&
      !e.target.closest('.chart-type-btn')
    )
      return;
    e.stopPropagation();
  });

  $('pptSlidesList')?.addEventListener('change', (e) => {
    const id = e.target?.dataset?.pptSlideId;
    if (!id) return;
    if (e.target.checked) state.pptSlideSelected.add(id);
    else state.pptSlideSelected.delete(id);
    renderPptStoryboard();
  });

  $('pptSlidesAll')?.addEventListener('click', () => {
    state.pptSlideSelected = new Set((state.pptSlidePlan || []).map((s) => s.id));
    renderPptStoryboard();
  });

  $('pptSlidesNone')?.addEventListener('click', () => {
    state.pptSlideSelected = new Set();
    renderPptStoryboard();
  });

  $('exportBtn')?.addEventListener('click', async () => {
    try {
      const templateOk = await probeTemplateApi();
      if (templateOk) {
        await exportPptViaTemplateApi();
        return;
      }
      if (TEMPLATE_EXPORT_ONLY) {
        throw new Error(
          'Template export server unavailable. Start app_server.py (with Oracle .potx) and reload this page.',
        );
      }
      await exportPptFromStoryboard();
    } catch (err) {
      setStatus(`PPT export failed: ${err?.message || err}`);
    }
  });

  $('previewBackBtn')?.addEventListener('click', () => {
    goToGlobal();
  });

  $('previewPrevBtn')?.addEventListener('click', () => {
    movePreviewSelection(-1);
  });

  $('previewNextBtn')?.addEventListener('click', () => {
    movePreviewSelection(1);
  });

  $('previewExportBtn')?.addEventListener('click', async () => {
    try {
      const templateOk = await probeTemplateApi();
      if (templateOk) {
        await exportPptViaTemplateApi();
        return;
      }
      if (TEMPLATE_EXPORT_ONLY) {
        throw new Error(
          'Template export server unavailable. Start app_server.py (with Oracle .potx) and reload this page.',
        );
      }
      await exportPptFromStoryboard();
    } catch (err) {
      setStatus(`PPT export failed: ${err?.message || err}`);
    }
  });

  $('fileInput').addEventListener('change', async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const text = await file.text();
      const raw = JSON.parse(text);
      if (!Array.isArray(raw.instances)) {
        throw new Error('Invalid AWR JSON: missing instances array');
      }
      setDefaultsForLoadedRaw(raw);
      enableReportActions();

      setStatus(
        `Loaded ${file.name} | ${state.graph.cohorts.length} cohorts | ${state.graph.dbs.length} DBs | ${state.graph.instances.length} instances | ${state.metricCatalog.length} metrics | ${Object.keys(state.graph.cpuSeriesByCohort || {}).length} CPU timelines`,
      );
      renderAll();
    } catch (err) {
      setStatus(`Load failed: ${err.message}`);
    }
  });

  $('templateInput')?.addEventListener('change', async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      setStatus(`Uploading template ${file.name}...`);
      await uploadTemplateToApi(file);
    } catch (err) {
      setStatus(`Template load failed: ${err?.message || err}`);
    } finally {
      e.target.value = '';
    }
  });

  $('referenceRefreshBtn')?.addEventListener('click', async () => {
    await loadReferenceLibrary();
    setStatus('Reference library refreshed.');
  });

  $('saveReportBtn')?.addEventListener('click', () => {
    if (!state.raw) {
      setStatus('No data loaded. Load JSON first.');
      return;
    }
    const payload = reportStateSnapshot();
    const fileName = downloadJsonFile(payload, 'awr_review_report');
    setStatus(`Report saved (${fileName}).`);
  });

  const advisorProfileSelect = $('advisorExportProfile');
  if (advisorProfileSelect) {
    advisorProfileSelect.value = state.advisorExportProfile === 'comprehensive' ? 'comprehensive' : 'compact';
    advisorProfileSelect.addEventListener('change', (e) => {
      state.advisorExportProfile = e.target.value === 'comprehensive' ? 'comprehensive' : 'compact';
    });
  }

  $('exportAdvisorJsonBtn')?.addEventListener('click', () => {
    if (!state.raw || !state.graph) {
      setStatus('No data loaded. Load JSON first.');
      return;
    }
    const profile = state.advisorExportProfile === 'comprehensive' ? 'comprehensive' : 'compact';
    const payloads = buildAdvisorExportPayloads('web', 1, profile);
    if (!payloads.length) {
      setStatus('Could not build Advisor export payload.');
      return;
    }
    const files = payloads.map((payload, idx) => {
      const cohortLabel = (payload?.scope?.selected_cohorts || [])[0] || `part${String(idx + 1).padStart(2, '0')}`;
      const safeCohort = String(cohortLabel).replace(/[^a-zA-Z0-9_-]+/g, '_');
      return downloadJsonFile(payload, `awr_advisor_context_${safeCohort}`, false);
    });
    setStatus(
      `Advisor JSON exported in ${files.length} file(s), profile ${profile}, split: 1 cohort per file. Upload each part to ChatGPT AWR DBA Advisor.`,
    );
  });

  if (CHATBOT_ENABLED) {
    $('aiFloatBtn')?.addEventListener('click', async () => {
      state.aiChat.panelOpen = !state.aiChat.panelOpen;
      renderAiChatAssistant();
      if (state.aiChat.panelOpen) {
        const ok = await probeAiConnection();
        setStatus(ok ? 'AI assistant connected.' : 'AI assistant not connected. Checking ~/.codex/auth.json and .env for OPENAI_API_KEY.');
      }
    });

    $('aiPanelCloseBtn')?.addEventListener('click', () => {
      state.aiChat.panelOpen = false;
      renderAiChatAssistant();
    });

    $('aiDocsInput')?.addEventListener('change', async (e) => {
      const files = [...(e.target.files || [])];
      if (!files.length) return;
      const accepted = [];
      for (const f of files) {
        const lower = String(f.name || '').toLowerCase();
        if (!lower.match(/\.(txt|md|json|csv|log)$/)) continue;
        try {
          const content = await f.text();
          accepted.push({ name: f.name, content: truncateText(content, 12000) });
        } catch {
          // ignore unreadable file
        }
      }
      if (accepted.length) {
        state.aiChat.docs = [...state.aiChat.docs, ...accepted];
        setStatus(`Loaded ${accepted.length} doc(s) for AI context.`);
        renderAiChatAssistant();
      } else {
        setStatus('No supported docs loaded. Use .txt, .md, .json, .csv, or .log');
      }
      e.target.value = '';
    });

    $('aiClearDocsBtn')?.addEventListener('click', () => {
      state.aiChat.docs = [];
      renderAiChatAssistant();
      setStatus('AI docs cleared.');
    });

    $('aiUserQuestion')?.addEventListener('input', (e) => {
      state.aiChat.question = e.target.value;
    });

    $('aiChatResponse')?.addEventListener('input', (e) => {
      state.aiChat.response = e.target.value;
    });

    $('aiBuildPromptBtn')?.addEventListener('click', () => {
      state.aiChat.prompt = buildAiPromptText();
      renderAiChatAssistant();
      setStatus('AI prompt generated.');
    });

    $('aiSendBtn')?.addEventListener('click', async () => {
      await sendAiMessage();
    });

    $('aiWebModeBtn')?.addEventListener('click', async () => {
      await sendAiViaWeb();
    });

    $('aiCopyPromptBtn')?.addEventListener('click', async () => {
      const prompt = state.aiChat.prompt || buildAiPromptText();
      state.aiChat.prompt = prompt;
      try {
        if (navigator?.clipboard?.writeText) {
          await navigator.clipboard.writeText(prompt);
        } else {
          const ta = document.createElement('textarea');
          ta.value = prompt;
          document.body.appendChild(ta);
          ta.select();
          document.execCommand('copy');
          ta.remove();
        }
        renderAiChatAssistant();
        setStatus('AI prompt copied to clipboard.');
      } catch {
        renderAiChatAssistant();
        setStatus('Could not copy automatically. Select and copy from the prompt box.');
      }
    });

    $('aiValidatePayloadBtn')?.addEventListener('click', () => {
      const raw = stripJsonCodeFence(state.aiChat.response || '');
      if (!raw) {
        state.aiChat.appPayload = null;
        state.aiChat.payloadErrors = ['Paste APP_PAYLOAD JSON first.'];
        renderAiChatAssistant();
        setStatus('APP_PAYLOAD validation failed: empty input.');
        return;
      }
      try {
        const parsed = JSON.parse(raw);
        const check = validateAppPayload(parsed);
        if (!check.valid) {
          state.aiChat.appPayload = null;
          state.aiChat.payloadErrors = check.errors;
          renderAiChatAssistant();
          setStatus(`APP_PAYLOAD invalid (${check.errors.length} issues).`);
          return;
        }
        state.aiChat.appPayload = parsed;
        state.aiChat.payloadErrors = [];
        renderAiChatAssistant();
        setStatus('APP_PAYLOAD validated and attached to report.');
      } catch (err) {
        state.aiChat.appPayload = null;
        state.aiChat.payloadErrors = [`JSON parse error: ${err?.message || err}`];
        renderAiChatAssistant();
        setStatus('APP_PAYLOAD validation failed: invalid JSON.');
      }
    });
  } else {
    $('aiPanel')?.remove();
    $('aiFloatBtn')?.remove();
  }

  $('reportInput')?.addEventListener('change', async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const text = await file.text();
      const parsed = JSON.parse(text);
      if (!parsed || typeof parsed !== 'object') {
        throw new Error('Invalid report format.');
      }
      if (!Array.isArray(parsed.raw_data?.instances)) {
        throw new Error('Invalid report: missing raw_data.instances.');
      }

      setDefaultsForLoadedRaw(parsed.raw_data);
      applySavedUiState(parsed.ui_state || {}, parsed.report_metadata || null);
      enableReportActions();
      renderAll();
      setStatus(
        `Report loaded ${file.name} | ${state.graph.cohorts.length} cohorts | ${state.graph.dbs.length} DBs | ${state.graph.instances.length} instances`,
      );
    } catch (err) {
      setStatus(`Load report failed: ${err.message}`);
    } finally {
      e.target.value = '';
    }
  });

  const metaFields = [
    ['metaCustomerName', 'customerName'],
    ['metaOpportunityNumber', 'opportunityNumber'],
    ['metaSalesRepName', 'salesRepName'],
    ['metaArchitectName', 'architectName'],
    ['metaEngineerName', 'engineerName'],
  ];
  metaFields.forEach(([id, key]) => {
    $(id)?.addEventListener('input', (e) => {
      state.reportMeta[key] = String(e.target.value || '');
    });
  });

  $('cpuScaleToggle')?.addEventListener('click', () => {
    state.cpuScaleMode = state.cpuScaleMode === 'dynamic' ? 'fixed100' : 'dynamic';
    renderAll();
  });

  $('cohortSelect').addEventListener('change', (e) => {
    const kind = e.target?.dataset?.scopeKind;
    if (kind !== 'cohort') return;
    applyScopeCheckboxChange('cohort', e.target.value, Boolean(e.target.checked));
  });

  $('dbSelect').addEventListener('change', (e) => {
    const kind = e.target?.dataset?.scopeKind;
    if (kind !== 'db') return;
    applyScopeCheckboxChange('db', e.target.value, Boolean(e.target.checked));
  });

  $('instanceSelect').addEventListener('change', (e) => {
    const kind = e.target?.dataset?.scopeKind;
    if (kind !== 'instance') return;
    applyScopeCheckboxChange('instance', e.target.value, Boolean(e.target.checked));
  });

  $('metricsAll').addEventListener('click', () => {
    const sel = activeSelection();
    state.metricCatalog.forEach((m) => sel.metrics.add(m.id));
    mirrorIfLinked('web');
    renderAll();
  });

  $('metricsNone').addEventListener('click', () => {
    const sel = activeSelection();
    sel.metrics.clear();
    mirrorIfLinked('web');
    renderAll();
  });

  $('metricsList').addEventListener('change', (e) => {
    const id = e.target?.dataset?.metric;
    if (!id) return;
    const sel = activeSelection();
    if (e.target.checked) sel.metrics.add(id);
    else sel.metrics.delete(id);
    mirrorIfLinked('web');
    renderAll();
  });

  initCollapsibleSections();
}

initEvents();
applyReportMetaToInputs();
renderAiChatAssistant();
void probeTemplateApi();
void loadReferenceLibrary();
