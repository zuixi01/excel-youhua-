<script setup lang="ts">
import { computed, onUnmounted, ref } from "vue";

type JobStatus = { job_id: string; status: string; progress: number; workflow?: string; summary?: Record<string, number>; warnings?: string[]; artifacts?: Record<string, string>; error_message_safe?: string; category_count?: number; unresolved_row_count?: number; review_count?: number; issue_count?: number; revision_number?: number };
type Difference = { difference_id: string; type: string; severity: string; sheet_id: string; sheet_name: string; cell?: string; excel_row?: number; canonical_field?: string; business_key?: Record<string, unknown>; excel_raw_value?: unknown; excel_normalized_value?: unknown; standard_raw_value?: unknown; standard_normalized_value?: unknown; rule_id?: string; message: string; repair_status: string };
type ReviewCandidate = { field_id: string; title: string; confidence: number; match_value: string };
type ProductReview = { review_id: string; review_type: string; review_key: string; status: string; payload: { message: string; raw_header?: string; physical_column?: number; excel_row?: number; candidates: ReviewCandidate[] }; decision?: Record<string, unknown> };
type ProductIssue = { issue_type: string; excel_row: number; category_id?: string; field_id?: string; raw_value?: unknown; message: string; color: string };

const tab = ref<"tasks" | "products" | "rules">("products");
const schemaId = ref("employee-roster");
const schemaVersion = ref("1.0.0");
const productSchemaId = ref("product-normalization");
const productSchemaVersion = ref("1.0.0");
const excel = ref<File>();
const standard = ref<File>();
const managedSource = ref(false);
const job = ref<JobStatus>();
const differences = ref<Difference[]>([]);
const differenceTotal = ref(0);
const filters = ref({ type: "", sheet_id: "", canonical_field: "", severity: "" });
const error = ref("");
const busy = ref(false);
const apiToken = ref(sessionStorage.getItem("excel-auditor-api-token") || "");
const authMessage = ref("");
const precheckResult = ref<Record<string, unknown>>();
const productReviews = ref<ProductReview[]>([]);
const productIssues = ref<ProductIssue[]>([]);
const productIssueTotal = ref(0);
const productIssueFilters = ref({ issue_type: "", category_id: "", field_id: "" });
const reviewBusy = ref("");
let timer: number | undefined;

const versions = ref<Array<{ version: string; config_sha256: string }>>([]);
const ruleJson = ref("");
const editorMode = ref<"visual" | "json">("visual");
const visualRule = ref<any>({ schema_id: "", schema_version: "1.0.0", name: "", workbook: {}, sheets: [] });
const draftId = ref("");
const ruleMessage = ref("");
const mapping = ref({ sheet_id: "", raw_header: "", canonical_field: "" });

const finished = computed(() => ["completed", "failed", "manual_review", "cancelled"].includes(job.value?.status ?? ""));
const allProductReviewsResolved = computed(() => productReviews.value.length > 0 && productReviews.value.every(item => item.status === "resolved"));

function choose(event: Event, target: "excel" | "standard") {
  const file = (event.target as HTMLInputElement).files?.[0];
  if (target === "excel") excel.value = file; else standard.value = file;
}

async function problem(response: Response) {
  try { const body = await response.json(); return body.detail || body.title || "请求失败"; }
  catch { return (await response.text()) || "请求失败"; }
}

function saveApiToken() {
  apiToken.value = apiToken.value.trim();
  if (apiToken.value) {
    sessionStorage.setItem("excel-auditor-api-token", apiToken.value);
    authMessage.value = "访问令牌已保存到当前浏览器会话。";
  } else {
    sessionStorage.removeItem("excel-auditor-api-token");
    authMessage.value = "访问令牌已清除。";
  }
}

function apiFetch(input: RequestInfo | URL, init: RequestInit = {}) {
  const headers = new Headers(init.headers);
  if (apiToken.value) headers.set("Authorization", `Bearer ${apiToken.value}`);
  return fetch(input, { ...init, headers });
}

async function submit() {
  error.value = "";
  if (!excel.value || (!managedSource.value && !standard.value)) { error.value = "请选择待核验 Excel，并为上传模式选择标准数据文件。"; return; }
  busy.value = true;
  const body = new FormData();
  body.append("excel_file", excel.value);
  if (standard.value && !managedSource.value) body.append("standard_data", standard.value);
  body.append("schema_id", schemaId.value); body.append("schema_version", schemaVersion.value);
  try {
    const response = await apiFetch("/api/v1/comparisons", { method: "POST", headers: { "Idempotency-Key": crypto.randomUUID() }, body });
    if (!response.ok) throw new Error(await problem(response));
    job.value = await response.json(); differences.value = [];
    timer = window.setInterval(refresh, 1000); await refresh();
  } catch (caught) { error.value = caught instanceof Error ? caught.message : "任务创建失败"; busy.value = false; }
}

async function submitProduct() {
  error.value = "";
  if (!excel.value) { error.value = "请选择需要整理的商品 Excel。"; return; }
  busy.value = true; productReviews.value = [];
  const body = new FormData();
  body.append("excel_file", excel.value);
  body.append("schema_id", productSchemaId.value); body.append("schema_version", productSchemaVersion.value);
  try {
    const response = await apiFetch("/api/v1/product-normalizations", { method: "POST", headers: { "Idempotency-Key": crypto.randomUUID() }, body });
    if (!response.ok) throw new Error(await problem(response));
    job.value = await response.json();
    timer = window.setInterval(refresh, 1000); await refresh();
  } catch (caught) { error.value = caught instanceof Error ? caught.message : "商品整理任务创建失败"; busy.value = false; }
}

async function precheck(productWorkflow = false) {
  error.value = ""; precheckResult.value = undefined;
  if (!excel.value) { error.value = "请先选择待核验 Excel。"; return; }
  const body = new FormData();
  body.append("excel_file", excel.value);
  body.append("schema_id", productWorkflow ? productSchemaId.value : schemaId.value);
  body.append("schema_version", productWorkflow ? productSchemaVersion.value : schemaVersion.value);
  const response = await apiFetch("/api/v1/workbooks/precheck", { method: "POST", body });
  if (!response.ok) { error.value = await problem(response); return; }
  precheckResult.value = await response.json();
}

async function refresh() {
  if (!job.value) return;
  const response = await apiFetch(`/api/v1/comparisons/${job.value.job_id}`);
  if (!response.ok) return;
  const current: JobStatus = await response.json();
  job.value = current;
  if (finished.value) {
    window.clearInterval(timer); busy.value = false;
    if (current.workflow === "product_normalization") { await loadProductReviews(); await loadProductIssues(); }
    else if (["completed", "manual_review"].includes(current.status)) await loadDifferences();
  }
}

async function loadProductReviews() {
  if (!job.value) return;
  const response = await apiFetch(`/api/v1/product-normalizations/${job.value.job_id}/reviews`);
  if (!response.ok) { error.value = await problem(response); return; }
  productReviews.value = (await response.json()).items;
}

async function loadProductIssues() {
  if (!job.value) return;
  const query = new URLSearchParams({ page: "1", page_size: "200" });
  Object.entries(productIssueFilters.value).forEach(([key, value]) => { if (value) query.set(key, value); });
  const response = await apiFetch(`/api/v1/product-normalizations/${job.value.job_id}/issues?${query}`);
  if (!response.ok) { error.value = await problem(response); return; }
  const payload = await response.json(); productIssues.value = payload.items; productIssueTotal.value = payload.total;
}

async function decideProductReview(review: ProductReview, action: "confirm" | "keep_extra" | "reject", candidate?: ReviewCandidate) {
  if (!job.value) return;
  reviewBusy.value = review.review_id; error.value = "";
  let decision: Record<string, unknown>;
  if (action === "confirm" && review.review_type === "category") decision = { action: "confirm_category", category_id: candidate?.field_id };
  else if (action === "confirm") decision = { action: "confirm_mapping", field_id: candidate?.field_id, raw_header: review.payload.raw_header };
  else decision = { action };
  const response = await apiFetch(`/api/v1/product-normalizations/${job.value.job_id}/reviews/${review.review_id}/decision`, {
    method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(decision),
  });
  reviewBusy.value = "";
  if (!response.ok) { error.value = await problem(response); return; }
  await loadProductReviews(); await refresh();
}

async function createProductRevision() {
  if (!job.value) return;
  busy.value = true; error.value = "";
  const response = await apiFetch(`/api/v1/product-normalizations/${job.value.job_id}/revisions`, { method: "POST" });
  if (!response.ok) { error.value = await problem(response); busy.value = false; return; }
  job.value = await response.json();
  timer = window.setInterval(refresh, 1000); await refresh();
}

async function cancelJob() {
  if (!job.value) return;
  const response = await apiFetch(`/api/v1/comparisons/${job.value.job_id}/cancel`, { method: "POST" });
  if (!response.ok) { error.value = await problem(response); return; }
  await refresh();
}

async function loadDifferences() {
  if (!job.value) return;
  const query = new URLSearchParams({ page: "1", page_size: "200" });
  Object.entries(filters.value).forEach(([key, value]) => { if (value) query.set(key, value); });
  const response = await apiFetch(`/api/v1/comparisons/${job.value.job_id}/differences?${query}`);
  if (!response.ok) { error.value = await problem(response); return; }
  const payload = await response.json(); differences.value = payload.items; differenceTotal.value = payload.total;
}

async function downloadArtifact(kind: string) {
  if (!job.value) return;
  error.value = "";
  const response = await apiFetch(`/api/v1/comparisons/${job.value.job_id}/artifacts/${kind}`);
  if (!response.ok) { error.value = await problem(response); return; }
  const extensions: Record<string, string> = { excel: "xlsx", json: "json", differences_jsonl: "jsonl", html: "html", manifest: "json", product_excel: "xlsx", product_result: "json", product_manifest: "json", product_issues: "jsonl" };
  const fallback = `${job.value.job_id}-${kind}.${extensions[kind] || "bin"}`;
  const disposition = response.headers.get("content-disposition") || "";
  const encoded = disposition.match(/filename\*=UTF-8''([^;]+)/i)?.[1];
  const quoted = disposition.match(/filename="([^"]+)"/i)?.[1];
  let filename = fallback;
  try { filename = encoded ? decodeURIComponent(encoded) : (quoted || fallback); } catch { filename = fallback; }
  const url = URL.createObjectURL(await response.blob());
  const link = document.createElement("a");
  link.href = url; link.download = filename; document.body.appendChild(link); link.click(); link.remove();
  URL.revokeObjectURL(url);
}

async function loadVersions() {
  ruleMessage.value = "";
  const response = await apiFetch(`/api/v1/schemas/${encodeURIComponent(schemaId.value)}/versions`);
  if (!response.ok) { ruleMessage.value = await problem(response); return; }
  versions.value = (await response.json()).items;
}

async function loadRule() {
  const response = await apiFetch(`/api/v1/schemas/${encodeURIComponent(schemaId.value)}/versions/${encodeURIComponent(schemaVersion.value)}`);
  if (!response.ok) { ruleMessage.value = await problem(response); return; }
  const loaded = await response.json(); ruleJson.value = JSON.stringify(loaded, null, 2); visualRule.value = loaded; draftId.value = ""; ruleMessage.value = "已加载发布版本；保存时会创建新草稿。";
}

function parsedRule() { try { return JSON.parse(ruleJson.value); } catch { throw new Error("规则必须是合法 JSON。Decimal 等精确值请使用字符串。") } }

function syncVisualFromJson() { visualRule.value = parsedRule(); visualRule.value.workbook ||= {}; visualRule.value.sheets ||= []; }
function syncJsonFromVisual() { ruleJson.value = JSON.stringify(visualRule.value, null, 2); schemaId.value = visualRule.value.schema_id || schemaId.value; schemaVersion.value = visualRule.value.schema_version || schemaVersion.value; }
function addSheet() { visualRule.value.sheets.push({ id: `sheet_${visualRule.value.sheets.length + 1}`, name: "", primary_key: [], columns: [] }); }
function addColumn(sheet: any) { sheet.columns ||= []; sheet.columns.push({ name: `field_${sheet.columns.length + 1}`, title: "", type: "string", required: false }); }
function updatePrimaryKey(sheet: any, event: Event) { sheet.primary_key = (event.target as HTMLInputElement).value.split(",").map(value => value.trim()).filter(Boolean); }

async function saveDraft() {
  try {
    if (editorMode.value === "visual") syncJsonFromVisual();
    const config = parsedRule();
    const url = draftId.value ? `/api/v1/schemas/${schemaId.value}/drafts/${draftId.value}` : `/api/v1/schemas/${schemaId.value}/drafts`;
    const response = await apiFetch(url, { method: draftId.value ? "PUT" : "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(draftId.value ? config : { config }) });
    if (!response.ok) throw new Error(await problem(response));
    const payload = await response.json(); draftId.value = payload.draft_id; ruleMessage.value = `草稿已保存：${draftId.value}`;
  } catch (caught) { ruleMessage.value = caught instanceof Error ? caught.message : "保存失败"; }
}

async function validateDraft() {
  if (!draftId.value) { ruleMessage.value = "请先保存草稿。"; return; }
  const response = await apiFetch(`/api/v1/schemas/${schemaId.value}/drafts/${draftId.value}/validate`, { method: "POST" });
  const payload = await response.json(); ruleMessage.value = payload.valid ? `规则有效，SHA-256：${payload.config_sha256}` : JSON.stringify(payload.errors, null, 2);
}

async function publishDraft() {
  if (!draftId.value) { ruleMessage.value = "请先保存并校验草稿。"; return; }
  const response = await apiFetch(`/api/v1/schemas/${schemaId.value}/drafts/${draftId.value}/publish`, { method: "POST" });
  if (!response.ok) { ruleMessage.value = await problem(response); return; }
  const payload = await response.json(); schemaVersion.value = payload.schema_version; ruleMessage.value = `已发布不可变版本 ${payload.schema_version}`; await loadVersions();
}

async function confirmMapping() {
  if (!draftId.value) { ruleMessage.value = "人工确认只能写入未发布草稿。"; return; }
  const response = await apiFetch("/api/v1/mappings/confirm", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ schema_id: schemaId.value, draft_id: draftId.value, ...mapping.value }) });
  ruleMessage.value = response.ok ? `已确认别名：${mapping.value.raw_header} → ${mapping.value.canonical_field}` : await problem(response);
}

onUnmounted(() => window.clearInterval(timer));
</script>

<template>
  <main>
    <header><p class="eyebrow">DATA QUALITY WORKBENCH</p><h1>Excel 标准核验平台</h1><p>基于不可变规则版本和标准数据快照，生成可追溯的标色工作簿与差异报告。</p></header>
    <section class="panel auth-panel"><div class="field"><label for="api-token">API Bearer 令牌（仅保存在当前浏览器会话）</label><input id="api-token" v-model="apiToken" type="password" autocomplete="off" @keyup.enter="saveApiToken" /></div><button class="secondary" @click="saveApiToken">保存/清除令牌</button><span v-if="authMessage" class="muted">{{ authMessage }}</span></section>
    <div class="tabs"><button :class="{ active: tab === 'products' }" @click="tab = 'products'">商品表整理</button><button :class="{ active: tab === 'tasks' }" @click="tab = 'tasks'">标准核验</button><button :class="{ active: tab === 'rules' }" @click="tab = 'rules'">规则管理</button></div>

    <template v-if="tab === 'tasks'">
      <section class="panel form-panel">
        <div class="field"><label for="schema">规则 ID</label><input id="schema" v-model="schemaId" /></div>
        <div class="field"><label for="version">规则版本</label><input id="version" v-model="schemaVersion" /></div>
        <label class="drop"><span>待核验 Excel</span><strong>{{ excel?.name || "选择 .xlsx / .xlsm 文件" }}</strong><input type="file" accept=".xlsx,.xlsm" @change="choose($event, 'excel')" /></label>
        <label class="drop" :class="{ disabled: managedSource }"><span>标准数据</span><strong>{{ managedSource ? "由已配置的受管 HTTP 连接获取" : (standard?.name || "选择 JSON 或 CSV 文件") }}</strong><input type="file" accept=".json,.csv" :disabled="managedSource" @change="choose($event, 'standard')" /></label>
        <label class="toggle"><input v-model="managedSource" type="checkbox" /> 使用规则中的受管 HTTP 标准源</label>
        <div class="actions"><button class="secondary" :disabled="busy" @click="precheck(false)">工作簿预检查</button><button class="primary" :disabled="busy" @click="submit">{{ busy ? "正在核验…" : "创建核验任务" }}</button></div>
        <p v-if="error" class="error">{{ error }}</p>
        <pre v-if="precheckResult" class="message">{{ JSON.stringify(precheckResult, null, 2) }}</pre>
      </section>
      <section v-if="job" class="panel result">
        <div class="status"><div><span>任务 {{ job.job_id }}</span><h2>{{ job.status }}</h2></div><strong>{{ job.progress }}%</strong></div>
        <div class="bar"><i :style="{ width: `${job.progress}%` }" /></div>
        <button v-if="!finished" class="secondary danger" @click="cancelJob">请求取消</button>
        <div v-if="job.summary" class="metrics"><article v-for="(value, key) in job.summary" :key="key"><strong>{{ value }}</strong><span>{{ key }}</span></article></div>
        <p v-if="job.error_message_safe" class="error">{{ job.error_message_safe }}</p>
        <ul v-if="job.warnings?.length" class="warnings"><li v-for="warning in job.warnings" :key="warning">{{ warning }}</li></ul>
        <nav v-if="job.artifacts"><button v-if="job.artifacts.excel" @click="downloadArtifact('excel')">下载标色 Excel</button><button @click="downloadArtifact('json')">JSON 报告</button><button v-if="job.artifacts.differences_jsonl" @click="downloadArtifact('differences_jsonl')">差异 JSONL</button><button @click="downloadArtifact('html')">HTML 报告</button></nav>
        <div v-if="job.artifacts?.json" class="differences">
          <div class="filter-grid"><input v-model="filters.type" placeholder="差异类型" /><input v-model="filters.sheet_id" placeholder="工作表 ID" /><input v-model="filters.canonical_field" placeholder="规范字段" /><select v-model="filters.severity"><option value="">全部级别</option><option>error</option><option>warning</option><option>info</option></select><button class="secondary" @click="loadDifferences">筛选</button></div>
          <p class="muted">共 {{ differenceTotal }} 条；当前最多展示 200 条。</p>
          <div class="table-wrap"><table><thead><tr><th>类型</th><th>位置</th><th>字段</th><th>级别</th><th>修复</th><th>说明</th><th>详情</th></tr></thead><tbody><tr v-for="item in differences" :key="item.difference_id"><td>{{ item.type }}</td><td>{{ item.sheet_name || item.sheet_id }} {{ item.cell || (item.excel_row ? `#${item.excel_row}` : "") }}</td><td>{{ item.canonical_field || "—" }}</td><td>{{ item.severity }}</td><td>{{ item.repair_status }}</td><td>{{ item.message }}</td><td><details><summary>查看</summary><dl class="diff-detail"><dt>业务主键</dt><dd>{{ JSON.stringify(item.business_key) }}</dd><dt>Excel 原值</dt><dd>{{ JSON.stringify(item.excel_raw_value) }}</dd><dt>Excel 归一值</dt><dd>{{ JSON.stringify(item.excel_normalized_value) }}</dd><dt>标准原值</dt><dd>{{ JSON.stringify(item.standard_raw_value) }}</dd><dt>标准归一值</dt><dd>{{ JSON.stringify(item.standard_normalized_value) }}</dd><dt>规则</dt><dd>{{ item.rule_id || "—" }}</dd><dt>差异 ID</dt><dd>{{ item.difference_id }}</dd></dl></details></td></tr></tbody></table></div>
        </div>
      </section>
    </template>

    <template v-else-if="tab === 'products'">
      <section class="panel form-panel product-intro">
        <div class="product-heading"><div><p class="eyebrow">PRODUCT NORMALIZATION</p><h2>平台类目驱动的商品表整理</h2><p>固定字段在前，平台属性与规格字段居中，商家自定义字段完整保留在最右侧。模糊匹配必须人工确认。</p></div><ol><li>上传商家商品表</li><li>解析类目与平台字段</li><li>处理待审核项</li><li>下载商品表与 SKU 表</li></ol></div>
        <div class="field"><label for="product-schema">商品规则 ID</label><input id="product-schema" v-model="productSchemaId" /></div>
        <div class="field"><label for="product-version">规则版本</label><input id="product-version" v-model="productSchemaVersion" /></div>
        <label class="drop"><span>商家商品 Excel</span><strong>{{ excel?.name || "选择 .xlsx / .xlsm 文件" }}</strong><input type="file" accept=".xlsx,.xlsm" @change="choose($event, 'excel')" /></label>
        <div class="actions"><button class="secondary" :disabled="busy" @click="precheck(true)">先做安全预检查</button><button class="primary" :disabled="busy" @click="submitProduct">{{ busy ? "正在处理…" : "开始整理商品表" }}</button></div>
        <p v-if="error" class="error">{{ error }}</p>
        <pre v-if="precheckResult" class="message">{{ JSON.stringify(precheckResult, null, 2) }}</pre>
      </section>
      <section v-if="job?.workflow === 'product_normalization'" class="panel result">
        <div class="status"><div><span>任务 {{ job.job_id }} · 修订 {{ job.revision_number || 1 }}</span><h2>{{ job.status }}</h2></div><strong>{{ job.progress || 0 }}%</strong></div>
        <div class="bar"><i :style="{ width: `${job.progress || 0}%` }" /></div>
        <div class="metrics product-metrics"><article><strong>{{ job.category_count || 0 }}</strong><span>已解析类目</span></article><article><strong>{{ job.unresolved_row_count || 0 }}</strong><span>未解析商品</span></article><article><strong>{{ productReviews.filter(item => item.status === 'pending').length }}</strong><span>待人工确认</span></article><article><strong>{{ job.issue_count || 0 }}</strong><span>字段质量问题</span></article></div>
        <p v-if="job.error_message_safe" class="error">{{ job.error_message_safe }}</p>
        <nav v-if="job.artifacts"><button v-if="job.artifacts.product_excel" @click="downloadArtifact('product_excel')">{{ job.status === 'manual_review' ? '下载待审核商品 Excel' : '下载最终商品 Excel' }}</button><button v-if="job.artifacts.product_result" @click="downloadArtifact('product_result')">下载完整结果 JSON</button><button v-if="job.artifacts.product_issues" @click="downloadArtifact('product_issues')">下载问题 JSONL</button><button v-if="job.artifacts.product_manifest" @click="downloadArtifact('product_manifest')">渲染清单</button></nav>
        <div v-if="productIssues.length || job.issue_count" class="differences">
          <div class="review-title"><div><h3>字段质量问题</h3><p class="muted">与 Excel 内“问题清单”一致；共 {{ productIssueTotal }} 条，当前最多展示 200 条。</p></div></div>
          <div class="filter-grid product-issue-filter"><input v-model="productIssueFilters.issue_type" placeholder="问题类型" /><input v-model="productIssueFilters.category_id" placeholder="类目 ID" /><input v-model="productIssueFilters.field_id" placeholder="字段 ID" /><button class="secondary" @click="loadProductIssues">筛选</button></div>
          <div class="table-wrap"><table><thead><tr><th>源行</th><th>类目</th><th>字段</th><th>类型</th><th>原值</th><th>说明</th></tr></thead><tbody><tr v-for="item in productIssues" :key="`${item.excel_row}-${item.category_id}-${item.field_id}-${item.issue_type}`"><td>{{ item.excel_row }}</td><td>{{ item.category_id || '—' }}</td><td>{{ item.field_id || '—' }}</td><td>{{ item.issue_type }}</td><td>{{ JSON.stringify(item.raw_value) }}</td><td>{{ item.message }}</td></tr></tbody></table></div>
        </div>
        <div v-if="productReviews.length" class="review-workbench">
          <div class="review-title"><div><h3>人工审核工作台</h3><p class="muted">候选项只提供依据，不会自动写入。每个决定都会进入修订历史。</p></div><button v-if="allProductReviewsResolved" class="primary" :disabled="busy" @click="createProductRevision">应用决定并生成新修订</button></div>
          <article v-for="review in productReviews" :key="review.review_id" class="review-card" :class="{ resolved: review.status === 'resolved' }">
            <header><div><span class="review-type">{{ review.review_type }}</span><strong>{{ review.payload.message }}</strong></div><span class="review-status">{{ review.status }}</span></header>
            <div v-if="review.status === 'pending'" class="candidate-list">
              <button v-for="candidate in review.payload.candidates" :key="candidate.field_id" :disabled="reviewBusy === review.review_id" @click="decideProductReview(review, 'confirm', candidate)"><strong>{{ candidate.title }}</strong><span>{{ candidate.field_id }} · {{ candidate.confidence.toFixed(1) }}%</span></button>
              <button v-if="review.review_type !== 'category'" class="secondary" :disabled="reviewBusy === review.review_id" @click="decideProductReview(review, 'keep_extra')">保留为商家扩展字段</button>
              <button class="danger" :disabled="reviewBusy === review.review_id" @click="decideProductReview(review, 'reject')">驳回并人工修改源表</button>
            </div>
            <pre v-else class="decision">{{ JSON.stringify(review.decision, null, 2) }}</pre>
          </article>
        </div>
      </section>
    </template>

    <template v-else>
      <section class="panel rules-panel">
        <div class="rule-toolbar"><div class="field"><label>规则 ID</label><input v-model="schemaId" /></div><div class="field"><label>版本</label><input v-model="schemaVersion" /></div><button class="secondary" @click="loadVersions">列出版本</button><button class="secondary" @click="loadRule">加载版本</button></div>
        <div v-if="versions.length" class="version-list"><button v-for="item in versions" :key="item.version" @click="schemaVersion = item.version; loadRule()">{{ item.version }} · {{ item.config_sha256.slice(0, 10) }}</button></div>
        <div class="editor-tabs"><button :class="{ active: editorMode === 'visual' }" @click="editorMode = 'visual'; syncVisualFromJson()">可视化编辑</button><button :class="{ active: editorMode === 'json' }" @click="editorMode = 'json'; syncJsonFromVisual()">高级 JSON</button></div>
        <div v-if="editorMode === 'visual'" class="visual-editor">
          <div class="rule-toolbar"><div class="field"><label>Schema ID</label><input v-model="visualRule.schema_id" /></div><div class="field"><label>语义版本</label><input v-model="visualRule.schema_version" /></div><div class="field"><label>规则名称</label><input v-model="visualRule.name" /></div><div class="field"><label>上传上限 MiB</label><input v-model.number="visualRule.workbook.max_upload_mib" type="number" min="1" /></div></div>
          <article v-for="(sheet, sheetIndex) in visualRule.sheets" :key="sheet.id || sheetIndex" class="sheet-card"><header><strong>工作表 {{ sheetIndex + 1 }}</strong><button class="danger" @click="visualRule.sheets.splice(sheetIndex, 1)">删除</button></header><div class="rule-toolbar"><div class="field"><label>ID</label><input v-model="sheet.id" /></div><div class="field"><label>工作表名</label><input v-model="sheet.name" /></div><div class="field"><label>主键（逗号分隔）</label><input :value="(sheet.primary_key || []).join(',')" @change="updatePrimaryKey(sheet, $event)" /></div></div><div class="table-wrap"><table><thead><tr><th>规范名</th><th>显示名</th><th>类型</th><th>必填</th><th></th></tr></thead><tbody><tr v-for="(column, columnIndex) in sheet.columns" :key="column.name || columnIndex"><td><input v-model="column.name" /></td><td><input v-model="column.title" /></td><td><select v-model="column.type"><option v-for="kind in ['string','integer','decimal','date','datetime','boolean','enum','set','json','phone','id_code','postal_code','fuzzy_string']" :key="kind">{{ kind }}</option></select></td><td><input v-model="column.required" type="checkbox" /></td><td><button class="danger" @click="sheet.columns.splice(columnIndex, 1)">删除</button></td></tr></tbody></table></div><button class="secondary" @click="addColumn(sheet)">添加字段</button></article>
          <button class="secondary" @click="addSheet">添加工作表</button>
        </div>
        <template v-else><label>规则 JSON（发布版本不可修改；编辑后保存为草稿）</label><textarea v-model="ruleJson" spellcheck="false" placeholder="加载已发布版本，或粘贴完整规则 JSON。" /></template>
        <div class="actions"><button class="secondary" @click="saveDraft">保存草稿</button><button class="secondary" @click="validateDraft">校验草稿</button><button class="primary" @click="publishDraft">发布不可变版本</button></div>
        <div class="mapping"><h3>人工确认表头别名</h3><input v-model="mapping.sheet_id" placeholder="工作表 ID" /><input v-model="mapping.raw_header" placeholder="原始表头" /><input v-model="mapping.canonical_field" placeholder="规范字段名" /><button class="secondary" @click="confirmMapping">确认并写入草稿</button></div>
        <pre v-if="ruleMessage" class="message">{{ ruleMessage }}</pre>
      </section>
    </template>
  </main>
</template>
