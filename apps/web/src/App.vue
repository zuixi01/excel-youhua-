<script setup lang="ts">
import { computed, onUnmounted, ref } from "vue";

type JobStatus = { job_id: string; status: string; progress: number; summary?: Record<string, number>; warnings?: string[]; artifacts?: Record<string, string>; error_message_safe?: string };
type Difference = { difference_id: string; type: string; severity: string; sheet_id: string; sheet_name: string; cell?: string; excel_row?: number; canonical_field?: string; business_key?: Record<string, unknown>; excel_raw_value?: unknown; excel_normalized_value?: unknown; standard_raw_value?: unknown; standard_normalized_value?: unknown; rule_id?: string; message: string; repair_status: string };

const tab = ref<"tasks" | "rules">("tasks");
const schemaId = ref("employee-roster");
const schemaVersion = ref("1.0.0");
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
let timer: number | undefined;

const versions = ref<Array<{ version: string; config_sha256: string }>>([]);
const ruleJson = ref("");
const editorMode = ref<"visual" | "json">("visual");
const visualRule = ref<any>({ schema_id: "", schema_version: "1.0.0", name: "", workbook: {}, sheets: [] });
const draftId = ref("");
const ruleMessage = ref("");
const mapping = ref({ sheet_id: "", raw_header: "", canonical_field: "" });

const finished = computed(() => ["completed", "failed", "manual_review", "cancelled"].includes(job.value?.status ?? ""));

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

async function precheck() {
  error.value = ""; precheckResult.value = undefined;
  if (!excel.value) { error.value = "请先选择待核验 Excel。"; return; }
  const body = new FormData();
  body.append("excel_file", excel.value); body.append("schema_id", schemaId.value); body.append("schema_version", schemaVersion.value);
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
  if (finished.value) { window.clearInterval(timer); busy.value = false; if (["completed", "manual_review"].includes(current.status)) await loadDifferences(); }
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
  const extensions: Record<string, string> = { excel: "xlsx", json: "json", differences_jsonl: "jsonl", html: "html", manifest: "json" };
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
    <div class="tabs"><button :class="{ active: tab === 'tasks' }" @click="tab = 'tasks'">核验任务</button><button :class="{ active: tab === 'rules' }" @click="tab = 'rules'">规则管理</button></div>

    <template v-if="tab === 'tasks'">
      <section class="panel form-panel">
        <div class="field"><label for="schema">规则 ID</label><input id="schema" v-model="schemaId" /></div>
        <div class="field"><label for="version">规则版本</label><input id="version" v-model="schemaVersion" /></div>
        <label class="drop"><span>待核验 Excel</span><strong>{{ excel?.name || "选择 .xlsx / .xlsm 文件" }}</strong><input type="file" accept=".xlsx,.xlsm" @change="choose($event, 'excel')" /></label>
        <label class="drop" :class="{ disabled: managedSource }"><span>标准数据</span><strong>{{ managedSource ? "由已配置的受管 HTTP 连接获取" : (standard?.name || "选择 JSON 或 CSV 文件") }}</strong><input type="file" accept=".json,.csv" :disabled="managedSource" @change="choose($event, 'standard')" /></label>
        <label class="toggle"><input v-model="managedSource" type="checkbox" /> 使用规则中的受管 HTTP 标准源</label>
        <div class="actions"><button class="secondary" :disabled="busy" @click="precheck">工作簿预检查</button><button class="primary" :disabled="busy" @click="submit">{{ busy ? "正在核验…" : "创建核验任务" }}</button></div>
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
