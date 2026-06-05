const LLM_BASE_URL = process.env.LLM_BASE_URL ?? "https://api.deepseek.com/v1";
const LLM_API_KEY = process.env.LLM_API_KEY ?? "";
const LLM_MODEL = process.env.LLM_MODEL ?? "deepseek-chat";

const TD_ACCESS_TOKEN = process.env.TD_ACCESS_TOKEN ?? "";
const TD_CLIENT_ID = process.env.TD_CLIENT_ID ?? "";
const TD_OPEN_ID = process.env.TD_OPEN_ID ?? "";
const TD_FILE_ID = process.env.TD_FILE_ID ?? "DZU1idU5oZ3hKckFs";
const TD_SHEET_ID = process.env.TD_SHEET_ID ?? "BB08J2";

const CORS_HEADERS = {
  "access-control-allow-origin": "*",
  "access-control-allow-methods": "POST, OPTIONS",
  "access-control-allow-headers": "content-type",
};

function json(body, statusCode = 200) {
  return {
    isBase64Encoded: false,
    statusCode,
    headers: {
      "content-type": "application/json; charset=utf-8",
      ...CORS_HEADERS,
    },
    body: JSON.stringify(body),
  };
}

function guessMappingByHeuristics(headers) {
  const norm = (s) => String(s ?? "").trim().toLowerCase();
  const candidates = headers.map((h) => ({ raw: h, n: norm(h) }));
  const pick = (rules) => {
    for (const r of rules) {
      const hit = candidates.find((c) => r.test(c.n));
      if (hit) return hit.raw;
    }
    return undefined;
  };
  return {
    name: pick([/^name$/, /merchant.*name/, /shop.*name/, /poi.*name/, /商户.*名/, /店铺.*名/, /门店.*名/, /名称/]),
    address: pick([/^address$/, /merchant.*address/, /shop.*address/, /poi.*address/, /地址/, /详细地址/, /联系地址/]),
    city: pick([/^(city|city_name|prefecture)$/, /地市/, /城市/, /地级市/, /所属地市/, /所在城市/, /city/, /地市.*名/, /城市.*名/]),
    district: pick([/^(district|district_name|county|area)$/, /区县/, /行政区/, /所属区县/, /所在区/, /district/, /区县.*名/, /区域/]),
    lng: pick([/^(lng|lon|long|longitude)$/, /经度/, /lng/, /longitude/, /lon/]),
    lat: pick([/^(lat|latitude)$/, /纬度/, /lat/, /latitude/]),
    merchantStatus: pick([/merchant.*status/, /biz.*status/, /商户.*状态/, /经营.*状态/, /营业.*状态/, /企业.*状态/]),
    phone: pick([/^(phone|tel|mobile|telephone)$/, /电话/, /联系电话/, /手机/, /联系.*话/, /联络.*话/]),
  };
}

function coerceMapping(headers, mapping) {
  const set = new Set(headers);
  const fixed = {};
  for (const k of ["name", "address", "city", "district", "lng", "lat", "merchantStatus", "phone"]) {
    const v = mapping?.[k];
    if (typeof v === "string" && set.has(v)) fixed[k] = v;
  }
  return fixed;
}

async function identifyColumnsWithLLM(headers, sampleRows) {
  if (!LLM_API_KEY) return null;
  const prompt = [
    "你是一个数据字段映射器。任务：从表头数组中找出以下字段对应的列名（精确返回原始表头字符串）：",
    "- name: 商户名称/门店名称/POI名称",
    "- address: 商户地址/门店地址/POI地址",
    "- city: 地市/城市/所属地市（如果存在）",
    "- district: 区县/行政区/所属区县（如果存在）",
    "- lng: 经度（数字）",
    "- lat: 纬度（数字）",
    "- merchantStatus: 商户经营状态/营业状态（如果存在）",
    "- phone: 联系电话/手机号（如果存在）",
    "",
    "要求：只输出严格JSON。",
    `headers=${JSON.stringify(headers)}`,
    sampleRows && sampleRows.length ? `sampleRows=${JSON.stringify(sampleRows.slice(0, 3))}` : "",
  ].join("\n");

  const resp = await fetch(`${LLM_BASE_URL.replace(/\/$/, "")}/chat/completions`, {
    method: "POST",
    headers: { authorization: `Bearer ${LLM_API_KEY}`, "content-type": "application/json" },
    body: JSON.stringify({ model: LLM_MODEL, temperature: 0, messages: [{ role: "system", content: "你只输出严格JSON。" }, { role: "user", content: prompt }] }),
  });
  if (!resp.ok) return null;
  const data = await resp.json();
  const content = data?.choices?.[0]?.message?.content;
  if (typeof content !== "string" || !content.trim()) return null;
  try { return JSON.parse(content); } catch {
    const m = content.match(/\{[\s\S]*\}/);
    if (!m) return null;
    try { return JSON.parse(m[0]); } catch { return null; }
  }
}

async function handleIdentifyColumns(body) {
  const headers = Array.isArray(body.headers) ? body.headers.map(String) : [];
  const sampleRows = Array.isArray(body.sampleRows) ? body.sampleRows : [];
  if (!headers.length) return json({ error: "headers_required" }, 400);

  const heuristic = guessMappingByHeuristics(headers);
  const llm = await identifyColumnsWithLLM(headers, sampleRows);
  const llmMapping = coerceMapping(headers, llm?.mapping);
  const merged = {
    name: llmMapping.name || heuristic.name,
    address: llmMapping.address || heuristic.address,
    city: llmMapping.city || heuristic.city,
    district: llmMapping.district || heuristic.district,
    lng: llmMapping.lng || heuristic.lng,
    lat: llmMapping.lat || heuristic.lat,
    merchantStatus: llmMapping.merchantStatus || heuristic.merchantStatus,
    phone: llmMapping.phone || heuristic.phone,
  };
  return json({
    mapping: merged,
    confidence: {
      name: typeof llm?.confidence?.name === "number" ? llm.confidence.name : merged.name ? 0.72 : 0,
      address: typeof llm?.confidence?.address === "number" ? llm.confidence.address : merged.address ? 0.72 : 0,
      lng: typeof llm?.confidence?.lng === "number" ? llm.confidence.lng : merged.lng ? 0.72 : 0,
      lat: typeof llm?.confidence?.lat === "number" ? llm.confidence.lat : merged.lat ? 0.72 : 0,
    },
    rationale: typeof llm?.rationale === "string" ? llm.rationale : "",
    llmEnabled: Boolean(LLM_API_KEY),
  });
}

async function handleTdSyncImport(body) {
  const headers = Array.isArray(body.headers) ? body.headers : [];
  const rows = Array.isArray(body.rows) ? body.rows : [];
  if (!rows.length) return json({ error: "data_required" }, 400);

  const allHeaders = [...headers, "id", "name", "address", "city", "district", "lng", "lat", "visitStatus", "merchantStatus", "phone", "remark", "updatedAt"];
  const allRows = [allHeaders];
  for (const r of rows) {
    allRows.push([
      ...headers.map((h) => String((r._orig || {})[h] ?? "")),
      String(r.id ?? ""), String(r.name ?? ""), String(r.address ?? ""),
      String(r.city ?? ""), String(r.district ?? ""),
      String(r.lng ?? ""), String(r.lat ?? ""),
      String(r.visitStatus ?? ""), String(r.merchantStatus ?? ""),
      String(r.phone ?? ""), String(r.remark ?? ""),
      String(r.updatedAt ?? ""),
    ]);
  }

  const resp = await fetch(`https://docs.qq.com/openapi/spreadsheet/v3/files/${TD_FILE_ID}/batchUpdate`, {
    method: "POST",
    headers: { "Access-Token": TD_ACCESS_TOKEN, "Client-Id": TD_CLIENT_ID, "Open-Id": TD_OPEN_ID, "content-type": "application/json" },
    body: JSON.stringify({ requests: [{ updateSheet: { sheetId: TD_SHEET_ID, resource: { values: allRows } } }] }),
  });
  if (!resp.ok) return json({ error: "td_sync_failed" }, 500);
  return json({ ok: true, count: rows.length });
}

async function handleTdSyncRecord(body) {
  const rowIndex = Number(body.rowIndex);
  const poi = body.poi;
  if (!Number.isFinite(rowIndex) || !poi) return json({ error: "data_required" }, 400);

  const r = rowIndex + 1;
  const cb = body.headerCount || 0;
  const cols = ["id", "name", "address", "city", "district", "lng", "lat", "visitStatus", "merchantStatus", "phone", "remark", "updatedAt"];
  const writeValues = [];
  for (let c = 0; c < (cb + cols.length); c++) writeValues.push("");
  for (let i = 0; i < cols.length; i++) writeValues[cb + i] = String(poi[cols[i]] ?? "");

  const resp = await fetch(`https://docs.qq.com/openapi/spreadsheet/v3/files/${TD_FILE_ID}/batchUpdate`, {
    method: "POST",
    headers: { "Access-Token": TD_ACCESS_TOKEN, "Client-Id": TD_CLIENT_ID, "Open-Id": TD_OPEN_ID, "content-type": "application/json" },
    body: JSON.stringify({ requests: [{ updateSheet: { sheetId: TD_SHEET_ID, resource: { row: r, values: [writeValues] } } }] }),
  });
  if (!resp.ok) return json({ error: "td_sync_failed" }, 500);
  return json({ ok: true });
}

exports.main_handler = async (event) => {
  try {
    if (event.httpMethod === "OPTIONS") {
      return { isBase64Encoded: false, statusCode: 204, headers: CORS_HEADERS, body: "" };
    }
    const path = (event.path || "").replace(/\/+$/, "") || "/";
    const body = event.body ? JSON.parse(event.body) : {};

    if (event.httpMethod === "POST" && path.endsWith("/api/columns/identify")) return handleIdentifyColumns(body);
    if (event.httpMethod === "POST" && path.endsWith("/api/td/sync-import")) return handleTdSyncImport(body);
    if (event.httpMethod === "POST" && path.endsWith("/api/td/sync-record")) return handleTdSyncRecord(body);

    return json({ error: "not_found" }, 404);
  } catch (e) {
    return json({ error: "internal_error", detail: e.message }, 500);
  }
};
