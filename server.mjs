import http from "node:http";
import os from "node:os";
import { readFile, stat } from "node:fs/promises";
import { extname, join, normalize } from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = normalize(join(__filename, ".."));

try {
  const dotenvPath = join(__dirname, ".env");
  const raw = await readFile(dotenvPath, "utf8");
  for (const line of raw.split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) continue;
    const idx = trimmed.indexOf("=");
    if (idx === -1) continue;
    const key = trimmed.slice(0, idx).trim();
    const value = trimmed.slice(idx + 1).trim();
    if (key && !process.env[key]) process.env[key] = value;
  }
} catch {}

const port = Number(process.env.PORT ?? 5173);
const publicRoot = join(__dirname, ".");

const LLM_BASE_URL = process.env.LLM_BASE_URL ?? "https://api.deepseek.com/v1";
const LLM_API_KEY = process.env.LLM_API_KEY ?? "";
const LLM_MODEL = process.env.LLM_MODEL ?? "deepseek-chat";

const TD_ACCESS_TOKEN = process.env.TD_ACCESS_TOKEN ?? "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJjbHQiOiJmYzQ4OTc0MjI1YzM0NDU2YjdiMDJlNzVlZWEwYzVjMyIsInR5cCI6MSwiZXhwIjoxNzc5OTU0MjI0Ljk5OTkyNywiaWF0IjoxNzc3MzYyMjI0Ljk5OTkyNywic3ViIjoiZjkwMmI1MjgyZGQ4NDc3NTliNjY5YTk1NzQ0YTk0YTcifQ.zvxLiLr4bxtBKDgvrD7OGMed-2KtuULHVIb9sgRG0wQ";
const TD_CLIENT_ID = process.env.TD_CLIENT_ID ?? "fc48974225c34456b7b02e75eea0c5c3";
const TD_OPEN_ID = process.env.TD_OPEN_ID ?? "f902b5282dd847759b669a95744a94a7";
const TD_FILE_ID = process.env.TD_FILE_ID ?? "DZU1idU5oZ3hKckFs";
const TD_SHEET_ID = process.env.TD_SHEET_ID ?? "BB08J2";

async function callTencentDocs(data) {
  const url = `https://docs.qq.com/openapi/spreadsheet/v3/files/${TD_FILE_ID}/batchUpdate`;
  const body = { requests: data };
  const resp = await fetch(url, {
    method: "POST",
    headers: {
      "Access-Token": TD_ACCESS_TOKEN,
      "Client-Id": TD_CLIENT_ID,
      "Open-Id": TD_OPEN_ID,
      "content-type": "application/json",
    },
    body: JSON.stringify(body),
  });
  const result = await resp.json().catch(() => ({}));
  if (!resp.ok) {
    console.error(`[TD] V3 batchUpdate error ${resp.status}:`, JSON.stringify(result));
    throw new Error(`TD_API_${resp.status}`);
  }
  return result;
}

async function handleTdSyncImport(req, res) {
  try {
    const raw = await readBody(req);
    const body = JSON.parse(raw || "{}");
    const headers = Array.isArray(body.headers) ? body.headers : [];
    const rows = Array.isArray(body.rows) ? body.rows : [];
    if (!rows.length) return json(res, 400, { error: "data_required" });

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

    await callTencentDocs([{
      updateSheet: {
        sheetId: TD_SHEET_ID,
        resource: { values: allRows },
      },
    }]);
    return json(res, 200, { ok: true, count: rows.length });
  } catch (e) {
    console.error("[TD] sync-import failed:", e.message);
    return json(res, 500, { error: "td_sync_failed", detail: e.message });
  }
}

async function handleTdSyncRecord(req, res) {
  try {
    const raw = await readBody(req);
    const body = JSON.parse(raw || "{}");
    const rowIndex = Number(body.rowIndex);
    const poi = body.poi;
    if (!Number.isFinite(rowIndex) || !poi) return json(res, 400, { error: "data_required" });

    const r = rowIndex + 1;
    const cb = body.headerCount || 0;
    const cols = ["id", "name", "address", "city", "district", "lng", "lat", "visitStatus", "merchantStatus", "phone", "remark", "updatedAt"];
    const writeValues = [];
    for (let c = 0; c < (cb + cols.length); c++) writeValues.push("");
    for (let i = 0; i < cols.length; i++) {
      writeValues[cb + i] = String(poi[cols[i]] ?? "");
    }

    await callTencentDocs([{
      updateSheet: {
        sheetId: TD_SHEET_ID,
        resource: { row: r, values: [writeValues] },
      },
    }]);
    return json(res, 200, { ok: true });
  } catch (e) {
    console.error("[TD] sync-record failed:", e.message);
    return json(res, 500, { error: "td_sync_failed", detail: e.message });
  }
}

function json(res, statusCode, data) {
  const body = JSON.stringify(data);
  res.writeHead(statusCode, {
    "content-type": "application/json; charset=utf-8",
    "content-length": Buffer.byteLength(body),
  });
  res.end(body);
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on("data", (c) => chunks.push(c));
    req.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
    req.on("error", reject);
  });
}

function getLocalIps() {
  const ifaces = os.networkInterfaces();
  const addrs = [];
  for (const [, list] of Object.entries(ifaces)) {
    for (const info of list || []) {
      if (info.family === "IPv4" && !info.internal) {
        addrs.push(info.address);
      }
    }
  }
  return addrs;
}

function guessMappingByHeuristics(headers) {
  const norm = (s) => String(s ?? "").trim().toLowerCase();
  const candidates = headers.map((h) => ({
    raw: h,
    n: norm(h),
  }));

  const pick = (rules) => {
    for (const r of rules) {
      const hit = candidates.find((c) => r.test(c.n));
      if (hit) return hit.raw;
    }
    return undefined;
  };

  return {
    name: pick([
      /^name$/,
      /merchant.*name/,
      /shop.*name/,
      /poi.*name/,
      /商户.*名/,
      /店铺.*名/,
      /门店.*名/,
      /名称/,
    ]),
    address: pick([
      /^address$/,
      /merchant.*address/,
      /shop.*address/,
      /poi.*address/,
      /地址/,
      /详细地址/,
      /联系地址/,
    ]),
    city: pick([
      /^(city|city_name|prefecture)$/,
      /地市/,
      /城市/,
      /地级市/,
      /所属地市/,
      /所在城市/,
      /city/,
      /地市.*名/,
      /城市.*名/,
    ]),
    district: pick([
      /^(district|district_name|county|area)$/,
      /区县/,
      /行政区/,
      /所属区县/,
      /所在区/,
      /district/,
      /区县.*名/,
      /区域/,
    ]),
    lng: pick([
      /^(lng|lon|long|longitude)$/,
      /经度/,
      /lng/,
      /longitude/,
      /lon/,
    ]),
    lat: pick([
      /^(lat|latitude)$/,
      /纬度/,
      /lat/,
      /latitude/,
    ]),
    merchantStatus: pick([
      /merchant.*status/,
      /biz.*status/,
      /商户.*状态/,
      /经营.*状态/,
      /营业.*状态/,
      /企业.*状态/,
    ]),
    phone: pick([
      /^(phone|tel|mobile|telephone)$/,
      /电话/,
      /联系电话/,
      /手机/,
      /联系.*话/,
      /联络.*话/,
    ]),
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
    "- merchantStatus: 商户经营状态/营业状态，如正常营业、暂停营业、关停等（如果存在）",
    "- phone: 联系电话/手机号（如果存在）",
    "",
    "要求：",
    "1) 只输出严格JSON，不要输出任何多余文本。",
    "2) JSON结构：{ \"mapping\": {\"name\": \"...\", \"address\": \"...\", \"city\": \"...\", \"district\": \"...\", \"lng\": \"...\", \"lat\": \"...\", \"merchantStatus\": \"...\", \"phone\": \"...\" }, \"confidence\": {\"name\":0-1,...}, \"rationale\": \"...\" }",
    "3) 不确定就留空字符串。",
    "",
    `headers=${JSON.stringify(headers)}`,
    sampleRows && sampleRows.length ? `sampleRows=${JSON.stringify(sampleRows.slice(0, 3))}` : "",
  ]
    .filter(Boolean)
    .join("\n");

  const resp = await fetch(`${LLM_BASE_URL.replace(/\/$/, "")}/chat/completions`, {
    method: "POST",
    headers: {
      authorization: `Bearer ${LLM_API_KEY}`,
      "content-type": "application/json",
    },
    body: JSON.stringify({
      model: LLM_MODEL,
      temperature: 0,
      messages: [
        { role: "system", content: "你只输出严格JSON。" },
        { role: "user", content: prompt },
      ],
    }),
  });

  if (!resp.ok) return null;

  const data = await resp.json();
  const content = data?.choices?.[0]?.message?.content;
  if (typeof content !== "string" || !content.trim()) return null;

  try {
    return JSON.parse(content);
  } catch {
    const m = content.match(/\{[\s\S]*\}/);
    if (!m) return null;
    try {
      return JSON.parse(m[0]);
    } catch {
      return null;
    }
  }
}

async function handleIdentifyColumns(req, res) {
  try {
    const raw = await readBody(req);
    const body = JSON.parse(raw || "{}");
    const headers = Array.isArray(body.headers) ? body.headers.map(String) : [];
    const sampleRows = Array.isArray(body.sampleRows) ? body.sampleRows : [];

    if (!headers.length) return json(res, 400, { error: "headers_required" });

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

    const confidence = {
      name: typeof llm?.confidence?.name === "number" ? llm.confidence.name : merged.name ? 0.72 : 0,
      address:
        typeof llm?.confidence?.address === "number" ? llm.confidence.address : merged.address ? 0.72 : 0,
      lng: typeof llm?.confidence?.lng === "number" ? llm.confidence.lng : merged.lng ? 0.72 : 0,
      lat: typeof llm?.confidence?.lat === "number" ? llm.confidence.lat : merged.lat ? 0.72 : 0,
    };

    return json(res, 200, {
      mapping: merged,
      confidence,
      rationale: typeof llm?.rationale === "string" ? llm.rationale : "",
      llmEnabled: Boolean(LLM_API_KEY),
    });
  } catch {
    return json(res, 500, { error: "internal_error" });
  }
}

const mimeByExt = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".mjs": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon",
};

function safeJoin(root, urlPath) {
  const cleaned = decodeURIComponent(urlPath.split("?")[0]).replace(/^\/+/, "");
  const p = normalize(join(root, cleaned));
  if (!p.startsWith(root)) return null;
  return p;
}

const server = http.createServer(async (req, res) => {
  const url = req.url || "/";
  if (req.method === "GET" && url === "/api/health") return json(res, 200, { ok: true });
  if (req.method === "GET" && url === "/api/info") {
    const ips = getLocalIps();
    return json(res, 200, { ips, hostname: os.hostname(), port, platform: os.platform() });
  }
  if (req.method === "POST" && url === "/api/columns/identify") return handleIdentifyColumns(req, res);
  if (req.method === "POST" && url === "/api/td/sync-import") return handleTdSyncImport(req, res);
  if (req.method === "POST" && url === "/api/td/sync-record") return handleTdSyncRecord(req, res);

  if (req.method !== "GET" && req.method !== "HEAD") {
    res.writeHead(405);
    return res.end();
  }

  const filePath = safeJoin(publicRoot, url === "/" ? "/index.html" : url);
  if (!filePath) {
    res.writeHead(400);
    return res.end();
  }

  try {
    const st = await stat(filePath);
    if (!st.isFile()) {
      res.writeHead(404);
      return res.end();
    }
  } catch {
    const fallback = safeJoin(publicRoot, "/index.html");
    if (!fallback) {
      res.writeHead(404);
      return res.end();
    }
    try {
      const html = await readFile(fallback);
      res.writeHead(200, { "content-type": mimeByExt[".html"] });
      return res.end(html);
    } catch {
      res.writeHead(404);
      return res.end();
    }
  }

  try {
    const data = await readFile(filePath);
    const ext = extname(filePath).toLowerCase();
    res.writeHead(200, { "content-type": mimeByExt[ext] ?? "application/octet-stream" });
    return res.end(data);
  } catch {
    res.writeHead(500);
    return res.end();
  }
});

server.on("error", () => {
  process.exitCode = 1;
});

server.listen(port, () => {
  const ips = getLocalIps();
  process.stdout.write(`Server running at http://localhost:${port}/\n`);
  for (const ip of ips) {
    process.stdout.write(`  LAN: http://${ip}:${port}/\n`);
  }
  process.stdout.write("\n微信小程序真机调试请使用上面的 LAN 地址\n");
});
