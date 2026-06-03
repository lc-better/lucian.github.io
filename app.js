const storageKeys = {
  userName: "visit_poi_user_name",
  state: "visit_poi_state_v1",
};

const visitStateColors = {
  unvisited: "#6B7AA6",
  visited: "#3AE6D1",
  exception: "#F7C454",
  selected: "#5BB1FF",
};

function byId(id) {
  const el = document.getElementById(id);
  if (!el) throw new Error(`missing_el:${id}`);
  return el;
}

function safeText(v) {
  if (v === null || v === undefined) return "";
  return String(v);
}

function nowIso() {
  return new Date().toISOString();
}

function nanoId() {
  const a = new Uint8Array(16);
  crypto.getRandomValues(a);
  return Array.from(a)
    .map((n) => n.toString(16).padStart(2, "0"))
    .join("");
}

function parseNumber(v) {
  const n = Number(String(v ?? "").trim());
  return Number.isFinite(n) ? n : null;
}

function normalizeStatus(s) {
  const t = String(s ?? "").trim();
  if (!t) return "未走访";
  return t;
}

const merchantStatusMap = {
  "1": "正常营业",
  "3": "暂停营业",
  "5": "尚未开业",
  "9": "停止营业",
  "15": "模糊上线",
  "16": "疑似下线",
  "20": "闭店",
  "99": "无法核实",
};

function normalizeMerchantStatus(v) {
  const t = String(v ?? "").trim();
  if (!t) return "";
  return merchantStatusMap[t] || t;
}

function toVisitState(status) {
  const t = String(status ?? "").trim();
  if (!t || t === "未走访") return "unvisited";
  if (t === "已走访") return "visited";
  if (t === "无法到访" || t === "关停") return "exception";
  if (/异常|无法|关停|停业|关闭/.test(t)) return "exception";
  return "unvisited";
}

function downloadBlob(filename, blob) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

async function poisToXlsxBlob(pois) {
  if (!window.XLSX) throw new Error("XLSX库未加载");

  const visitStateMap = { unvisited: "未走访", visited: "已走访", exception: "异常/无法到访" };
  const accLabel = (v) => (v && v !== "unknown" ? ({ accurate: "准确", inaccurate: "不准确" }[v] || v) : "");
  const origHeaders = state.headers.length ? state.headers : [];
  const extraHeaders = [
    "id",
    "name",
    "address",
    "city",
    "district",
    "lng",
    "lat",
    "visitStatus",
    "merchantStatus",
    "phone",
    "remark",
    "nameAccuracy",
    "addressAccuracy",
    "coordAccuracy",
    "visitStatusAccuracy",
    "updatedAt",
  ];
  const headers = [...origHeaders, ...extraHeaders];
  const photoColBase = headers.length;
  const maxPhotos = 5;
  for (let k = 1; k <= maxPhotos; k++) headers.push(`照片${k}`);

  if (!state.db) state.db = await createDb();

  const photoByRow = [];
  const rows = [];
  rows.push(headers);

  for (let i = 0; i < pois.length; i++) {
    const p = pois[i];
    const acc = p.accuracy || {};
    const origRow = typeof p.rowIdx === "number" && state.rawRows[p.rowIdx] ? state.rawRows[p.rowIdx] : {};
    const origVals = origHeaders.map((h) => {
      const v = origRow[h];
      return v != null ? String(v) : "";
    });
    const photos = p.photos || [];
    const extraVals = [
      p.id,
      p.name,
      p.address,
      p.city || "",
      p.district || "",
      p.lng,
      p.lat,
      visitStateMap[p.visitState] || p.visitState || "",
      p.merchantStatus || "",
      p.phone || "",
      p.remark || "",
      accLabel(acc.name),
      accLabel(acc.address),
      accLabel(acc.coord),
      accLabel(acc.status),
      p.updatedAt || "",
    ];
    for (let k = 0; k < maxPhotos; k++) {
      extraVals.push(k < photos.length ? photos[k].filename : "");
    }
    rows.push([...origVals, ...extraVals]);

    if (photos.length > 0) {
      const rowBases = [];
      for (let j = 0; j < Math.min(photos.length, maxPhotos); j++) {
        try {
          const blob = await idbGet(state.db, "photos", photos[j].blobKey);
          if (!blob) continue;
          const b64 = await blobToBase64(blob);
          const mime = photos[j].mime || "image/jpeg";
          rowBases.push({ b64, mime, col: photoColBase + j });
        } catch {}
      }
      if (rowBases.length > 0) photoByRow.push({ row: i + 1, images: rowBases });
    }
  }

  if (photoByRow.length > 0) {
    return buildHtmlExcelBlob(rows, photoByRow);
  }

  const ws = XLSX.utils.aoa_to_sheet(rows);
  const colWidths = [];
  for (let c = 0; c < headers.length; c++) {
    if (c >= photoColBase && c < photoColBase + maxPhotos) {
      colWidths.push({ wch: 22 });
    } else {
      colWidths.push({ wch: 18 });
    }
  }
  ws["!cols"] = colWidths;
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "POI导出");
  const buf = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  return new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

function buildHtmlExcelBlob(rows, photoByRow) {
  const photoLookup = {};
  for (const pr of photoByRow) {
    photoLookup[pr.row] = pr.images;
  }

  const escapeHtml = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

  let html = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
<head><meta charset="UTF-8">
<!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>POI导出</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->
<style>
table { border-collapse: collapse; }
th { background: #e8e8e8; font-weight: bold; padding: 6px 10px; border: 1px solid #999; text-align: center; }
td { padding: 6px 10px; border: 1px solid #ccc; vertical-align: top; }
img { display: block; max-width: 120px; height: auto; margin: 2px 0; }
</style>
</head><body><table>\n`;

  for (let i = 0; i < rows.length; i++) {
    html += "<tr>";
    const isHead = i === 0;
    const tag = isHead ? "th" : "td";
    for (let c = 0; c < rows[i].length; c++) {
      const lookup = photoLookup[i] || [];
      const imgEntry = lookup.find((x) => x.col === c);
      html += `<${tag}>`;
      if (imgEntry) {
        html += `<img src="data:${imgEntry.mime};base64,${imgEntry.b64}" />`;
      }
      html += escapeHtml(rows[i][c] || "");
      html += `</${tag}>`;
    }
    html += "</tr>\n";
  }

  html += `</table></body></html>`;
  return new Blob([html], { type: "application/vnd.ms-excel;charset=utf-8" });
}

function blobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => {
      const result = reader.result;
      if (typeof result === "string") {
        const comma = result.indexOf(",");
        resolve(comma >= 0 ? result.slice(comma + 1) : result);
      } else {
        reject(new Error("read_failed"));
      }
    };
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

function isMobileLike() {
  return /Android|iPhone|iPad|iPod/i.test(navigator.userAgent);
}

function buildBaiduNavUrl(origin, dest, destName) {
  const o = `latlng:${origin.lat},${origin.lng}|name:当前位置`;
  const d = `latlng:${dest.lat},${dest.lng}|name:${encodeURIComponent(destName || "目的地")}`;
  const coordType = (window.__APP_CONFIG__?.coordType || "bd09ll").toLowerCase();

  if (isMobileLike()) {
    return `baidumap://map/direction?origin=${o}&destination=${d}&mode=driving&coord_type=${coordType}&src=webapp.visitpoi`;
  }
  const https = `https://api.map.baidu.com/direction?origin=${encodeURIComponent(
    `latlng:${origin.lat},${origin.lng}|name:当前位置`
  )}&destination=${encodeURIComponent(
    `latlng:${dest.lat},${dest.lng}|name:${destName || "目的地"}`
  )}&mode=driving&coord_type=${encodeURIComponent(coordType)}&output=html&src=webapp.visitpoi`;
  return https;
}

function createDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open("visit_poi_db", 1);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains("photos")) db.createObjectStore("photos");
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function idbPut(db, store, key, value) {
  return new Promise((resolve, reject) => {
    const tx = db.transaction(store, "readwrite");
    tx.objectStore(store).put(value, key);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

async function idbGet(db, store, key) {
  return new Promise((resolve, reject) => {
    const tx = db.transaction(store, "readonly");
    const req = tx.objectStore(store).get(key);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function idbDel(db, store, key) {
  return new Promise((resolve, reject) => {
    const tx = db.transaction(store, "readwrite");
    tx.objectStore(store).delete(key);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

const state = {
  file: null,
  rawRows: [],
  headers: [],
  mapping: null,
  pois: [],
  selectedPoiId: null,
  navPlannedForPoiId: null,
  driving: null,
  mapReady: false,
  map: null,
  markers: new Map(),
  mapping: null,
  db: null,
  pendingPhotoPoiId: null,
  pendingPhotoFiles: [],
  filters: {
    city: "",
    district: "",
    merchantStatus: "",
    visitState: "",
  },
};

const photoPreviewUrls = [];

function clearPhotoPreviewUrls() {
  while (photoPreviewUrls.length) {
    const u = photoPreviewUrls.pop();
    try {
      URL.revokeObjectURL(u);
    } catch {}
  }
}

function saveState() {
  const payload = {
    pois: state.pois,
    selectedPoiId: state.selectedPoiId,
    navPlannedForPoiId: state.navPlannedForPoiId,
    mapping: state.mapping || null,
    headers: state.headers || [],
    rawRows: state.pois.length < 2000 ? state.rawRows : [],
    filters: state.filters || { city: "", district: "", merchantStatus: "", visitState: "" },
  };
  localStorage.setItem(storageKeys.state, JSON.stringify(payload));
  postStateToMiniProgram(payload);
}

function postStateToMiniProgram(payload) {
  if (typeof wx !== "undefined" && wx.miniProgram && wx.miniProgram.postMessage) {
    try {
      wx.miniProgram.postMessage({ data: { type: "state", payload } });
    } catch (e) { }
  }
}

function tryRestoreFromMiniProgram() {
  if (typeof __MP_BACKUP_STATE__ !== "undefined" && __MP_BACKUP_STATE__) {
    try {
      const v = typeof __MP_BACKUP_STATE__ === "string" ? JSON.parse(__MP_BACKUP_STATE__) : __MP_BACKUP_STATE__;
      if (v && Array.isArray(v.pois)) {
        state.pois = v.pois;
        state.selectedPoiId = v.selectedPoiId || null;
        state.navPlannedForPoiId = v.navPlannedForPoiId || null;
        state.mapping = v.mapping || null;
        state.headers = Array.isArray(v.headers) ? v.headers : [];
        state.rawRows = Array.isArray(v.rawRows) ? v.rawRows : [];
        state.filters = v.filters || { city: "", district: "", merchantStatus: "", visitState: "" };
        saveState();
        return true;
      }
    } catch (e) { }
  }
  return false;
}

async function loadState() {
  const raw = localStorage.getItem(storageKeys.state);
  if (!raw) return;
  try {
    const v = JSON.parse(raw);
    state.pois = Array.isArray(v?.pois) ? v.pois : [];
    state.selectedPoiId = typeof v?.selectedPoiId === "string" ? v.selectedPoiId : null;
    state.navPlannedForPoiId = typeof v?.navPlannedForPoiId === "string" ? v.navPlannedForPoiId : null;
    state.mapping = v?.mapping || null;
    state.headers = Array.isArray(v?.headers) ? v.headers : [];
    state.rawRows = Array.isArray(v?.rawRows) ? v.rawRows : [];
    state.filters = v?.filters || { city: "", district: "", merchantStatus: "", visitState: "" };
  } catch {}
}

function setOverlay(message) {
  const el = byId("mapOverlay");
  if (!message) {
    el.classList.add("hidden");
    el.textContent = "";
    return;
  }
  el.textContent = message;
  el.classList.remove("hidden");
}

function showLoading() {
  byId("loadingOverlay").classList.remove("hidden");
}

function hideLoading() {
  byId("loadingOverlay").classList.add("hidden");
}

function updateStats() {
  const filtered = getFilteredPois();
  const total = state.pois.length;
  const filteredTotal = filtered.length;
  const visited = filtered.filter((p) => p.visitState !== "unvisited").length;
  const unvisited = filteredTotal - visited;
  byId("statTotal").textContent = String(total);
  byId("statVisited").textContent = String(visited);
  byId("statUnvisited").textContent = String(unvisited);
  byId("exportBtn").disabled = total === 0;
  updateFilterCounts();
}

function statusBadgeText(poi) {
  const t = poi?.status || "未走访";
  return t;
}

function getFilteredPois() {
  const f = state.filters;
  return state.pois.filter((p) => {
    if (f.city && p.city !== f.city) return false;
    if (f.district && p.district !== f.district) return false;
    if (f.merchantStatus && p.merchantStatus !== f.merchantStatus) return false;
    if (f.visitState) {
      if (f.visitState === "visited" && p.visitState === "unvisited") return false;
      if (f.visitState === "unvisited" && p.visitState !== "unvisited") return false;
    }
    return true;
  });
}

function updateFilterCounts() {
  const filterCities = byId("filterCity");
  const filterDistricts = byId("filterDistrict");
  const filterMerchantStatus = byId("filterMerchantStatus");
  const filterVisitState = byId("filterVisitState");
  if (!filterCities) return;

  const cities = new Set();
  const districts = new Set();
  const merchantStatuses = new Set();
  state.pois.forEach((p) => {
    if (p.city) cities.add(p.city);
    if (p.district) districts.add(p.district);
    if (p.merchantStatus) merchantStatuses.add(p.merchantStatus);
  });

  populateFilterSelect(filterCities, cities, state.filters.city || "");
  populateFilterSelect(filterDistricts, districts, state.filters.district || "");
  populateFilterSelect(filterMerchantStatus, merchantStatuses, state.filters.merchantStatus || "");
  populateFilterSelect(filterVisitState, ["已走访", "未走访"], state.filters.visitState === "visited" ? "已走访" : state.filters.visitState === "unvisited" ? "未走访" : "");
}

function populateFilterSelect(sel, values, currentValue) {
  if (!sel) return;
  const val = sel.value || "";
  sel.innerHTML = '<option value="">全部</option>';
  const arr = Array.isArray(values) ? values : Array.from(values).sort();
  for (const v of arr) {
    const opt = document.createElement("option");
    opt.value = v;
    opt.textContent = v;
    sel.appendChild(opt);
  }
  sel.value = currentValue || "";
}

function applyFilters() {
  saveState();
  updateStats();
  renderMarkers();
  fitToPois();
  updateFilterCounts();
}

function restoreFilterUi() {
  updateFilterCounts();
  const filterVisitState = byId("filterVisitState");
  if (filterVisitState) {
    const fv = state.filters.visitState;
    filterVisitState.value = fv === "visited" ? "已走访" : fv === "unvisited" ? "未走访" : "";
  }
}

function ensurePanels() {
  const empty = byId("emptyRight");
  const hasSelected = Boolean(state.selectedPoiId);
  if (hasSelected) empty.classList.add("hidden");
  else empty.classList.remove("hidden");
}

function renderDetail() {
  const card = byId("detailCard");
  const poi = state.pois.find((p) => p.id === state.selectedPoiId);
  if (!poi) {
    card.classList.add("hidden");
    byId("recordBtn").disabled = true;
    byId("navBtn").disabled = true;
    ensurePanels();
    return;
  }
  byId("detailName").textContent = poi.name || "";
  byId("detailAddress").textContent = poi.address || "";
  byId("detailCityDistrict").textContent = [poi.city, poi.district].filter(Boolean).join(" / ") || "-";
  byId("detailLngLat").textContent = `${poi.lng}, ${poi.lat}`;
  byId("detailStatusBadge").textContent = statusBadgeText(poi);
  byId("recordBtn").disabled = false;
  byId("navBtn").disabled = false;
  card.classList.remove("hidden");

  const navHint = byId("navHint");
  if (state.navPlannedForPoiId === poi.id) {
    navHint.textContent = "已在站内规划路线，再次点击“导航”可跳转百度地图App进行导航。";
    byId("navBtn").textContent = "导航（跳转百度地图）";
  } else {
    navHint.textContent = "首次点击“导航”会在应用内显示路线。";
    byId("navBtn").textContent = "导航（站内规划）";
  }
  ensurePanels();
}

function setSelectedPoi(poiId) {
  state.selectedPoiId = poiId;
  if (document.body.classList.contains("sidebarOpen")) {
    document.body.classList.remove("sidebarOpen");
    byId("sidebarBackdrop").classList.add("hidden");
  }
  renderMarkers();
  renderDetail();
  saveState();
  const recordCard = byId("recordCard");
  if (!recordCard.classList.contains("hidden")) {
    const newPoi = state.pois.find((p) => p.id === poiId);
    if (newPoi) openRecord(newPoi);
  }
}

function markerSvg(color) {
  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="36" height="36" viewBox="0 0 36 36">` +
    `<path d="M18 3c6.075 0 11 4.702 11 10.5 0 7.563-9.225 17.41-10.238 18.466a1.05 1.05 0 0 1-1.524 0C16.225 30.91 7 21.063 7 13.5 7 7.702 11.925 3 18 3z" fill="${color}"/>` +
    `<circle cx="18" cy="13.5" r="4.4" fill="rgba(5,16,21,.85)"/>` +
    `</svg>`;
  return `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
}

function getMarkerColor(poi) {
  if (poi.id === state.selectedPoiId) return visitStateColors.selected;
  if (poi.visitState === "visited") return visitStateColors.visited;
  if (poi.visitState === "exception") return visitStateColors.exception;
  return visitStateColors.unvisited;
}

function initMap() {
  if (!window.BMapGL) {
    const ak = window.__APP_CONFIG__?.baiduMapAk || "";
    const reason = window.__BMAP_LOAD_ERROR__ || "";
    const msg = !ak
      ? "请在 config.js 中配置 baiduMapAk（百度地图AK），刷新后即可加载地图并打点。"
      : reason === "load_failed"
        ? `百度地图脚本加载失败：请确认网络可访问 https://api.map.baidu.com/ ，并在百度控制台将 Referer 白名单配置为 ${location.origin}/* 后刷新。`
        : reason === "load_timeout"
          ? `百度地图初始化超时：如果你看到“APP服务被禁用了”的弹窗，说明该AK在百度控制台被禁用或未开通JSAPI GL权限；请到百度控制台启用应用并配置 Referer 白名单为 ${location.origin}/* 后刷新。`
          : "百度地图加载中…（如长时间无变化，请检查网络是否可访问百度开放平台脚本，并确认AK与Referer白名单配置正确）";
    setOverlay(msg);
    return;
  }
  state.mapReady = true;
  setOverlay("");
  const map = new BMapGL.Map("map");
  state.map = map;
  const center = new BMapGL.Point(116.404, 39.915);
  map.centerAndZoom(center, 12);
  map.enableScrollWheelZoom(true);
  try {
    map.addControl(new BMapGL.NavigationControl3D());
  } catch {}
  try {
    map.addControl(new BMapGL.ScaleControl());
  } catch {}
  renderMarkers();

  if (state.pois.length) fitToPois();
}

function ensureMapReady() {
  if (state.mapReady && state.map) return true;
  return false;
}

function fitToPois() {
  if (!ensureMapReady()) return;
  const pois = getFilteredPois();
  const pts = pois
    .map((p) => ({ p, pt: new BMapGL.Point(p.lng, p.lat) }))
    .filter((x) => Number.isFinite(x.p.lng) && Number.isFinite(x.p.lat));
  if (!pts.length) return;
  const view = state.map.getViewport(pts.map((x) => x.pt));
  state.map.centerAndZoom(view.center, view.zoom);
}

function renderMarkers() {
  if (!ensureMapReady()) return;
  for (const [, marker] of state.markers) state.map.removeOverlay(marker);
  state.markers.clear();

  const pois = getFilteredPois();
  for (const poi of pois) {
    if (!Number.isFinite(poi.lng) || !Number.isFinite(poi.lat)) continue;
    const pt = new BMapGL.Point(poi.lng, poi.lat);
    const icon = new BMapGL.Icon(markerSvg(getMarkerColor(poi)), new BMapGL.Size(36, 36), {
      anchor: new BMapGL.Size(18, 34),
    });
    const marker = new BMapGL.Marker(pt, { icon });
    marker.addEventListener("click", () => setSelectedPoi(poi.id));
    state.map.addOverlay(marker);
    state.markers.set(poi.id, marker);
  }
}

async function parseFile(file) {
  const name = file.name.toLowerCase();
  if (name.endsWith(".csv")) {
    const buf = await file.arrayBuffer();
    const text = decodeTextBestEffort(buf);
    if (window.Papa) {
      const parsed = Papa.parse(text, { header: true, skipEmptyLines: true });
      const rows = Array.isArray(parsed.data) ? parsed.data : [];
      const headers = parsed.meta?.fields || Object.keys(rows[0] || {});
      return { headers, rows };
    }
    const lines = text.split(/\r?\n/).filter((x) => x.trim().length);
    const headers = lines[0].split(",").map((x) => x.trim());
    const rows = lines.slice(1).map((line) => {
      const parts = line.split(",");
      const obj = {};
      headers.forEach((h, i) => (obj[h] = parts[i] ?? ""));
      return obj;
    });
    return { headers, rows };
  }

  if (name.endsWith(".xlsx") || name.endsWith(".xls")) {
    if (!window.XLSX) {
      throw new Error(
        "xlsx_lib_missing：解析Excel需要XLSX库，但当前未加载（常见原因：网络拦截CDN）。请改用CSV，或允许访问 jsdelivr/unpkg，或将 xlsx.full.min.js 放到 /vendor 后刷新。"
      );
    }
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const sheetName = wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(ws, { defval: "" });
    const headers = json.length ? Object.keys(json[0]) : [];
    return { headers, rows: json };
  }

  throw new Error("unsupported_file");
}

async function identifyColumns(headers, sampleRows) {
  const resp = await fetch("/api/columns/identify", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ headers, sampleRows: sampleRows.slice(0, 3), locale: "zh-CN" }),
  });
  if (!resp.ok) throw new Error("identify_failed");
  const data = await resp.json();
  return data;
}

function fillSelect(sel, headers, value) {
  sel.innerHTML = "";
  const opt0 = document.createElement("option");
  opt0.value = "";
  opt0.textContent = "未选择";
  sel.appendChild(opt0);
  for (const h of headers) {
    const opt = document.createElement("option");
    opt.value = h;
    opt.textContent = h;
    sel.appendChild(opt);
  }
  sel.value = value || "";
}

function setMappingUi(headers, mapping, confidence, llmEnabled) {
  const box = byId("mappingBox");
  box.classList.remove("hidden");

  fillSelect(byId("mapName"), headers, mapping?.name);
  fillSelect(byId("mapAddress"), headers, mapping?.address);
  fillSelect(byId("mapCity"), headers, mapping?.city);
  fillSelect(byId("mapDistrict"), headers, mapping?.district);
  fillSelect(byId("mapLng"), headers, mapping?.lng);
  fillSelect(byId("mapLat"), headers, mapping?.lat);
  fillSelect(byId("mapMerchantStatus"), headers, mapping?.merchantStatus);
  fillSelect(byId("mapPhone"), headers, mapping?.phone);

  const hint = byId("mappingHint");
  const cn = (x) => (typeof x === "number" ? `${Math.round(x * 100)}%` : "-");
  hint.textContent = llmEnabled
    ? `已使用大模型识别（置信度：名称${cn(confidence?.name)} / 地址${cn(confidence?.address)} / 经度${cn(
        confidence?.lng
      )} / 纬度${cn(confidence?.lat)}）`
    : "未配置大模型Key，已使用本地规则进行识别（可手动纠正映射）";
  validateMapping();
}

function getMappingFromUi() {
  return {
    name: byId("mapName").value || "",
    address: byId("mapAddress").value || "",
    city: byId("mapCity").value || "",
    district: byId("mapDistrict").value || "",
    lng: byId("mapLng").value || "",
    lat: byId("mapLat").value || "",
    merchantStatus: byId("mapMerchantStatus").value || "",
    phone: byId("mapPhone").value || "",
  };
}

function validateMapping() {
  const m = getMappingFromUi();
  const ok = Boolean(m.name && m.address && m.lng && m.lat);
  byId("applyMappingBtn").disabled = !ok;
  return ok;
}

function buildPois(rows, mapping) {
  const pois = [];
  const seen = new Set();
  rows.forEach((r, idx) => {
    const name = safeText(r[mapping.name]).trim();
    const address = safeText(r[mapping.address]).trim();
    const city = mapping.city ? safeText(r[mapping.city]).trim() : "";
    const district = mapping.district ? safeText(r[mapping.district]).trim() : "";
    const lng = parseNumber(r[mapping.lng]);
    const lat = parseNumber(r[mapping.lat]);
    const status = "未走访";
    const merchantStatus = mapping.merchantStatus ? normalizeMerchantStatus(r[mapping.merchantStatus]) : "";
    const phone = mapping.phone ? safeText(r[mapping.phone]).trim() : "";

    if (lng === null || lat === null) return;

    const key = `${name}|${address}|${lng}|${lat}`;
    const id = `${idx + 1}_${key}`.replaceAll(/\s+/g, "_").slice(0, 160);
    const stable = id || nanoId();
    if (seen.has(stable)) return;
    seen.add(stable);

    const poi = {
      id: stable,
      rowIdx: idx,
      name,
      address,
      city,
      district,
      lng,
      lat,
      status,
      merchantStatus,
      phone,
      remark: "",
      visitState: toVisitState(status),
      accuracy: { name: "unknown", address: "unknown", coord: "unknown", status: "unknown" },
      photos: [],
      updatedAt: nowIso(),
    };
    pois.push(poi);
  });
  return pois;
}

async function openRecord(poi) {
  byId("recordCard").classList.remove("hidden");
  byId("detailCard").classList.remove("hidden");

  byId("editName").value = poi.name || "";
  byId("editAddress").value = poi.address || "";
  byId("editStatus").value = poi.status || "未走访";

  let merchantStatus = poi.merchantStatus;
  let phone = poi.phone;
  if (!merchantStatus && state.mapping?.merchantStatus && typeof poi.rowIdx === "number" && state.rawRows[poi.rowIdx]) {
    merchantStatus = normalizeMerchantStatus(state.rawRows[poi.rowIdx][state.mapping.merchantStatus]);
  }
  if (!phone && state.mapping?.phone && typeof poi.rowIdx === "number" && state.rawRows[poi.rowIdx]) {
    phone = safeText(state.rawRows[poi.rowIdx][state.mapping.phone]).trim();
  }
  byId("editMerchantStatus").value = merchantStatus || "正常营业";
  byId("editPhone").value = phone || "";
  byId("editLng").value = String(poi.lng ?? "");
  byId("editLat").value = String(poi.lat ?? "");
  byId("editRemark").value = poi.remark || "";

  const acc = poi.accuracy || {};
  byId("accName").value = acc.name || "unknown";
  byId("accAddress").value = acc.address || "unknown";
  byId("accCoord").value = acc.coord || "unknown";
  byId("accStatus").value = acc.status || "unknown";

  byId("photoInput").value = "";
  state.pendingPhotoPoiId = poi.id;
  state.pendingPhotoFiles = [];
  clearPhotoPreviewUrls();
  await renderPhotoList(poi);
  byId("recordHint").textContent = "";
}

function closeRecord() {
  byId("recordCard").classList.add("hidden");
  state.pendingPhotoPoiId = null;
  state.pendingPhotoFiles = [];
  clearPhotoPreviewUrls();
}

async function renderPhotoList(poi) {
  const list = byId("photoList");
  list.innerHTML = "";
  clearPhotoPreviewUrls();

  const saved = poi.photos || [];
  const pending =
    state.pendingPhotoPoiId === poi.id ? state.pendingPhotoFiles.map((f) => ({ file: f, filename: f.name })) : [];

  const renderTile = (url, filename, onRemove) => {
    const tile = document.createElement("div");
    tile.className = "photoTile";

    const img = document.createElement("img");
    img.className = "photoImg";
    img.src = url;
    img.alt = filename;
    tile.appendChild(img);

    const cap = document.createElement("div");
    cap.className = "photoCap";
    cap.textContent = filename;
    tile.appendChild(cap);

    const rm = document.createElement("button");
    rm.className = "photoRemove";
    rm.type = "button";
    rm.textContent = "×";
    rm.addEventListener("click", onRemove);
    tile.appendChild(rm);

    list.appendChild(tile);
  };

  for (const ph of saved) {
    let blob = null;
    try {
      if (!state.db) state.db = await createDb();
      if (ph.blobKey) blob = await idbGet(state.db, "photos", ph.blobKey);
    } catch {}
    if (!blob) continue;
    const url = URL.createObjectURL(blob);
    photoPreviewUrls.push(url);
    renderTile(url, ph.filename, async () => {
      try {
        if (state.db && ph.blobKey) await idbDel(state.db, "photos", ph.blobKey);
      } catch {}
      poi.photos = (poi.photos || []).filter((x) => x.id !== ph.id);
      poi.updatedAt = nowIso();
      saveState();
      await renderPhotoList(poi);
      updateStats();
    });
  }

  for (const p of pending) {
    const url = URL.createObjectURL(p.file);
    photoPreviewUrls.push(url);
    renderTile(url, p.filename, async () => {
      state.pendingPhotoFiles = (state.pendingPhotoFiles || []).filter((x) => x !== p.file);
      await renderPhotoList(poi);
    });
  }
}

async function addPhotos(poi, files) {
  if (!files.length) return;
  if (!state.db) state.db = await createDb();
  const added = [];
  for (const f of files) {
    const blobKey = nanoId();
    await idbPut(state.db, "photos", blobKey, f);
    added.push({
      id: nanoId(),
      filename: f.name,
      mime: f.type,
      size: f.size,
      blobKey,
      createdAt: nowIso(),
    });
  }
  poi.photos = [...(poi.photos || []), ...added];
  poi.updatedAt = nowIso();
}

async function handleSaveRecord() {
  const poi = state.pois.find((p) => p.id === state.selectedPoiId);
  if (!poi) return;

  const name = safeText(byId("editName").value).trim();
  const address = safeText(byId("editAddress").value).trim();
  const status = normalizeStatus(byId("editStatus").value);
  const merchantStatus = byId("editMerchantStatus").value || "正常营业";
  const phone = safeText(byId("editPhone").value).trim();
  const lng = parseNumber(byId("editLng").value);
  const lat = parseNumber(byId("editLat").value);
  const remark = safeText(byId("editRemark").value).trim();
  const accuracy = {
    name: byId("accName").value || "unknown",
    address: byId("accAddress").value || "unknown",
    coord: byId("accCoord").value || "unknown",
    status: byId("accStatus").value || "unknown",
  };

  if (!name || !address || lng === null || lat === null) {
    byId("recordHint").textContent = "请补齐名称、地址与有效经纬度。";
    return;
  }

  const files =
    state.pendingPhotoPoiId === poi.id && state.pendingPhotoFiles.length
      ? state.pendingPhotoFiles
      : Array.from(byId("photoInput").files || []);
  await addPhotos(poi, files);

  poi.name = name;
  poi.address = address;
  poi.status = status;
  poi.merchantStatus = merchantStatus;
  poi.phone = phone;
  poi.lng = lng;
  poi.lat = lat;
  poi.remark = remark;
  const vs = toVisitState(status);
  poi.visitState = vs === "unvisited" ? "visited" : vs;
  poi.accuracy = accuracy;
  poi.updatedAt = nowIso();

  byId("recordHint").textContent = "已提交并更新点位。";
  state.pendingPhotoPoiId = null;
  state.pendingPhotoFiles = [];
  byId("photoInput").value = "";
  renderMarkers();
  fitToPois();
  renderDetail();
  updateStats();
  saveState();
  closeRecord();
  syncRecordToTencentDocs(poi);
}

async function getCurrentPositionBd09() {
  if (window.BMapGL && BMapGL.Geolocation && state.map) {
    return new Promise((resolve, reject) => {
      const geo = new BMapGL.Geolocation();
      geo.getCurrentPosition((r) => {
        if (geo.getStatus() === 0 && r?.point) {
          resolve({ lng: r.point.lng, lat: r.point.lat });
        } else reject(new Error("geo_failed"));
      });
    });
  }

  if (!navigator.geolocation) throw new Error("geo_unsupported");
  return new Promise((resolve, reject) => {
    navigator.geolocation.getCurrentPosition(
      (pos) => resolve({ lng: pos.coords.longitude, lat: pos.coords.latitude }),
      reject,
      { enableHighAccuracy: true, timeout: 8000 }
    );
  });
}

async function planRouteToSelectedPoi() {
  const poi = state.pois.find((p) => p.id === state.selectedPoiId);
  if (!poi) return;
  if (!ensureMapReady()) return;

  const origin = await getCurrentPositionBd09();
  const start = new BMapGL.Point(origin.lng, origin.lat);
  const end = new BMapGL.Point(poi.lng, poi.lat);

  if (!state.driving) {
    state.driving = new BMapGL.DrivingRoute(state.map, { renderOptions: { map: state.map, autoViewport: true } });
  }
  state.driving.search(start, end);
  state.navPlannedForPoiId = poi.id;
  saveState();
  renderDetail();
}

async function jumpToBaiduMapNav() {
  const poi = state.pois.find((p) => p.id === state.selectedPoiId);
  if (!poi) return;
  const origin = await getCurrentPositionBd09();
  const url = buildBaiduNavUrl(origin, { lng: poi.lng, lat: poi.lat }, poi.name);
  location.href = url;
}

async function handleNavClick() {
  const poi = state.pois.find((p) => p.id === state.selectedPoiId);
  if (!poi) return;
  byId("navBtn").disabled = true;
  try {
    if (state.navPlannedForPoiId === poi.id) await jumpToBaiduMapNav();
    else await planRouteToSelectedPoi();
  } catch {
    byId("navHint").textContent = "导航失败：请确认已授权定位，并检查网络与百度地图组件加载情况。";
  } finally {
    byId("navBtn").disabled = false;
  }
}

async function handleExport() {
  byId("exportBtn").disabled = true;
  byId("exportBtn").textContent = "正在导出...";
  try {
    const blob = await poisToXlsxBlob(state.pois);
    const isHtml = blob.type.startsWith("application/vnd.ms-excel");
    const ext = isHtml ? ".xls" : ".xlsx";
    const filename = `POI导出_${new Date().toISOString().slice(0, 10)}${ext}`;
    downloadBlob(filename, blob);
  } catch (e) {
    alert("导出失败：" + (e?.message || "未知错误"));
  } finally {
    byId("exportBtn").disabled = state.pois.length === 0;
    byId("exportBtn").textContent = "导出所有POI信息";
  }
}

function setFileMeta(text) {
  byId("fileMeta").textContent = text;
}

function setFilePill(file) {
  const pill = byId("fileNamePill");
  if (!file) {
    pill.textContent = "拖拽文件到此处，或点击选择";
    pill.classList.remove("filePillActive");
    return;
  }
  pill.textContent = file.name;
  pill.classList.add("filePillActive");
}

function decodeTextBestEffort(buf) {
  const bytes = new Uint8Array(buf);
  if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    return new TextDecoder("utf-8").decode(bytes.subarray(3));
  }
  if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe) {
    return new TextDecoder("utf-16le").decode(bytes.subarray(2));
  }
  if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) {
    return new TextDecoder("utf-16be").decode(bytes.subarray(2));
  }

  const tryDecode = (enc) => {
    try {
      return new TextDecoder(enc, { fatal: false }).decode(bytes);
    } catch {
      return "";
    }
  };
  const score = (s) => {
    if (!s) return -1;
    const rep = (s.match(/\uFFFD/g) || []).length;
    const ratio = rep / Math.max(1, s.length);
    return 1 - ratio;
  };
  const utf8 = tryDecode("utf-8");
  const gb18030 = tryDecode("gb18030");
  const gbk = tryDecode("gbk");
  const candidates = [
    { enc: "utf-8", s: utf8 },
    { enc: "gb18030", s: gb18030 },
    { enc: "gbk", s: gbk },
  ].sort((a, b) => score(b.s) - score(a.s));
  return candidates[0]?.s || utf8 || "";
}

function handleFileSelected(file) {
  state.file = file;
  byId("importBtn").disabled = !file;
  byId("applyMappingBtn").disabled = true;
  byId("mappingBox").classList.add("hidden");
  setFileMeta("");
  setFilePill(file);
  if (file) {
    setFileMeta(`已选择：${file.name}（${Math.round(file.size / 1024)} KB），点击"导入并识别"开始解析`);
  }
}

async function handleImportAndIdentify() {
  const file = state.file;
  if (!file) return;
  byId("importBtn").disabled = true;
  setFileMeta(`正在解析 ${file.name}...`);

  try {
    const { headers, rows } = await parseFile(file);
    state.headers = headers || [];
    state.rawRows = rows || [];

    showLoading();
    const sampleRows = state.rawRows.slice(0, 3);
    const r = await identifyColumns(state.headers, sampleRows);
    hideLoading();
    state.mapping = r?.mapping || {};
    setMappingUi(state.headers, r?.mapping, r?.confidence, r?.llmEnabled);
  } catch (e) {
    hideLoading();
    setFileMeta(`文件解析或字段识别失败：${String(e?.message || e)}`);
  }
}

function clearAllData() {
  state.file = null;
  state.rawRows = [];
  state.headers = [];
  state.mapping = null;
  state.pois = [];
  state.selectedPoiId = null;
  state.navPlannedForPoiId = null;
  if (state.driving) {
    try {
      state.driving.clearResults();
    } catch {}
  }
  state.driving = null;
  state.filters = { city: "", district: "", merchantStatus: "", visitState: "" };
  saveState();
  updateStats();
  renderMarkers();
  renderDetail();
  closeRecord();
  byId("mappingBox").classList.add("hidden");
  byId("uploadCard").classList.remove("compact");
  byId("fileInput").value = "";
  byId("importBtn").disabled = true;
  setFileMeta("");
  setFilePill(null);
}

async function applyMappingAndImport() {
  const ok = validateMapping();
  if (!ok) return;
  const mapping = getMappingFromUi();
  state.mapping = mapping;
  const poiList = buildPois(state.rawRows, mapping);
  state.pois = poiList;
  state.selectedPoiId = poiList[0]?.id || null;
  state.navPlannedForPoiId = null;
  saveState();
  updateStats();
  renderMarkers();
  fitToPois();
  renderDetail();
  byId("mappingBox").classList.add("hidden");
  byId("uploadCard").classList.add("compact");
  setFileMeta("已导入并打点。需要重新映射可点击\u201c清空当前数据\u201d后重新上传。");
  syncPoisToTencentDocs();
}

async function syncPoisToTencentDocs() {
  try {
    const visitStateMap = { unvisited: "\u672a\u8d70\u8bbf", visited: "\u5df2\u8d70\u8bbf", exception: "\u5f02\u5e38/\u65e0\u6cd5\u5230\u8bbf" };
    const rows = state.pois.map((p) => ({
      id: p.id,
      name: p.name,
      address: p.address,
      city: p.city || "",
      district: p.district || "",
      lng: String(p.lng ?? ""),
      lat: String(p.lat ?? ""),
      visitStatus: visitStateMap[p.visitState] || p.visitState || "",
      merchantStatus: p.merchantStatus || "",
      phone: p.phone || "",
      remark: p.remark || "",
      updatedAt: p.updatedAt || "",
      _orig: typeof p.rowIdx === "number" && state.rawRows[p.rowIdx] ? state.rawRows[p.rowIdx] : {},
    }));
    const resp = await fetch("/api/td/sync-import", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ headers: state.headers, rows }),
    });
    const data = await resp.json().catch(() => ({}));
    if (data.queued) {
      setFileMeta("已导入并打点。需要重新映射可点击\u201c清空当前数据\u201d后重新上传。\u2705 数据已排队同步腾讯文档");
    }
  } catch {}
}

async function syncRecordToTencentDocs(poi) {
  try {
    const idx = state.pois.findIndex((p) => p.id === poi.id);
    if (idx < 0) return;
    const visitStateMap = { unvisited: "\u672a\u8d70\u8bbf", visited: "\u5df2\u8d70\u8bbf", exception: "\u5f02\u5e38/\u65e0\u6cd5\u5230\u8bbf" };
    await fetch("/api/td/sync-record", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        rowIndex: idx,
        headerCount: state.headers ? state.headers.length : 0,
        poi: {
          id: poi.id,
          name: poi.name,
          address: poi.address,
          city: poi.city || "",
          district: poi.district || "",
          lng: String(poi.lng ?? ""),
          lat: String(poi.lat ?? ""),
          visitStatus: visitStateMap[poi.visitState] || poi.visitState || "",
          merchantStatus: poi.merchantStatus || "",
          phone: poi.phone || "",
          remark: poi.remark || "",
          updatedAt: poi.updatedAt || "",
        },
      }),
    });
  } catch {}
}

function bindEvents() {
  const nameInput = byId("userNameInput");
  nameInput.value = localStorage.getItem(storageKeys.userName) || "";
  nameInput.addEventListener("input", () => localStorage.setItem(storageKeys.userName, nameInput.value));

  byId("exportBtn").addEventListener("click", handleExport);
  byId("resetBtn").addEventListener("click", clearAllData);

  byId("fileInput").addEventListener("change", (e) => handleFileSelected(e.target.files?.[0] || null));
  byId("importBtn").addEventListener("click", handleImportAndIdentify);

  const pill = byId("fileNamePill");
  pill.addEventListener("click", () => byId("fileInput").click());
  pill.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") byId("fileInput").click();
  });
  const prevent = (e) => {
    e.preventDefault();
    e.stopPropagation();
  };
  ["dragenter", "dragover"].forEach((t) => {
    pill.addEventListener(t, (e) => {
      prevent(e);
      pill.classList.add("filePillDrag");
    });
  });
  ["dragleave", "drop"].forEach((t) => {
    pill.addEventListener(t, (e) => {
      prevent(e);
      pill.classList.remove("filePillDrag");
    });
  });
  pill.addEventListener("drop", async (e) => {
    const f = e.dataTransfer?.files?.[0];
    if (!f) return;
    const dt = new DataTransfer();
    dt.items.add(f);
    byId("fileInput").files = dt.files;
    handleFileSelected(f);
  });

  ["mapName", "mapAddress", "mapCity", "mapDistrict", "mapLng", "mapLat"].forEach((id) => {
    byId(id).addEventListener("change", validateMapping);
  });

  byId("applyMappingBtn").addEventListener("click", applyMappingAndImport);

  byId("recordBtn").addEventListener("click", () => {
    const poi = state.pois.find((p) => p.id === state.selectedPoiId);
    if (poi) openRecord(poi);
  });
  byId("closeRecordBtn").addEventListener("click", closeRecord);
  byId("saveRecordBtn").addEventListener("click", handleSaveRecord);
  byId("photoInput").addEventListener("change", () => {
    const poi = state.pois.find((p) => p.id === state.selectedPoiId);
    if (!poi) return;
    state.pendingPhotoPoiId = poi.id;
    state.pendingPhotoFiles = Array.from(byId("photoInput").files || []);
    renderPhotoList(poi);
  });

  byId("navBtn").addEventListener("click", handleNavClick);

  byId("filterCity").addEventListener("change", () => {
    state.filters.city = byId("filterCity").value;
    applyFilters();
  });
  byId("filterDistrict").addEventListener("change", () => {
    state.filters.district = byId("filterDistrict").value;
    applyFilters();
  });
  byId("filterMerchantStatus").addEventListener("change", () => {
    state.filters.merchantStatus = byId("filterMerchantStatus").value;
    applyFilters();
  });
  byId("filterVisitState").addEventListener("change", () => {
    const v = byId("filterVisitState").value;
    state.filters.visitState = v === "已走访" ? "visited" : v === "未走访" ? "unvisited" : "";
    applyFilters();
  });

  document.addEventListener("visibilitychange", () => {
    if (!document.hidden) renderDetail();
  });
  window.addEventListener("pageshow", () => renderDetail());

  const backdrop = byId("sidebarBackdrop");
  const toggleBtn = byId("sidebarToggleBtn");

  toggleBtn.addEventListener("click", () => {
    document.body.classList.toggle("sidebarOpen");
    const open = document.body.classList.contains("sidebarOpen");
    if (open) backdrop.classList.remove("hidden");
    else backdrop.classList.add("hidden");
  });

  backdrop.addEventListener("click", () => {
    document.body.classList.remove("sidebarOpen");
    backdrop.classList.add("hidden");
  });
}

async function boot() {
  const params = new URLSearchParams(location.search);
  const mpRestore = params.get("_mp_restore");
  if (mpRestore) {
    try {
      const decoded = decodeURIComponent(mpRestore);
      window.__MP_BACKUP_STATE__ = JSON.parse(decoded);
    } catch (e) { }
    if (history.replaceState) {
      const clean = location.href.replace(/[?&]_mp_restore=[^&]*/, "").replace(/\?$/, "");
      history.replaceState(null, "", clean);
    }
  }

  await loadState();
  if (!state.pois.length) {
    if (tryRestoreFromMiniProgram()) {
      console.log("[boot] 已从小程序备份恢复走访数据");
    }
  }
  updateStats();
  restoreFilterUi();
  bindEvents();
  renderDetail();
  ensurePanels();

  if (state.pois.length) {
    byId("mappingBox").classList.add("hidden");
    byId("importBtn").disabled = true;
    let msg = "已恢复上次数据，可直接点选地图点位继续走访。";
    if (typeof wx !== "undefined" && wx.miniProgram) {
      msg += " 数据已双重保护（浏览器本地 + 小程序存储）。";
    }
    setFileMeta(msg);
  }

  const initOnce = () => {
    initMap();
    document.removeEventListener("readystatechange", initOnce);
  };
  if (document.readyState === "complete" || document.readyState === "interactive") initMap();
  else document.addEventListener("readystatechange", initOnce);

  const waitMap = () => {
    if (state.mapReady) return;
    if (window.BMapGL) {
      initMap();
      return;
    }
    setTimeout(waitMap, 350);
  };
  waitMap();
}

boot();
