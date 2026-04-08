import fs from "node:fs";
import path from "node:path";
import https from "node:https";
import { URL, fileURLToPath } from "node:url";

const PORT = Number(process.env.PORT || 3000);
const HOME = process.env.USERPROFILE || process.env.HOME || "";
const CERT_PATH =
  process.env.NUBRA_CERT_PATH || path.join(HOME, ".office-addin-dev-certs", "localhost.crt");
const KEY_PATH =
  process.env.NUBRA_KEY_PATH || path.join(HOME, ".office-addin-dev-certs", "localhost.key");
const LOOPBACK_HOST = "localhost";
const PROXY_ORIGIN = "https://uatapi.nubra.io";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const MIME_TYPES = {
  ".css": "text/css; charset=UTF-8",
  ".html": "text/html; charset=UTF-8",
  ".ico": "image/x-icon",
  ".js": "application/javascript; charset=UTF-8",
  ".json": "application/json; charset=UTF-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".txt": "text/plain; charset=UTF-8",
};

const ALLOWED_CORS_ORIGINS = new Set([
  `https://localhost:${PORT}`,
  `https://127.0.0.1:${PORT}`,
  `https://[::1]:${PORT}`,
]);
const instrumentCache = {
  date: "",
  count: 0,
  headers: [],
  rows: [],
};
const futureStockCache = {
  date: "",
  items: [],
};

function normalizeOrigin(value) {
  const raw = String(value || "").trim();
  if (!raw || raw === "null") return "";
  try {
    const url = new URL(raw);
    return `${url.protocol}//${url.host}`;
  } catch (_error) {
    return "";
  }
}

function resolveCorsOrigin(req) {
  const origin = normalizeOrigin(req?.headers?.origin);
  return ALLOWED_CORS_ORIGINS.has(origin) ? origin : "";
}

function corsHeaders(res) {
  const headers = {
    "Access-Control-Allow-Methods": "GET,POST,OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, Authorization, x-device-id",
    "Access-Control-Max-Age": "600",
  };
  if (res?._corsOrigin) {
    headers["Access-Control-Allow-Origin"] = res._corsOrigin;
    headers.Vary = "Origin";
  }
  return headers;
}

function writeJson(res, statusCode, payload) {
  res.writeHead(statusCode, {
    ...corsHeaders(res),
    "Content-Type": "application/json; charset=UTF-8",
  });
  res.end(JSON.stringify(payload));
}

function clean(value) {
  return String(value || "").trim();
}

function upper(value) {
  return clean(value).toUpperCase();
}

function todayIst() {
  try {
    return new Intl.DateTimeFormat("en-CA", {
      timeZone: "Asia/Kolkata",
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
    }).format(new Date());
  } catch (_error) {
    return new Date().toISOString().slice(0, 10);
  }
}

function nowIstDate() {
  const now = new Date();
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: "Asia/Kolkata",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(now);
  const lookup = Object.fromEntries(parts.map((part) => [part.type, part.value]));
  return `${lookup.year}-${lookup.month}-${lookup.day}`;
}

function istTimeToUtcIso(dateText, timeMinutes) {
  const safeDate = clean(dateText || todayIst());
  const minutes = Number(timeMinutes);
  const hh = String(Math.floor(minutes / 60)).padStart(2, "0");
  const mm = String(minutes % 60).padStart(2, "0");
  const iso = `${safeDate}T${hh}:${mm}:00+05:30`;
  return new Date(iso).toISOString();
}

function readJsonBody(req) {
  return new Promise((resolve, reject) => {
    let body = "";
    req.on("data", (chunk) => {
      body += chunk.toString("utf8");
      if (body.length > 1_000_000) {
        reject(new Error("Payload too large"));
      }
    });
    req.on("end", () => {
      try {
        resolve(body.trim() ? JSON.parse(body) : {});
      } catch (error) {
        reject(error);
      }
    });
    req.on("error", reject);
  });
}

function proxyRequest(upstreamPath, options = {}) {
  return new Promise((resolve, reject) => {
    const target = new URL(PROXY_ORIGIN);
    const payload = options.body ? JSON.stringify(options.body) : "";
    const headers = { ...(options.headers || {}) };
    if (payload) {
      headers["Content-Type"] = "application/json";
      headers["Content-Length"] = Buffer.byteLength(payload);
    }

    const request = https.request(
      {
        protocol: target.protocol,
        hostname: target.hostname,
        port: target.port || 443,
        method: options.method || "GET",
        path: upstreamPath,
        headers,
      },
      (response) => {
        let raw = "";
        response.setEncoding("utf8");
        response.on("data", (chunk) => {
          raw += chunk;
        });
        response.on("end", () => {
          let data = {};
          try {
            data = raw ? JSON.parse(raw) : {};
          } catch (_error) {
            data = { raw };
          }
          resolve({
            statusCode: response.statusCode || 500,
            body: raw,
            data,
          });
        });
      }
    );

    request.on("error", reject);
    if (payload) request.write(payload);
    request.end();
  });
}

async function fetchRefdata(sessionToken, deviceId, date) {
  const upstream = await proxyRequest(`/refdata/refdata/${encodeURIComponent(date)}`, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${sessionToken}`,
      "x-device-id": deviceId,
    },
  });

  if (upstream.statusCode < 200 || upstream.statusCode >= 300) {
    const message = clean(upstream.data?.error || upstream.data?.message || upstream.body || `HTTP ${upstream.statusCode}`);
    throw new Error(`Refdata fetch failed: ${message}`);
  }

  return Array.isArray(upstream.data?.refdata) ? upstream.data.refdata : [];
}

function pickHeaders(items) {
  const preferred = [
    "symbol",
    "trading_symbol",
    "display_name",
    "stock_name",
    "name",
    "instrument_type",
    "series",
    "expiry",
    "lot_size",
    "tick_size",
    "tick_size_paise",
    "ref_id",
    "inst_id",
    "isin",
  ];

  const discovered = new Set();
  for (const item of items) {
    for (const key of Object.keys(item || {})) {
      if (key === "exchange") continue;
      discovered.add(key);
    }
  }

  const ordered = [];
  for (const key of preferred) {
    if (discovered.has(key)) {
      ordered.push(key);
      discovered.delete(key);
    }
  }

  return ordered.concat(Array.from(discovered).sort());
}

function numberOrBlank(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n : "";
}

function numberOrNull(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

async function buildInstrumentDump(sessionToken, deviceId, date = todayIst()) {
  const items = await fetchRefdata(sessionToken, deviceId, date);
  const headers = pickHeaders(items);
  const rows = items.map((item) => {
    const row = {};
    for (const key of headers) {
      row[key] = item?.[key] ?? "";
    }
    return row;
  });

  return { date, count: rows.length, headers, rows };
}

function preferredKey(headers, patterns) {
  const lowerHeaders = Array.isArray(headers) ? headers.map((key) => String(key || "")) : [];
  for (const pattern of patterns) {
    const exact = lowerHeaders.find((key) => key.toLowerCase() === pattern);
    if (exact) return exact;
  }
  for (const pattern of patterns) {
    const partial = lowerHeaders.find((key) => key.toLowerCase().includes(pattern));
    if (partial) return partial;
  }
  return "";
}

function symbolKey(headers) {
  return preferredKey(headers, [
    "trading_symbol",
    "tradingsymbol",
    "symbol",
    "ticker",
    "scrip",
    "security",
    "instrument",
    "name",
  ]);
}

function displayKey(headers) {
  return preferredKey(headers, [
    "display_name",
    "stock_name",
    "company_name",
    "security_name",
    "instrument_name",
    "description",
    "desc",
    "name",
  ]);
}

function refIdKey(headers) {
  return preferredKey(headers, [
    "ref_id",
    "refid",
    "inst_id",
    "instrument_id",
    "token",
    "security_id",
  ]);
}

function fallbackText(row) {
  const values = Object.values(row || {})
    .map((value) => clean(value))
    .filter(Boolean);
  const candidates = values.filter((value) =>
    /^[A-Z0-9\-_.]{3,40}$/i.test(value) || /[A-Za-z]/.test(value)
  );
  return candidates.sort((a, b) => a.length - b.length)[0] || values[0] || "";
}

function exchangeKey(row) {
  return upper(row?.exchange || row?.segment || row?.market || "");
}

function derivativeKey(row) {
  return upper(row?.derivative_type || row?.instrument_type || row?.security_type || row?.segment || row?.series || "");
}

function optionTypeKey(row) {
  return upper(row?.option_type || row?.right || row?.side || "");
}

function normalizeExpiry(value) {
  const raw = clean(value);
  if (!raw) return "";
  const digits = raw.replace(/\D/g, "");
  if (digits.length >= 8) return digits.slice(0, 8);
  const ms = Date.parse(raw);
  if (!Number.isNaN(ms)) {
    return new Date(ms).toISOString().slice(0, 10).replace(/-/g, "");
  }
  return upper(raw);
}

function displayExpiry(expiryKey) {
  if (!/^\d{8}$/.test(expiryKey)) return expiryKey;
  return `${expiryKey.slice(0, 4)}-${expiryKey.slice(4, 6)}-${expiryKey.slice(6, 8)}`;
}

function inferUnderlying(row, headers) {
  const values = [
    row?.asset,
    row?.underlying,
    row?.stock_name,
    row?.display_name,
    row?.company_name,
    row?.name,
    row?.symbol,
    row?.[symbolKey(headers)],
  ].map((value) => upper(value)).filter(Boolean);
  return values[0] || "";
}

function isFutureRow(row) {
  const derivative = derivativeKey(row);
  const symbol = upper(row?.symbol || row?.trading_symbol || "");
  const optionType = optionTypeKey(row);
  if (optionType === "CE" || optionType === "PE") return false;
  return derivative.includes("FUT") || symbol.endsWith("FUT");
}

function isCashStockRow(row) {
  const derivative = derivativeKey(row);
  const expiry = normalizeExpiry(row?.expiry);
  const optionType = optionTypeKey(row);
  const exchange = exchangeKey(row);
  if (optionType === "CE" || optionType === "PE") return false;
  if (derivative.includes("FUT") || derivative.includes("OPT")) return false;
  if (expiry) return false;
  return exchange === "NSE" || exchange === "BSE" || !exchange;
}

function sortByExpiryAsc(a, b) {
  return normalizeExpiry(a?.expiry).localeCompare(normalizeExpiry(b?.expiry));
}

function chooseCashRow(rows, underlying, headers) {
  const matches = rows.filter((row) => isCashStockRow(row) && inferUnderlying(row, headers) === underlying);
  if (!matches.length) return null;
  const preferred = matches.find((row) => exchangeKey(row) === "NSE")
    || matches.find((row) => exchangeKey(row) === "BSE")
    || matches[0];
  return preferred || null;
}

function buildFutureStockCache() {
  const headers = Array.isArray(instrumentCache.headers) ? instrumentCache.headers : [];
  const rows = Array.isArray(instrumentCache.rows) ? instrumentCache.rows : [];
  const futuresByUnderlying = new Map();

  for (const row of rows) {
    if (!isFutureRow(row)) continue;
    const underlying = inferUnderlying(row, headers);
    if (!underlying) continue;
    const bucket = futuresByUnderlying.get(underlying) || [];
    bucket.push(row);
    futuresByUnderlying.set(underlying, bucket);
  }

  const items = [];
  for (const [underlying, futureRows] of futuresByUnderlying.entries()) {
    const sortedFutures = futureRows.slice().sort(sortByExpiryAsc);
    if (!sortedFutures.length) continue;
    const stockRow = chooseCashRow(rows, underlying, headers);
    const stockSymbol = clean(stockRow?.[symbolKey(headers)] || stockRow?.symbol || stockRow?.trading_symbol || underlying);
    const display = clean(stockRow?.[displayKey(headers)] || stockRow?.stock_name || stockRow?.name || underlying);
    items.push({
      underlying,
      display,
      label: display && upper(display) !== underlying ? `${underlying} | ${display}` : underlying,
      stock: stockRow ? {
        symbol: stockSymbol || underlying,
        ref_id: numberOrNull(stockRow?.[refIdKey(headers)] ?? stockRow?.ref_id ?? ""),
        exchange: exchangeKey(stockRow) || "NSE",
      } : {
        symbol: underlying,
        ref_id: null,
        exchange: "NSE",
      },
      futures: sortedFutures.slice(0, 2).map((row) => ({
        symbol: clean(row?.[symbolKey(headers)] || row?.symbol || row?.trading_symbol || ""),
        ref_id: numberOrNull(row?.[refIdKey(headers)] ?? row?.ref_id ?? ""),
        expiry: displayExpiry(normalizeExpiry(row?.expiry)),
        exchange: exchangeKey(row),
      })),
    });
  }

  futureStockCache.date = instrumentCache.date || "";
  futureStockCache.items = items.sort((a, b) => a.underlying.localeCompare(b.underlying));
}

function normalizeInstrumentRow(row, headers) {
  const symbol = clean(row?.[symbolKey(headers)]) || fallbackText(row);
  const display = clean(row?.[displayKey(headers)] || row?.display_name || row?.stock_name || row?.name || symbol);
  const label = display && display !== symbol ? `${symbol} | ${display}` : symbol;
  return {
    symbol,
    display,
    label,
    ref_id: row?.[refIdKey(headers)] ?? row?.ref_id ?? "",
  };
}

function searchCachedInstruments(query, limit = 25) {
  const probe = upper(query);
  const rows = Array.isArray(futureStockCache.items) ? futureStockCache.items : [];
  const startsWithMatches = [];
  const containsMatches = [];
  const seen = new Set();

  for (const row of rows) {
    const item = {
      symbol: row.underlying,
      display: row.display || row.underlying,
      label: row.label || row.underlying,
      stock: row.stock,
      futures: row.futures,
    };
    if (!item.symbol) continue;
    if (probe) {
      const haystacks = [
        item.symbol,
        item.display,
        item.label,
        clean(item.stock?.symbol || ""),
        ...(Array.isArray(item.futures) ? item.futures.flatMap((future) => [future.symbol, future.expiry]) : []),
      ].filter(Boolean).map((value) => upper(value));
      const hit = haystacks.some((value) => value.includes(probe));
      if (!hit) continue;
    }

    const token = upper(item.symbol);
    if (seen.has(token)) continue;
    seen.add(token);

    const startsWith = upper(item.symbol).startsWith(probe)
      || upper(item.display).startsWith(probe)
      || upper(item.label).startsWith(probe);

    if (!probe || startsWith) {
      startsWithMatches.push(item);
    } else {
      containsMatches.push(item);
    }

    if (startsWithMatches.length + containsMatches.length >= limit) break;
  }

  return startsWithMatches.concat(containsMatches).slice(0, limit);
}

async function mapWithConcurrency(items, limit, worker) {
  const source = Array.isArray(items) ? items : [];
  const max = Math.max(1, Number(limit) || 1);
  const results = new Array(source.length);
  let cursor = 0;

  async function runOne() {
    while (cursor < source.length) {
      const index = cursor++;
      results[index] = await worker(source[index], index);
    }
  }

  await Promise.all(Array.from({ length: Math.min(max, source.length || 1) }, () => runOne()));
  return results;
}

function findFutureStock(query) {
  const probe = upper(query);
  return (futureStockCache.items || []).find((item) =>
    upper(item.underlying) === probe
    || upper(item.label) === probe
    || upper(item.display) === probe
    || upper(item.stock?.symbol || "") === probe
  ) || null;
}

function timestampToIso(value) {
  const n = Number(value);
  if (Number.isFinite(n) && n > 0) {
    const ms = n > 10_000_000_000 ? n : n * 1000;
    return new Date(ms).toISOString();
  }
  const raw = clean(value);
  if (!raw) return "";
  const parsed = Date.parse(raw);
  return Number.isNaN(parsed) ? "" : new Date(parsed).toISOString();
}

function snapshotFromBook(raw, fallbackSymbol) {
  const book = raw?.orderBook || raw || {};
  return {
    symbol: clean(book?.symbol || fallbackSymbol),
    ref_id: numberOrNull(book?.ref_id ?? book?.refId ?? book?.instrument_id),
    ltp: numberOrNull(book?.ltp ?? book?.last_traded_price ?? book?.price),
    prev_close: numberOrNull(book?.prev_close ?? book?.prevClose ?? book?.previous_close ?? book?.previousClose),
    oi: numberOrNull(book?.oi ?? book?.open_interest ?? book?.openInterest ?? book?.volume_oi ?? book?.volumeOi),
    prev_oi: numberOrNull(book?.prev_oi ?? book?.previous_open_interest ?? book?.previousOpenInterest),
    as_of: timestampToIso(book?.ts ?? book?.timestamp ?? raw?.ts ?? raw?.timestamp) || new Date().toISOString(),
  };
}

async function fetchMarketSnapshot(sessionToken, deviceId, instrument) {
  const refId = Number(instrument?.ref_id);
  const symbol = clean(instrument?.symbol);
  if (Number.isInteger(refId) && refId > 0) {
    try {
      const upstream = await proxyRequest(`/orderbooks/${encodeURIComponent(String(refId))}?levels=1`, {
        method: "GET",
        headers: {
          Authorization: `Bearer ${sessionToken}`,
          "x-device-id": deviceId,
        },
      });
      if (upstream.statusCode >= 200 && upstream.statusCode < 300) {
        return snapshotFromBook(upstream.data, symbol);
      }
    } catch (_error) {
      // fallback below
    }
  }

  if (symbol) {
    try {
      const upstream = await proxyRequest(`/optionchains/${encodeURIComponent(symbol)}/price`, {
        method: "GET",
        headers: {
          Authorization: `Bearer ${sessionToken}`,
          "x-device-id": deviceId,
        },
      });
      if (upstream.statusCode >= 200 && upstream.statusCode < 300) {
        return {
          symbol,
          ref_id: Number.isInteger(refId) && refId > 0 ? refId : null,
          ltp: numberOrNull(upstream.data?.price ?? upstream.data?.ltp),
          prev_close: numberOrNull(upstream.data?.prev_close ?? upstream.data?.prevClose ?? upstream.data?.previous_close),
          oi: numberOrNull(upstream.data?.oi ?? upstream.data?.open_interest ?? upstream.data?.volume_oi),
          prev_oi: numberOrNull(upstream.data?.prev_oi ?? upstream.data?.previous_open_interest),
          as_of: new Date().toISOString(),
        };
      }
    } catch (_error) {
      // ignore
    }
  }

  return {
    symbol,
    ref_id: Number.isInteger(refId) && refId > 0 ? refId : null,
    ltp: null,
    prev_close: null,
    oi: null,
    prev_oi: null,
    as_of: new Date().toISOString(),
  };
}

async function fetchHistoricalSeries(sessionToken, deviceId, requestBody) {
  const upstream = await proxyRequest("/charts/timeseries", {
    method: "POST",
    body: requestBody,
    headers: {
      Authorization: `Bearer ${sessionToken}`,
      "x-device-id": deviceId,
    },
  });

  if (upstream.statusCode < 200 || upstream.statusCode >= 300) {
    const message = clean(upstream.data?.error || upstream.data?.message || upstream.body || `HTTP ${upstream.statusCode}`);
    throw new Error(`Historical fetch failed: ${message}`);
  }

  return upstream.data || {};
}

function tsToMs(value) {
  const n = Number(value);
  if (!Number.isFinite(n)) return null;
  if (Math.abs(n) >= 1e18) return Math.round(n / 1e6);
  if (Math.abs(n) >= 1e15) return Math.round(n / 1e3);
  if (Math.abs(n) >= 1e12) return Math.round(n);
  if (Math.abs(n) >= 1e9) return Math.round(n * 1000);
  return Math.round(n);
}

function extractSeriesMap(payload) {
  const out = new Map();
  const result = Array.isArray(payload?.result) ? payload.result : [];
  for (const block of result) {
    const values = Array.isArray(block?.values) ? block.values : [];
    for (const entry of values) {
      if (!entry || typeof entry !== "object") continue;
      for (const [symbol, stockChart] of Object.entries(entry)) {
        out.set(upper(symbol), stockChart || {});
      }
    }
  }
  return out;
}

function extractPointValue(points, targetMs) {
  const list = Array.isArray(points) ? points : [];
  for (const point of list) {
    const ms = tsToMs(point?.ts ?? point?.timestamp);
    if (ms === targetMs) {
      return numberOrNull(point?.v ?? point?.value);
    }
  }
  return null;
}

function extractLastPointAtOrBefore(points, targetMs) {
  const list = Array.isArray(points) ? points : [];
  let best = null;
  for (const point of list) {
    const ms = tsToMs(point?.ts ?? point?.timestamp);
    if (!Number.isFinite(ms) || ms > targetMs) continue;
    const value = numberOrNull(point?.v ?? point?.value);
    if (value === null) continue;
    if (!best || ms > best.ms) {
      best = { ms, value };
    }
  }
  return best;
}

function intervalSeriesValue(series, endMs) {
  const closePoint = extractLastPointAtOrBefore(series?.close, endMs);
  const oiPoint = extractLastPointAtOrBefore(series?.cumulative_oi, endMs)
    || extractLastPointAtOrBefore(series?.cumulative_fut_oi, endMs);
  return {
    close: closePoint?.value ?? null,
    close_ts: closePoint?.ms ?? null,
    cumulative_oi: oiPoint?.value ?? null,
    oi_ts: oiPoint?.ms ?? null,
  };
}

function diffOrNull(currentValue, previousValue) {
  if (!Number.isFinite(currentValue) || !Number.isFinite(previousValue)) return null;
  return currentValue - previousValue;
}

function mapByUnderlying(rows) {
  const out = new Map();
  for (const row of Array.isArray(rows) ? rows : []) {
    out.set(upper(row?.underlying || row?.symbol || ""), row);
  }
  return out;
}

async function fetchUnderlyingOverview(sessionToken, deviceId, payload) {
  const underlying = upper(payload?.underlying);
  const cached = findFutureStock(underlying);
  if (!cached) {
    throw new Error("underlying not found in futures cache");
  }

  const stockSnapshot = await fetchMarketSnapshot(sessionToken, deviceId, cached.stock);
  const currentFuture = cached.futures?.[0] || null;
  const nextFuture = cached.futures?.[1] || null;
  const thirdFuture = cached.futures?.[2] || null;
  const currentFutureSnapshot = currentFuture ? await fetchMarketSnapshot(sessionToken, deviceId, currentFuture) : null;
  const nextFutureSnapshot = nextFuture ? await fetchMarketSnapshot(sessionToken, deviceId, nextFuture) : null;
  const thirdFutureSnapshot = thirdFuture ? await fetchMarketSnapshot(sessionToken, deviceId, thirdFuture) : null;

  return {
    underlying: cached.underlying,
    display: cached.display,
    stock: {
      symbol: cached.stock?.symbol || cached.underlying,
      ref_id: cached.stock?.ref_id ?? null,
      prev_close: stockSnapshot?.prev_close ?? null,
      curr_ltp: stockSnapshot?.ltp ?? null,
      ltp_as_of: stockSnapshot?.as_of || "",
    },
    current_future: currentFuture ? {
      symbol: currentFuture.symbol,
      expiry: currentFuture.expiry,
      ref_id: currentFuture.ref_id ?? null,
      prev_close: currentFutureSnapshot?.prev_close ?? null,
      curr_ltp: currentFutureSnapshot?.ltp ?? null,
      ltp_as_of: currentFutureSnapshot?.as_of || "",
      oi_yest_close: currentFutureSnapshot?.prev_oi ?? null,
      oi_current: currentFutureSnapshot?.oi ?? null,
      oi_as_of: currentFutureSnapshot?.as_of || "",
    } : null,
    next_future: nextFuture ? {
      symbol: nextFuture.symbol,
      expiry: nextFuture.expiry,
      ref_id: nextFuture.ref_id ?? null,
      prev_close: nextFutureSnapshot?.prev_close ?? null,
      curr_ltp: nextFutureSnapshot?.ltp ?? null,
      ltp_as_of: nextFutureSnapshot?.as_of || "",
      oi_yest_close: nextFutureSnapshot?.prev_oi ?? null,
      oi_current: nextFutureSnapshot?.oi ?? null,
      oi_as_of: nextFutureSnapshot?.as_of || "",
    } : null,
    third_future: thirdFuture ? {
      symbol: thirdFuture.symbol,
      expiry: thirdFuture.expiry,
      ref_id: thirdFuture.ref_id ?? null,
      prev_close: thirdFutureSnapshot?.prev_close ?? null,
      curr_ltp: thirdFutureSnapshot?.ltp ?? null,
      ltp_as_of: thirdFutureSnapshot?.as_of || "",
      oi_yest_close: thirdFutureSnapshot?.prev_oi ?? null,
      oi_current: thirdFutureSnapshot?.oi ?? null,
      oi_as_of: thirdFutureSnapshot?.as_of || "",
    } : null,
    updated_at: new Date().toISOString(),
  };
}

async function captureIntervalUniverse(sessionToken, deviceId) {
  const items = Array.isArray(futureStockCache.items) ? futureStockCache.items : [];
  const snapshots = await mapWithConcurrency(items, 6, async (item) => {
    try {
      return await fetchUnderlyingOverview(sessionToken, deviceId, { underlying: item.underlying });
    } catch (error) {
      return {
        underlying: item.underlying,
        display: item.display,
        stock: item.stock || null,
        current_future: item.futures?.[0] || null,
        next_future: item.futures?.[1] || null,
        third_future: item.futures?.[2] || null,
        updated_at: new Date().toISOString(),
        error: error.message || String(error),
      };
    }
  });

  return {
    date: futureStockCache.date || todayIst(),
    count: snapshots.length,
    items: snapshots,
    captured_at: new Date().toISOString(),
  };
}

function intervalMinutesLabelToApi(intervalMinutes) {
  const n = Number(intervalMinutes);
  const mapping = new Map([
    [1, "1m"],
    [2, "2m"],
    [3, "3m"],
    [5, "5m"],
    [15, "15m"],
    [30, "30m"],
    [60, "1h"],
  ]);
  return mapping.get(n) || `${n}m`;
}

function chunkArray(items, size) {
  const out = [];
  for (let i = 0; i < items.length; i += size) {
    out.push(items.slice(i, i + size));
  }
  return out;
}

function currentIstMinutes() {
  const parts = new Intl.DateTimeFormat("en-IN", {
    timeZone: "Asia/Kolkata",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  }).formatToParts(new Date());
  const lookup = Object.fromEntries(parts.map((part) => [part.type, part.value]));
  return Number(lookup.hour || 0) * 60 + Number(lookup.minute || 0);
}

async function fetchHistoricalSeriesBatched(sessionToken, deviceId, type, exchange, symbols, fields, startDateUtc, endDateUtc, intervalMinutes) {
  const apiInterval = intervalMinutesLabelToApi(intervalMinutes);
  const chunks = chunkArray(Array.from(new Set((symbols || []).map((symbol) => upper(symbol)).filter(Boolean))), 5);
  const merged = new Map();
  for (const chunk of chunks) {
    if (!chunk.length) continue;
    try {
      const payload = {
        query: [
          {
            exchange,
            type,
            values: chunk,
            fields,
            startDate: startDateUtc,
            endDate: endDateUtc,
            interval: apiInterval,
            intraDay: true,
            realTime: false,
          },
        ],
      };
      const data = await fetchHistoricalSeries(sessionToken, deviceId, payload);
      const seriesMap = extractSeriesMap(data);
      for (const [symbol, series] of seriesMap.entries()) {
        merged.set(symbol, series);
      }
    } catch (error) {
      if (!String(error.message || error).toLowerCase().includes("ticker not found")) {
        throw error;
      }
      for (const symbol of chunk) {
        try {
          const payload = {
            query: [
              {
                exchange,
                type,
                values: [symbol],
                fields,
                startDate: startDateUtc,
                endDate: endDateUtc,
                interval: apiInterval,
                intraDay: true,
                realTime: false,
              },
            ],
          };
          const data = await fetchHistoricalSeries(sessionToken, deviceId, payload);
          const seriesMap = extractSeriesMap(data);
          for (const [resolvedSymbol, series] of seriesMap.entries()) {
            merged.set(resolvedSymbol, series);
          }
        } catch (_innerError) {
          // Skip unsupported historical ticker and continue building the sheet.
        }
      }
    }
  }
  return merged;
}

async function fetchHistoricalSeriesWithFallbacks(sessionToken, deviceId, attempts, fields, startDateUtc, endDateUtc, intervalMinutes) {
  let lastError = null;
  for (const attempt of attempts) {
    const symbols = Array.from(new Set((attempt?.symbols || []).map((symbol) => upper(symbol)).filter(Boolean)));
    if (!symbols.length) continue;
    try {
      const data = await fetchHistoricalSeriesBatched(
        sessionToken,
        deviceId,
        attempt.type,
        attempt.exchange,
        symbols,
        fields,
        startDateUtc,
        endDateUtc,
        intervalMinutes
      );
      if (data.size) {
        return { data, resolved: { type: attempt.type, exchange: attempt.exchange, symbols } };
      }
    } catch (error) {
      lastError = error;
      const text = String(error?.message || error).toLowerCase();
      if (!text.includes("invalid exchange") && !text.includes("ticker not found")) {
        throw error;
      }
    }
  }
  if (lastError) throw lastError;
  return { data: new Map(), resolved: null };
}

async function buildIntervalUniverse(sessionToken, deviceId, payload) {
  const cachedItems = Array.isArray(futureStockCache.items) ? futureStockCache.items : [];
  const requestedUnderlyings = new Set(
    (Array.isArray(payload?.underlyings) ? payload.underlyings : [])
      .map((value) => upper(value))
      .filter(Boolean)
  );
  const items = requestedUnderlyings.size
    ? cachedItems.filter((item) => requestedUnderlyings.has(upper(item?.underlying)))
    : cachedItems;
  const intervals = Array.isArray(payload?.intervals) ? payload.intervals : [];
  const intervalMinutes = Number(payload?.interval_minutes || 15);
  const tradeDate = clean(payload?.date || nowIstDate());
  if (!items.length) {
    throw new Error(requestedUnderlyings.size ? "No selected LiveTracker stocks are available in cache" : "F&O stock cache is empty");
  }
  if (!intervals.length) {
    throw new Error("No intervals provided");
  }

  const nowMinutes = currentIstMinutes();
  const latestIntervalIndex = intervals.findIndex((interval) => nowMinutes >= Number(interval.start) && nowMinutes < Number(interval.end));
  const completedIntervals = intervals.filter((interval) => interval.index !== latestIntervalIndex);

  const earliestStart = Math.min(...intervals.map((interval) => Number(interval.start)));
  const latestCompletedEnd = completedIntervals.length
    ? Math.max(...completedIntervals.map((interval) => Number(interval.end)))
    : Math.min(...intervals.map((interval) => Number(interval.end)));

  const startDateUtc = istTimeToUtcIso(tradeDate, earliestStart);
  const endDateUtc = istTimeToUtcIso(tradeDate, latestCompletedEnd);

  const stockSymbols = items.map((item) => item.stock?.symbol).filter(Boolean);
  const frontFutureSymbols = items.map((item) => item.futures?.[0]?.symbol).filter(Boolean);
  const futureUnderlyingSymbols = items.map((item) => item.underlying).filter(Boolean);
  const stockSeries = completedIntervals.length
    ? await fetchHistoricalSeriesBatched(sessionToken, deviceId, "STOCK", "NSE", stockSymbols, ["close"], startDateUtc, endDateUtc, intervalMinutes)
    : new Map();
  const futureSeriesResult = completedIntervals.length
    ? await fetchHistoricalSeriesWithFallbacks(
      sessionToken,
      deviceId,
      [
        { type: "FUT", exchange: "NSE", symbols: frontFutureSymbols },
        { type: "FUT", exchange: "NSE", symbols: futureUnderlyingSymbols },
        { type: "FUT", exchange: "NFO", symbols: frontFutureSymbols },
        { type: "FUT", exchange: "NFO", symbols: futureUnderlyingSymbols },
      ],
      ["close", "cumulative_oi"],
      startDateUtc,
      endDateUtc,
      intervalMinutes
    )
    : { data: new Map(), resolved: null };
  const futureSeries = futureSeriesResult.data;
  const futureSeriesResolvedSymbols = new Set((futureSeriesResult.resolved?.symbols || []).map((symbol) => upper(symbol)));

  const liveSnapshots = latestIntervalIndex >= 0
    ? mapByUnderlying(await mapWithConcurrency(items, 6, async (item) => fetchUnderlyingOverview(sessionToken, deviceId, { underlying: item.underlying })))
    : new Map();
  const baselineSnapshots = liveSnapshots.size
    ? liveSnapshots
    : mapByUnderlying(await mapWithConcurrency(items, 6, async (item) => fetchUnderlyingOverview(sessionToken, deviceId, { underlying: item.underlying })));

  const snapshotsByInterval = {};
  for (const interval of intervals) {
    const key = String(interval.index);
    snapshotsByInterval[key] = {};
    const intervalEndMs = Date.parse(istTimeToUtcIso(tradeDate, Number(interval.end)));
    for (const item of items) {
      if (interval.index === latestIntervalIndex) {
        const live = liveSnapshots.get(upper(item.underlying));
        if (live) snapshotsByInterval[key][item.underlying] = live;
        continue;
      }
      const stockData = intervalSeriesValue(stockSeries.get(upper(item.stock?.symbol || "")), intervalEndMs);
      const futureLookupKey = futureSeriesResolvedSymbols.has(upper(item.futures?.[0]?.symbol || ""))
        ? upper(item.futures?.[0]?.symbol || "")
        : upper(item.underlying || "");
      const futureData = intervalSeriesValue(futureSeries.get(futureLookupKey), intervalEndMs);
      const previousFutureData = intervalSeriesValue(futureSeries.get(futureLookupKey), intervalEndMs - intervalMinutes * 60 * 1000);
      const baseline = baselineSnapshots.get(upper(item.underlying));
      snapshotsByInterval[key][item.underlying] = {
        underlying: item.underlying,
        display: item.display,
        stock: {
          symbol: item.stock?.symbol || item.underlying,
          ref_id: item.stock?.ref_id ?? null,
          prev_close: baseline?.stock?.prev_close ?? null,
          curr_ltp: stockData.close,
          ltp_as_of: stockData.close_ts ? new Date(stockData.close_ts).toISOString() : "",
        },
        current_future: item.futures?.[0] ? {
          symbol: item.futures[0].symbol,
          expiry: item.futures[0].expiry,
          ref_id: item.futures[0].ref_id ?? null,
          prev_close: baseline?.current_future?.prev_close ?? null,
          curr_ltp: futureData.close,
          ltp_as_of: futureData.close_ts ? new Date(futureData.close_ts).toISOString() : "",
          oi_yest_close: baseline?.current_future?.oi_yest_close ?? previousFutureData.cumulative_oi,
          oi_current: futureData.cumulative_oi,
          oi_as_of: futureData.oi_ts ? new Date(futureData.oi_ts).toISOString() : "",
        } : null,
        next_future: item.futures?.[1] ? { ...item.futures[1] } : null,
        third_future: item.futures?.[2] ? { ...item.futures[2] } : null,
        updated_at: new Date().toISOString(),
      };
    }
  }

  return {
    date: tradeDate,
    interval_minutes: intervalMinutes,
    latest_interval_index: latestIntervalIndex,
    count: items.length,
    items,
    intervals,
    snapshots: snapshotsByInterval,
    built_at: new Date().toISOString(),
  };
}

async function fetchTrackerQuote(sessionToken, deviceId, payload) {
  const refId = Number(payload?.ref_id);
  const symbol = clean(payload?.symbol);
  let ltp = "";
  let oi = "";
  let source = "";

  if (Number.isInteger(refId) && refId > 0) {
    try {
      const upstream = await proxyRequest(`/orderbooks/${encodeURIComponent(String(refId))}?levels=1`, {
        method: "GET",
        headers: {
          Authorization: `Bearer ${sessionToken}`,
          "x-device-id": deviceId,
        },
      });

      if (upstream.statusCode >= 200 && upstream.statusCode < 300) {
        const book = upstream.data?.orderBook || upstream.data || {};
        ltp = numberOrBlank(book?.ltp ?? book?.last_traded_price ?? book?.price);
        oi = numberOrBlank(
          book?.oi
          ?? book?.open_interest
          ?? book?.openInterest
          ?? book?.volume_oi
          ?? book?.volumeOi
        );
        source = "orderbook";
      }
    } catch (_error) {
      // fall through to symbol price
    }
  }

  if (ltp === "" && symbol) {
    try {
      const upstream = await proxyRequest(`/optionchains/${encodeURIComponent(symbol)}/price`, {
        method: "GET",
        headers: {
          Authorization: `Bearer ${sessionToken}`,
          "x-device-id": deviceId,
        },
      });

      if (upstream.statusCode >= 200 && upstream.statusCode < 300) {
        ltp = numberOrBlank(upstream.data?.price ?? upstream.data?.ltp);
        oi = oi === "" ? numberOrBlank(
          upstream.data?.oi
          ?? upstream.data?.open_interest
          ?? upstream.data?.openInterest
          ?? upstream.data?.volume_oi
          ?? upstream.data?.volumeOi
        ) : oi;
        source = source || "symbol_price";
      }
    } catch (_error) {
      // ignore
    }
  }

  return {
    symbol,
    ref_id: Number.isInteger(refId) && refId > 0 ? refId : "",
    ltp,
    oi,
    source,
    updated_at: new Date().toISOString(),
  };
}

async function routeProxy(req, res, urlObj) {
  const pathname = urlObj.pathname || "/";
  if (!pathname.startsWith("/proxy/insti")) {
    return false;
  }

  const upstreamPath = pathname.replace("/proxy/insti", "") + (urlObj.search || "");
  const body = req.method === "POST" ? await readJsonBody(req) : undefined;
  const upstream = await proxyRequest(upstreamPath, {
    method: req.method || "GET",
    body,
    headers: {
      Authorization: req.headers.authorization || "",
      "x-device-id": req.headers["x-device-id"] || "",
    },
  });

  res.writeHead(upstream.statusCode, {
    ...corsHeaders(res),
    "Content-Type": "application/json; charset=UTF-8",
  });
  res.end(upstream.body);
  return true;
}

async function routeLocalApi(req, res, urlObj) {
  const pathname = urlObj.pathname || "/";
  const method = String(req.method || "GET").toUpperCase();

  if (pathname === "/api/instruments/fno-universe" && method === "POST") {
    const payload = await readJsonBody(req);
    const sessionToken = clean((req.headers.authorization || "").replace(/^Bearer\s+/i, ""));
    const deviceId = clean(req.headers["x-device-id"] || "");
    const date = clean(payload?.date || todayIst());

    if (!sessionToken) {
      writeJson(res, 401, { error: "session token is required" });
      return true;
    }

    const result = await buildInstrumentDump(sessionToken, deviceId, date);
    instrumentCache.date = result.date;
    instrumentCache.count = result.count;
    instrumentCache.headers = result.headers;
    instrumentCache.rows = result.rows;
    buildFutureStockCache();
    writeJson(res, 200, {
      date: result.date,
      count: result.count,
      stocks_with_futures: futureStockCache.items.length,
      cached: true,
    });
    return true;
  }

  if (pathname === "/api/instruments/search" && method === "GET") {
    const query = clean(urlObj.searchParams.get("q") || "");
    const limit = Math.max(1, Math.min(50, Number(urlObj.searchParams.get("limit") || 25)));
    writeJson(res, 200, {
      date: futureStockCache.date || "",
      count: futureStockCache.items.length || 0,
      cached: Array.isArray(futureStockCache.items) && futureStockCache.items.length > 0,
      items: searchCachedInstruments(query, limit),
    });
    return true;
  }

  if (pathname === "/api/instruments/resolve" && method === "GET") {
    const query = clean(urlObj.searchParams.get("q") || "");
    const match = findFutureStock(query) || searchCachedInstruments(query, 1)[0] || null;
    if (!match) {
      writeJson(res, 404, { error: "instrument not found in cache" });
      return true;
    }
    writeJson(res, 200, {
      underlying: match.underlying || match.symbol,
      symbol: match.underlying || match.symbol,
      display: match.display || match.label || match.symbol,
      label: match.label || match.symbol,
      stock: match.stock || null,
      futures: Array.isArray(match.futures) ? match.futures : [],
    });
    return true;
  }

  if (pathname === "/api/instruments/all-fno-stocks" && method === "GET") {
    writeJson(res, 200, {
      date: futureStockCache.date || "",
      count: futureStockCache.items.length || 0,
      cached: Array.isArray(futureStockCache.items) && futureStockCache.items.length > 0,
      items: futureStockCache.items,
    });
    return true;
  }

  if (pathname === "/api/tracker/overview" && method === "POST") {
    const payload = await readJsonBody(req);
    const sessionToken = clean((req.headers.authorization || "").replace(/^Bearer\s+/i, ""));
    const deviceId = clean(req.headers["x-device-id"] || "");

    if (!sessionToken) {
      writeJson(res, 401, { error: "session token is required" });
      return true;
    }

    const result = await fetchUnderlyingOverview(sessionToken, deviceId, payload);
    writeJson(res, 200, result);
    return true;
  }

  if (pathname === "/api/intervals/capture" && method === "POST") {
    const sessionToken = clean((req.headers.authorization || "").replace(/^Bearer\s+/i, ""));
    const deviceId = clean(req.headers["x-device-id"] || "");

    if (!sessionToken) {
      writeJson(res, 401, { error: "session token is required" });
      return true;
    }

    const result = await captureIntervalUniverse(sessionToken, deviceId);
    writeJson(res, 200, result);
    return true;
  }

  if (pathname === "/api/intervals/build" && method === "POST") {
    const payload = await readJsonBody(req);
    const sessionToken = clean((req.headers.authorization || "").replace(/^Bearer\s+/i, ""));
    const deviceId = clean(req.headers["x-device-id"] || "");

    if (!sessionToken) {
      writeJson(res, 401, { error: "session token is required" });
      return true;
    }

    const result = await buildIntervalUniverse(sessionToken, deviceId, payload);
    writeJson(res, 200, result);
    return true;
  }

  if (pathname === "/api/tracker/quote" && method === "POST") {
    const payload = await readJsonBody(req);
    const sessionToken = clean((req.headers.authorization || "").replace(/^Bearer\s+/i, ""));
    const deviceId = clean(req.headers["x-device-id"] || "");

    if (!sessionToken) {
      writeJson(res, 401, { error: "session token is required" });
      return true;
    }

    const result = await fetchTrackerQuote(sessionToken, deviceId, payload);
    writeJson(res, 200, result);
    return true;
  }

  return false;
}

function serveStatic(res, urlObj) {
  let pathname = urlObj.pathname || "/";
  if (pathname === "/") pathname = "/taskpane.html";
  const safePath = path.normalize(path.join(__dirname, pathname));
  if (!safePath.startsWith(__dirname)) {
    writeJson(res, 403, { error: "Forbidden" });
    return;
  }
  if (!fs.existsSync(safePath) || fs.statSync(safePath).isDirectory()) {
    writeJson(res, 404, { error: "Not found" });
    return;
  }

  const ext = path.extname(safePath).toLowerCase();
  res.writeHead(200, {
    ...corsHeaders(res),
    "Content-Type": MIME_TYPES[ext] || "application/octet-stream",
  });
  fs.createReadStream(safePath).pipe(res);
}

async function requestHandler(req, res) {
  try {
    res._corsOrigin = resolveCorsOrigin(req);
    const urlObj = new URL(req.url || "/", "https://localhost");
    if (req.method === "OPTIONS") {
      res.writeHead(204, corsHeaders(res));
      res.end();
      return;
    }
    if (await routeProxy(req, res, urlObj)) {
      return;
    }
    if (await routeLocalApi(req, res, urlObj)) {
      return;
    }
    serveStatic(res, urlObj);
  } catch (error) {
    writeJson(res, 500, { error: error.message || String(error) });
  }
}

function main() {
  if (!fs.existsSync(CERT_PATH) || !fs.existsSync(KEY_PATH)) {
    console.error("Missing dev cert files.");
    process.exit(1);
  }

  const server = https.createServer(
    {
      cert: fs.readFileSync(CERT_PATH),
      key: fs.readFileSync(KEY_PATH),
    },
    requestHandler
  );

  server.listen(PORT, LOOPBACK_HOST, () => {
    console.log(`Insti Excel dev server running: https://localhost:${PORT}`);
    console.log(`Proxy target: ${PROXY_ORIGIN}`);
  });
}

main();
