(function () {
  "use strict";

  const STORAGE = {
    device: "nubra.insti.device_id",
    auth: "nubra.insti.auth_token",
    session: "nubra.insti.session_token",
    tracked: "nubra.insti.tracked_instruments",
    intervalState: "nubra.insti.interval_state",
  };

  let U = null;
  let trackerTimer = null;
  let activeSuggestionIndex = -1;
  let searchTimer = null;
  let searchResults = [];
  let pendingHistoricalRefreshTimer = null;
  let pendingHistoricalRefreshIndex = -1;
  const ROLLOVER_HISTORICAL_REFRESH_DELAY_MS = 10_000;

  function emptyIntervalState() {
    return {
      intervals: [],
      rows: [],
      captures: {},
      latestIntervalIndex: -1,
    };
  }

  function loadIntervalState() {
    return getJsonStorage(STORAGE.intervalState, emptyIntervalState());
  }

  function saveIntervalState(state) {
    setJsonStorage(STORAGE.intervalState, state || emptyIntervalState());
  }

  function clearPendingHistoricalRefresh() {
    if (pendingHistoricalRefreshTimer) {
      clearTimeout(pendingHistoricalRefreshTimer);
      pendingHistoricalRefreshTimer = null;
    }
    pendingHistoricalRefreshIndex = -1;
  }

  function scheduleHistoricalRefreshAfterRollover(nextIntervalIndex, options = {}) {
    const targetIndex = Number(nextIntervalIndex);
    if (!Number.isInteger(targetIndex) || targetIndex < 0) return;
    if (pendingHistoricalRefreshTimer && pendingHistoricalRefreshIndex === targetIndex) return;

    clearPendingHistoricalRefresh();
    pendingHistoricalRefreshIndex = targetIndex;
    pendingHistoricalRefreshTimer = window.setTimeout(() => {
      pendingHistoricalRefreshTimer = null;
      const scheduledIndex = pendingHistoricalRefreshIndex;
      pendingHistoricalRefreshIndex = -1;
      refreshIntervalSheet({ silent: options.silent }).catch((error) => {
        const message = error?.message || String(error);
        log(`Delayed interval refresh failed for interval ${scheduledIndex + 1}: ${message}`, true);
      });
    }, ROLLOVER_HISTORICAL_REFRESH_DELAY_MS);
  }

  function now() {
    return new Date().toLocaleTimeString("en-IN", {
      timeZone: "Asia/Kolkata",
      hour12: true,
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
    });
  }

  function clean(value) {
    return String(value || "").trim();
  }

  function getStorage(key, fallback = "") {
    try {
      const value = localStorage.getItem(key);
      return value == null ? fallback : value;
    } catch (_error) {
      return fallback;
    }
  }

  function setStorage(key, value) {
    try {
      localStorage.setItem(key, String(value));
      return true;
    } catch (_error) {
      return false;
    }
  }

  function delStorage(key) {
    try {
      localStorage.removeItem(key);
    } catch (_error) {
      // ignore
    }
  }

  function getJsonStorage(key, fallback) {
    const raw = getStorage(key, "");
    if (!raw) return fallback;
    try {
      return JSON.parse(raw);
    } catch (_error) {
      return fallback;
    }
  }

  function setJsonStorage(key, value) {
    try {
      return setStorage(key, JSON.stringify(value));
    } catch (_error) {
      return false;
    }
  }

  function deviceId() {
    let value = getStorage(STORAGE.device, "");
    if (!value) {
      value = typeof crypto !== "undefined" && crypto.randomUUID
        ? `EXCEL-INSTI-${crypto.randomUUID()}`
        : `EXCEL-INSTI-${Date.now()}`;
      setStorage(STORAGE.device, value);
    }
    return value;
  }

  function shortToken(value) {
    const text = clean(value);
    if (!text) return "-";
    if (text.length <= 20) return text;
    return `${text.slice(0, 8)}...${text.slice(-8)}`;
  }

  function log(message, isError) {
    const line = `[${now()}] ${isError ? "ERROR: " : ""}${message}`;
    U.statusLog.textContent = `${line}\n${U.statusLog.textContent}`.slice(0, 30000);
  }

  function setMessage(message, tone = "bad") {
    const text = clean(message);
    U.actionMessage.textContent = text;
    U.actionMessage.classList.toggle("hidden", !text);
    U.actionMessage.classList.toggle("good", tone === "good");
    U.actionMessage.classList.toggle("bad", tone !== "good");
  }

  function refreshUi() {
    const authToken = getStorage(STORAGE.auth, "");
    const sessionToken = getStorage(STORAGE.session, "");
    const loggedIn = Boolean(sessionToken);
    U.deviceIdText.textContent = deviceId();
    U.authTokenText.textContent = shortToken(authToken);
    U.sessionTokenText.textContent = shortToken(sessionToken);
    U.authBadge.textContent = loggedIn ? "Logged in" : authToken ? "Auth token ready" : "Not logged in";
    U.authBadge.classList.toggle("good", loggedIn);
    U.authBadge.classList.toggle("bad", !loggedIn);
    U.mpinStage.classList.toggle("hidden", !authToken || loggedIn);
    U.loginInstiButton.disabled = loggedIn;
    U.verifyPinButton.disabled = loggedIn || !authToken;
    U.syncInstrumentsButton.disabled = !loggedIn;
    U.addTrackerInstrumentButton.disabled = !loggedIn;
    if (U.createIntervalSheetButton) U.createIntervalSheetButton.disabled = !loggedIn;
    if (U.captureIntervalButton) U.captureIntervalButton.disabled = !loggedIn;
  }

  function hideInstrumentDropdown() {
    if (!U?.instrumentDropdown) return;
    U.instrumentDropdown.innerHTML = "";
    U.instrumentDropdown.classList.add("hidden");
    activeSuggestionIndex = -1;
  }

  function currentSuggestions() {
    if (!U?.instrumentDropdown) return [];
    return Array.from(U.instrumentDropdown.querySelectorAll(".combo-option"));
  }

  function setActiveSuggestion(index) {
    const options = currentSuggestions();
    if (!options.length) {
      activeSuggestionIndex = -1;
      return;
    }

    const bounded = Math.max(0, Math.min(index, options.length - 1));
    activeSuggestionIndex = bounded;
    options.forEach((option, optionIndex) => {
      option.classList.toggle("active", optionIndex === bounded);
    });
    options[bounded].scrollIntoView({ block: "nearest" });
  }

  function applySuggestion(value) {
    U.trackerInstrumentInput.value = clean(value);
    hideInstrumentDropdown();
  }

  function asDisplayNumber(value, decimals = 2) {
    const n = Number(value);
    return Number.isFinite(n) ? Number(n.toFixed(decimals)) : "";
  }

  function asDisplayPrice(value) {
    const n = Number(value);
    return Number.isFinite(n) ? Number((n / 100).toFixed(2)) : "";
  }

  function toFiniteNumber(value) {
    const n = Number(value);
    return Number.isFinite(n) ? n : null;
  }

  function asDisplayTime(value) {
    const raw = clean(value);
    if (!raw) return "";
    const parsed = Date.parse(raw);
    if (Number.isNaN(parsed)) return raw;
    return new Date(parsed).toLocaleTimeString("en-IN", {
      timeZone: "Asia/Kolkata",
      hour12: true,
      hour: "2-digit",
      minute: "2-digit",
    });
  }

  function uniformFormatMatrix(rowCount, format) {
    return Array.from({ length: Math.max(0, rowCount) }, () => [format]);
  }

  function parseTimeToMinutes(value) {
    const text = clean(value);
    const parts = text.split(":").map((part) => Number(part));
    if (parts.length < 2 || parts.some((part) => !Number.isFinite(part))) return null;
    return parts[0] * 60 + parts[1];
  }

  function minutesToTimeLabel(minutes) {
    const hours = Math.floor(minutes / 60);
    const mins = minutes % 60;
    const dt = new Date();
    dt.setHours(hours, mins, 0, 0);
    return dt.toLocaleTimeString("en-IN", {
      timeZone: "Asia/Kolkata",
      hour12: true,
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
    });
  }

  function buildIntervals(startText, endText, stepMinutes) {
    const start = parseTimeToMinutes(startText);
    const end = parseTimeToMinutes(endText);
    const step = Number(stepMinutes);
    if (!Number.isFinite(start) || !Number.isFinite(end) || !Number.isFinite(step) || step <= 0 || end <= start) {
      return [];
    }

    const intervals = [];
    for (let cursor = start, index = 0; cursor < end; cursor += step, index += 1) {
      const next = Math.min(cursor + step, end);
      intervals.push({
        index,
        start: cursor,
        end: next,
        label: `${minutesToTimeLabel(cursor)} to ${minutesToTimeLabel(next)}`,
      });
    }
    return intervals;
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

  function resolveCurrentIntervalIndex(intervals) {
    const items = Array.isArray(intervals) ? intervals : [];
    if (!items.length) return -1;
    const nowMinutes = currentIstMinutes();
    return items.findIndex((interval) => nowMinutes >= Number(interval.start) && nowMinutes < Number(interval.end));
  }

  function intervalMetricFields() {
    return [
      { key: "last_price", header: "Last_Price", format: "price" },
      { key: "rt_px_chg_pct_1d", header: "% vs Prev Close", format: "percent" },
      { key: "oi_chng", header: "OI Chng", format: "whole" },
      { key: "oi_pct", header: "OI %", format: "percent" },
      { key: "basis", header: "Basis", format: "price" },
      { key: "basis_change", header: "Basis Change", format: "price" },
      { key: "prev_basis", header: "Prev. Basis", format: "price" },
    ];
  }

  function metricCount() {
    return intervalMetricFields().length;
  }

  function populateIntervalSlotSelect(intervals, selectedIndex = 0) {
    if (!U?.intervalSlotSelect) return;
    U.intervalSlotSelect.innerHTML = "";
    const items = Array.isArray(intervals) ? intervals : [];
    for (const interval of items) {
      const option = document.createElement("option");
      option.value = String(interval.index);
      option.textContent = interval.label;
      U.intervalSlotSelect.appendChild(option);
    }
    if (items.length) {
      U.intervalSlotSelect.value = String(Math.max(0, Math.min(selectedIndex, items.length - 1)));
    }
  }

  function renderInstrumentDropdown(items, query) {
    if (!U?.instrumentDropdown) return;

    const probe = clean(query).toUpperCase();
    const matches = Array.isArray(items) ? items : [];

    U.instrumentDropdown.innerHTML = "";
    if (!matches.length) {
      if (probe) {
        const empty = document.createElement("div");
        empty.className = "combo-empty";
        empty.textContent = "No instruments match your search.";
        U.instrumentDropdown.appendChild(empty);
        U.instrumentDropdown.classList.remove("hidden");
      } else {
        U.instrumentDropdown.classList.add("hidden");
      }
      activeSuggestionIndex = -1;
      return;
    }

    for (const item of matches) {
      const option = document.createElement("div");
      option.className = "combo-option";
      option.setAttribute("role", "option");
      option.dataset.value = item.symbol;
      option.dataset.index = String(searchResults.indexOf(item));
      const currentFuture = Array.isArray(item.futures) && item.futures.length ? clean(item.futures[0]?.symbol) : "";
      option.textContent = currentFuture ? `${item.label} | ${currentFuture}` : item.label;
      option.addEventListener("mousedown", (event) => {
        event.preventDefault();
        applySuggestion(item.symbol);
      });
      U.instrumentDropdown.appendChild(option);
    }

    U.instrumentDropdown.classList.remove("hidden");
    setActiveSuggestion(0);
  }

  function clearSearchTimer() {
    if (searchTimer) {
      clearTimeout(searchTimer);
      searchTimer = null;
    }
  }

  async function searchInstruments(query) {
    const sessionToken = getStorage(STORAGE.session, "");
    const data = await localApi(`/api/instruments/search?q=${encodeURIComponent(clean(query))}&limit=25`, {
      method: "GET",
      headers: {
        Authorization: sessionToken ? `Bearer ${sessionToken}` : "",
      },
    });
    searchResults = Array.isArray(data?.items) ? data.items : [];
    renderInstrumentDropdown(searchResults, query);
  }

  function scheduleInstrumentSearch(query) {
    clearSearchTimer();
    const probe = clean(query);
    if (!probe) {
      searchResults = [];
      hideInstrumentDropdown();
      return;
    }

    searchTimer = window.setTimeout(() => {
      searchInstruments(probe).catch((error) => {
        searchResults = [];
        hideInstrumentDropdown();
        log(error.message || String(error), true);
      });
    }, 120);
  }

  async function requestJson(url, options = {}) {
    const headers = {
      "Content-Type": "application/json",
      "x-device-id": deviceId(),
      ...(options.headers || {}),
    };

    const response = await fetch(url, {
      method: options.method || "GET",
      headers,
      body: options.body ? JSON.stringify(options.body) : undefined,
    });

    const text = await response.text();
    let data = {};
    try {
      data = text ? JSON.parse(text) : {};
    } catch (_error) {
      data = { raw: text };
    }

    if (!response.ok) {
      const message = clean(data.error || data.message || data.raw || `HTTP ${response.status}`);
      throw new Error(message);
    }

    return data;
  }

  function api(path, options = {}) {
    return requestJson(`/proxy/insti${path}`, options);
  }

  function localApi(path, options = {}) {
    return requestJson(path, options);
  }

  async function handleInstiLogin() {
    const exchangeClientCode = clean(U.exchangeClientCodeInput.value);
    const clientCode = clean(U.clientCodeInput.value);
    const username = clean(U.usernameInput.value);
    const password = clean(U.passwordInput.value);

    if (!exchangeClientCode || !clientCode || !username || !password) {
      throw new Error("All login fields are required.");
    }

    setMessage("Submitting institutional login...");
    const data = await api("/login-insti", {
      method: "POST",
      body: {
        exchange_client_code: exchangeClientCode,
        client_code: clientCode,
        username,
        password,
      },
    });

    const authToken = clean(data.auth_token || data.authToken);
    if (!authToken) {
      throw new Error("`auth_token` missing in login response.");
    }

    setStorage(STORAGE.auth, authToken);
    delStorage(STORAGE.session);
    U.pinInput.focus();
    setMessage("Institutional login successful. Enter MPIN.", "good");
    log("`login-insti` succeeded. Awaiting MPIN verification.");
    refreshUi();
  }

  async function handleVerifyPin() {
    const pin = clean(U.pinInput.value);
    const authToken = getStorage(STORAGE.auth, "");

    if (!authToken) {
      throw new Error("No `auth_token` available. Run Insti login first.");
    }
    if (!/^\d{4}$/.test(pin)) {
      throw new Error("MPIN must be 4 digits.");
    }

    setMessage("Verifying MPIN...");
    const data = await api("/verifypin", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${authToken}`,
      },
      body: { pin },
    });

    const sessionToken = clean(data.session_token || data.sessionToken);
    if (!sessionToken) {
      throw new Error("`session_token` missing in MPIN response.");
    }

    setStorage(STORAGE.session, sessionToken);
    setMessage("Login successful. Session token stored.", "good");
    log("`verifypin` succeeded. Session token stored.");
    refreshUi();
    syncTrackerPoller();
  }

  async function writeLiveTrackerSheet(trackedItems) {
    if (typeof Excel === "undefined") {
      throw new Error("Excel host is not available.");
    }

    const items = Array.isArray(trackedItems) ? trackedItems : [];
    await Excel.run(async (context) => {
      let sheet = context.workbook.worksheets.getItemOrNullObject("LiveTracker");
      sheet.load("name,isNullObject");
      await context.sync();

      if (sheet.isNullObject) {
        sheet = context.workbook.worksheets.add("LiveTracker");
      }

      const usedRange = sheet.getUsedRangeOrNullObject(true);
      usedRange.load("isNullObject,rowCount,columnCount");
      await context.sync();

      const headers = [
        "stock",
        "current_future",
        "next_future",
        "spot_ltp",
        "fut_1_ltp",
        "fut_1_oi",
        "fut_2_ltp",
        "fut_2_oi",
        "updated_at",
      ];
      const values = [
        headers,
        ...items.map((item) => [
          item.underlying || item.symbol || "",
          item.current_future?.symbol || "",
          item.next_future?.symbol || "",
          asDisplayPrice(item.stock?.curr_ltp),
          asDisplayPrice(item.current_future?.curr_ltp),
          asDisplayNumber(item.current_future?.oi_current, 0),
          asDisplayPrice(item.next_future?.curr_ltp),
          asDisplayNumber(item.next_future?.oi_current, 0),
          asDisplayTime(item.updated_at),
        ]),
      ];
      const bodyValues = values.slice(1);
      const expectedRows = values.length;
      const expectedCols = headers.length;
      const sameShape = !usedRange.isNullObject
        && Number(usedRange.rowCount || 0) === expectedRows
        && Number(usedRange.columnCount || 0) === expectedCols;

      if (!sameShape) {
        if (!usedRange.isNullObject) {
          usedRange.clear();
        }

        const range = sheet.getRangeByIndexes(0, 0, values.length, headers.length);
        range.values = values;
        sheet.getRangeByIndexes(0, 0, 1, headers.length).format.font.bold = true;
        sheet.getRangeByIndexes(0, 0, 1, headers.length).format.fill.color = "#DDEAFB";
        for (const columnIndex of [3, 4, 6]) {
          if (columnIndex < headers.length && values.length > 1) {
            sheet.getRangeByIndexes(1, columnIndex, values.length - 1, 1).numberFormat = uniformFormatMatrix(values.length - 1, "0.00");
          }
        }
        for (const columnIndex of [5, 7]) {
          if (columnIndex < headers.length && values.length > 1) {
            sheet.getRangeByIndexes(1, columnIndex, values.length - 1, 1).numberFormat = uniformFormatMatrix(values.length - 1, "0");
          }
        }
        range.format.autofitColumns();
        sheet.freezePanes.freezeRows(1);
        await context.sync();
        return;
      }

      if (bodyValues.length) {
        sheet.getRangeByIndexes(1, 0, bodyValues.length, headers.length).values = bodyValues;
      } else if (!usedRange.isNullObject) {
        usedRange.clear();
        sheet.getRangeByIndexes(0, 0, 1, headers.length).values = [headers];
        sheet.getRangeByIndexes(0, 0, 1, headers.length).format.font.bold = true;
        sheet.getRangeByIndexes(0, 0, 1, headers.length).format.fill.color = "#DDEAFB";
      }
      await context.sync();
    });
  }

  async function removeObsoleteSheets() {
    if (typeof Excel === "undefined") {
      return;
    }

    await Excel.run(async (context) => {
      const workbook = context.workbook;
      const worksheets = workbook.worksheets;
      worksheets.load("items/name");
      await context.sync();

      const names = new Set(worksheets.items.map((sheet) => sheet.name));
      for (const name of ["Instruments", "Tracker"]) {
        if (!names.has(name)) continue;
        const sheet = worksheets.getItem(name);
        sheet.delete();
      }
      await context.sync();
    });
  }

  function computeIntervalMetrics(snapshot, prevCapture, allowSessionOiFallback = false) {
    const stockLtp = toFiniteNumber(snapshot?.stock?.curr_ltp);
    const futureLtp = toFiniteNumber(snapshot?.current_future?.curr_ltp);
    const futurePrevClose = toFiniteNumber(snapshot?.current_future?.prev_close);
    const futureOi = toFiniteNumber(snapshot?.current_future?.oi_current);
    const previousIntervalOi = toFiniteNumber(prevCapture?.current_future?.oi_current);
    const futurePrevOi = toFiniteNumber(snapshot?.current_future?.oi_yest_close);
    const prevStockLtp = toFiniteNumber(prevCapture?.stock?.curr_ltp);
    const prevFutureLtp = toFiniteNumber(prevCapture?.current_future?.curr_ltp);
    const basisValue = Number.isFinite(futureLtp) && Number.isFinite(stockLtp) ? futureLtp - stockLtp : null;
    const prevBasisValue = Number.isFinite(prevFutureLtp) && Number.isFinite(prevStockLtp)
      ? prevFutureLtp - prevStockLtp
      : null;
    const oiBaseValue = Number.isFinite(previousIntervalOi)
      ? previousIntervalOi
      : (allowSessionOiFallback && Number.isFinite(futurePrevOi) ? futurePrevOi : null);
    const useLiveWebsocketOi = clean(snapshot?.current_future?.quote_source).toLowerCase() === "websocket:index";
    const prevBasis = Number.isFinite(prevBasisValue) ? prevBasisValue : "";
    return {
      last_price: Number.isFinite(futureLtp) ? futureLtp : "",
      rt_px_chg_pct_1d: Number.isFinite(futureLtp) && Number.isFinite(futurePrevClose) && futurePrevClose !== 0
        ? ((futureLtp - futurePrevClose) / futurePrevClose) * 100
        : "",
      oi_chng: useLiveWebsocketOi
        ? (Number.isFinite(futureOi) ? futureOi : "")
        : (Number.isFinite(futureOi) && Number.isFinite(oiBaseValue) ? futureOi - oiBaseValue : ""),
      oi_pct: useLiveWebsocketOi
        ? ""
        : (Number.isFinite(futureOi) && Number.isFinite(oiBaseValue) && oiBaseValue !== 0
        ? ((futureOi - oiBaseValue) / oiBaseValue) * 100
        : ""),
      basis: Number.isFinite(basisValue) ? basisValue : "",
      basis_change: Number.isFinite(basisValue) && Number.isFinite(prevBasisValue) ? basisValue - prevBasisValue : "",
      prev_basis: prevBasis,
    };
  }

  function rowMapByUnderlying(rows) {
    const map = new Map();
    for (const row of Array.isArray(rows) ? rows : []) {
      map.set(clean(row?.underlying || row?.symbol).toUpperCase(), row);
    }
    return map;
  }

  function buildLiveIntervalSnapshot(item) {
    const underlying = clean(item?.underlying || item?.symbol);
    if (!underlying) return null;
    return {
      underlying,
      display: item?.display || underlying,
      stock: item?.stock ? {
        symbol: item.stock.symbol || underlying,
        ref_id: item.stock.ref_id ?? null,
        prev_close: item.stock.prev_close ?? null,
        curr_ltp: item.stock.curr_ltp ?? null,
        ltp_as_of: item.stock.ltp_as_of || item.updated_at || "",
        quote_source: item.stock.quote_source || "",
      } : null,
      current_future: item?.current_future ? {
        symbol: item.current_future.symbol || "",
        expiry: item.current_future.expiry || "",
        ref_id: item.current_future.ref_id ?? null,
        prev_close: item.current_future.prev_close ?? null,
        curr_ltp: item.current_future.curr_ltp ?? null,
        ltp_as_of: item.current_future.ltp_as_of || item.updated_at || "",
        oi_yest_close: item.current_future.oi_yest_close ?? null,
        oi_current: item.current_future.oi_current ?? null,
        oi_as_of: item.current_future.oi_as_of || item.updated_at || "",
        quote_source: item.current_future.quote_source || "",
      } : null,
      next_future: item?.next_future ? {
        symbol: item.next_future.symbol || "",
        expiry: item.next_future.expiry || "",
        ref_id: item.next_future.ref_id ?? null,
        prev_close: item.next_future.prev_close ?? null,
        curr_ltp: item.next_future.curr_ltp ?? null,
        ltp_as_of: item.next_future.ltp_as_of || item.updated_at || "",
        oi_yest_close: item.next_future.oi_yest_close ?? null,
        oi_current: item.next_future.oi_current ?? null,
        oi_as_of: item.next_future.oi_as_of || item.updated_at || "",
        quote_source: item.next_future.quote_source || "",
      } : null,
      updated_at: item?.updated_at || new Date().toISOString(),
    };
  }

  async function writeIntervalTrackerSheet(state, options = {}) {
    if (typeof Excel === "undefined") {
      throw new Error("Excel host is not available.");
    }

    const intervalState = state || loadIntervalState();
    const intervals = Array.isArray(intervalState?.intervals) ? intervalState.intervals : [];
    const rows = Array.isArray(intervalState?.rows) ? intervalState.rows : [];
    const captures = intervalState?.captures || {};
    const metrics = intervalMetricFields();
    const staticHeaders = [
      "No",
      "SYMBOL",
      "Futs",
      "FUTURES 1",
      "FUTURES 2",
      "Spot Prev Close",
      "Spot Curr LTP",
      "Fut Prev Close",
      "Fut Curr LTP",
    ];

    const headerRow1 = staticHeaders.map(() => "");
    const headerRow2 = staticHeaders.slice();
    for (const interval of intervals) {
      headerRow1.push(interval.label);
      for (let i = 1; i < metrics.length; i += 1) headerRow1.push("");
      for (const metric of metrics) headerRow2.push(metric.header);
    }

    const valueRows = rows.map((row, rowIndex) => {
      const base = [
        rowIndex + 1,
        row.underlying || "",
        row.front_future_symbol || "",
        row.future_1 || "",
        row.future_2 || "",
        "",
        "",
        "",
        "",
      ];
      for (const interval of intervals) {
        const snapshot = captures?.[String(interval.index)]?.[row.underlying] || null;
        const prevSnapshot = interval.index > 0 ? captures?.[String(interval.index - 1)]?.[row.underlying] || null : null;
        const intervalMetrics = snapshot ? computeIntervalMetrics(snapshot, prevSnapshot, interval.index === 0) : {};
        const stockPrevClose = snapshot?.stock?.prev_close;
        const stockCurrLtp = snapshot?.stock?.curr_ltp;
        const futurePrevClose = snapshot?.current_future?.prev_close;
        const futureCurrLtp = snapshot?.current_future?.curr_ltp;
        if (base[5] === "" && stockPrevClose !== undefined && stockPrevClose !== null) base[5] = asDisplayPrice(stockPrevClose);
        if (base[6] === "" && stockCurrLtp !== undefined && stockCurrLtp !== null) base[6] = asDisplayPrice(stockCurrLtp);
        if (base[7] === "" && futurePrevClose !== undefined && futurePrevClose !== null) base[7] = asDisplayPrice(futurePrevClose);
        if (base[8] === "" && futureCurrLtp !== undefined && futureCurrLtp !== null) base[8] = asDisplayPrice(futureCurrLtp);
        for (const metric of metrics) {
          if (metric.format === "price") {
            base.push(asDisplayPrice(intervalMetrics[metric.key]));
          } else if (metric.format === "whole") {
            base.push(asDisplayNumber(intervalMetrics[metric.key], 0));
          } else {
            base.push(asDisplayNumber(intervalMetrics[metric.key]));
          }
        }
      }
      return base;
    });

    const values = [headerRow1, headerRow2, ...valueRows];

    await Excel.run(async (context) => {
      let sheet = context.workbook.worksheets.getItemOrNullObject("IntervalTracker");
      sheet.load("name,isNullObject");
      await context.sync();
      if (sheet.isNullObject) {
        sheet = context.workbook.worksheets.add("IntervalTracker");
      }

      const usedRange = sheet.getUsedRangeOrNullObject(true);
      usedRange.load("isNullObject,rowCount,columnCount");
      await context.sync();
      const expectedRows = values.length;
      const expectedCols = values[0].length;
      const forceFullRewrite = Boolean(options?.forceFullRewrite);
      const sameShape = !usedRange.isNullObject
        && Number(usedRange.rowCount || 0) === expectedRows
        && Number(usedRange.columnCount || 0) === expectedCols;

      if (forceFullRewrite || !sameShape) {
        if (!usedRange.isNullObject) {
          usedRange.clear();
        }

        sheet.getRangeByIndexes(0, 0, values.length, values[0].length).values = values;
        sheet.getRangeByIndexes(0, 0, 2, values[0].length).format.font.bold = true;
        sheet.getRangeByIndexes(0, 0, 2, values[0].length).format.fill.color = "#DDEAFB";
        const fullRange = sheet.getRangeByIndexes(0, 0, values.length, values[0].length);
        for (const borderName of ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal", "InsideVertical"]) {
          const border = fullRange.format.borders.getItem(borderName);
          border.style = "Continuous";
          border.color = "#B8C2CC";
        }

        if (intervals.length) {
          for (const interval of intervals) {
            const startCol = staticHeaders.length + interval.index * metrics.length;
            const labelRange = sheet.getRangeByIndexes(0, startCol, 1, metrics.length);
            labelRange.merge();
            sheet.getCell(0, startCol).values = [[interval.label]];
            labelRange.format.horizontalAlignment = "Center";
            const sectionRange = sheet.getRangeByIndexes(0, startCol, values.length, metrics.length);
            for (const borderName of ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"]) {
              const border = sectionRange.format.borders.getItem(borderName);
              border.style = "Continuous";
              border.color = "#5B6B7A";
            }
          }
        }

        if (values.length > 2) {
          for (const columnIndex of [5, 6, 7, 8]) {
            if (columnIndex < values[0].length) {
              sheet.getRangeByIndexes(2, columnIndex, values.length - 2, 1).numberFormat = uniformFormatMatrix(values.length - 2, "0.00");
            }
          }
          for (const interval of intervals) {
            const startCol = staticHeaders.length + interval.index * metrics.length;
            for (let i = 0; i < metrics.length; i += 1) {
              const metric = metrics[i];
              const columnIndex = startCol + i;
              if (columnIndex >= values[0].length) continue;
              const range = sheet.getRangeByIndexes(2, columnIndex, values.length - 2, 1);
              if (metric.format === "price") {
                range.numberFormat = uniformFormatMatrix(values.length - 2, "0.00");
              } else if (metric.format === "whole") {
                range.numberFormat = uniformFormatMatrix(values.length - 2, "0");
              } else {
                range.numberFormat = uniformFormatMatrix(values.length - 2, "0.00");
              }
            }
          }
        }

        sheet.getRangeByIndexes(0, 0, values.length, values[0].length).format.autofitColumns();
        sheet.freezePanes.freezeRows(2);
        await context.sync();
        return;
      }

      if (valueRows.length) {
        sheet.getRangeByIndexes(2, 0, valueRows.length, values[0].length).values = valueRows;
      }
      await context.sync();
    });
  }

  function trackedInstruments() {
    const rows = getJsonStorage(STORAGE.tracked, []);
    return Array.isArray(rows)
      ? rows.filter((row) => clean(row?.underlying || row?.symbol))
      : [];
  }

  function saveTrackedInstruments(rows) {
    setJsonStorage(STORAGE.tracked, Array.isArray(rows) ? rows : []);
  }

  async function resolveInstrumentSelection(inputValue) {
    const probe = clean(inputValue);
    if (!probe) return null;
    const existing = searchResults.find((item) =>
      clean(item.symbol).toUpperCase() === probe.toUpperCase()
      || clean(item.label).toUpperCase() === probe.toUpperCase()
    );
    if (existing) return existing;

    const sessionToken = getStorage(STORAGE.session, "");
    const data = await localApi(`/api/instruments/resolve?q=${encodeURIComponent(probe)}`, {
      method: "GET",
      headers: {
        Authorization: sessionToken ? `Bearer ${sessionToken}` : "",
      },
    });
    return data || null;
  }

  function rowsFromIntervalItems(items) {
    return (Array.isArray(items) ? items : []).map((item) => ({
      underlying: item.underlying || "",
      display: item.display || "",
      stock_symbol: item.stock?.symbol || item.underlying || "",
      front_future_symbol: item.futures?.[0]?.symbol || "",
      future_1: item.futures?.[0]?.symbol || "",
      future_2: item.futures?.[1]?.symbol || "",
    }));
  }

  async function handleCreateIntervalSheet() {
    const sessionToken = getStorage(STORAGE.session, "");
    if (!sessionToken) {
      throw new Error("Login is required before creating the interval sheet.");
    }

    const intervals = buildIntervals(
      U.intervalStartInput?.value,
      U.intervalEndInput?.value,
      U.intervalMinutesInput?.value
    );
    if (!intervals.length) {
      throw new Error("Enter a valid start time, end time, and interval minutes.");
    }

    const tradeDate = new Date().toLocaleDateString("en-CA", { timeZone: "Asia/Kolkata" });
    const tracked = trackedInstruments();
    const underlyings = tracked
      .map((item) => clean(item?.underlying || item?.symbol).toUpperCase())
      .filter(Boolean);
    if (!underlyings.length) {
      throw new Error("Add one or more stocks to LiveTracker first.");
    }
    const data = await localApi("/api/intervals/build", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${sessionToken}`,
      },
      body: {
        date: tradeDate,
        interval_minutes: Number(U.intervalMinutesInput?.value || 15),
        intervals,
        underlyings,
      },
    });
    const items = Array.isArray(data?.items) ? data.items : [];
    if (!items.length) {
      throw new Error("No selected LiveTracker stocks are available in cache.");
    }

    const rows = rowsFromIntervalItems(items);

    const state = {
      intervals,
      rows,
      captures: data?.snapshots || {},
      latestIntervalIndex: Number.isInteger(Number(data?.latest_interval_index)) ? Number(data.latest_interval_index) : -1,
    };
    saveIntervalState(state);
    populateIntervalSlotSelect(intervals, Math.max(0, state.latestIntervalIndex));
    await writeIntervalTrackerSheet(state, { forceFullRewrite: true });
    setMessage(`IntervalTracker built with ${rows.length} stocks and ${intervals.length} intervals.`, "good");
    log(`Built IntervalTracker with ${rows.length} stocks and ${intervals.length} intervals using historical + live data.`);
  }

  async function handleCaptureInterval() {
    const sessionToken = getStorage(STORAGE.session, "");
    if (!sessionToken) {
      throw new Error("Login is required before capturing interval data.");
    }

    const state = loadIntervalState();
    const intervals = Array.isArray(state?.intervals) ? state.intervals : [];
    if (!intervals.length) {
      throw new Error("Create the interval sheet first.");
    }

    const tradeDate = new Date().toLocaleDateString("en-CA", { timeZone: "Asia/Kolkata" });
    const tracked = trackedInstruments();
    const underlyings = tracked
      .map((item) => clean(item?.underlying || item?.symbol).toUpperCase())
      .filter(Boolean);
    if (!underlyings.length) {
      throw new Error("Add one or more stocks to LiveTracker first.");
    }
    setMessage("Refreshing interval sheet...");
    const data = await localApi("/api/intervals/build", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${sessionToken}`,
      },
      body: {
        date: tradeDate,
        interval_minutes: Number(U.intervalMinutesInput?.value || 15),
        intervals,
        underlyings,
      },
    });
    const rows = rowsFromIntervalItems(data?.items);
    const nextState = {
      ...state,
      rows: rows.length ? rows : state?.rows || [],
      captures: data?.snapshots || state?.captures || {},
      latestIntervalIndex: Number.isInteger(Number(data?.latest_interval_index)) ? Number(data.latest_interval_index) : state?.latestIntervalIndex ?? -1,
    };
    saveIntervalState(nextState);
    populateIntervalSlotSelect(intervals, Math.max(0, nextState.latestIntervalIndex));
    await writeIntervalTrackerSheet(nextState, { forceFullRewrite: true });
    setMessage(`Interval sheet refreshed for ${Number(data?.count || 0)} stocks.`, "good");
    log(`Refreshed interval sheet for ${Number(data?.count || 0)} stocks.`);
  }

  async function refreshIntervalSheet(options = {}) {
    clearPendingHistoricalRefresh();

    const sessionToken = getStorage(STORAGE.session, "");
    if (!sessionToken) return;

    const state = loadIntervalState();
    const intervals = Array.isArray(state?.intervals) ? state.intervals : [];
    if (!intervals.length) return;

    const tracked = trackedInstruments();
    const underlyings = tracked
      .map((item) => clean(item?.underlying || item?.symbol).toUpperCase())
      .filter(Boolean);
    if (!underlyings.length) return;

    const tradeDate = new Date().toLocaleDateString("en-CA", { timeZone: "Asia/Kolkata" });
    const data = await localApi("/api/intervals/build", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${sessionToken}`,
      },
      body: {
        date: tradeDate,
        interval_minutes: Number(U.intervalMinutesInput?.value || 15),
        intervals,
        underlyings,
      },
    });
    const rows = rowsFromIntervalItems(data?.items);

    const nextState = {
      ...state,
      rows: rows.length ? rows : state?.rows || [],
      captures: data?.snapshots || state?.captures || {},
      latestIntervalIndex: Number.isInteger(Number(data?.latest_interval_index)) ? Number(data.latest_interval_index) : state?.latestIntervalIndex ?? -1,
    };
    saveIntervalState(nextState);
    populateIntervalSlotSelect(intervals, Math.max(0, nextState.latestIntervalIndex));
    await writeIntervalTrackerSheet(nextState, { forceFullRewrite: true });
    if (!options.silent) {
      setMessage(`Interval sheet refreshed for ${Number(data?.count || 0)} stocks.`, "good");
      log(`Refreshed interval sheet for ${Number(data?.count || 0)} stocks.`);
    }
  }

  async function syncIntervalSheetWithLiveData(trackedItems, options = {}) {
    const sessionToken = getStorage(STORAGE.session, "");
    if (!sessionToken) return;

    const state = loadIntervalState();
    const intervals = Array.isArray(state?.intervals) ? state.intervals : [];
    if (!intervals.length) return;

    const currentIntervalIndex = resolveCurrentIntervalIndex(intervals);
    const previousIntervalIndex = Number.isInteger(Number(state?.latestIntervalIndex))
      ? Number(state.latestIntervalIndex)
      : -1;
    const rolloverDetected = currentIntervalIndex !== previousIntervalIndex;

    if (rolloverDetected && currentIntervalIndex < 0) {
      const nextState = {
        ...state,
        latestIntervalIndex: -1,
      };
      saveIntervalState(nextState);
      return;
    }

    if (rolloverDetected) {
      scheduleHistoricalRefreshAfterRollover(currentIntervalIndex, { silent: options.silent });
      if (!options.silent) {
        setMessage("New interval started. Historical data will refresh in 10 seconds.", "good");
        log(`Interval rolled over to ${currentIntervalIndex + 1}. Scheduled delayed historical refresh.`);
      }
    }
    if (currentIntervalIndex < 0) return;

    const liveRows = Array.isArray(trackedItems) ? trackedItems : trackedInstruments();
    if (!liveRows.length) return;

    const trackedMap = rowMapByUnderlying(liveRows);
    const captureKey = String(currentIntervalIndex);
    const nextCaptures = {
      ...(state?.captures || {}),
      [captureKey]: {
        ...((state?.captures || {})[captureKey] || {}),
      },
    };

    for (const row of Array.isArray(state?.rows) ? state.rows : []) {
      const tracked = trackedMap.get(clean(row?.underlying).toUpperCase());
      if (!tracked) continue;
      const snapshot = buildLiveIntervalSnapshot(tracked);
      if (snapshot) {
        nextCaptures[captureKey][row.underlying] = snapshot;
      }
    }

    const nextState = {
      ...state,
      captures: nextCaptures,
      latestIntervalIndex: currentIntervalIndex,
    };
    saveIntervalState(nextState);
    populateIntervalSlotSelect(intervals, Math.max(0, currentIntervalIndex));
    await writeIntervalTrackerSheet(nextState);

    if (!options.silent) {
      setMessage("Live interval updated from websocket-backed quotes.", "good");
      log(`Updated open interval ${currentIntervalIndex + 1} with live data.`);
    }
  }

  async function refreshTrackedQuotes(options = {}) {
    const sessionToken = getStorage(STORAGE.session, "");
    if (!sessionToken) return [];

    const current = trackedInstruments();
    if (!current.length) return [];

    const next = await Promise.all(current.map(async (item) => {
      try {
        const quote = await localApi("/api/tracker/overview", {
          method: "POST",
          headers: {
            Authorization: `Bearer ${sessionToken}`,
          },
          body: item,
        });

        return {
          ...item,
          underlying: quote.underlying || item.underlying || item.symbol || "",
          display: quote.display || item.display || "",
          stock: quote.stock || item.stock || null,
          current_future: quote.current_future || item.current_future || null,
          next_future: quote.next_future || item.next_future || null,
          updated_at: quote.updated_at || item.updated_at || "",
        };
      } catch (_error) {
        return item;
      }
    }));

    saveTrackedInstruments(next);
    await writeLiveTrackerSheet(next);
    if (!options.silent) {
      setMessage(`LiveTracker updated for ${next.length} instrument${next.length === 1 ? "" : "s"}.`, "good");
    }
    return next;
  }

  function stopTrackerPoller() {
    if (trackerTimer) {
      clearInterval(trackerTimer);
      trackerTimer = null;
    }
  }

  function syncTrackerPoller() {
    stopTrackerPoller();
    if (!getStorage(STORAGE.session, "") || !trackedInstruments().length) return;
    trackerTimer = setInterval(() => {
      refreshTrackedQuotes({ silent: true })
        .then((next) => syncIntervalSheetWithLiveData(next, { silent: true }))
        .catch(() => null);
    }, 3000);
  }

  async function handleAddTrackerInstrument() {
    const selection = await resolveInstrumentSelection(U.trackerInstrumentInput.value);
    if (!selection) {
      throw new Error("Select a valid instrument from the dropdown first.");
    }

    const underlying = clean(selection?.underlying || selection?.symbol);
    if (!underlying) {
      throw new Error("Selected stock has no usable underlying.");
    }

    const existing = trackedInstruments();
    const already = existing.find((item) => clean(item.underlying || item.symbol).toUpperCase() === underlying.toUpperCase());
    const next = already
      ? existing
      : existing.concat([{
          underlying,
          display: selection.display || selection.label || underlying,
          stock: selection.stock || null,
          current_future: Array.isArray(selection.futures) && selection.futures[0] ? selection.futures[0] : null,
          next_future: Array.isArray(selection.futures) && selection.futures[1] ? selection.futures[1] : null,
          updated_at: "",
        }]);

    saveTrackedInstruments(next);
    const refreshed = await refreshTrackedQuotes({ silent: true });
    await refreshIntervalSheet({ silent: true }).catch(() => null);
    await syncIntervalSheetWithLiveData(refreshed, { silent: true }).catch(() => null);
    syncTrackerPoller();
    U.trackerInstrumentInput.value = underlying;
    setMessage(`Added ${underlying} to LiveTracker.`, "good");
    log(`Added ${underlying} to LiveTracker.`);
  }

  async function handleSyncInstruments() {
    const sessionToken = getStorage(STORAGE.session, "");
    if (!sessionToken) {
      throw new Error("Login is required before loading F&O stocks.");
    }

    setMessage("Fetching stocks that have futures into backend cache...");
    const data = await localApi("/api/instruments/fno-universe", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${sessionToken}`,
      },
      body: {},
    });

    await removeObsoleteSheets();
    clearPendingHistoricalRefresh();
    saveTrackedInstruments([]);
    await writeLiveTrackerSheet([]);
    saveIntervalState(emptyIntervalState());
    populateIntervalSlotSelect([], 0);
    syncTrackerPoller();
    searchResults = [];
    hideInstrumentDropdown();
    setMessage(`Backend F&O stocks cache loaded with ${Number(data?.stocks_with_futures || 0)} stocks.`, "good");
    log(`Loaded backend F&O stocks cache with ${Number(data?.stocks_with_futures || 0)} stocks from ${Number(data?.count || 0)} refdata rows.`);
  }

  function clearSession() {
    clearPendingHistoricalRefresh();
    delStorage(STORAGE.auth);
    delStorage(STORAGE.session);
    U.pinInput.value = "";
    setMessage("Stored auth state cleared.", "good");
    log("Cleared local auth/session tokens.");
    refreshUi();
    syncTrackerPoller();
  }

  function bind() {
    U = {
      exchangeClientCodeInput: document.getElementById("exchangeClientCodeInput"),
      clientCodeInput: document.getElementById("clientCodeInput"),
      usernameInput: document.getElementById("usernameInput"),
      passwordInput: document.getElementById("passwordInput"),
      pinInput: document.getElementById("pinInput"),
      loginInstiButton: document.getElementById("loginInstiButton"),
      verifyPinButton: document.getElementById("verifyPinButton"),
      syncInstrumentsButton: document.getElementById("syncInstrumentsButton"),
      trackerInstrumentInput: document.getElementById("trackerInstrumentInput"),
      instrumentDropdown: document.getElementById("instrumentDropdown"),
      addTrackerInstrumentButton: document.getElementById("addTrackerInstrumentButton"),
      intervalStartInput: document.getElementById("intervalStartInput"),
      intervalEndInput: document.getElementById("intervalEndInput"),
      intervalMinutesInput: document.getElementById("intervalMinutesInput"),
      createIntervalSheetButton: document.getElementById("createIntervalSheetButton"),
      intervalSlotSelect: document.getElementById("intervalSlotSelect"),
      captureIntervalButton: document.getElementById("captureIntervalButton"),
      clearSessionButton: document.getElementById("clearSessionButton"),
      mpinStage: document.getElementById("mpinStage"),
      actionMessage: document.getElementById("actionMessage"),
      authBadge: document.getElementById("authBadge"),
      deviceIdText: document.getElementById("deviceIdText"),
      authTokenText: document.getElementById("authTokenText"),
      sessionTokenText: document.getElementById("sessionTokenText"),
      statusLog: document.getElementById("statusLog"),
    };

    U.loginInstiButton.addEventListener("click", async () => {
      try {
        await handleInstiLogin();
      } catch (error) {
        setMessage(error.message || String(error));
        log(error.message || String(error), true);
        refreshUi();
      }
    });

    U.verifyPinButton.addEventListener("click", async () => {
      try {
        await handleVerifyPin();
      } catch (error) {
        setMessage(error.message || String(error));
        log(error.message || String(error), true);
        refreshUi();
      }
    });

    U.clearSessionButton.addEventListener("click", clearSession);

    U.syncInstrumentsButton.addEventListener("click", async () => {
      try {
        await handleSyncInstruments();
      } catch (error) {
        setMessage(error.message || String(error));
        log(error.message || String(error), true);
        refreshUi();
      }
    });

    U.addTrackerInstrumentButton.addEventListener("click", async () => {
      try {
        await handleAddTrackerInstrument();
      } catch (error) {
        setMessage(error.message || String(error));
        log(error.message || String(error), true);
      }
    });

    U.createIntervalSheetButton.addEventListener("click", async () => {
      try {
        await handleCreateIntervalSheet();
      } catch (error) {
        setMessage(error.message || String(error));
        log(error.message || String(error), true);
      }
    });

    U.captureIntervalButton.addEventListener("click", async () => {
      try {
        await handleCaptureInterval();
      } catch (error) {
        setMessage(error.message || String(error));
        log(error.message || String(error), true);
      }
    });

    U.trackerInstrumentInput.addEventListener("input", () => {
      scheduleInstrumentSearch(U.trackerInstrumentInput.value);
    });

    U.trackerInstrumentInput.addEventListener("focus", () => {
      scheduleInstrumentSearch(U.trackerInstrumentInput.value);
    });

    U.trackerInstrumentInput.addEventListener("keydown", (event) => {
      const options = currentSuggestions();
      if (!options.length) return;

      if (event.key === "ArrowDown") {
        event.preventDefault();
        setActiveSuggestion(activeSuggestionIndex + 1);
        return;
      }

      if (event.key === "ArrowUp") {
        event.preventDefault();
        setActiveSuggestion(activeSuggestionIndex - 1);
        return;
      }

      if (event.key === "Enter" && activeSuggestionIndex >= 0) {
        event.preventDefault();
        applySuggestion(options[activeSuggestionIndex].dataset.value || options[activeSuggestionIndex].textContent);
        return;
      }

      if (event.key === "Escape") {
        hideInstrumentDropdown();
      }
    });

    U.trackerInstrumentInput.addEventListener("blur", () => {
      window.setTimeout(() => hideInstrumentDropdown(), 120);
    });

    document.addEventListener("mousedown", (event) => {
      if (!U?.instrumentDropdown || !U?.trackerInstrumentInput) return;
      if (U.instrumentDropdown.contains(event.target) || U.trackerInstrumentInput.contains(event.target)) return;
      hideInstrumentDropdown();
    });
  }

  function init() {
    bind();
    const state = loadIntervalState();
    const latestIntervalIndex = Number.isInteger(Number(state?.latestIntervalIndex))
      ? Number(state.latestIntervalIndex)
      : 0;
    populateIntervalSlotSelect(state?.intervals || [], Math.max(0, latestIntervalIndex));
    refreshUi();
    syncTrackerPoller();
    log("Insti login task pane initialized.");
  }

  window.addEventListener("beforeunload", () => {
    clearSearchTimer();
    clearPendingHistoricalRefresh();
    stopTrackerPoller();
  });

  if (typeof Office === "undefined") return;
  Office.onReady((info) => {
    if (!info || info.host !== Office.HostType.Excel) return;
    init();
  });
})();
