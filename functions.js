/* MXChange custom functions — Version 1.3.4
 * MX.TICKER(ticker, fields...)  MX.ISIN(isin, fields...)  MX.FIELDS()
 * Tickers are exchange-suffixed (ASML, ASML.AS) so a ticker matches one listing.
 * An ISIN names the security, not the listing, so MX.ISIN returns ONE ROW PER
 * LISTING — add "ticker" or "Market" to the fields to see which row is which.
 * All work shares three workers. Deadlines include headers and JSON response bodies.
 */
/* global CustomFunctions, OfficeRuntime */
(function () {
  "use strict";
  const API_BASE = "https://muslimxchange.com/wp-json/mx-api/v1";
  const DEFAULT_BATCH_MAX = 1000, FLUSH_MS = 80, TIMEOUT_MS = 20000, PARALLEL = 3;
  const STORAGE_SESSION = "mx_session_v2", STORAGE_AUTH = "mx_auth", STORAGE_BATCH = "mx_batch_max";
  let queue = [], timer = null, active = 0;
  const jobs = [];

  async function readStorage(key) {
    try { return await OfficeRuntime.storage.getItem(key); } catch (e) { return null; }
  }
  async function readSession() {
    // A tombstone takes precedence over legacy credentials, even if removal failed.
    try {
      const raw = await OfficeRuntime.storage.getItem(STORAGE_SESSION);
      if (raw !== null && raw !== undefined) {
        const session = JSON.parse(raw);
        return session && session.id && session.auth ? session : null;
      }
      const auth = await OfficeRuntime.storage.getItem(STORAGE_AUTH);
      return auth ? {id: "legacy:" + auth, auth: auth} : null;
    } catch (e) { return null; }
  }
  async function isCurrent(session) {
    const now = await readSession();
    return !!(session && now && now.id === session.id && now.auth === session.auth);
  }
  function makeError(kind, message) {
    if (typeof CustomFunctions !== "undefined" && CustomFunctions.Error && CustomFunctions.ErrorCode) {
      return new CustomFunctions.Error(kind === "notAvailable" ? CustomFunctions.ErrorCode.notAvailable : CustomFunctions.ErrorCode.invalidValue, message);
    }
    return [[message]];
  }
  function signedOut() { return makeError("invalidValue", "Session changed — sign in and recalculate from the MXChange task pane"); }
  function normalizeFields(args) {
    const out = [];
    (function walk(v) {
      if (v === null || v === undefined) return;
      if (Array.isArray(v)) { v.forEach(walk); return; }
      if (typeof v === "object") return; // Excel's Invocation object (appended after the fields array) — never a field name
      String(v).split(",").forEach(function (s) { s = s.trim(); if (s && out.indexOf(s) < 0) out.push(s); });
    })(args);
    return out.length ? out : ["Result"];
  }
  function statusMessage(status, data) {
    if (status === 401) return "Sign in again from the MXChange task pane";
    if (status === 403) return (data && data.code === "mx_api_field_restricted") ? "Field not in your plan — see MX.FIELDS()" : ((data && data.message) || "Your membership does not include API access");
    if (status === 429) return (data && data.code === "mx_api_quota_exceeded") ? (data.message || "Daily quota reached") : "Rate limit reached — recalculate in a minute";
    if (status === 503) return "Data service temporarily unavailable — recalculate shortly";
    return (data && data.message) || ("API error " + status);
  }
  async function fetchJson(url, opts) {
    const ctrl = typeof AbortController === "undefined" ? null : new AbortController();
    let deadline;
    const timeout = new Promise(function (_, reject) {
      deadline = setTimeout(function () {
        const error = new Error("Request timed out"); error.name = "TimeoutError";
        reject(error);
        if (ctrl) ctrl.abort();
      }, TIMEOUT_MS);
    });
    const work = (async function () {
      const res = await fetch(url, Object.assign({}, opts, ctrl ? {signal: ctrl.signal} : {}));
      const data = await res.json();
      return {res: res, data: data};
    })();
    try { return await Promise.race([work, timeout]); }
    finally { clearTimeout(deadline); }
  }
  function networkError(e) {
    return e && (e.name === "TimeoutError" || e.name === "AbortError") ? "Request timed out — recalculate to retry" : "Request failed — check your connection or try again";
  }
  function sleep(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }

  // One pool for every flush and MX.FIELDS call, including retries and split batches.
  function schedule(work) {
    return new Promise(function (resolve, reject) { jobs.push({work: work, resolve: resolve, reject: reject}); pump(); });
  }
  function pump() {
    while (active < PARALLEL && jobs.length) {
      const job = jobs.shift(); active++;
      Promise.resolve().then(job.work).then(job.resolve, job.reject).finally(function () { active--; pump(); });
    }
  }
  function enqueue(kind, id, fields) {
    id = String(id === null || id === undefined ? "" : id).trim().toUpperCase();
    if (!id) return Promise.resolve(makeError("invalidValue", "Identifier is empty"));
    return new Promise(function (resolve) {
      queue.push({kind: kind, id: id, fields: fields, resolve: resolve});
      if (!timer) timer = setTimeout(flush, FLUSH_MS);
    });
  }
  function groupItems(items, max) {
    const groups = [];
    let g = null;
    items.forEach(function (it) {
      const key = it.kind + ":" + it.id;
      if (!g || (!g.ids.has(key) && g.ids.size >= max)) {
        g = {ids: new Set(), fields: new Set(), items: []};
        groups.push(g);
      }
      g.ids.add(key); it.fields.forEach(function (f) { g.fields.add(f); }); g.items.push(it);
    });
    return groups;
  }
  function resolveAll(items, result) { items.forEach(function (it) { it.resolve(result); }); }
  async function post(body, session) {
    for (let attempt = 0; attempt < 2; attempt++) {
      if (!(await isCurrent(session))) return {cancelled: true};
      let r;
      try {
        r = await fetchJson(API_BASE + "/batch", {
          method: "POST", headers: {"Content-Type": "application/json", Authorization: session.auth}, body: JSON.stringify(body)
        });
      } catch (e) { return {error: networkError(e)}; }
      if (!(await isCurrent(session))) return {cancelled: true};
      if (r.res.status === 429 && attempt === 0) {
        const wait = parseInt(r.res.headers.get("Retry-After") || "0", 10);
        if (wait > 0 && wait <= 10) { await sleep(wait * 1000); continue; }
      }
      return r;
    }
  }
  async function runGroup(group, depth, session) {
    const body = {tickers: [], isins: [], fields: Array.from(group.fields)};
    group.ids.forEach(function (key) { const i = key.indexOf(":"); body[key.slice(0, i)].push(key.slice(i + 1)); });
    const r = await post(body, session);
    if (r.cancelled) { resolveAll(group.items, signedOut()); return; }
    if (r.error) { resolveAll(group.items, makeError("notAvailable", r.error)); return; }
    const res = r.res, data = r.data;
    if (!res.ok) {
      if (res.status === 400 && data && data.code === "mx_api_batch_too_large" && depth < 1) {
        const limit = Math.max(1, parseInt(data.data && data.data.batch_max, 10) || Math.floor(group.ids.size / 2));
        try { await OfficeRuntime.storage.setItem(STORAGE_BATCH, String(limit)); } catch (e) { /* best effort */ }
        for (const subgroup of groupItems(group.items, limit)) await runGroup(subgroup, depth + 1, session);
        return;
      }
      resolveAll(group.items, makeError("invalidValue", statusMessage(res.status, data))); return;
    }
    if (!data || !Array.isArray(data.fields) || !data.results) {
      resolveAll(group.items, makeError("notAvailable", "Invalid API response — recalculate to retry")); return;
    }
    const canonical = data.fields, invalid = data.invalid_fields || [], restricted = data.restricted_fields || [];
    function cells(row, fields) {
      return fields.map(function (f) {
        const canon = canonical.find(function (c) { return c.toLowerCase() === f.toLowerCase(); });
        if (!canon) {
          if (restricted.some(function (r) { return r.toLowerCase() === f.toLowerCase(); })) return "Not in your plan: " + f;
          return invalid.indexOf(f) >= 0 ? "Invalid field: " + f : "";
        }
        return row[canon] === null || row[canon] === undefined ? "" : row[canon];
      });
    }
    group.items.forEach(function (it) {
      const bucket = data.results[it.kind], row = bucket ? bucket[it.id] : undefined;
      if (row === null || row === undefined) { it.resolve(makeError("notAvailable", it.id + " not found")); return; }
      // Several listings (an ISIN listed on more than one exchange): one row per listing, spilled down.
      const listings = row.ambiguous && Array.isArray(row.alternatives) && row.alternatives.length > 1 ? row.alternatives : [row];
      it.resolve(listings.map(function (r) { return cells(r, it.fields); }));
    });
  }
  async function flush() {
    timer = null;
    const items = queue; queue = [];
    if (!items.length) return;
    const session = await readSession();
    if (!session) { resolveAll(items, signedOut()); return; }
    const stored = parseInt(await readStorage(STORAGE_BATCH), 10);
    const groups = groupItems(items, stored > 0 ? stored : DEFAULT_BATCH_MAX);
    await Promise.all(groups.map(function (g) {
      return schedule(function () { return runGroup(g, 0, session); }).catch(function (e) { resolveAll(g.items, makeError("notAvailable", networkError(e))); });
    }));
  }
  // Excel passes a repeating parameter as ONE array, then its Invocation object as an extra last argument.
  function TICKER(ticker, fields) { return enqueue("tickers", ticker, normalizeFields(fields)); }
  function ISIN(isin, fields) { return enqueue("isins", isin, normalizeFields(fields)); }
  async function FIELDS() {
    const session = await readSession();
    if (!session) return signedOut();
    return schedule(async function () {
      if (!(await isCurrent(session))) return signedOut();
      try {
        const r = await fetchJson(API_BASE + "/fields", {headers: {Authorization: session.auth}});
        if (!(await isCurrent(session))) return signedOut();
        if (!r.res.ok) return makeError("invalidValue", statusMessage(r.res.status, r.data));
        const rows = ((r.data && r.data.fields) || []).map(function (f) { return [f.name, f.type, f.tier || "", f.available === false ? "not in your plan" : "yes"]; });
        return rows.length ? rows : makeError("notAvailable", "No fields returned");
      } catch (e) { return makeError("notAvailable", networkError(e)); }
    });
  }
  CustomFunctions.associate("TICKER", TICKER);
  CustomFunctions.associate("ISIN", ISIN);
  CustomFunctions.associate("FIELDS", FIELDS);
})();
