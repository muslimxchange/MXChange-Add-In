/* MXChange task pane — Version 1.3.4. Credentials are committed as one session record.
 * Restore never writes credentials; stale responses cannot undo sign-out or account changes.
 */
/* global Office, Excel, OfficeRuntime, document */
(function () {
  "use strict";
  const API_BASE = "https://muslimxchange.com/wp-json/mx-api/v1";
  const STORAGE_SESSION = "mx_session_v2", STORAGE_BATCH = "mx_batch_max";
  const LEGACY_KEYS = ["mx_auth", "mx_name", "mx_tier"];
  let generation = 0, mutations = Promise.resolve();
  const $ = function (id) { return document.getElementById(id); };
  Office.onReady(function () {
    $("sign-in").addEventListener("click", signIn);
    $("sign-out").addEventListener("click", signOut);
    $("recalc").addEventListener("click", recalc);
    $("key").addEventListener("keydown", function (e) { if (e.key === "Enter" && !$("sign-in").disabled) signIn(); });
    restore();
  });
  function show(state) { $("form").hidden = state !== "out"; $("session").hidden = state !== "in"; }
  function note(text, kind) { $("note").textContent = text || ""; $("note").className = "note" + (kind ? " note-" + kind : ""); }
  function sessionId() { return Date.now().toString(36) + ":" + Math.random().toString(36).slice(2) + ":" + generation; }
  function mutate(fn) {
    const next = mutations.then(fn, fn);
    mutations = next.catch(function () {});
    return next;
  }
  async function readSession() {
    try {
      const raw = await OfficeRuntime.storage.getItem(STORAGE_SESSION);
      if (raw !== null && raw !== undefined) {
        const session = JSON.parse(raw);
        return session && session.id && session.auth ? session : null;
      }
      const auth = await OfficeRuntime.storage.getItem("mx_auth");
      return auth ? {id: "legacy:" + auth, auth: auth, name: (await OfficeRuntime.storage.getItem("mx_name")) || ""} : null;
    } catch (e) { return null; }
  }
  async function current(op, session) {
    if (op !== generation) return false;
    const now = await readSession();
    return op === generation && !!(now && now.id === session.id && now.auth === session.auth);
  }
  async function restore() {
    const op = generation, session = await readSession();
    if (op !== generation) return;
    if (!session) { show("out"); return; }
    const me = await fetchMe(session.auth);
    if (!(await current(op, session))) return;
    if (me.ok) {
      // Deliberately read-only: an old response never rewrites saved auth or batch settings.
      showSession(me.data);
    } else if (me.status === 401 || me.status === 403) {
      show("in"); $("who").textContent = session.name || "";
      $("tier").textContent = me.message || "Your saved API key was rejected. Sign out and sign in again.";
      note("Sign out to remove this key, then sign in with an active key.", "error");
    } else {
      show("in"); $("who").textContent = session.name || "";
      $("tier").textContent = me.message || "Could not reach the API. Recalculate when connected.";
    }
  }
  async function signIn() {
    if ($("sign-in").disabled) return;
    const username = $("username").value.trim(), key = $("key").value.trim();
    if (!username || !key) { note("Enter your username and API key.", "error"); return; }
    const op = ++generation;
    let auth;
    try { auth = "Basic " + btoa(unescape(encodeURIComponent(username + ":" + key))); }
    catch (e) { note("The username or key could not be encoded.", "error"); return; }
    note("Signing in…"); $("sign-in").disabled = true;
    const me = await fetchMe(auth);
    if (op !== generation) return;
    $("sign-in").disabled = false;
    if (!me.ok) { note(me.status === 401 ? "Invalid username or API key." : me.message, "error"); return; }
    const stored = await mutate(async function () {
      if (op !== generation) return false;
      const max = me.data.limits && me.data.limits.batch_max;
      const session = {id: sessionId(), auth: auth, name: (me.data.user && me.data.user.name) || "", tier: me.data.tier, batch: max > 0 ? max : 1000}; // server 0 = no cap; 1,000 per request keeps responses quick
      try {
        for (const k of LEGACY_KEYS) await OfficeRuntime.storage.removeItem(k);
        await OfficeRuntime.storage.setItem(STORAGE_BATCH, String(session.batch));
        if (op !== generation) return false;
        await OfficeRuntime.storage.setItem(STORAGE_SESSION, JSON.stringify(session));
        return await current(op, session);
      } catch (e) { return false; }
    });
    if (op !== generation) return;
    if (!stored) { note("Excel could not save the key. Check add-in storage and try again.", "error"); return; }
    $("key").value = ""; showSession(me.data); recalc();
  }
  async function signOut() {
    const op = ++generation;
    $("sign-in").disabled = false;
    const removed = await mutate(async function () {
      const tombstone = {id: sessionId(), auth: null};
      let ok = true;
      try { await OfficeRuntime.storage.setItem(STORAGE_SESSION, JSON.stringify(tombstone)); }
      catch (e) { ok = false; }
      for (const k of LEGACY_KEYS.concat([STORAGE_BATCH])) {
        try { await OfficeRuntime.storage.removeItem(k); } catch (e) { ok = false; }
      }
      try {
        const raw = await OfficeRuntime.storage.getItem(STORAGE_SESSION);
        return ok && raw === JSON.stringify(tombstone) && !(await OfficeRuntime.storage.getItem("mx_auth"));
      } catch (e) { return false; }
    });
    if (op !== generation) return;
    if (!removed) { note("Could not remove all saved credentials. Revoke the key on the API keys page, then restart Excel.", "error"); return; }
    show("out"); $("key").value = "";
    await recalc();
    if (op === generation) note("Signed out. MX formulas require sign-in to refresh.");
  }
  function showSession(me) {
    show("in"); $("who").textContent = me.user && me.user.name ? me.user.name : "";
    const t = me.tier;
    $("tier").className = "tier" + (t === "none" ? " tier-none" : "");
    const scope = me.fields === "full" ? "all fields" : "basic fields";
    const quota = me.limits && me.limits.lookup_quota_per_day;
    $("tier").textContent = t === "none" ? "Your membership does not include API lookups. Upgrade on muslimxchange.com."
      : t === "trial" ? "Free trial — " + (quota || 1) + " lookup" + ((quota || 1) === 1 ? "" : "s") + " per day of the basic fields. Upgrade on muslimxchange.com for full access."
      : t === "bulk" ? "Lookups and bulk export enabled (" + scope + ")." : "Lookups enabled — MX.TICKER, MX.ISIN and MX.FIELDS (" + scope + ").";
    note("");
  }
  async function fetchMe(auth) {
    const ctrl = typeof AbortController === "undefined" ? null : new AbortController();
    let deadline;
    const timeout = new Promise(function (_, reject) {
      deadline = setTimeout(function () {
        const error = new Error("Sign-in request timed out. Try again.");
        reject(error); if (ctrl) ctrl.abort();
      }, 20000);
    });
    try {
      return await Promise.race([(async function () {
        const res = await fetch(API_BASE + "/me", Object.assign({headers: {Authorization: auth}}, ctrl ? {signal: ctrl.signal} : {}));
        const data = await res.json();
        const valid = data && typeof data === "object" && data.user && typeof data.tier === "string";
        return {ok: res.ok && !!valid, status: res.status, data: data, message: (data && data.message) || "Invalid API response. Try again."};
      })(), timeout]);
    } catch (e) { return {ok: false, status: 0, data: null, message: e.message || "Network error. Try again."}; }
    finally { clearTimeout(deadline); }
  }
  async function recalc() {
    try {
      await Excel.run(async function (ctx) { ctx.workbook.application.calculate(Excel.CalculationType.full); await ctx.sync(); });
    } catch (e) { note("Press Ctrl+Alt+F9 to recalculate."); }
  }
})();
