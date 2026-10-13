/* shared.js - syncs the local product history (storage.js) with the shared
 * team database behind /api/store, so everyone who logs in sees the same
 * products and version history.
 *
 * Design: the app always works from the local copy (fast, works offline).
 * This layer
 *   - pulls what the team has saved and merges it into the local copy,
 *   - pushes new saves up as they happen,
 *   - retries anything that failed to upload (e.g. while offline).
 *
 * If the database hasn't been connected in Vercel yet, the API answers
 * "not configured" and the app simply stays in local-only mode, exactly as
 * before. Nothing breaks.
 *
 * Images are never uploaded (they're large and come from the PD sheet);
 * they stay on whichever computer imported the sheet.
 *
 * No DOM code in here on purpose - app.js renders the status.
 */
(function (global) {
  const ENDPOINT = "/api/store";
  const NAME_KEY = "pbg_username";
  const MAX_HISTORY = 25;

  // state: unknown | checking | connected | local | error
  const status = { state: "unknown", lastSync: null, message: "" };
  const listeners = [];
  let pulling = null;
  let dataVersion = 0; // bumps whenever the local copy changes because of a sync

  function emit() {
    listeners.forEach((fn) => { try { fn(status); } catch (e) { /* a UI listener must never break syncing */ } });
  }

  function setStatus(state, message) {
    status.state = state;
    status.message = message || "";
    emit();
  }

  async function api(method, query, body) {
    const resp = await fetch(ENDPOINT + (query || ""), {
      method,
      headers: body ? { "Content-Type": "application/json" } : undefined,
      body: body ? JSON.stringify(body) : undefined,
      credentials: "same-origin",
      cache: "no-store",
    });
    let json = null;
    try { json = await resp.json(); } catch (e) { /* non-JSON body */ }
    return { status: resp.status, json };
  }

  function idOf(entry) {
    return entry.id || `L${entry.savedAt}`;
  }

  function withoutImages(brief) {
    return PBStorage.stripImages(brief);
  }

  function unsyncedEntries(product) {
    return product.history.filter((e) => e.synced !== true);
  }

  // Folds a record from the shared database into the local copy. The two
  // histories are combined by entry id (nothing is lost on either side), and
  // the product's current brief is its newest saved entry.
  function mergeRemote(remote) {
    const local = PBStorage.load(remote.key);
    const localEntries = local ? local.history : [];
    const localById = new Map(localEntries.map((e) => [idOf(e), e]));
    const seen = new Set();
    const merged = [];

    remote.history.forEach((re) => {
      const id = idOf(re);
      seen.add(id);
      const le = localById.get(id);
      if (le) {
        le.id = id;
        le.synced = true; // keep the local copy: it may still carry images
        merged.push(le);
      } else {
        merged.push({ ...re, id, synced: true });
      }
    });
    // Saves that exist only here (not uploaded yet) stay in the history.
    localEntries.forEach((le) => { if (!seen.has(idOf(le))) merged.push(le); });

    merged.sort((a, b) => a.savedAt - b.savedAt);
    const history = merged.length > MAX_HISTORY ? merged.slice(-MAX_HISTORY) : merged;

    const newest = history[history.length - 1];
    const brief = JSON.parse(JSON.stringify(newest.brief));
    // Pictures aren't shared, so a brief arriving from the team has none;
    // keep this computer's own where the instruction / QC row is unchanged.
    PBStorage.carryImages(brief, local && local.brief);
    if (!brief.images) brief.images = { hero: [], feature: [] };

    dataVersion++;
    PBStorage.putRecord({
      key: remote.key,
      item: remote.item || (local && local.item) || "",
      articleNo: remote.articleNo || (local && local.articleNo) || "",
      createdAt: remote.createdAt,
      updatedAt: remote.updatedAt,
      sharedUpdatedAt: remote.updatedAt,
      remote: true,
      dirty: !!(local && local.dirty),
      brief,
      history,
    });
  }

  async function pull() {
    if (pulling) return pulling;
    pulling = (async () => {
      setStatus("checking");
      let list;
      try {
        list = await api("GET");
      } catch (e) {
        setStatus("error", "Can't reach the shared history (offline?). Showing this browser's copy.");
        return false;
      }
      if (list.status === 503 && list.json && list.json.configured === false) {
        status.lastSync = null;
        setStatus("local", "Shared history isn't connected yet - saves stay on this browser only.");
        return false;
      }
      if (list.status === 401) {
        setStatus("error", "Login expired - reload the page to log in again.");
        return false;
      }
      if (list.status !== 200 || !list.json) {
        setStatus("error", (list.json && list.json.error) || `Shared history error (HTTP ${list.status}).`);
        return false;
      }

      try {
        const remote = list.json.products || [];
        const remoteKeys = new Set(remote.map((p) => p.key));
        const local = PBStorage.getAll();

        const stale = remote.filter((s) => !local[s.key] || local[s.key].sharedUpdatedAt !== s.updatedAt);
        await Promise.all(stale.map(async (s) => {
          const r = await api("GET", "?key=" + encodeURIComponent(s.key));
          if (r.status === 200 && r.json && r.json.product) mergeRemote(r.json.product);
        }));

        // A product we know the team had, which is now gone, was deleted by
        // someone else - remove our copy too (unless it holds unshared work).
        Object.values(PBStorage.getAll()).forEach((p) => {
          if (p.remote && !remoteKeys.has(p.key)) {
            if (unsyncedEntries(p).length) {
              p.remote = false;
              PBStorage.putRecord(p);
            } else {
              PBStorage.remove(p.key);
              dataVersion++;
            }
          }
        });

        // Retry anything whose upload failed earlier.
        const retry = Object.values(PBStorage.getAll()).filter((p) => p.dirty);
        for (const p of retry) await pushProduct(p.key, { quiet: true });
      } catch (e) {
        setStatus("error", "Shared history sync failed: " + e.message);
        return false;
      }

      status.lastSync = Date.now();
      setStatus("connected");
      return true;
    })().finally(() => { pulling = null; });
    return pulling;
  }

  // Uploads a product's not-yet-shared saves. Returns { ok, shared }.
  async function pushProduct(key, opts) {
    const quiet = opts && opts.quiet;
    if (status.state === "local") return { ok: false, shared: false };
    const product = PBStorage.load(key);
    if (!product) return { ok: true, shared: true };
    const pending = unsyncedEntries(product);
    if (!pending.length) return { ok: true, shared: true };

    pending.forEach((e) => { if (!e.id) e.id = idOf(e); });
    PBStorage.putRecord(product);

    try {
      const r = await api("POST", "", {
        key,
        entries: pending.map((e) => ({
          id: e.id, savedAt: e.savedAt, version: e.version, label: e.label,
          savedBy: e.savedBy || "", brief: withoutImages(e.brief),
        })),
      });
      if (r.status === 503 && r.json && r.json.configured === false) {
        status.lastSync = null;
        setStatus("local", "Shared history isn't connected yet - saves stay on this browser only.");
        return { ok: false, shared: false };
      }
      if (r.status !== 200 || !r.json || !r.json.product) throw new Error((r.json && r.json.error) || `HTTP ${r.status}`);

      const fresh = PBStorage.load(key) || product;
      fresh.dirty = false;
      PBStorage.putRecord(fresh);
      mergeRemote(r.json.product);
      status.lastSync = Date.now();
      if (status.state !== "connected") setStatus("connected"); else emit();
      return { ok: true, shared: true };
    } catch (e) {
      const fresh = PBStorage.load(key) || product;
      fresh.dirty = true;
      PBStorage.putRecord(fresh);
      if (!quiet) setStatus("error", "Saved on this browser, but couldn't reach the shared history. It will retry automatically.");
      return { ok: false, shared: false, error: e.message };
    }
  }

  // Deletes a product for everyone. Returns { ok, error }.
  async function removeProduct(key) {
    if (status.state === "local") return { ok: true };
    try {
      const r = await api("DELETE", "?key=" + encodeURIComponent(key));
      if (r.status === 503) return { ok: true };
      if (r.status !== 200) return { ok: false, error: (r.json && r.json.error) || `HTTP ${r.status}` };
      return { ok: true };
    } catch (e) {
      return { ok: false, error: "Can't reach the shared history (offline?)." };
    }
  }

  function localOnlyProducts() {
    return Object.values(PBStorage.getAll()).filter((p) => unsyncedEntries(p).length > 0);
  }

  // One-time catch-up for products saved before sharing existed.
  async function uploadLocalOnly() {
    let uploaded = 0;
    for (const p of localOnlyProducts()) {
      const r = await pushProduct(p.key, { quiet: true });
      if (r.ok) uploaded++;
    }
    return uploaded;
  }

  global.PBShared = {
    status,
    get dataVersion() { return dataVersion; },
    // True once the team database has answered successfully (and not since
    // reported itself missing) - what product cards use to decide whether an
    // unuploaded save is worth flagging.
    get isShared() { return status.state === "connected" || !!status.lastSync; },
    pull,
    pushProduct,
    removeProduct,
    localOnlyProducts,
    uploadLocalOnly,
    onChange(fn) { listeners.push(fn); },
    getName() { try { return localStorage.getItem(NAME_KEY) || ""; } catch (e) { return ""; } },
    setName(name) { try { localStorage.setItem(NAME_KEY, (name || "").trim()); } catch (e) { /* ignore */ } },
  };
})(window);
