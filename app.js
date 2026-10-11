/* The Ultraspeaker Image Lab – add-in PowerPoint per regolare le immagini (istogramma, livelli, colore, zoom, ritaglio) */
(function () {
  "use strict";

  /* ---------- Utilità ---------- */
  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => Array.from(r.querySelectorAll(s));
  const clamp = (v, a, b) => (v < a ? a : v > b ? b : v);
  const store = {
    get(k, d) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch (e) { return d; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* archiviazione non disponibile */ } },
  };
  let toastT;
  function toast(msg) {
    const el = $("#toast"); el.textContent = msg; el.classList.add("show");
    clearTimeout(toastT); toastT = setTimeout(() => el.classList.remove("show"), 2600);
  }
  const wait = (ms) => new Promise((r) => setTimeout(r, ms));
  const mq = (q) => { try { return window.matchMedia(q).matches; } catch (e) { return false; } };
  // Telefono o tablet: schermo touch con menu di condivisione del sistema (Foto, WhatsApp, Messaggi…)
  const isMobile = () => mq("(pointer: coarse)") && typeof navigator.share === "function";
  // Solo touch (nessun mouse o trackpad): lo zoom si fa con le dita
  const touchOnly = () => mq("(hover: none) and (pointer: coarse)");
  function canvas(w, h) { const c = document.createElement("canvas"); c.width = Math.max(1, w); c.height = Math.max(1, h); return c; }

  /* ---------- Lingua ---------- */
  const I18N = window.IL_I18N;
  const LANGS = I18N.langs.map((l) => l[0]);
  let LANG = "it";
  const t = (k, vars) => {
    let v = (I18N[LANG] && I18N[LANG][k]) || I18N.it[k] || k;
    if (vars) for (const n in vars) v = v.replace("{" + n + "}", vars[n]);
    return v;
  };
  function detectLang() {
    const saved = store.get("imagelab.lang", null);
    if (saved && LANGS.includes(saved)) return saved;
    let l = "";
    try { if (inOffice && Office.context && Office.context.displayLanguage) l = Office.context.displayLanguage; } catch (e) { /* ignora */ }
    if (!l) l = (navigator.languages && navigator.languages[0]) || navigator.language || "it";
    l = l.slice(0, 2).toLowerCase();
    return LANGS.includes(l) ? l : "en";
  }
  function applyStatic() {
    document.documentElement.lang = LANG;
    $$("[data-i18n]").forEach((el) => { el.textContent = t(el.dataset.i18n); });
    $$("[data-i18n-aria]").forEach((el) => el.setAttribute("aria-label", t(el.dataset.i18nAria)));
    $$("[data-i18n-title]").forEach((el) => el.setAttribute("title", t(el.dataset.i18nTitle)));
    buildSliders();
    buildPresets();
    buildFormats();
    updateTexts();
    updateZoomUI();
    if (cropMode) updateCropInfo();
  }

  /* ---------- Stato ---------- */
  const D = () => Object.assign(IL.defaults(), { crop: null, rq: 0, rot: 0 }); // regolazioni di partenza (con ritaglio)
  const snap = (p) => JSON.stringify(Object.assign(D(), p));
  let inOffice = false;
  let mode = "boot";           // demo | office | unsupported
  let P = D();                 // regolazioni correnti (compreso il ritaglio)
  let baseline = D();          // regolazioni già applicate alla slide
  let preset = "nat";
  let ch = "rgb", before = false, picking = false, busy = false, restoreArmed = false;
  let orig = null;             // foto di partenza intera, dritta: {canvas, w, h}
  let base = null;             // foto di partenza ruotata (quarti di giro + raddrizzamento): {canvas, w, h}
  let baseKey = null, workCrop = null;
  let full = null, prev = null;// zona di lavoro (ritagliata): {data, w, h} a piena risoluzione e ridotta
  let fullCanvas = null, prevCanvas = null, prevSrcCanvas = null;
  let prevOut = null, prevHist = null, srcHist = null, workKey = null;
  const cachePrev = {}, cacheFull = {};
  let shape = null;            // {id, slideId, origId, edited}

  /* ---------- Annulla / Ripeti ---------- */
  const hist = { s: [], i: -1 };
  function histReset() { hist.s = [{ v: snap(P), preset }]; hist.i = 0; updateTexts(); }
  function commit() {
    const v = snap(P);
    if (hist.i >= 0 && hist.s[hist.i].v === v) return;
    hist.s = hist.s.slice(0, hist.i + 1);
    hist.s.push({ v, preset });
    if (hist.s.length > 120) hist.s.shift();
    hist.i = hist.s.length - 1;
    updateTexts();
  }
  function goHistory(d) {
    if (cropMode) exitCrop(false);
    const j = hist.i + d;
    if (j < 0 || j >= hist.s.length) return;
    hist.i = j;
    P = Object.assign(D(), JSON.parse(hist.s[j].v));
    preset = hist.s[j].preset;
    syncSliders(); changed();
  }

  /* ---------- Cursori ---------- */
  const GROUPS = [
    ["g.light", [["exp", -100, 100, "bi"], ["con", -100, 100, "bi"], ["sat", -100, 100, "bi"]]],
    ["g.tone", [["hl", 0, 100, ""], ["sh", 0, 100, ""], ["sharp", 0, 100, ""], ["nr", 0, 100, ""]]],
    ["g.wb", [["temp", -100, 100, "temp"], ["tint", -100, 100, "tint"]]],
  ];
  const BIPOLAR = { exp: 1, con: 1, sat: 1, temp: 1, tint: 1 };
  const fmt = (k, v) => (BIPOLAR[k] && v > 0 ? "+" : "") + v + "%";

  function buildSliders() {
    const box = $("#sliders");
    box.innerHTML = "";
    GROUPS.forEach(([title, items]) => {
      const g = document.createElement("div");
      g.className = "group";
      const h = document.createElement("div");
      h.className = "h"; h.style.marginBottom = "2px"; h.textContent = t(title);
      g.appendChild(h);
      items.forEach(([k, min, max, cls]) => {
        const row = document.createElement("div");
        row.className = "srow";
        row.innerHTML = `<label for="s-${k}"></label><input id="s-${k}" class="sl ${cls}" type="range" min="${min}" max="${max}" step="1"><output for="s-${k}"></output>`;
        row.querySelector("label").textContent = t("s." + k);
        const inp = row.querySelector("input");
        inp.value = P[k];
        row.querySelector("output").textContent = fmt(k, P[k]);
        inp.addEventListener("input", () => { P[k] = +inp.value; preset = "custom"; row.querySelector("output").textContent = fmt(k, P[k]); changed(); });
        inp.addEventListener("change", commit);
        inp.addEventListener("dblclick", () => { P[k] = 0; inp.value = 0; preset = "custom"; row.querySelector("output").textContent = fmt(k, 0); changed(); commit(); });
        g.appendChild(row);
      });
      box.appendChild(g);
    });
  }
  function syncSliders() {
    GROUPS.forEach(([, items]) => items.forEach(([k]) => {
      const inp = $("#s-" + k); if (!inp) return;
      inp.value = P[k]; inp.parentElement.querySelector("output").textContent = fmt(k, P[k]);
    }));
    $("#lvB").value = P.b; $("#lvW").value = P.w; $("#lvM").value = Math.round(P.b + P.p * (P.w - P.b));
    // Migliora qualità
    $("#qBtn").setAttribute("aria-pressed", String(P.q > 0));
    $("#qRow").hidden = !(P.q > 0);
    if (P.q > 0) { $("#qAmt").value = P.q; $("#qOut").textContent = P.q + "%"; }
    syncRotUI();
  }

  /* ---------- Stili rapidi ---------- */
  const PRESETS = {
    nat: () => ({}),
    stage: () => ({ con: 25, sat: 12, sh: 25, hl: 15, sharp: 30 }),
    bw: () => ({ sat: -100, con: 20, sh: 10 }),
    warm: () => ({ temp: 35, tint: 5, sat: 8 }),
    auto: () => Object.assign(srcHist ? IL.autoLevels(srcHist) : {}, { con: 10, sat: 8, sh: 15, hl: 10 }),
  };
  function buildPresets() {
    const box = $("#presets");
    box.innerHTML = "";
    Object.keys(PRESETS).forEach((k) => {
      const b = document.createElement("button");
      b.type = "button"; b.className = "chip"; b.dataset.p = k;
      b.textContent = t("p." + k);
      b.setAttribute("aria-pressed", String(preset === k));
      // lo stile cambia le regolazioni, il ritaglio resta
      b.addEventListener("click", () => { P = Object.assign(D(), PRESETS[k](), { crop: P.crop, q: P.q, qn: P.qn, rq: P.rq, rot: P.rot }); preset = k; syncSliders(); changed(); commit(); });
      box.appendChild(b);
    });
  }

  /* ---------- Testi che cambiano con lo stato ---------- */
  function setStatus(key, cls) {
    const s = $("#status");
    s.className = "status " + (cls || "");
    $("#statusText").textContent = t(key);
  }
  function updateTexts() {
    $$("#presets .chip").forEach((b) => b.setAttribute("aria-pressed", String(b.dataset.p === preset)));
    $("#note").textContent = mode === "demo" ? (window.IL_ARTIFACT ? t("noteArtifact") : isMobile() ? t("noteMobile") : t("noteDemo")) : t("note");
    $("#applyBtn").hidden = !!(mode === "demo" && window.IL_ARTIFACT);
    const apply = $("#applyBtn");
    apply.textContent = busy ? t("applying") : mode === "demo" ? (shareFile ? t("shareReady") : isMobile() ? t("shareMobile") : t("applyDemo")) : t("apply");
    const dirty = snap(P) !== snap(baseline);
    apply.disabled = busy || cropMode || (mode === "office" && !dirty);
    const rb = $("#restoreBtn");
    const showRestore = mode === "office" && shape && shape.edited;
    rb.hidden = !showRestore;
    $("#foot").classList.toggle("three", !!showRestore);
    rb.textContent = restoreArmed ? t("restoreConfirm") : t("restore");
    rb.classList.toggle("warn", restoreArmed);
    $("#beforeBtn").setAttribute("aria-pressed", String(before));
    $("#pickBtn").setAttribute("aria-pressed", String(picking));
    $("#pickHint").hidden = !picking;
    $("#preview").classList.toggle("picking", picking);
    $("#preview").classList.toggle("before", before);
    $$(".pill").forEach((b) => b.setAttribute("aria-pressed", String(b.dataset.ch === ch)));
    $("#undoBtn").disabled = busy || hist.i <= 0;
    $("#redoBtn").disabled = busy || hist.i < 0 || hist.i >= hist.s.length - 1;
    $("#cropBtn").setAttribute("aria-pressed", String(cropMode));
    $("#rotBtn").setAttribute("aria-pressed", String(!$("#rotPanel").hidden));
    $("#rotGrid").hidden = $("#rotPanel").hidden || cropMode;
    if (!$("#rotPanel").hidden) syncRotUI();
    if (mode === "demo") setStatus("status.demo", "");
    else if (shape) setStatus(shape.edited ? "status.edited" : "status.selected", "");
  }

  /* ---------- Viste ---------- */
  function show(view, titleKey, textKey) {
    $("#editor").hidden = view !== "editor";
    $("#foot").hidden = view !== "editor";
    $("#empty").hidden = view !== "empty";
    $("#loading").hidden = view !== "loading";
    if (view === "empty") {
      $("#emptyTitle").textContent = t(titleKey);
      $("#emptyText").textContent = t(textKey);
    }
    if (view === "editor") requestAnimationFrame(drawView);
  }

  /* ---------- Zona di lavoro: foto intera o ritagliata ---------- */
  const PREVIEW_MAX = 760;
  function cropRectPx(c, w, h) {
    if (!c) return { x: 0, y: 0, w, h };
    const x = clamp(Math.round(c.x * w), 0, w - 1), y = clamp(Math.round(c.y * h), 0, h - 1);
    return { x, y, w: clamp(Math.round(c.w * w), 1, w - x), h: clamp(Math.round(c.h * h), 1, h - y) };
  }
  /* Rotazione: quarti di giro (P.rq) più un angolo fine (P.rot, ±45°).
     Con l'angolo fine la foto si ingrandisce quel tanto che basta a riempire il riquadro, senza angoli vuoti:
     le misure restano quelle della foto, così il ritaglio resta valido mentre si ruota. */
  const rotQ = () => (((P.rq | 0) % 4) + 4) % 4;
  const rotA = () => clamp(+P.rot || 0, -45, 45);
  function ensureBase() {
    if (!orig) { base = null; return false; }
    const q = rotQ(), a = rotA(), k = q + "/" + a.toFixed(2);
    if (base && baseKey === k) return false;
    baseKey = k;
    if (!q && !a) { base = orig; return true; }
    const W = q & 1 ? orig.h : orig.w, H = q & 1 ? orig.w : orig.h;
    const th = Math.abs(a) * Math.PI / 180, c = Math.cos(th), s = Math.sin(th);
    const m = a ? Math.max((W * c + H * s) / W, (W * s + H * c) / H) * 1.002 : 1;
    const cv = canvas(W, H), x = cv.getContext("2d", { willReadFrequently: true });
    x.imageSmoothingEnabled = true; x.imageSmoothingQuality = "high";
    x.translate(W / 2, H / 2); x.rotate((q * 90 + a) * Math.PI / 180); x.scale(m, m);
    x.drawImage(orig.canvas, -orig.w / 2, -orig.h / 2);
    base = { canvas: cv, w: W, h: H };
    return true;
  }
  function ensureWorking() {
    if (!orig) return false;
    ensureBase();
    const want = cropMode ? null : P.crop;
    const ck = JSON.stringify(want);
    const key = ck + "|" + baseKey;
    if (key === workKey && full) return false;
    const r = cropRectPx(want, base.w, base.h);
    // stessa zona e stesse misure (si sta solo raddrizzando): zoom e posizione della vista restano
    const keepView = !!full && workCrop === ck && full.w === r.w && full.h === r.h;
    if (!want) fullCanvas = base.canvas;
    else {
      fullCanvas = canvas(r.w, r.h);
      fullCanvas.getContext("2d", { willReadFrequently: true }).drawImage(base.canvas, r.x, r.y, r.w, r.h, 0, 0, r.w, r.h);
    }
    // i pixel a piena risoluzione si leggono solo quando servono (zoom, esportazione)
    const fc = fullCanvas;
    full = { w: r.w, h: r.h, _d: null, get data() { return this._d || (this._d = fc.getContext("2d", { willReadFrequently: true }).getImageData(0, 0, this.w, this.h).data); } };
    const k = Math.min(1, PREVIEW_MAX / Math.max(r.w, r.h));
    const pw = Math.max(1, Math.round(r.w * k)), ph = Math.max(1, Math.round(r.h * k));
    prevSrcCanvas = canvas(pw, ph);
    const px = prevSrcCanvas.getContext("2d", { willReadFrequently: true });
    px.imageSmoothingQuality = "high";
    px.drawImage(base.canvas, r.x, r.y, r.w, r.h, 0, 0, pw, ph);
    prev = { data: px.getImageData(0, 0, pw, ph).data, w: pw, h: ph };
    prevOut = new ImageData(pw, ph);
    prevCanvas = canvas(pw, ph);
    delete cachePrev.blur; delete cacheFull.blur; cachePrev.src = cacheFull.src = null;
    srcHist = IL.histogram(prev.data);
    // il riquadro segue la forma della foto fino a 21:9; oltre, la foto resta intera con bande sopra e sotto
    $("#preview").style.setProperty("--ar", String(Math.min(r.w / r.h, 21 / 9)));
    workKey = key; workCrop = ck;
    if (!keepView) { view.z = 1; view.cx = r.w / 2; view.cy = r.h / 2; }
    return true;
  }

  /* ---------- Anteprima e istogramma ---------- */
  let raf = 0;
  function changed() {
    restoreArmed = false;
    shareFile = null;
    updateTexts();
    if (!raf) raf = requestAnimationFrame(() => { raf = 0; render(); });
  }
  function render() {
    if (!orig) return;
    ensureWorking();
    IL.process(prev.data, prev.w, prev.h, P, prevOut.data, cachePrev);
    prevCanvas.getContext("2d").putImageData(prevOut, 0, 0);
    prevHist = IL.histogram(before ? prev.data : prevOut.data);
    drawView();
    drawHist();
    placeLevelLabels();
  }

  /* ---------- Vista con zoom ---------- */
  const view = { z: 1, cx: 0, cy: 0, s0: 1, zmax: 8 };
  let vt = null;               // trasformazione corrente: {x0, y0, s} (coordinate della foto ↔ pixel CSS)
  let tmpC = null;
  function fitScale(bw, bh) { return Math.min(bw / full.w, bh / full.h); }
  function clampView(bw, bh, s) {
    const vw = bw / s, vh = bh / s;
    view.cx = vw >= full.w ? full.w / 2 : clamp(view.cx, vw / 2, full.w - vw / 2);
    view.cy = vh >= full.h ? full.h / 2 : clamp(view.cy, vh / 2, full.h - vh / 2);
  }
  let vraf = 0;
  function requestDraw() { if (!vraf) vraf = requestAnimationFrame(() => { vraf = 0; drawView(); }); }
  function drawView() {
    if (!full || !prevCanvas) return;
    const pv = $("#pv"), box = $("#preview");
    const bw = box.clientWidth, bh = box.clientHeight;
    if (!bw || !bh) return;
    const dpr = window.devicePixelRatio || 1;
    const W = Math.round(bw * dpr), H = Math.round(bh * dpr);
    if (pv.width !== W || pv.height !== H) { pv.width = W; pv.height = H; }
    const ctx = pv.getContext("2d");
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.clearRect(0, 0, W, H);
    view.s0 = fitScale(bw, bh);
    view.zmax = clamp(8 / view.s0, 2, 64);   // fino a 8 pixel dello schermo per ogni pixel della foto
    view.z = clamp(view.z, 1, view.zmax);
    const s = view.s0 * view.z;
    clampView(bw, bh, s);
    const vw = bw / s, vh = bh / s;
    if (cropMode && draft && frameLock && frame) lockFrame(s, vw, vh);
    const x0 = view.cx - vw / 2, y0 = view.cy - vh / 2;
    vt = { x0, y0, s, bw, bh };
    const sx = Math.max(0, x0), sy = Math.max(0, y0), ex = Math.min(full.w, x0 + vw), ey = Math.min(full.h, y0 + vh);
    const sw = ex - sx, sh = ey - sy;
    if (sw > 0 && sh > 0) {
      if (view.z <= 1.001) {
        // foto intera: si usa l'anteprima già calcolata
        const k = prev.w / full.w;
        ctx.imageSmoothingEnabled = true; ctx.imageSmoothingQuality = "high";
        ctx.drawImage(before ? prevSrcCanvas : prevCanvas, sx * k, sy * k, sw * k, sh * k, (sx - x0) * s * dpr, (sy - y0) * s * dpr, sw * s * dpr, sh * s * dpr);
      } else {
        // ingrandimento: si elabora solo la zona visibile, alla piena risoluzione
        const ix = Math.floor(sx), iy = Math.floor(sy);
        const iw = Math.min(full.w - ix, Math.ceil(ex) - ix), ih = Math.min(full.h - iy, Math.ceil(ey) - iy);
        const dev = s * dpr;
        let tw = iw, th = ih;
        if (dev < 1) { tw = Math.max(1, Math.round(iw * dev)); th = Math.max(1, Math.round(ih * dev)); }
        if (!tmpC) tmpC = canvas(tw, th);
        if (tmpC.width !== tw || tmpC.height !== th) { tmpC.width = tw; tmpC.height = th; }
        const tx = tmpC.getContext("2d", { willReadFrequently: true });
        tx.imageSmoothingEnabled = true; tx.imageSmoothingQuality = "high";
        tx.clearRect(0, 0, tw, th);
        tx.drawImage(fullCanvas, ix, iy, iw, ih, 0, 0, tw, th);
        if (!before) {
          const id = tx.getImageData(0, 0, tw, th);
          const out = new ImageData(tw, th);
          IL.process(id.data, tw, th, P, out.data, {}, Math.min(full.w, full.h) * (tw / iw));
          tx.putImageData(out, 0, 0);
        }
        ctx.imageSmoothingEnabled = dev < 3; // oltre 3×: pixel netti per vedere i dettagli
        ctx.drawImage(tmpC, 0, 0, tw, th, (ix - x0) * s * dpr, (iy - y0) * s * dpr, iw * s * dpr, ih * s * dpr);
      }
    }
    $("#preview").classList.toggle("zoomed", view.z > 1.001);
    updateZoomUI();
    positionCrop();
  }
  // Ritaglio: mentre si ingrandisce o si sposta la vista il riquadro resta fermo sullo schermo e la foto si muove sotto
  let frame = null, frameLock = false; // frame: riquadro in pixel CSS dentro l'anteprima
  function lockFrame(s, vw, vh) {
    let dw = frame.w / s, dh = frame.h / s;
    const f = Math.min(1, base.w / dw, base.h / dh); // allontanandosi il riquadro non può superare la foto
    dw *= f; dh *= f;
    const x0 = view.cx - vw / 2, y0 = view.cy - vh / 2;
    let dx = x0 + (frame.l + frame.w / 2) / s - dw / 2, dy = y0 + (frame.t + frame.h / 2) / s - dh / 2;
    // ai bordi si ferma la foto, non il riquadro
    if (vw < base.w) { const c = dx < 0 ? -dx : dx + dw > base.w ? base.w - (dx + dw) : 0; view.cx += c; dx += c; }
    if (vh < base.h) { const c = dy < 0 ? -dy : dy + dh > base.h ? base.h - (dy + dh) : 0; view.cy += c; dy += c; }
    draft.w = dw; draft.h = dh;
    draft.x = clamp(dx, 0, base.w - dw); draft.y = clamp(dy, 0, base.h - dh);
  }
  function setZoom(z, ax, ay) {
    // ax, ay: punto dello schermo (relativo al riquadro) che resta fermo; senza, il centro
    if (!full || !vt) return;
    const bw = vt.bw, bh = vt.bh;
    const nz = clamp(z, 1, view.zmax);
    if (ax == null && cropMode && frame) { ax = frame.l + frame.w / 2; ay = frame.t + frame.h / 2; } // si ingrandisce attorno al riquadro
    if (ax == null) { ax = bw / 2; ay = bh / 2; }
    if (cropMode) frameLock = true;
    const fx = vt.x0 + ax / vt.s, fy = vt.y0 + ay / vt.s;
    const ns = view.s0 * nz;
    view.z = nz;
    view.cx = fx - ax / ns + bw / (2 * ns);
    view.cy = fy - ay / ns + bh / (2 * ns);
    requestDraw();
  }
  function updateZoomUI() {
    const v = view.zmax > 1 ? Math.log(view.z) / Math.log(view.zmax) : 0;
    const pct = Math.round(v * 100);
    $("#zthumb").style.bottom = pct + "%";
    $("#zfill").style.height = pct + "%";
    const zb = $("#zoombar");
    zb.setAttribute("aria-valuenow", String(pct));
    const lbl = view.z <= 1.001 ? t("fit") : Math.round(view.s0 * view.z * 100) + "%";
    zb.setAttribute("aria-valuetext", lbl);
    $("#zoomLbl").textContent = lbl;
    $("#zoomLbl").disabled = view.z <= 1.001;
    $("#zoomHint").textContent = touchOnly() ? t("zoomHintTouch") : view.z > 1.001 ? t("zoomHintMouse") : "";
  }
  function bindZoom() {
    const zb = $("#zoombar"), track = zb.querySelector(".ztrack"), pv = $("#preview");
    const fromY = (clientY) => {
      const r = track.getBoundingClientRect();
      const v = clamp(1 - (clientY - r.top) / r.height, 0, 1);
      setZoom(Math.pow(view.zmax, v));
    };
    let drag = false;
    zb.addEventListener("pointerdown", (e) => { drag = true; zb.setPointerCapture(e.pointerId); fromY(e.clientY); e.preventDefault(); });
    zb.addEventListener("pointermove", (e) => { if (drag) fromY(e.clientY); });
    zb.addEventListener("pointerup", () => { drag = false; });
    zb.addEventListener("pointercancel", () => { drag = false; });
    zb.addEventListener("dblclick", () => setZoom(1));
    zb.addEventListener("keydown", (e) => {
      const v = view.zmax > 1 ? Math.log(view.z) / Math.log(view.zmax) : 0;
      const step = e.shiftKey ? 0.2 : 0.05;
      let nv = null;
      if (e.key === "ArrowUp" || e.key === "ArrowRight" || e.key === "PageUp") nv = v + step;
      else if (e.key === "ArrowDown" || e.key === "ArrowLeft" || e.key === "PageDown") nv = v - step;
      else if (e.key === "Home") nv = 0; else if (e.key === "End") nv = 1;
      if (nv == null) return;
      e.preventDefault(); setZoom(Math.pow(view.zmax, clamp(nv, 0, 1)));
    });
    $("#zoomLbl").addEventListener("click", () => setZoom(1));

    // Mouse: tasto destro e trascina per spostare la foto ingrandita
    let pan = null;
    pv.addEventListener("contextmenu", (e) => { if (view.z > 1.001 || pan) e.preventDefault(); });
    pv.addEventListener("pointerdown", (e) => {
      if (e.pointerType !== "mouse" || e.button !== 2 || view.z <= 1.001) return;
      e.preventDefault();
      pan = { x: e.clientX, y: e.clientY, cx: view.cx, cy: view.cy };
      pv.setPointerCapture(e.pointerId); pv.classList.add("panning");
    });
    pv.addEventListener("pointermove", (e) => {
      if (!pan) return;
      if (cropMode) frameLock = true;
      view.cx = pan.cx - (e.clientX - pan.x) / vt.s;
      view.cy = pan.cy - (e.clientY - pan.y) / vt.s;
      requestDraw();
    });
    const endPan = () => { if (pan) { pan = null; pv.classList.remove("panning"); } };
    pv.addEventListener("pointerup", endPan);
    pv.addEventListener("pointercancel", endPan);

    // Touch: pizzica per ingrandire, due dita per spostare
    let pinch = null;
    const mid = (a, b) => ({ x: (a.clientX + b.clientX) / 2, y: (a.clientY + b.clientY) / 2 });
    const dist = (a, b) => Math.hypot(a.clientX - b.clientX, a.clientY - b.clientY);
    pv.addEventListener("touchstart", (e) => {
      if (e.touches.length !== 2 || !vt) return;
      e.preventDefault();
      cropOp = null; // due dita: si ingrandisce, il riquadro del ritaglio resta fermo
      const r = pv.getBoundingClientRect(), m = mid(e.touches[0], e.touches[1]);
      pinch = { d0: dist(e.touches[0], e.touches[1]) || 1, z0: view.z, fx: vt.x0 + (m.x - r.left) / vt.s, fy: vt.y0 + (m.y - r.top) / vt.s };
    }, { passive: false });
    pv.addEventListener("touchmove", (e) => {
      if (!pinch || e.touches.length !== 2) return;
      e.preventDefault();
      const r = pv.getBoundingClientRect(), m = mid(e.touches[0], e.touches[1]);
      if (cropMode) frameLock = true;
      const z = clamp(pinch.z0 * dist(e.touches[0], e.touches[1]) / pinch.d0, 1, view.zmax);
      const s = view.s0 * z, ax = m.x - r.left, ay = m.y - r.top;
      view.z = z;
      view.cx = pinch.fx - ax / s + vt.bw / (2 * s);
      view.cy = pinch.fy - ay / s + vt.bh / (2 * s);
      requestDraw();
    }, { passive: false });
    pv.addEventListener("touchend", (e) => { if (e.touches.length < 2) pinch = null; });
    pv.addEventListener("touchcancel", () => { pinch = null; });

    if (window.ResizeObserver) new ResizeObserver(() => { frameLock = false; requestDraw(); }).observe(pv);
  }

  function smoothBins(a) {
    const o = new Float32Array(256);
    for (let i = 0; i < 256; i++) {
      let s = 0, n = 0;
      for (let d = -2; d <= 2; d++) { const j = i + d; if (j >= 0 && j < 256) { const wgt = 3 - Math.abs(d); s += a[j] * wgt; n += wgt; } }
      o[i] = s / n;
    }
    return o;
  }
  function drawHist() {
    const cv = $("#hist"), box = cv.parentElement;
    const dpr = window.devicePixelRatio || 1;
    const W = Math.max(10, box.clientWidth), H = box.clientHeight;
    if (cv.width !== Math.round(W * dpr) || cv.height !== Math.round(H * dpr)) { cv.width = Math.round(W * dpr); cv.height = Math.round(H * dpr); }
    const ctx = cv.getContext("2d");
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.clearRect(0, 0, W, H);
    ctx.fillStyle = "#0c0c0c"; ctx.fillRect(0, 0, W, H);
    ctx.strokeStyle = "#1c1c1c"; ctx.lineWidth = 1;
    for (let i = 1; i < 4; i++) { const x = Math.round((W * i) / 4) + 0.5; ctx.beginPath(); ctx.moveTo(x, 0); ctx.lineTo(x, H); ctx.stroke(); }
    const h = prevHist; if (!h) return;
    const sets = { rgb: [["r", "#ff5252"], ["g", "#5cff6e"], ["b", "#4f7dff"]], r: [["r", "#ff5252"]], g: [["g", "#5cff6e"]], b: [["b", "#4f7dff"]], l: [["l", "#e8e8e8"]] }[ch];
    const data = sets.map(([k, c]) => [smoothBins(h[k]), c]);
    let mx = 1;
    data.forEach(([a]) => { for (let i = 3; i < 253; i++) mx = Math.max(mx, a[i]); });
    mx *= 1.08;
    const path = (a) => {
      ctx.beginPath(); ctx.moveTo(0, H);
      for (let i = 0; i < 256; i++) ctx.lineTo((i / 255) * W, H - Math.min(1, a[i] / mx) * (H - 6));
      ctx.lineTo(W, H); ctx.closePath();
    };
    ctx.globalCompositeOperation = ch === "rgb" ? "screen" : "source-over";
    data.forEach(([a, c]) => { path(a); ctx.globalAlpha = 0.85; ctx.fillStyle = c; ctx.fill(); });
    ctx.globalAlpha = 1; ctx.globalCompositeOperation = "source-over";
    // Forma originale, tratteggiata, per vedere quanto si è spostato l'istogramma
    if (srcHist && !before) {
      const a = smoothBins(srcHist.l);
      let m2 = 1; for (let i = 3; i < 253; i++) m2 = Math.max(m2, a[i]); m2 *= 1.15;
      ctx.setLineDash([2, 3]); ctx.strokeStyle = "rgba(255,255,255,.45)"; ctx.lineWidth = 1;
      ctx.beginPath();
      for (let i = 0; i < 256; i++) { const x = (i / 255) * W, y = H - Math.min(1, a[i] / m2) * (H - 6); i ? ctx.lineTo(x, y) : ctx.moveTo(x, y); }
      ctx.stroke(); ctx.setLineDash([]);
    }
    // Zone escluse dai livelli
    if (!before) {
      ctx.fillStyle = "rgba(0,0,0,.55)";
      ctx.fillRect(0, 0, (P.b / 255) * W, H);
      ctx.fillRect((P.w / 255) * W, 0, W - (P.w / 255) * W, H);
    }
    $("#clipLo").classList.toggle("on", h.clipLo);
    $("#clipHi").classList.toggle("on", h.clipHi);
  }

  /* ---------- Livelli: tre cursori alla base dell'istogramma ---------- */
  function placeLevelLabels() {
    const SPAN = $("#lvLabels").clientWidth, BW = 96, HALF = 7;
    const put = (el, f, txt) => {
      const x = f * SPAN;
      let left, align;
      if (x - BW / 2 < -HALF) { left = -HALF; align = "left"; }
      else if (x + BW / 2 > SPAN + HALF) { left = SPAN + HALF - BW; align = "right"; }
      else { left = x - BW / 2; align = "center"; }
      el.style.left = left + "px"; el.style.textAlign = align;
      el.querySelector("b").textContent = txt;
    };
    const mid = P.b + P.p * (P.w - P.b);
    put($("#lbB"), P.b / 255, Math.round((P.b / 255) * 100) + "%");
    put($("#lbW"), P.w / 255, Math.round((P.w / 255) * 100) + "%");
    put($("#lbM"), mid / 255, Math.round(P.p * 100) + "%");
    $("#lvB").setAttribute("aria-valuetext", Math.round((P.b / 255) * 100) + "%");
    $("#lvW").setAttribute("aria-valuetext", Math.round((P.w / 255) * 100) + "%");
    $("#lvM").setAttribute("aria-valuetext", Math.round(P.p * 100) + "%");
  }
  function bindLevels() {
    const b = $("#lvB"), w = $("#lvW"), m = $("#lvM");
    [b, w, m].forEach((el) => el.addEventListener("pointerdown", () => { [b, w, m].forEach((x) => (x.style.zIndex = x === el ? 3 : 1)); }, true));
    [b, w, m].forEach((el) => el.addEventListener("focus", () => { [b, w, m].forEach((x) => (x.style.zIndex = x === el ? 3 : 1)); }));
    [b, w, m].forEach((el) => el.addEventListener("change", commit));
    b.addEventListener("input", () => { P.b = Math.min(+b.value, P.w - 10); b.value = P.b; m.value = Math.round(P.b + P.p * (P.w - P.b)); preset = "custom"; changed(); });
    w.addEventListener("input", () => { P.w = Math.max(+w.value, P.b + 10); w.value = P.w; m.value = Math.round(P.b + P.p * (P.w - P.b)); preset = "custom"; changed(); });
    m.addEventListener("input", () => {
      const v = Math.min(P.w - 3, Math.max(P.b + 3, +m.value));
      m.value = v; P.p = (v - P.b) / (P.w - P.b); preset = "custom"; changed();
    });
    m.addEventListener("dblclick", () => { P.p = 0.5; syncSliders(); changed(); commit(); });
  }

  /* ---------- Ritaglio ---------- */
  const SOCIAL = [
    ["ig-post", () => "Instagram · post", 1080, 1080],
    ["ig-portrait", () => "Instagram · " + t("f.portrait"), 1080, 1350],
    ["ig-story", () => "Instagram / Facebook · " + t("f.story"), 1080, 1920],
    ["fb-post", () => "Facebook · post", 1200, 630],
    ["fb-cover", () => "Facebook · " + t("f.cover"), 851, 315],
    ["li-post", () => "LinkedIn · post", 1200, 627],
    ["li-cover", () => "LinkedIn · " + t("f.cover"), 1584, 396],
    ["x-post", () => "X · post", 1600, 900],
    ["x-header", () => "X · " + t("f.header"), 1500, 500],
    ["yt-thumb", () => "YouTube · " + t("f.thumb"), 1280, 720],
    ["pin", () => "Pinterest · pin", 1000, 1500],
    ["wa-status", () => "WhatsApp · " + t("f.status"), 1080, 1920],
  ];
  const SLIDES = [
    ["slide-169", () => t("f.slide") + " 16:9", 1920, 1080],
    ["slide-43", () => t("f.slide") + " 4:3", 1440, 1080],
    ["slide-31", () => t("f.slide") + " 3:1", 1920, 640],
  ];
  const PRINT = [ // centimetri a 300 dpi
    ["p-10x15", () => t("f.photo") + " 10 × 15 cm", 10, 15],
    ["p-13x18", () => t("f.photo") + " 13 × 18 cm", 13, 18],
    ["p-20x30", () => t("f.photo") + " 20 × 30 cm", 20, 30],
    ["p-a4", () => "A4 21 × 29,7 cm", 21, 29.7],
  ];
  const RATIOS = [["1:1", 1, 1], ["4:5", 4, 5], ["3:2", 3, 2], ["2:3", 2, 3], ["4:3", 4, 3], ["3:4", 3, 4], ["16:9", 16, 9], ["9:16", 9, 16], ["21:9", 21, 9]];
  const cm2px = (cm, dpi) => Math.max(1, Math.round((cm / 2.54) * dpi));
  function fmtInfo(id) {
    // {ratio: w/h o null, out: {w,h} o null}
    if (id === "free") return { ratio: null, out: null };
    if (id === "orig") return { ratio: base.w / base.h, out: null };
    const r = RATIOS.find((x) => x[0] === id); if (r) return { ratio: r[1] / r[2], out: null };
    const s = SOCIAL.concat(SLIDES).find((x) => x[0] === id); if (s) return { ratio: s[2] / s[3], out: { w: s[2], h: s[3] } };
    const p = PRINT.find((x) => x[0] === id); if (p) { const w = cm2px(p[2], 300), h = cm2px(p[3], 300); return { ratio: w / h, out: { w, h } }; }
    if (id === "custom") {
      const w = +$("#cW").value || 1, h = +$("#cH").value || 1, cm = $("#cU").value === "cm", dpi = clamp(+$("#cDpi").value || 300, 72, 1200);
      const ow = cm ? cm2px(w, dpi) : Math.round(w), oh = cm ? cm2px(h, dpi) : Math.round(h);
      return { ratio: ow / oh, out: { w: Math.max(1, ow), h: Math.max(1, oh) } };
    }
    return { ratio: null, out: null };
  }
  function buildFormats() {
    const sel = $("#cropFmt"); if (!sel) return;
    const cur = sel.value || (draft && draft.id) || "free";
    const opt = (v, l) => `<option value="${v}">${l}</option>`;
    const esc = (x) => String(x).replace(/&/g, "&amp;").replace(/</g, "&lt;");
    const sz = (w, h) => w + " × " + h;
    sel.innerHTML =
      `<optgroup label="${esc(t("crop.g.ratio"))}">` + opt("free", esc(t("crop.free"))) + opt("orig", esc(t("crop.original"))) + RATIOS.map((r) => opt(r[0], r[0])).join("") + "</optgroup>" +
      `<optgroup label="${esc(t("crop.g.social"))}">` + SOCIAL.map((s) => opt(s[0], esc(s[1]()) + " — " + sz(s[2], s[3]))).join("") + "</optgroup>" +
      `<optgroup label="${esc(t("crop.g.slides"))}">` + SLIDES.map((s) => opt(s[0], esc(s[1]()) + " — " + sz(s[2], s[3]))).join("") + "</optgroup>" +
      `<optgroup label="${esc(t("crop.g.print"))}">` + PRINT.map((s) => opt(s[0], esc(s[1]()))).join("") + "</optgroup>" +
      `<optgroup label="${esc(t("crop.g.custom"))}">` + opt("custom", esc(t("crop.custom"))) + "</optgroup>";
    sel.value = cur;
  }

  let cropMode = false, draft = null, cropOp = null; // draft: {id, swap, x, y, w, h} in pixel della foto intera
  function draftRatio() {
    const f = fmtInfo(draft.id);
    if (!f.ratio) return null;
    return draft.swap && draft.id !== "custom" ? 1 / f.ratio : f.ratio;
  }
  function draftOut() {
    const f = fmtInfo(draft.id);
    if (!f.out) return null;
    return draft.swap && draft.id !== "custom" ? { w: f.out.h, h: f.out.w } : f.out;
  }
  function fitDraftToRatio() {
    const r = draftRatio(); if (!r) return;
    // il riquadro più grande con queste proporzioni, centrato sulla zona attuale
    let w = base.w, h = w / r;
    if (h > base.h) { h = base.h; w = h * r; }
    const cx = draft.x + draft.w / 2, cy = draft.y + draft.h / 2;
    draft.w = w; draft.h = h;
    draft.x = clamp(cx - w / 2, 0, base.w - w); draft.y = clamp(cy - h / 2, 0, base.h - h);
  }
  function enterCrop() {
    if (!orig || busy) return;
    ensureBase();
    picking = false;
    const c = P.crop;
    draft = c ? { id: c.id || "free", swap: !!c.swap, x: c.x * base.w, y: c.y * base.h, w: c.w * base.w, h: c.h * base.h }
              : { id: "free", swap: false, x: 0, y: 0, w: base.w, h: base.h };
    if (c && c.id === "custom") { $("#cW").value = c.cw; $("#cH").value = c.ch; $("#cU").value = c.unit || "px"; $("#cDpi").value = c.dpi || 300; }
    else if (!$("#cW").value) { $("#cW").value = 1920; $("#cH").value = 1080; $("#cU").value = "px"; }
    cropMode = true; frameLock = false;
    view.z = 1; // si parte dalla foto intera
    buildFormats(); $("#cropFmt").value = draft.id;
    $("#cropPanel").hidden = false; $("#cropBox").hidden = false;
    $("#cropRemove").hidden = !P.crop;
    syncCustomUI();
    updateTexts(); render(); updateCropInfo();
    // il pannello del ritaglio sta subito sotto l'anteprima: si torna in cima per vederlo
    try { window.scrollTo({ top: 0, behavior: "smooth" }); } catch (e) { window.scrollTo(0, 0); }
  }
  function exitCrop(applyIt) {
    if (!cropMode) return;
    if (applyIt) {
      const d = draft, out = draftOut();
      const whole = d.x < 1 && d.y < 1 && d.w > base.w - 1 && d.h > base.h - 1 && !out;
      P.crop = whole ? null : {
        x: d.x / base.w, y: d.y / base.h, w: d.w / base.w, h: d.h / base.h,
        ow: out ? out.w : null, oh: out ? out.h : null, id: d.id, swap: d.swap,
        cw: d.id === "custom" ? +$("#cW").value : undefined, ch: d.id === "custom" ? +$("#cH").value : undefined,
        unit: d.id === "custom" ? $("#cU").value : undefined, dpi: d.id === "custom" ? +$("#cDpi").value : undefined,
      };
    }
    cropMode = false; draft = null;
    $("#cropPanel").hidden = true; $("#cropBox").hidden = true;
    changed();
    if (applyIt) commit();
  }
  function syncCustomUI() {
    const custom = draft && draft.id === "custom";
    $("#cropCustom").hidden = !custom;
    $("#cDpiWrap").hidden = $("#cU").value !== "cm";
    const r = draft && fmtInfo(draft.id).ratio;
    $("#cropSwap").hidden = !draft || !r || draft.id === "orig" || draft.id === "custom" || Math.abs(r - 1) < 1e-6;
  }
  function updateCropInfo() {
    if (!draft) return;
    const a = Math.round(draft.w) + " × " + Math.round(draft.h);
    const out = draftOut();
    const b = out ? out.w + " × " + out.h : a;
    $("#cropInfo").textContent = t("crop.area", { a }) + " · " + t("crop.output", { b });
    $("#cropWarn").textContent = t("crop.upscale");
    $("#cropWarn").hidden = !(out && (out.w > draft.w * 1.02 || out.h > draft.h * 1.02));
  }
  function positionCrop() {
    const cb = $("#cropBox");
    if (!cropMode || !draft || !vt) { cb.hidden = true; return; }
    cb.hidden = false;
    cb.style.left = (draft.x - vt.x0) * vt.s + "px";
    cb.style.top = (draft.y - vt.y0) * vt.s + "px";
    cb.style.width = draft.w * vt.s + "px";
    cb.style.height = draft.h * vt.s + "px";
    frame = { l: (draft.x - vt.x0) * vt.s, t: (draft.y - vt.y0) * vt.s, w: draft.w * vt.s, h: draft.h * vt.s };
    if (frameLock) updateCropInfo();
  }
  function bindCrop() {
    $("#cropBtn").addEventListener("click", () => (cropMode ? exitCrop(false) : enterCrop()));
    $("#cropCancel").addEventListener("click", () => exitCrop(false));
    $("#cropApply").addEventListener("click", () => exitCrop(true));
    $("#cropRemove").addEventListener("click", () => { cropMode = false; draft = null; $("#cropPanel").hidden = true; $("#cropBox").hidden = true; P.crop = null; changed(); commit(); });
    $("#cropFmt").addEventListener("change", () => {
      draft.id = $("#cropFmt").value; draft.swap = false; frameLock = false;
      if (draft.id === "free") { /* resta la zona attuale */ }
      fitDraftToRatio(); syncCustomUI(); positionCrop(); updateCropInfo();
    });
    ["#cW", "#cH", "#cU", "#cDpi"].forEach((s) => $(s).addEventListener("input", () => {
      if (!draft || draft.id !== "custom") return;
      frameLock = false;
      syncCustomUI(); fitDraftToRatio(); positionCrop(); updateCropInfo();
    }));
    $("#cU").addEventListener("change", () => {
      // converte i valori quando si cambia unità, così la proporzione resta uguale
      const dpi = +$("#cDpi").value || 300, toCm = $("#cU").value === "cm";
      ["#cW", "#cH"].forEach((s) => { const v = +$(s).value || 1; $(s).value = toCm ? +((v / dpi) * 2.54).toFixed(2) : Math.round((v / 2.54) * dpi); });
      syncCustomUI(); updateCropInfo();
    });
    $("#cropSwap").addEventListener("click", () => { draft.swap = !draft.swap; frameLock = false; fitDraftToRatio(); positionCrop(); updateCropInfo(); });

    // trascina il riquadro o i suoi angoli (mouse, penna e dita)
    const cb = $("#cropBox");
    const touches = new Set();
    cb.addEventListener("pointerdown", (e) => {
      if (!cropMode || !vt) return;
      if (e.pointerType === "mouse" && e.button !== 0) return; // tasto destro: sposta la foto ingrandita
      if (e.pointerType === "touch") { touches.add(e.pointerId); if (touches.size > 1) { cropOp = null; return; } }
      e.preventDefault(); e.stopPropagation();
      frameLock = false;
      const hd = e.target.dataset && e.target.dataset.h;
      const d = draft;
      cropOp = { h: hd || null, x: e.clientX, y: e.clientY, d0: { x: d.x, y: d.y, w: d.w, h: d.h } };
      if (hd) cropOp.anchor = { x: hd.includes("w") ? d.x + d.w : d.x, y: hd.includes("n") ? d.y + d.h : d.y };
      cb.setPointerCapture(e.pointerId);
    });
    cb.addEventListener("pointermove", (e) => {
      if (!cropOp) return;
      const d = draft, r = draftRatio(), MIN = 16;
      if (!cropOp.h) {
        d.x = clamp(cropOp.d0.x + (e.clientX - cropOp.x) / vt.s, 0, base.w - d.w);
        d.y = clamp(cropOp.d0.y + (e.clientY - cropOp.y) / vt.s, 0, base.h - d.h);
      } else {
        const pr = $("#preview").getBoundingClientRect();
        const px = clamp(vt.x0 + (e.clientX - pr.left - 1) / vt.s, 0, base.w);
        const py = clamp(vt.y0 + (e.clientY - pr.top - 1) / vt.s, 0, base.h);
        let w = Math.max(MIN, Math.abs(px - cropOp.anchor.x)), h = Math.max(MIN, Math.abs(py - cropOp.anchor.y));
        if (r) { if (w / h > r) w = h * r; else h = w / r; }
        const left = px < cropOp.anchor.x, up = py < cropOp.anchor.y;
        w = Math.min(w, left ? cropOp.anchor.x : base.w - cropOp.anchor.x);
        h = Math.min(h, up ? cropOp.anchor.y : base.h - cropOp.anchor.y);
        if (r) { if (w / h > r) w = h * r; else h = w / r; }
        d.w = w; d.h = h;
        d.x = left ? cropOp.anchor.x - w : cropOp.anchor.x;
        d.y = up ? cropOp.anchor.y - h : cropOp.anchor.y;
      }
      positionCrop(); updateCropInfo();
    });
    const end = (e) => { cropOp = null; if (e && e.pointerType === "touch") touches.delete(e.pointerId); };
    cb.addEventListener("pointerup", end);
    cb.addEventListener("pointercancel", end);
  }

  /* ---------- Rotazione ---------- */
  const fmtDeg = (v) => (Math.round(v * 10) / 10).toLocaleString(LANG, { minimumFractionDigits: 0, maximumFractionDigits: 1 }) + "°";
  // ruota di un quarto di giro in senso orario un rettangolo normalizzato (0…1)
  const turnNorm = (c) => ({ x: 1 - c.y - c.h, y: c.x, w: c.h, h: c.w });
  const swapsRatio = (id) => id && id !== "free" && id !== "orig" && id !== "custom";
  function setRot(q, a, live) {
    if (!orig) return;
    q = ((q % 4) + 4) % 4;
    a = Math.round(clamp(+a || 0, -45, 45) * 10) / 10;
    const dq = (q - rotQ() + 4) % 4;
    if (dq) {
      // il ritaglio gira insieme alla foto
      ensureBase();
      let W = base.w, H = base.h;
      for (let i = 0; i < dq; i++) {
        if (P.crop) {
          const c = P.crop;
          P.crop = Object.assign({}, c, turnNorm(c), { ow: c.oh, oh: c.ow, swap: swapsRatio(c.id) ? !c.swap : c.swap,
            cw: c.id === "custom" ? c.ch : c.cw, ch: c.id === "custom" ? c.cw : c.ch });
        }
        if (cropMode && draft) {
          const d = draft;
          Object.assign(d, { x: H - d.y - d.h, y: d.x, w: d.h, h: d.w });
          if (swapsRatio(d.id)) d.swap = !d.swap;
          if (d.id === "custom") { const w = $("#cW").value; $("#cW").value = $("#cH").value; $("#cH").value = w; }
        }
        const t0 = W; W = H; H = t0;
      }
      frameLock = false;
    }
    P.rq = q; P.rot = a;
    syncRotUI();
    changed();
    if (cropMode) updateCropInfo();
    if (!live) commit();
  }
  function syncRotUI() {
    const a = rotA(), q = rotQ();
    const f = $("#rotFine"), n = $("#rotNum");
    if (!f) return;
    if (document.activeElement !== f) f.value = a;
    if (document.activeElement !== n) n.value = a;
    let tot = q * 90 + a; if (tot > 180) tot -= 360;
    $("#rotInfo").textContent = t("rot.total", { a: fmtDeg(tot) });
  }
  function toggleRot(open) {
    const p = $("#rotPanel");
    p.hidden = open == null ? !p.hidden : !open;
    if (!p.hidden) { syncRotUI(); try { window.scrollTo({ top: 0, behavior: "smooth" }); } catch (e) { window.scrollTo(0, 0); } }
    updateTexts();
  }
  function bindRot() {
    $("#rotBtn").addEventListener("click", () => toggleRot());
    $("#rotDone").addEventListener("click", () => toggleRot(false));
    $("#rotL").addEventListener("click", () => setRot(rotQ() - 1, rotA()));
    $("#rotR").addEventListener("click", () => setRot(rotQ() + 1, rotA()));
    $("#rot180").addEventListener("click", () => setRot(rotQ() + 2, rotA()));
    $("#rotReset").addEventListener("click", () => setRot(0, 0));
    const f = $("#rotFine"), n = $("#rotNum");
    f.addEventListener("input", () => setRot(rotQ(), +f.value, true));
    f.addEventListener("change", commit);
    f.addEventListener("dblclick", () => setRot(rotQ(), 0));
    n.addEventListener("input", () => { if (n.value !== "" && isFinite(+n.value)) setRot(rotQ(), +n.value, true); });
    n.addEventListener("change", () => { setRot(rotQ(), +n.value || 0); n.value = rotA(); });
    n.addEventListener("keydown", (e) => {
      if (e.key !== "ArrowUp" && e.key !== "ArrowDown") return;
      e.preventDefault();
      const st = e.shiftKey ? 0.1 : 0.5;
      setRot(rotQ(), rotA() + (e.key === "ArrowUp" ? st : -st), true);
      n.value = rotA();
    });
    n.addEventListener("keyup", (e) => { if (e.key === "ArrowUp" || e.key === "ArrowDown") commit(); });
    // frecce su e giù: un clic sposta di 0,5° (0,1° con Maiusc), tenendo premuto l'angolo continua a cambiare
    [["#rotUp", 1], ["#rotDown", -1]].forEach(([sel, dir]) => {
      const b = $(sel); let tm = null, iv = null;
      const step = (e) => setRot(rotQ(), rotA() + dir * (e && e.shiftKey ? 0.1 : 0.5), true);
      const stop = () => { if (tm || iv) { clearTimeout(tm); clearInterval(iv); tm = iv = null; commit(); } };
      b.addEventListener("pointerdown", (e) => {
        if (e.button !== 0) return;
        e.preventDefault(); step(e);
        tm = setTimeout(() => { iv = setInterval(() => step(e), 70); }, 380);
      });
      ["pointerup", "pointerleave", "pointercancel"].forEach((ev) => b.addEventListener(ev, stop));
      b.addEventListener("keydown", (e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); step(e); commit(); } });
    });
  }

  /* ---------- Caricamento immagine ---------- */
  async function useImage(img) {
    let w = img.naturalWidth || img.width, h = img.naturalHeight || img.height;
    if (mode === "demo") {
      // foto grandi del telefono: lato lungo al massimo 2560 px su mobile, 4096 px su computer
      const cap = isMobile() ? 2560 : 4096;
      const k = Math.min(1, cap / Math.max(w, h));
      w = Math.max(1, Math.round(w * k)); h = Math.max(1, Math.round(h * k));
    }
    const c = canvas(w, h);
    const x = c.getContext("2d");
    x.imageSmoothingQuality = "high";
    x.drawImage(img, 0, 0, w, h);
    orig = { canvas: c, w, h }; base = null; baseKey = null; workCrop = null;
    if (cropMode) { cropMode = false; draft = null; $("#cropPanel").hidden = true; $("#cropBox").hidden = true; }
    full = null; workKey = null;
  }
  function loadImg(src) {
    return new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = rej; i.src = src; });
  }

  /* ---------- Prova nel browser ---------- */
  const DEMO_SVG = `<svg xmlns="http://www.w3.org/2000/svg" width="1600" height="900" viewBox="0 0 880 495"><defs><linearGradient id="s" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#5f9fd8"/><stop offset="1" stop-color="#d8ecf7"/></linearGradient><linearGradient id="m" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#3a86b4"/><stop offset="1" stop-color="#123f5e"/></linearGradient><linearGradient id="a" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#f0dcb2"/><stop offset="1" stop-color="#c9a878"/></linearGradient></defs><rect width="880" height="300" fill="url(#s)"/><circle cx="700" cy="104" r="40" fill="#fff4cf"/><ellipse cx="220" cy="80" rx="90" ry="18" fill="#fff" opacity=".75"/><ellipse cx="300" cy="96" rx="70" ry="14" fill="#fff" opacity=".6"/><ellipse cx="560" cy="60" rx="60" ry="12" fill="#fff" opacity=".55"/><path d="M0 262C120 220 220 240 330 252C450 262 560 236 680 250C760 258 830 252 880 256V300H0Z" fill="#7d9c8c"/><rect y="296" width="880" height="199" fill="url(#m)"/><path d="M0 170C60 150 120 158 170 190C220 220 250 262 300 300V495H0Z" fill="#2c5a36"/><path d="M0 230C70 214 140 230 190 262C230 290 250 330 262 380V495H0Z" fill="#1d4027"/><circle cx="60" cy="176" r="22" fill="#23492d"/><circle cx="100" cy="170" r="18" fill="#2d5c38"/><circle cx="138" cy="184" r="20" fill="#23492d"/><path d="M500 495C560 430 680 400 880 392V495Z" fill="url(#a)"/><rect x="340" y="336" width="300" height="10" fill="#5a3d28"/><g fill="#3f2a1b"><rect x="352" y="346" width="6" height="40"/><rect x="420" y="346" width="6" height="40"/><rect x="490" y="346" width="6" height="40"/><rect x="560" y="346" width="6" height="40"/><rect x="628" y="346" width="6" height="40"/></g><path d="M380 420h60M520 450h80M300 470h50M620 318h70M420 312h40" stroke="#bfe3f5" stroke-width="2" opacity=".55"/></svg>`;
  function freshStart(params, isEdited) {
    P = Object.assign(D(), params || {});
    baseline = Object.assign(D(), params || {});
    preset = isEdited ? "custom" : "nat";
    before = false; picking = false; restoreArmed = false;
    histReset();
    syncSliders(); show("editor"); updateTexts(); render();
  }
  async function startDemo() {
    mode = "demo";
    shape = null;
    show("loading");
    try {
      const img = await loadImg("data:image/svg+xml;charset=utf-8," + encodeURIComponent(DEMO_SVG));
      await useImage(img);
    } catch (e) {
      console.error(e);
      show("empty", "empty.title", "empty.text");
      return;
    }
    if (mode !== "demo") return; // nel frattempo si è collegato PowerPoint
    $("#loadRow").hidden = false;
    $("#preview").classList.add("droppable");
    freshStart(null, false);
  }
  async function loadOwnPhoto(file) {
    if (!file) return;
    const url = URL.createObjectURL(file);
    try {
      const img = await loadImg(url);
      await useImage(img);
      freshStart(null, false);
    } catch (e) { toast(t("toast.loadError")); }
    finally { URL.revokeObjectURL(url); }
  }

  /* ---------- Trascina e rilascia una foto (solo versione web) ---------- */
  function bindDrop() {
    const drop = $("#drop");
    let depth = 0;
    const hasFile = (e) => e.dataTransfer && Array.from(e.dataTransfer.types || []).includes("Files");
    const hide = () => { depth = 0; drop.hidden = true; };
    window.addEventListener("dragenter", (e) => {
      if (mode !== "demo" || !hasFile(e)) return;
      e.preventDefault(); depth++; drop.hidden = false;
    });
    window.addEventListener("dragover", (e) => {
      if (mode !== "demo" || !hasFile(e)) return;
      e.preventDefault(); e.dataTransfer.dropEffect = "copy";
    });
    window.addEventListener("dragleave", (e) => {
      if (mode !== "demo" || !hasFile(e)) return;
      depth = Math.max(0, depth - 1); if (!depth) drop.hidden = true;
    });
    window.addEventListener("drop", (e) => {
      if (mode !== "demo") return;
      e.preventDefault(); hide();
      const f = Array.from((e.dataTransfer && e.dataTransfer.files) || []).find((x) => /^image\//.test(x.type));
      if (f) loadOwnPhoto(f); else toast(t("toast.loadError"));
    });
    window.addEventListener("blur", hide);
  }

  /* ---------- PowerPoint: lettura della selezione ---------- */
  let selT = 0, loadSeq = 0, lastKey = "";
  function onSelectionChanged() {
    if (busy) return;
    clearTimeout(selT);
    selT = setTimeout(() => loadFromSlide(false), 250);
  }

  function readTags(tags) {
    const o = {};
    tags.items.forEach((x) => { o[x.key.toUpperCase()] = x.value; });
    return o;
  }

  async function loadFromSlide(force) {
    const seq = ++loadSeq;
    try {
      const info = await PowerPoint.run(async (ctx) => {
        const sel = ctx.presentation.getSelectedShapes();
        sel.load("items/id,items/type");
        await ctx.sync();
        if (sel.items.length === 0) return { none: true };
        if (sel.items.length > 1) return { multi: true };
        const sh = sel.items[0];
        if (sh.type !== "Image") {
          // segnaposto che contiene una foto
          let isPic = false;
          if (sh.type === "Placeholder") {
            try { sh.placeholderFormat.load("containedType"); await ctx.sync(); isPic = sh.placeholderFormat.containedType === "Image"; } catch (e) { isPic = false; }
          }
          if (!isPic) return { notImage: true };
        }
        const slide = sh.getParentSlide();
        slide.load("id");
        sh.tags.load("items/key,items/value");
        sh.load("id,name,width");
        await ctx.sync();
        const tags = readTags(sh.tags);
        return { id: sh.id, slideId: slide.id, tags };
      });
      if (seq !== loadSeq) return;
      if (info.none || info.multi || info.notImage) {
        shape = null; lastKey = ""; orig = base = full = prev = null;
        if (cropMode) { cropMode = false; draft = null; $("#cropPanel").hidden = true; }
        setStatus("status.none", "off");
        show("empty", "empty.title", info.multi ? "empty.multi" : info.notImage ? "empty.notImage" : "empty.text");
        return;
      }
      const key = info.slideId + "/" + info.id;
      if (!force && key === lastKey && orig) return; // stessa immagine: niente da ricaricare
      lastKey = key;
      setStatus("status.loading", "busy");
      show("loading");

      const edited = !!info.tags.ULAB_ORIG;
      const res = await PowerPoint.run(async (ctx) => {
        const slide = ctx.presentation.slides.getItem(info.slideId);
        let src = slide.shapes.getItem(info.id);
        let origId = null;
        if (edited) {
          const o = slide.shapes.getItemOrNullObject(info.tags.ULAB_ORIG);
          o.load("id");
          await ctx.sync();
          if (!o.isNullObject) { src = o; origId = o.id; }
        }
        src.load("id,width,height,rotation,visible");
        await ctx.sync();
        const rot = src.rotation, vis = src.visible;
        if (rot) src.rotation = 0;           // la foto si legge dritta: la rotazione viene rimessa sulla copia
        if (!vis) src.visible = true;        // l'originale nascosto si rende visibile solo per la lettura
        // circa 300 dpi, con il lato lungo al massimo di 3200 px
        const longSide = Math.max(src.width, src.height) || 1;
        const scale = Math.min(4, 3200 / longSide);
        const px = Math.max(200, Math.round(src.width * scale));
        const img = src.getImageAsBase64({ format: "Png", width: px });
        await ctx.sync();
        if (rot) src.rotation = rot;
        if (!vis) src.visible = false;
        await ctx.sync();
        return { b64: img.value, origId };
      });
      if (seq !== loadSeq) return;
      const im = await loadImg("data:image/png;base64," + res.b64);
      await useImage(im);
      shape = { id: info.id, slideId: info.slideId, origId: res.origId, edited: edited && !!res.origId };
      let saved = null;
      if (shape.edited) { try { saved = JSON.parse(info.tags.ULAB_PARAMS || "null"); } catch (e) { saved = null; } }
      freshStart(saved, shape.edited);
    } catch (e) {
      if (seq !== loadSeq) return;
      console.error(e);
      lastKey = "";
      setStatus("status.none", "off");
      show("empty", "empty.title", "empty.text");
      toast(t("toast.loadError"));
    }
  }

  /* ---------- Risultato finale (regolazioni, ritaglio e misura d'uscita) ---------- */
  let shareFile = null; // foto pronta da condividere se il primo tentativo è stato bloccato dal browser
  function finalCanvas() {
    ensureWorking();
    const out = new ImageData(full.w, full.h);
    IL.process(full.data, full.w, full.h, P, out.data, cacheFull);
    let c = canvas(full.w, full.h);
    c.getContext("2d").putImageData(out, 0, 0);
    const alpha = IL.histogram(out.data).alpha;
    const ow = P.crop && P.crop.ow, oh = P.crop && P.crop.oh;
    if (ow && oh && (ow !== c.width || oh !== c.height)) {
      // riduzione a passi per una misura finale nitida
      while (c.width / 2 >= ow && c.height / 2 >= oh) {
        const h2 = canvas(Math.round(c.width / 2), Math.round(c.height / 2)), x2 = h2.getContext("2d");
        x2.imageSmoothingQuality = "high"; x2.drawImage(c, 0, 0, h2.width, h2.height); c = h2;
      }
      const f = canvas(ow, oh), xf = f.getContext("2d");
      xf.imageSmoothingQuality = "high"; xf.drawImage(c, 0, 0, ow, oh); c = f;
    }
    return { c, alpha };
  }
  function renderFullBlob() {
    const { c, alpha } = finalCanvas();
    const type = alpha ? "image/png" : "image/jpeg";
    return new Promise((res, rej) => c.toBlob((b) => (b ? res({ blob: b, ext: alpha ? "png" : "jpg", type }) : rej(new Error("toBlob"))), type, 0.92));
  }
  function renderFull() {
    const { c, alpha } = finalCanvas();
    return { url: alpha ? c.toDataURL("image/png") : c.toDataURL("image/jpeg", 0.92), w: c.width, h: c.height };
  }
  async function shareOrSave(file) {
    // Telefono: menu di condivisione del sistema → "Salva immagine" (Foto), WhatsApp, Messaggi, Mail…
    if (isMobile() && navigator.canShare && navigator.canShare({ files: [file] })) {
      try {
        await navigator.share({ files: [file], title: "Image Lab" });
        shareFile = null; toast(t("toast.shared"));
      } catch (e) {
        if (e && e.name === "AbortError") { shareFile = null; }            // chiuso dall'utente
        else { shareFile = file; toast(t("shareReady")); }                 // il browser chiede un nuovo tocco
      }
      updateTexts();
      return;
    }
    const url = URL.createObjectURL(file);
    const a = document.createElement("a");
    a.href = url; a.download = file.name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
    toast(t("toast.downloaded"));
  }

  function insertImage(b64, box) {
    return new Promise((res, rej) => Office.context.document.setSelectedDataAsync(
      b64,
      { coercionType: Office.CoercionType.Image, imageLeft: box.left, imageTop: box.top, imageWidth: box.width, imageHeight: box.height },
      (r) => (r.status === Office.AsyncResultStatus.Succeeded ? res() : rej(r.error))
    ));
  }

  async function apply() {
    if (busy || !orig || cropMode) return;
    if (mode === "demo") {
      if (shareFile) { const f = shareFile; shareFile = null; await shareOrSave(f); return; }
      busy = true; updateTexts();
      await wait(30);
      try {
        const r = await renderFullBlob();
        const file = new File([r.blob], "image-lab-" + new Date().toISOString().slice(0, 10) + "." + r.ext, { type: r.type });
        busy = false; updateTexts();
        await shareOrSave(file);
      } catch (e) {
        console.error(e); toast(t("toast.error"));
      } finally { busy = false; updateTexts(); }
      return;
    }
    busy = true; updateTexts(); setStatus("status.loading", "busy");
    await wait(30); // lascia aggiornare il pulsante prima del calcolo
    try {
      const fin = renderFull();
      const b64 = fin.url.split(",")[1];
      const cur = shape;
      const origId = cur.edited ? cur.origId : cur.id;

      // 1) misure dell'originale (riferimento fisso anche dopo più modifiche), ordine della foto attuale, forme presenti
      const box = await PowerPoint.run(async (ctx) => {
        const slide = ctx.presentation.slides.getItem(cur.slideId);
        const s = slide.shapes.getItem(cur.id);
        const o = slide.shapes.getItem(origId);
        s.load("rotation,zOrderPosition");
        o.load("left,top,width,height,rotation,name");
        slide.shapes.load("items/id");
        await ctx.sync();
        try { ctx.presentation.setSelectedShapes([]); await ctx.sync(); } catch (e) { /* alcune versioni non accettano la selezione vuota */ }
        // con il ritaglio la foto mantiene le proporzioni nuove dentro lo spazio dell'originale, centrata
        let left = o.left, top = o.top, width = o.width, height = o.height;
        if (rotQ() & 1) { // un quarto di giro: lo spazio si gira attorno al centro
          left = o.left + (o.width - o.height) / 2; top = o.top + (o.height - o.width) / 2; width = o.height; height = o.width;
        }
        if (P.crop || rotQ()) {
          const bx = { left, top, width, height };
          const a = fin.w / fin.h;
          if (a > bx.width / bx.height) { height = bx.width / a; top = bx.top + (bx.height - height) / 2; }
          else { width = bx.height * a; left = bx.left + (bx.width - width) / 2; }
        }
        return { left, top, width, height, rotation: s.rotation, z: s.zOrderPosition, ids: slide.shapes.items.map((x) => x.id) };
      });

      // 2) inserimento della foto regolata
      await insertImage(b64, box);

      // 3) sistemazione: rotazione, nome, etichette, originale nascosto, ordine di livello
      const newId = await PowerPoint.run(async (ctx) => {
       try {
        const slide = ctx.presentation.slides.getItem(cur.slideId);
        slide.shapes.load("items/id,items/type");
        await ctx.sync();
        const fresh = slide.shapes.items.filter((x) => !box.ids.includes(x.id) && x.type === "Image");
        if (!fresh.length) throw new Error("Nuova immagine non trovata");
        const ns = fresh[fresh.length - 1];
        const orig = slide.shapes.getItemOrNullObject(origId);
        orig.load("id,name");
        await ctx.sync();
        if (orig.isNullObject) throw new Error("Originale non trovato");
        ns.rotation = box.rotation || 0;
        ns.name = (orig.name || "Immagine") + " · Image Lab";
        ns.tags.add("ULAB_ORIG", orig.id);
        ns.tags.add("ULAB_PARAMS", JSON.stringify(P));
        if (cur.edited) slide.shapes.getItem(cur.id).delete();       // la versione precedente non serve più
        else { orig.visible = false; orig.tags.add("ULAB_ROLE", "original"); }
        await ctx.sync();
        // porta la nuova foto al livello di quella di partenza (sopra l'originale nascosto)
        const target = cur.edited ? box.z : box.z + 1;
        for (let i = 0; i < 80; i++) {
          ns.load("zOrderPosition");
          await ctx.sync();
          if (ns.zOrderPosition <= target) break;
          ns.setZOrder("SendBackward");
        }
        await ctx.sync();
        try { ctx.presentation.setSelectedShapes([ns.id]); await ctx.sync(); } catch (e) { /* ignora */ }
        return ns.id;
       } catch (err) {
        // pulizia: toglie eventuali copie inserite e lascia la slide com'era
        try {
          const sl = ctx.presentation.slides.getItem(cur.slideId);
          sl.shapes.load("items/id");
          const o = sl.shapes.getItemOrNullObject(origId);
          o.load("id");
          await ctx.sync();
          // la copia nuova si toglie solo se l'originale c'è ancora: così la foto non sparisce mai dalla slide
          if (!o.isNullObject) {
            sl.shapes.items.filter((x) => !box.ids.includes(x.id)).forEach((x) => x.delete());
            if (!cur.edited) o.visible = true;
            await ctx.sync();
          }
        } catch (e2) { /* ignora */ }
        throw err;
       }
      });

      shape = { id: newId, slideId: cur.slideId, origId, edited: true };
      lastKey = cur.slideId + "/" + newId;
      baseline = Object.assign(D(), JSON.parse(snap(P)));
      toast(t("toast.applied"));
    } catch (e) {
      console.error(e);
      toast(t("toast.error"));
    } finally {
      busy = false;
      updateTexts();
    }
  }

  /* ---------- Torna all'originale ---------- */
  let armT = 0;
  async function restore() {
    if (busy || !shape || !shape.edited) return;
    if (!restoreArmed) { restoreArmed = true; updateTexts(); clearTimeout(armT); armT = setTimeout(() => { restoreArmed = false; updateTexts(); }, 3500); return; }
    restoreArmed = false; clearTimeout(armT);
    busy = true; updateTexts();
    try {
      const cur = shape;
      await PowerPoint.run(async (ctx) => {
        const slide = ctx.presentation.slides.getItem(cur.slideId);
        const orig = slide.shapes.getItem(cur.origId);
        orig.visible = true;
        const role = orig.tags.getItemOrNullObject("ULAB_ROLE");
        await ctx.sync();
        if (!role.isNullObject) orig.tags.delete("ULAB_ROLE");
        slide.shapes.getItem(cur.id).delete();
        await ctx.sync();
        try { ctx.presentation.setSelectedShapes([cur.origId]); await ctx.sync(); } catch (e) { /* ignora */ }
      });
      toast(t("toast.restored"));
    } catch (e) {
      console.error(e);
      toast(t("toast.error"));
    } finally {
      busy = false;
      await loadFromSlide(true);
    }
  }

  /* ---------- Punto neutro ---------- */
  function pickAt(ev) {
    if (!picking || !full || !vt || cropMode) return;
    const r = $("#pv").getBoundingClientRect();
    const fx = vt.x0 + (ev.clientX - r.left) / vt.s, fy = vt.y0 + (ev.clientY - r.top) / vt.s;
    if (fx < 0 || fy < 0 || fx >= full.w || fy >= full.h) return; // clic nelle bande vuote
    const x = Math.floor(fx), y = Math.floor(fy);
    const rad = clamp(Math.round(2 / vt.s), 2, 12); // circa 5 pixel dello schermo
    let R = 0, G = 0, B = 0, n = 0;
    for (let dy = -rad; dy <= rad; dy++) for (let dx = -rad; dx <= rad; dx++) {
      const xx = clamp(x + dx, 0, full.w - 1), yy = clamp(y + dy, 0, full.h - 1);
      const i = (yy * full.w + xx) * 4;
      R += full.data[i]; G += full.data[i + 1]; B += full.data[i + 2]; n++;
    }
    const nt = IL.neutralFrom(R / n, G / n, B / n);
    P.temp = nt.temp; P.tint = nt.tint; preset = "custom";
    picking = false; syncSliders(); changed(); commit();
    toast(t("toast.picked"));
  }

  /* ---------- Avvio ---------- */
  function bindUI() {
    const sel = $("#lang");
    sel.innerHTML = I18N.langs.map(([k, n]) => `<option value="${k}">${n}</option>`).join("");
    sel.value = LANG;
    sel.addEventListener("change", () => { LANG = sel.value; store.set("imagelab.lang", LANG); applyStatic(); if (mode === "unsupported") showUnsupported(); else if (!shape && mode === "office") { setStatus("status.none", "off"); if ($("#loading").hidden) show("empty", "empty.title", "empty.text"); } render(); });
    $$(".pill").forEach((b) => b.addEventListener("click", () => { ch = b.dataset.ch; updateTexts(); drawHist(); }));
    bindLevels();
    bindZoom();
    bindCrop();
    bindRot();
    $("#qBtn").addEventListener("click", () => {
      if (!full) return;
      if (P.q > 0) { P.q = 0; }
      else { P.qn = IL.estimateNoise(full.data, full.w, full.h); P.q = 60; toast(t("toast.quality")); }
      syncSliders(); changed(); commit();
    });
    $("#qAmt").addEventListener("input", () => { P.q = +$("#qAmt").value; $("#qOut").textContent = P.q + "%"; changed(); });
    $("#qAmt").addEventListener("change", commit);
    $("#qAmt").addEventListener("dblclick", () => { P.q = 60; syncSliders(); changed(); commit(); });
    $("#autoLv").addEventListener("click", () => { if (!srcHist) return; Object.assign(P, IL.autoLevels(srcHist)); preset = "custom"; syncSliders(); changed(); commit(); });
    $("#resetLv").addEventListener("click", () => { P.b = 0; P.w = 255; P.p = 0.5; syncSliders(); changed(); commit(); });
    $("#pickBtn").addEventListener("click", () => { if (cropMode) return; picking = !picking; updateTexts(); });
    $("#pv").addEventListener("click", pickAt);
    $("#beforeBtn").addEventListener("click", () => { before = !before; updateTexts(); render(); });
    // Azzera: tutto come all'inizio, ritaglio compreso
    $("#resetAll").addEventListener("click", () => {
      if (cropMode) { cropMode = false; draft = null; $("#cropPanel").hidden = true; $("#cropBox").hidden = true; }
      P = D(); preset = "nat"; syncSliders(); changed(); commit();
      frameLock = false;
    });
    $("#restoreBtn").addEventListener("click", restore);
    $("#applyBtn").addEventListener("click", apply);
    $("#undoBtn").addEventListener("click", () => goHistory(-1));
    $("#redoBtn").addEventListener("click", () => goHistory(1));
    document.addEventListener("keydown", (e) => {
      if (!(e.metaKey || e.ctrlKey) || e.altKey) return;
      const tg = e.target, tag = tg && tg.tagName;
      if (tag === "SELECT" || (tag === "INPUT" && tg.type !== "range") || tag === "TEXTAREA") return;
      const k = e.key.toLowerCase();
      if (k === "z" && !e.shiftKey) { e.preventDefault(); goHistory(-1); }
      else if ((k === "z" && e.shiftKey) || k === "y") { e.preventDefault(); goHistory(1); }
    });
    $("#loadBtn").addEventListener("click", () => $("#fileIn").click());
    bindDrop();
    $("#fileIn").addEventListener("change", (e) => { loadOwnPhoto(e.target.files[0]); e.target.value = ""; });
    // In PowerPoint i link si aprono nel browser di sistema, non dentro il pannello
    $$(".links a, a.ext").forEach((el) => el.addEventListener("click", (e) => {
      if (!inOffice) return;
      try {
        if (Office.context.requirements.isSetSupported("OpenBrowserWindowApi", "1.1")) { e.preventDefault(); Office.context.ui.openBrowserWindow(el.href); }
      } catch (err) { /* il link si apre normalmente */ }
    }));
    let rz = 0;
    window.addEventListener("resize", () => { cancelAnimationFrame(rz); rz = requestAnimationFrame(() => { drawHist(); placeLevelLabels(); drawView(); }); });
  }

  function showUnsupported() {
    setStatus("status.none", "off");
    show("empty", "unsupported.title", "unsupported.text");
  }

  let started = false;
  function start() {
    if (started) return; started = true;
    LANG = detectLang();
    bindUI();
    applyStatic();
    if (!inOffice) { document.body.classList.add("web"); startDemo(); return; }
    startOffice();
  }
  function startOffice() {
    let ok = false;
    try { ok = Office.context.requirements.isSetSupported("PowerPointApi", "1.10"); } catch (e) { ok = false; }
    if (!ok) { mode = "unsupported"; showUnsupported(); return; }
    mode = "office";
    setStatus("status.none", "off");
    show("empty", "empty.title", "empty.text");
    try { Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChanged); } catch (e) { console.error(e); }
    loadFromSlide(true);
  }

  if (window.IL_DEMO || !window.Office || !Office.onReady) {
    document.readyState === "loading" ? document.addEventListener("DOMContentLoaded", start) : start();
  } else {
    Office.onReady((info) => {
      inOffice = !!(info && info.host === Office.HostType.PowerPoint);
      if (!started) { start(); return; }
      // PowerPoint ha risposto dopo il tempo di attesa (computer lento): lascia la prova e collega la presentazione
      if (inOffice && mode === "demo") {
        document.body.classList.remove("web");
        $("#loadRow").hidden = true;
        $("#preview").classList.remove("droppable");
        P = D(); baseline = D(); preset = "nat"; orig = base = full = prev = null;
        LANG = detectLang(); $("#lang").value = LANG; applyStatic();
        startOffice();
      }
    });
    setTimeout(start, 2500); // se Office.js non risponde (pagina aperta nel browser) parte la prova
  }
})();
