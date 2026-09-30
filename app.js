/* The Ultraspeaker Image Lab – add-in PowerPoint per regolare le immagini (istogramma, livelli, colore) */
(function () {
  "use strict";

  /* ---------- Utilità ---------- */
  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => Array.from(r.querySelectorAll(s));
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
  // Telefono o tablet: schermo touch con menu di condivisione del sistema (Foto, WhatsApp, Messaggi…)
  const isMobile = () => {
    try { return window.matchMedia("(pointer: coarse)").matches && typeof navigator.share === "function"; } catch (e) { return false; }
  };

  /* ---------- Lingua ---------- */
  const I18N = window.IL_I18N;
  const LANGS = I18N.langs.map((l) => l[0]);
  let LANG = "it";
  const t = (k) => (I18N[LANG] && I18N[LANG][k]) || I18N.it[k] || k;
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
    updateTexts();
  }

  /* ---------- Stato ---------- */
  let inOffice = false;
  let mode = "boot";           // demo | office | unsupported
  let P = IL.defaults();       // regolazioni correnti
  let baseline = IL.defaults();// regolazioni già applicate (per sapere se ci sono modifiche)
  let preset = "nat";
  let ch = "rgb", before = false, picking = false, busy = false, restoreArmed = false;
  let full = null, prev = null;          // {data, w, h}: immagine intera e anteprima ridotta
  let prevOut = null, prevHist = null, srcHist = null;
  const cachePrev = {}, cacheFull = {};
  let shape = null;                      // {id, slideId, origId, edited}

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
        inp.addEventListener("dblclick", () => { P[k] = 0; inp.value = 0; preset = "custom"; row.querySelector("output").textContent = fmt(k, 0); changed(); });
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
      b.addEventListener("click", () => { P = Object.assign(IL.defaults(), PRESETS[k]()); preset = k; syncSliders(); changed(); });
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
    const dirty = !sameParams(P, baseline);
    apply.disabled = busy || (mode === "office" && !dirty);
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
    if (mode === "demo") setStatus("status.demo", "");
    else if (shape) setStatus(shape.edited ? "status.edited" : "status.selected", "");
  }
  function sameParams(a, b) { for (const k in IL.defaults()) if (Math.abs(a[k] - b[k]) > 1e-6) return false; return true; }

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
    if (!prev) return;
    IL.process(prev.data, prev.w, prev.h, P, prevOut.data, cachePrev);
    const cv = $("#pv"), ctx = cv.getContext("2d");
    ctx.putImageData(before ? new ImageData(prev.data, prev.w, prev.h) : prevOut, 0, 0);
    prevHist = IL.histogram(before ? prev.data : prevOut.data);
    drawHist();
    placeLevelLabels();
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
    b.addEventListener("input", () => { P.b = Math.min(+b.value, P.w - 10); b.value = P.b; m.value = Math.round(P.b + P.p * (P.w - P.b)); preset = "custom"; changed(); });
    w.addEventListener("input", () => { P.w = Math.max(+w.value, P.b + 10); w.value = P.w; m.value = Math.round(P.b + P.p * (P.w - P.b)); preset = "custom"; changed(); });
    m.addEventListener("input", () => {
      const v = Math.min(P.w - 3, Math.max(P.b + 3, +m.value));
      m.value = v; P.p = (v - P.b) / (P.w - P.b); preset = "custom"; changed();
    });
    m.addEventListener("dblclick", () => { P.p = 0.5; syncSliders(); changed(); });
  }

  /* ---------- Caricamento immagine ---------- */
  const PREVIEW_MAX = 760;
  function toCanvas(img, w, h) {
    const c = document.createElement("canvas"); c.width = w; c.height = h;
    const x = c.getContext("2d", { willReadFrequently: true });
    x.imageSmoothingQuality = "high";
    x.drawImage(img, 0, 0, w, h);
    return x.getImageData(0, 0, w, h);
  }
  async function useImage(img) {
    let w = img.naturalWidth || img.width, h = img.naturalHeight || img.height;
    if (mode === "demo") {
      // foto grandi del telefono: lato lungo al massimo 2560 px su mobile, 4096 px su computer
      const cap = isMobile() ? 2560 : 4096;
      const k = Math.min(1, cap / Math.max(w, h));
      w = Math.max(1, Math.round(w * k)); h = Math.max(1, Math.round(h * k));
    }
    // il riquadro segue la forma della foto fino a 21:9; oltre, la foto resta intera con bande sopra e sotto
    $("#preview").style.setProperty("--ar", String(Math.min(w / h, 21 / 9)));
    const fd = toCanvas(img, w, h);
    full = { data: fd.data, w, h };
    const s = Math.min(1, PREVIEW_MAX / Math.max(w, h));
    const pw = Math.max(1, Math.round(w * s)), ph = Math.max(1, Math.round(h * s));
    const pd = toCanvas(img, pw, ph);
    prev = { data: pd.data, w: pw, h: ph };
    prevOut = new ImageData(pw, ph);
    delete cachePrev.blur; delete cacheFull.blur;
    const cv = $("#pv"); cv.width = pw; cv.height = ph;
    srcHist = IL.histogram(prev.data);
  }
  function loadImg(src) {
    return new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = rej; i.src = src; });
  }

  /* ---------- Prova nel browser ---------- */
  const DEMO_SVG = `<svg xmlns="http://www.w3.org/2000/svg" width="1600" height="900" viewBox="0 0 880 495"><defs><linearGradient id="s" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#5f9fd8"/><stop offset="1" stop-color="#d8ecf7"/></linearGradient><linearGradient id="m" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#3a86b4"/><stop offset="1" stop-color="#123f5e"/></linearGradient><linearGradient id="a" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#f0dcb2"/><stop offset="1" stop-color="#c9a878"/></linearGradient></defs><rect width="880" height="300" fill="url(#s)"/><circle cx="700" cy="104" r="40" fill="#fff4cf"/><ellipse cx="220" cy="80" rx="90" ry="18" fill="#fff" opacity=".75"/><ellipse cx="300" cy="96" rx="70" ry="14" fill="#fff" opacity=".6"/><ellipse cx="560" cy="60" rx="60" ry="12" fill="#fff" opacity=".55"/><path d="M0 262C120 220 220 240 330 252C450 262 560 236 680 250C760 258 830 252 880 256V300H0Z" fill="#7d9c8c"/><rect y="296" width="880" height="199" fill="url(#m)"/><path d="M0 170C60 150 120 158 170 190C220 220 250 262 300 300V495H0Z" fill="#2c5a36"/><path d="M0 230C70 214 140 230 190 262C230 290 250 330 262 380V495H0Z" fill="#1d4027"/><circle cx="60" cy="176" r="22" fill="#23492d"/><circle cx="100" cy="170" r="18" fill="#2d5c38"/><circle cx="138" cy="184" r="20" fill="#23492d"/><path d="M500 495C560 430 680 400 880 392V495Z" fill="url(#a)"/><rect x="340" y="336" width="300" height="10" fill="#5a3d28"/><g fill="#3f2a1b"><rect x="352" y="346" width="6" height="40"/><rect x="420" y="346" width="6" height="40"/><rect x="490" y="346" width="6" height="40"/><rect x="560" y="346" width="6" height="40"/><rect x="628" y="346" width="6" height="40"/></g><path d="M380 420h60M520 450h80M300 470h50M620 318h70M420 312h40" stroke="#bfe3f5" stroke-width="2" opacity=".55"/></svg>`;
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
    P = IL.defaults(); baseline = IL.defaults(); preset = "nat";
    $("#loadRow").hidden = false;
    $("#preview").classList.add("droppable");
    syncSliders(); show("editor"); updateTexts(); render();
  }
  async function loadOwnPhoto(file) {
    if (!file) return;
    const url = URL.createObjectURL(file);
    try {
      const img = await loadImg(url);
      await useImage(img);
      P = IL.defaults(); baseline = IL.defaults(); preset = "nat"; before = false; picking = false;
      syncSliders(); updateTexts(); render();
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
        shape = null; lastKey = ""; full = prev = null;
        setStatus("status.none", "off");
        show("empty", "empty.title", info.multi ? "empty.multi" : info.notImage ? "empty.notImage" : "empty.text");
        return;
      }
      const key = info.slideId + "/" + info.id;
      if (!force && key === lastKey && prev) return; // stessa immagine: niente da ricaricare
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
      P = Object.assign(IL.defaults(), saved || {});
      baseline = Object.assign({}, P);
      preset = shape.edited ? "custom" : "nat";
      before = false; picking = false; restoreArmed = false;
      syncSliders(); show("editor"); updateTexts(); render();
    } catch (e) {
      if (seq !== loadSeq) return;
      console.error(e);
      lastKey = "";
      setStatus("status.none", "off");
      show("empty", "empty.title", "empty.text");
      toast(t("toast.loadError"));
    }
  }

  /* ---------- Applica ---------- */
  let shareFile = null; // foto pronta da condividere se il primo tentativo è stato bloccato dal browser
  function renderFullBlob() {
    const out = new ImageData(full.w, full.h);
    IL.process(full.data, full.w, full.h, P, out.data, cacheFull);
    const c = document.createElement("canvas"); c.width = full.w; c.height = full.h;
    c.getContext("2d").putImageData(out, 0, 0);
    const png = IL.histogram(out.data).alpha;
    const type = png ? "image/png" : "image/jpeg";
    return new Promise((res, rej) => c.toBlob((b) => (b ? res({ blob: b, ext: png ? "png" : "jpg", type }) : rej(new Error("toBlob"))), type, 0.92));
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
  function renderFull() {
    const out = new ImageData(full.w, full.h);
    IL.process(full.data, full.w, full.h, P, out.data, cacheFull);
    const c = document.createElement("canvas"); c.width = full.w; c.height = full.h;
    c.getContext("2d").putImageData(out, 0, 0);
    const alpha = IL.histogram(out.data).alpha;
    return alpha ? c.toDataURL("image/png") : c.toDataURL("image/jpeg", 0.92);
  }

  function insertImage(b64, box) {
    return new Promise((res, rej) => Office.context.document.setSelectedDataAsync(
      b64,
      { coercionType: Office.CoercionType.Image, imageLeft: box.left, imageTop: box.top, imageWidth: box.width, imageHeight: box.height },
      (r) => (r.status === Office.AsyncResultStatus.Succeeded ? res() : rej(r.error))
    ));
  }

  async function apply() {
    if (busy || !full) return;
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
      const dataUrl = renderFull();
      const b64 = dataUrl.split(",")[1];
      const cur = shape;
      const origId = cur.edited ? cur.origId : cur.id;

      // 1) posizione e ordine della foto attuale; elenco delle forme per riconoscere quella nuova
      const box = await PowerPoint.run(async (ctx) => {
        const slide = ctx.presentation.slides.getItem(cur.slideId);
        const s = slide.shapes.getItem(cur.id);
        s.load("left,top,width,height,rotation,name,zOrderPosition");
        slide.shapes.load("items/id");
        await ctx.sync();
        try { ctx.presentation.setSelectedShapes([]); await ctx.sync(); } catch (e) { /* alcune versioni non accettano la selezione vuota */ }
        return { left: s.left, top: s.top, width: s.width, height: s.height, rotation: s.rotation, name: s.name, z: s.zOrderPosition, ids: slide.shapes.items.map((x) => x.id) };
      });

      // 2) inserimento della foto regolata nella stessa posizione
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
      baseline = Object.assign({}, P);
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
    if (!picking || !prev) return;
    const cv = $("#pv"), r = cv.getBoundingClientRect();
    // area effettiva della foto dentro il riquadro (object-fit: contain)
    const k = Math.min(r.width / prev.w, r.height / prev.h);
    const dw = prev.w * k, dh = prev.h * k, ox = r.left + (r.width - dw) / 2, oy = r.top + (r.height - dh) / 2;
    if (ev.clientX < ox || ev.clientX > ox + dw || ev.clientY < oy || ev.clientY > oy + dh) return;
    const x = Math.floor(((ev.clientX - ox) / dw) * prev.w), y = Math.floor(((ev.clientY - oy) / dh) * prev.h);
    let R = 0, G = 0, B = 0, n = 0;
    for (let dy = -2; dy <= 2; dy++) for (let dx = -2; dx <= 2; dx++) {
      const xx = Math.min(prev.w - 1, Math.max(0, x + dx)), yy = Math.min(prev.h - 1, Math.max(0, y + dy));
      const i = (yy * prev.w + xx) * 4;
      R += prev.data[i]; G += prev.data[i + 1]; B += prev.data[i + 2]; n++;
    }
    const nt = IL.neutralFrom(R / n, G / n, B / n);
    P.temp = nt.temp; P.tint = nt.tint; preset = "custom";
    picking = false; syncSliders(); changed();
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
    $("#autoLv").addEventListener("click", () => { if (!srcHist) return; Object.assign(P, IL.autoLevels(srcHist)); preset = "custom"; syncSliders(); changed(); });
    $("#resetLv").addEventListener("click", () => { P.b = 0; P.w = 255; P.p = 0.5; syncSliders(); changed(); });
    $("#pickBtn").addEventListener("click", () => { picking = !picking; updateTexts(); });
    $("#pv").addEventListener("click", pickAt);
    $("#beforeBtn").addEventListener("click", () => { before = !before; updateTexts(); render(); });
    $("#resetAll").addEventListener("click", () => { P = IL.defaults(); preset = "nat"; syncSliders(); changed(); });
    $("#restoreBtn").addEventListener("click", restore);
    $("#applyBtn").addEventListener("click", apply);
    $("#loadBtn").addEventListener("click", () => $("#fileIn").click());
    bindDrop();
    $("#fileIn").addEventListener("change", (e) => { loadOwnPhoto(e.target.files[0]); e.target.value = ""; });
    // In PowerPoint i link si aprono nel browser di sistema, non dentro il pannello
    $$(".links a").forEach((el) => el.addEventListener("click", (e) => {
      if (!inOffice) return;
      try {
        if (Office.context.requirements.isSetSupported("OpenBrowserWindowApi", "1.1")) { e.preventDefault(); Office.context.ui.openBrowserWindow(el.href); }
      } catch (err) { /* il link si apre normalmente */ }
    }));
    let rz = 0;
    window.addEventListener("resize", () => { cancelAnimationFrame(rz); rz = requestAnimationFrame(() => { drawHist(); placeLevelLabels(); }); });
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
        P = IL.defaults(); baseline = IL.defaults(); preset = "nat"; full = prev = null;
        LANG = detectLang(); $("#lang").value = LANG; applyStatic();
        startOffice();
      }
    });
    setTimeout(start, 2500); // se Office.js non risponde (pagina aperta nel browser) parte la prova
  }
})();
