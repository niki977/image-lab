/* The Ultraspeaker Image Lab – elaborazione dei pixel (nessuna dipendenza). I parametri p sono quelli di app.js. */
(function (root) {
  "use strict";

  const clamp01 = (v) => (v < 0 ? 0 : v > 1 ? 1 : v);
  const clamp255 = (v) => (v < 0 ? 0 : v > 255 ? 255 : v);
  const hexRGB = (h) => { const n = parseInt(String(h).slice(1), 16) || 0; return [(n >> 16) & 255, (n >> 8) & 255, n & 255]; };
  const luma = (r, g, b) => 0.299 * r + 0.587 * g + 0.114 * b;

  // Curva tonale: ombre, luci, livelli (nero/bianco) e mezzitoni. Cursore centrale a metà = gamma 1
  function toneLUT(p) {
    const N = 1024, lut = new Float32Array(N + 1);
    const sh = p.sh / 100, hl = p.hl / 100, bb = p.b / 255, span = Math.max(0.01, p.w / 255 - bb);
    const ig = Math.log(0.5) / Math.log(Math.min(0.95, Math.max(0.05, p.p)));
    for (let i = 0; i <= N; i++) {
      let x = i / N;
      x = x + sh * 1.2 * x * (1 - x) * (1 - x);      // schiarisce le ombre
      x = x - hl * 1.2 * x * x * (1 - x);            // recupera le luci
      lut[i] = Math.pow(clamp01((clamp01(x) - bb) / span), ig);
    }
    return lut;
  }

  // Esposizione, contrasto e bilanciamento del bianco: una tabella per canale (0–255 → 0–1)
  function linLUT(p) {
    const E = Math.pow(2, (p.exp / 100) * 1.5);
    const C = p.con >= 0 ? 1 + p.con / 100 : 1 + (p.con / 100) * 0.8;
    const t = p.temp / 100, m = p.tint / 100;
    let k = [1 + 0.22 * t, 1 - 0.18 * m, 1 - 0.22 * t];
    if (p.dt) {
      // con DuoTone temperatura e tinta sono filtri colore: spostano la luminosità tra i colori, i grigi restano uguali
      k = [(1 + 0.6 * t) * (1 + 0.3 * m), (1 + 0.1 * t) * (1 - 0.55 * m), (1 - 0.6 * t) * (1 + 0.3 * m)];
      const lk = 0.2126 * k[0] + 0.7152 * k[1] + 0.0722 * k[2];
      k = k.map((v) => Math.max(0.05, v / lk));
    }
    return k.map((kc) => {
      const a = new Float32Array(256);
      for (let v = 0; v < 256; v++) a[v] = ((v / 255) * kc * E - 0.5) * C + 0.5;
      return a;
    });
  }

  // DuoTone / TriTone: dalla luminosità (0–1) al colore RGBA.
  // La parte trasparente prende il colore vicino, così il bordo sfuma senza aloni
  function duoLUT(p) {
    const N = 1024, lut = new Float32Array((N + 1) * 4), tri = p.dt === 3;
    let S = hexRGB(p.dS).concat(255), M = hexRGB(p.dM).concat(255), H = hexRGB(p.dH).concat(255);
    if (p.dTr === "s") S = (tri ? M : H).slice(0, 3).concat(0);
    else if (p.dTr === "h") H = (tri ? M : S).slice(0, 3).concat(0);
    const mid = Math.min(0.95, Math.max(0.05, p.dMid / 100));
    for (let i = 0; i <= N; i++) {
      const t = i / N;
      let a = S, b = H, f = t;
      if (tri) { if (t < mid) { b = M; f = t / mid; } else { a = M; f = (t - mid) / (1 - mid); } }
      for (let c = 0; c < 4; c++) lut[i * 4 + c] = a[c] + (b[c] - a[c]) * f;
    }
    return lut;
  }

  // Sfocatura a scatola separabile (due passaggi ≈ gaussiana) su 3 canali, o sulla sola luminosità con C = 1
  function blur(src, w, h, r, C) {
    r = Math.max(1, Math.round(r));
    const a = new Float32Array(w * h * C), b = new Float32Array(w * h * C), win = 2 * r + 1;
    for (let i = 0, j = 0; i < src.length; i += 4, j += C) {
      if (C === 1) a[j] = luma(src[i], src[i + 1], src[i + 2]);
      else { a[j] = src[i]; a[j + 1] = src[i + 1]; a[j + 2] = src[i + 2]; }
    }
    const pass = (inp, out, horiz) => {
      const len = horiz ? w : h, lines = horiz ? h : w;
      for (let L = 0; L < lines; L++) for (let c = 0; c < C; c++) {
        const idx = (i) => C * (horiz ? L * w + i : i * w + L) + c;
        let s = 0;
        for (let i = -r; i <= r; i++) s += inp[idx(Math.min(len - 1, Math.max(0, i)))];
        for (let i = 0; i < len; i++) {
          out[idx(i)] = s / win;
          s += inp[idx(Math.min(len - 1, i + r + 1))] - inp[idx(Math.max(0, i - r))];
        }
      }
    };
    for (let k = 0; k < 2; k++) { pass(a, b, true); pass(b, a, false); }
    return a;
  }
  // la stessa sfocatura si riusa finché sorgente e raggio non cambiano (mentre si muovono i cursori)
  function cachedBlur(cache, name, src, w, h, r, C) {
    const e = cache[name];
    if (e && e.src === src && e.r === r) return e.v;
    const v = blur(src, w, h, r, C);
    cache[name] = { src, r, v };
    return v;
  }

  /* Applica tutte le regolazioni. src e out: RGBA della stessa misura; cache: oggetto per riusare le sfocature;
     ref: lato corto dell'intera foto nella scala di src (serve quando src è solo una zona ingrandita) */
  function process(src, w, h, p, out, cache = {}, ref = Math.min(w, h)) {
    const [L0, L1, L2] = linLUT(p), tone = toneLUT(p), N = tone.length - 1;
    const S = p.dt ? 1 : 1 + p.sat / 100; // con DuoTone la saturazione non avrebbe effetto
    const duo = p.dt ? duoLUT(p) : null, dA = clamp01(p.dA / 100);
    const nrA = (p.nr / 100) * 0.85, shA = (p.sharp / 100) * 1.6;
    // "Migliora qualità": leviga il rumore rispettando i contorni, rinforza solo i dettagli veri, aggiunge contrasto locale
    const qa = p.q / 100;
    const qDen = Math.min(0.95, qa * (0.65 + 0.006 * p.qn)), qT = 6 + 0.5 * p.qn, qShp = qa * (1.1 - 0.006 * p.qn), qCla = qa * 0.55;
    const bl = nrA > 0 || shA > 0 ? cachedBlur(cache, "nr", src, w, h, Math.max(1, ref / 500), 3) : null;
    const qbl = qa > 0 ? cachedBlur(cache, "q", src, w, h, Math.max(1, (ref / 500) * (1 + p.qn / 35)), 3) : null;
    const lb = qa > 0 ? cachedBlur(cache, "local", src, w, h, Math.max(2, ref / 45), 1) : null;
    for (let i = 0, j = 0, q = 0; i < src.length; i += 4, j += 3, q++) {
      let r = src[i], g = src[i + 1], b = src[i + 2];
      if (qbl) {
        const br = qbl[j], bg = qbl[j + 1], bb = qbl[j + 2];
        const dr = r - br, dg = g - bg, db = b - bb;
        const d = (Math.abs(dr) + Math.abs(dg) + Math.abs(db)) / 3;
        // piccole differenze = rumore (si leviga), grandi = contorni (si conservano e si rinforzano)
        const k = d / qT, wDen = qDen * Math.exp(-k * k);
        const sh = qShp * clamp01((d - 0.6 * qT) / (0.8 * qT));
        r += wDen * (br - r) + sh * dr;
        g += wDen * (bg - g) + sh * dg;
        b += wDen * (bb - b) + sh * db;
        const Lp = luma(r, g, b), t = Lp / 127.5 - 1, c = qCla * (Lp - lb[q]) * (1 - t * t);
        r = clamp255(r + c); g = clamp255(g + c); b = clamp255(b + c);
      }
      if (bl) {
        const br = bl[j], bg = bl[j + 1], bb = bl[j + 2];
        r = clamp255(r + nrA * (br - r) + shA * (r - br));
        g = clamp255(g + nrA * (bg - g) + shA * (g - bg));
        b = clamp255(b + nrA * (bb - b) + shA * (b - bb));
      }
      let x = L0[r | 0], y = L1[g | 0], z = L2[b | 0];
      if (S !== 1) {
        const l = 0.2126 * x + 0.7152 * y + 0.0722 * z;
        x = l + (x - l) * S; y = l + (y - l) * S; z = l + (z - l) * S;
      }
      const R = tone[(clamp01(x) * N + 0.5) | 0] * 255, G = tone[(clamp01(y) * N + 0.5) | 0] * 255, B = tone[(clamp01(z) * N + 0.5) | 0] * 255;
      if (duo) {
        const k = ((clamp01((0.2126 * R + 0.7152 * G + 0.0722 * B) / 255) * N + 0.5) | 0) * 4;
        const dn = (Math.imul(q + 1, 2654435761) >>> 24) / 256 - 0.5; // un filo di rumore contro le bande nei passaggi lisci
        out[i] = R + (duo[k] - R) * dA + dn;
        out[i + 1] = G + (duo[k + 1] - G) * dA + dn;
        out[i + 2] = B + (duo[k + 2] - B) * dA + dn;
        out[i + 3] = src[i + 3] * (1 - dA * (1 - duo[k + 3] / 255)) + 0.5;
      } else {
        out[i] = R + 0.5; out[i + 1] = G + 0.5; out[i + 2] = B + 0.5; out[i + 3] = src[i + 3];
      }
    }
    return out;
  }

  // Istogramma a 256 livelli per R, G, B e luminosità (solo pixel visibili)
  function histogram(data) {
    const r = new Float32Array(256), g = new Float32Array(256), b = new Float32Array(256), l = new Float32Array(256);
    let n = 0, lo = 0, hi = 0, alpha = false;
    for (let i = 0; i < data.length; i += 4) {
      if (data[i + 3] < 255) { alpha = true; if (data[i + 3] < 8) continue; }
      const R = data[i], G = data[i + 1], B = data[i + 2];
      const L = (0.2126 * R + 0.7152 * G + 0.0722 * B + 0.5) | 0;
      r[R]++; g[G]++; b[B]++; l[L]++; n++;
      if (L <= 1) lo++; else if (L >= 254) hi++;
    }
    return { r, g, b, l, n, alpha, clipLo: n > 0 && lo / n > 0.01, clipHi: n > 0 && hi / n > 0.01 };
  }

  // Livelli automatici: taglia lo 0,2% più scuro e più chiaro della luminosità
  function autoLevels(h) {
    const cut = h.n * 0.002;
    let s = 0, b = 0, w = 255;
    for (let i = 0; i < 256; i++) { s += h.l[i]; if (s > cut) { b = i; break; } }
    s = 0;
    for (let i = 255; i >= 0; i--) { s += h.l[i]; if (s > cut) { w = i; break; } }
    b = Math.min(b, 120); w = Math.max(w, 135);
    return w - b < 20 ? { b: 0, w: 255, p: 0.5 } : { b, w, p: 0.5 };
  }

  // Punto neutro: temperatura e tinta che rendono grigio il colore campionato (valori 0–255)
  function neutralFrom(R, G, B) {
    R = Math.max(1, R); G = Math.max(1, G); B = Math.max(1, B);
    const t = Math.max(-1, Math.min(1, (B - R) / (0.22 * (R + B))));
    const m = Math.max(-1, Math.min(1, (1 - (R * (1 + 0.22 * t)) / G) / 0.18));
    return { temp: Math.round(t * 100), tint: Math.round(m * 100) };
  }

  // Stima del rumore (0–100) su una zona centrale: scarto mediano tra ogni pixel e la media dei vicini
  function estimateNoise(data, w, h) {
    const sz = Math.min(480, w - 2, h - 2);
    if (sz < 8) return 20;
    const x0 = Math.floor((w - sz) / 2), y0 = Math.floor((h - sz) / 2);
    const L = (x, y) => { const i = (y * w + x) * 4; return luma(data[i], data[i + 1], data[i + 2]); };
    const res = [];
    for (let y = y0 + 1; y < y0 + sz - 1; y += 2) for (let x = x0 + 1; x < x0 + sz - 1; x += 2) {
      res.push(Math.abs(L(x, y) - (L(x - 1, y) + L(x + 1, y) + L(x, y - 1) + L(x, y + 1)) / 4));
    }
    res.sort((a, b) => a - b);
    return Math.max(0, Math.min(100, Math.round(res[res.length >> 1] * 1.4826 * 8)));
  }

  const api = { process, histogram, autoLevels, neutralFrom, estimateNoise, hexRGB };
  if (typeof module !== "undefined") module.exports = api;
  else root.IL = api;
})(this);
