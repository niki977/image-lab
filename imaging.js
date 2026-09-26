/* The Ultraspeaker Image Lab – elaborazione delle immagini (nessuna dipendenza, gira nel pannello) */
(function (root) {
  "use strict";

  const DEFAULTS = { exp: 0, con: 0, sat: 0, hl: 0, sh: 0, sharp: 0, nr: 0, temp: 0, tint: 0, b: 0, w: 255, p: 0.5 };

  function defaults() { return Object.assign({}, DEFAULTS); }
  function isNeutral(p) { for (const k in DEFAULTS) if (Math.abs((p[k] ?? DEFAULTS[k]) - DEFAULTS[k]) > 1e-6) return false; return true; }
  const clamp01 = (v) => (v < 0 ? 0 : v > 1 ? 1 : v);

  // Gamma dei mezzitoni: il cursore centrale in posizione relativa p (0–1) tra nero e bianco. p = 0,5 → gamma 1
  function gamma(p) { const q = Math.min(0.95, Math.max(0.05, p)); return Math.log(q) / Math.log(0.5); }

  // Curva tonale finale: ombre, luci, livelli (nero/bianco) e mezzitoni
  function toneLUT(p) {
    const N = 1024, lut = new Float32Array(N + 1);
    const sh = p.sh / 100, hl = p.hl / 100, bb = p.b / 255, ww = p.w / 255, ig = 1 / gamma(p.p);
    const span = Math.max(0.01, ww - bb);
    for (let i = 0; i <= N; i++) {
      let x = i / N;
      x = x + sh * 1.2 * x * (1 - x) * (1 - x);      // schiarisce le ombre
      x = x - hl * 1.2 * x * x * (1 - x);            // recupera le luci
      let y = clamp01((clamp01(x) - bb) / span);     // livelli
      lut[i] = Math.pow(y, ig);                      // mezzitoni
    }
    return lut;
  }

  // Esposizione, contrasto e bilanciamento del bianco: una tabella per canale (0–255 → valore lineare)
  function linLUT(p) {
    const E = Math.pow(2, (p.exp / 100) * 1.5);
    const C = p.con >= 0 ? 1 + p.con / 100 : 1 + (p.con / 100) * 0.8;
    const t = p.temp / 100, m = p.tint / 100;
    const k = [1 + 0.22 * t, 1 - 0.18 * m, 1 - 0.22 * t];
    const luts = [0, 1, 2].map((c) => {
      const a = new Float32Array(256);
      for (let v = 0; v < 256; v++) a[v] = ((v / 255) * k[c] * E - 0.5) * C + 0.5;
      return a;
    });
    return luts;
  }

  // Sfocatura a scatola separabile (due passaggi ≈ gaussiana), usata per riduzione rumore e nitidezza
  function blur(src, w, h) {
    const r = Math.max(1, Math.round(Math.min(w, h) / 500));
    let a = new Float32Array(w * h * 3), b = new Float32Array(w * h * 3);
    for (let i = 0, j = 0; i < src.length; i += 4, j += 3) { a[j] = src[i]; a[j + 1] = src[i + 1]; a[j + 2] = src[i + 2]; }
    const pass = (inp, out, horiz) => {
      const len = horiz ? w : h, lines = horiz ? h : w, win = 2 * r + 1;
      for (let L = 0; L < lines; L++) {
        for (let c = 0; c < 3; c++) {
          const idx = (i) => 3 * (horiz ? L * w + i : i * w + L) + c;
          let s = 0;
          for (let i = -r; i <= r; i++) s += inp[idx(Math.min(len - 1, Math.max(0, i)))];
          for (let i = 0; i < len; i++) {
            out[idx(i)] = s / win;
            s += inp[idx(Math.min(len - 1, i + r + 1))] - inp[idx(Math.max(0, i - r))];
          }
        }
      }
    };
    for (let k = 0; k < 2; k++) { pass(a, b, true); pass(b, a, false); }
    return a;
  }

  /* Applica tutte le regolazioni.
     src: Uint8ClampedArray RGBA, out: Uint8ClampedArray RGBA della stessa misura, cache: oggetto per riusare la sfocatura */
  function process(src, w, h, p, out, cache) {
    p = Object.assign(defaults(), p);
    const lin = linLUT(p), tone = toneLUT(p), N = tone.length - 1;
    const S = 1 + p.sat / 100;
    const nrA = (p.nr / 100) * 0.85, shA = (p.sharp / 100) * 1.6;
    let bl = null;
    if (nrA > 0 || shA > 0) {
      if (cache && cache.blur && cache.src === src) bl = cache.blur;
      else { bl = blur(src, w, h); if (cache) { cache.blur = bl; cache.src = src; } }
    }
    const L0 = lin[0], L1 = lin[1], L2 = lin[2];
    for (let i = 0, j = 0; i < src.length; i += 4, j += 3) {
      let r = src[i], g = src[i + 1], b = src[i + 2];
      if (bl) {
        const br = bl[j], bg = bl[j + 1], bb = bl[j + 2];
        const dr = r - br, dg = g - bg, db = b - bb;
        r = r + nrA * (br - r) + shA * dr;
        g = g + nrA * (bg - g) + shA * dg;
        b = b + nrA * (bb - b) + shA * db;
        r = r < 0 ? 0 : r > 255 ? 255 : r; g = g < 0 ? 0 : g > 255 ? 255 : g; b = b < 0 ? 0 : b > 255 ? 255 : b;
      }
      let x = L0[r | 0], y = L1[g | 0], z = L2[b | 0];
      if (S !== 1) {
        const l = 0.2126 * x + 0.7152 * y + 0.0722 * z;
        x = l + (x - l) * S; y = l + (y - l) * S; z = l + (z - l) * S;
      }
      x = x < 0 ? 0 : x > 1 ? 1 : x; y = y < 0 ? 0 : y > 1 ? 1 : y; z = z < 0 ? 0 : z > 1 ? 1 : z;
      out[i] = tone[(x * N + 0.5) | 0] * 255 + 0.5;
      out[i + 1] = tone[(y * N + 0.5) | 0] * 255 + 0.5;
      out[i + 2] = tone[(z * N + 0.5) | 0] * 255 + 0.5;
      out[i + 3] = src[i + 3];
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
    if (w - b < 20) { b = 0; w = 255; }
    return { b, w, p: 0.5 };
  }

  // Punto neutro: temperatura e tinta che rendono grigio il colore campionato (valori 0–255)
  function neutralFrom(R, G, B) {
    R = Math.max(1, R); G = Math.max(1, G); B = Math.max(1, B);
    let t = (B - R) / (0.22 * (R + B));
    t = Math.max(-1, Math.min(1, t));
    const target = R * (1 + 0.22 * t);
    let m = (1 - target / G) / 0.18;
    m = Math.max(-1, Math.min(1, m));
    return { temp: Math.round(t * 100), tint: Math.round(m * 100) };
  }

  const api = { defaults, isNeutral, gamma, toneLUT, linLUT, blur, process, histogram, autoLevels, neutralFrom };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.IL = api;
})(typeof self !== "undefined" ? self : this);
