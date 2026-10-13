/* flow.js – mesin bagan alir SPO
   Tiga lapis terpisah:
   1) parse()    : teks prosedur  -> langkah (aksi, sublangkah, keputusan, jalur)
   2) chart()    : langkah        -> model alur (makro/mikro) + validasi
   3) render()   : model alur     -> tata letak otomatis + gambar (canvas)
   Struktur diagram mengikuti isi prosedur, tidak ada cabang yang diada-adakan. */
const FLOW = (() => {
  const NUM = /^\d+[.)]\s*/;
  const SUB = /^(?:[-•*–]|[a-z][.)])\s+/i;
  const LOOP = /^kembali\s+ke\s+(?:langkah\s+)?(\d+)\.?$/i;
  const cap = s => s.charAt(0).toUpperCase() + s.slice(1);
  const short = s => (s.length > 32 ? s.slice(0, 30) + "…" : s);

  function classAct(a) {
    if (!a || /^(selesai|berhenti|stop)\.?$/i.test(a)) return { k: "end" };
    const m = a.match(LOOP);
    return m ? { k: "loop", n: +m[1] } : { k: "box", t: a };
  }

  /* ---------- 1) PARSER ---------- */
  function parse(text) {
    const raw = text.split(/\r?\n/).filter(l => l.trim());
    const ind = l => l.match(/^\s*/)[0].length;
    const base = raw.length ? Math.min(...raw.map(ind)) : 0;
    const steps = [], warn = [];
    let ref = 0;

    raw.forEach(l => {
      const t = l.trim();
      const isSub = steps.length && !NUM.test(t) && (ind(l) > base || SUB.test(t));
      if (isSub) {
        steps.push({ ref, sub: true, pre: "", text: t.replace(SUB, "").trim(), q: false, branches: [] });
        return;
      }
      const pre = (t.match(NUM) || [""])[0];
      const s = t.replace(NUM, "");
      ref++;
      const parts = s.split("|").map(x => x.trim());

      if (/\?$/.test(parts[0])) {
        const branches = parts.slice(1).filter(Boolean).map(seg => {
          const m = seg.match(/^([^:]{1,24}):\s*(.*)$/);
          return m ? { label: cap(m[1].trim()), act: m[2].trim() } : { label: "Tidak", act: seg };
        });
        steps.push({ ref, sub: false, pre, text: parts[0], q: true, branches });
      } else {
        if (/^(jika|apabila|bila)\b/i.test(s))
          warn.push(`Langkah ${ref} berawalan "jika/apabila" tetapi tidak berakhiran "?", jadi tidak dijadikan keputusan. Ubah menjadi pertanyaan bila ingin percabangan.`);
        steps.push({ ref, sub: false, pre, text: s, q: false, branches: [] });
      }
    });
    return { steps, warn };
  }

  /* teks untuk dokumen Word; sublangkah diberi penanda \u0001 */
  function docLines(text) {
    const act = a => {
      const c = classAct(a);
      return c.k === "end" ? "selesai" : c.k === "loop" ? "kembali ke langkah " + c.n : a;
    };
    return parse(text).steps.map(s => {
      if (s.sub) return "\u0001" + s.text;
      const j = s.branches.length
        ? " " + s.branches.map(b => "Jika " + b.label.toLowerCase() + ": " + act(b.act)).join("; ") + "."
        : "";
      return s.pre + s.text + j;
    });
  }

  /* ---------- peringkas (Mode Makro) ---------- */
  const STOP_END = /^(dan|atau|yang|di|ke|dari|dengan|untuk|pada|serta|oleh|agar|supaya|sebagai|dalam|bagi|kepada|terhadap|sesuai|sampai|hingga)$/i;
  const CUT_RE = /\s*(?:[;:]|,\s|\s(?:sehingga|agar|supaya|untuk memastikan|untuk mengetahui|karena|kemudian|lalu|sesuai dengan|sesuai|berdasarkan|yaitu|yakni|dimana|di mana)\s)/gi;

  function gist(t, maxW) {
    let s = t.replace(/\([^)]*\)/g, " ").replace(/\s+/g, " ").trim();
    const q = /\?$/.test(s);
    s = s.replace(/\?$/, "").replace(/[.;:,]+$/, "").trim();
    CUT_RE.lastIndex = 0;
    let m;
    while ((m = CUT_RE.exec(s))) {
      if (m.index >= 12) { s = s.slice(0, m.index).trim(); break; }
      if (m[0].length === 0) CUT_RE.lastIndex++;
    }
    let w = s.split(" ");
    if (w.length > maxW) w = w.slice(0, maxW);
    while (w.length > 2 && STOP_END.test(w[w.length - 1])) w.pop();
    s = cap(w.join(" "));
    return q ? s + "?" : s;
  }

  function joinGists(src, w) {
    const g = src.map(x => gist(x, w));
    const subj = g[0].split(" ")[0].toLowerCase();
    return g.map((x, i) => {
      if (i === 0) return x;
      const p = x.split(" ");
      if (p.length > 2 && p[0].toLowerCase() === subj) p.shift();
      return p.join(" ").replace(/^./, c => c.toLowerCase());
    }).join(" & ");
  }

  function condense(steps, maxBoxes) {
    const items = steps.map(s => ({
      q: s.q, src: [s.text], text: gist(s.text, s.q ? 10 : 9),
      branches: s.branches.map(b => ({ label: b.label, act: classAct(b.act).k === "box" ? gist(b.act, 7) : b.act })),
      orig: [s.ref]
    }));
    while (items.length > maxBoxes) {
      let best = -1, bestCost = Infinity;
      for (let i = 0; i < items.length - 1; i++) {
        const a = items[i], b = items[i + 1];
        if (a.q || b.q) continue;
        const cost = (a.orig.length + b.orig.length) * 1000 + a.text.length + b.text.length;
        if (cost < bestCost) { bestCost = cost; best = i; }
      }
      if (best < 0) break;
      const a = items[best], b = items[best + 1], src = [...a.src, ...b.src];
      items.splice(best, 2, {
        q: false, src, text: joinGists(src, Math.max(3, Math.floor(14 / src.length))),
        branches: [], orig: [...a.orig, ...b.orig]
      });
    }
    const map = {};
    items.forEach((it, ni) => it.orig.forEach(o => (map[o] = ni + 1)));
    return items.map((it, ni) => ({
      ref: ni + 1, sub: false, text: it.text, q: it.q,
      branches: it.branches.map(b => {
        const m = b.act.match(LOOP);
        return { label: b.label, act: m && map[+m[1]] ? "kembali ke " + map[+m[1]] : b.act };
      })
    }));
  }

  /* ---------- 2) MODEL ALUR + VALIDASI ---------- */
  function validate(nodes) {
    const w = [];
    nodes.forEach((n, i) => {
      if (!n.q) return;
      const tag = `Keputusan "${short(n.text)}"`;
      if (!n.branches.length) {
        w.push(`${tag}: jalur belum ditulis, sementara diarahkan ke Selesai. Tambahkan "| Tidak: ...".`);
        n.branches = [{ label: "Tidak", act: "" }];
      }
      const seen = {};
      n.branches.forEach(b => {
        const k = b.label.toLowerCase();
        if (seen[k]) w.push(`${tag}: label jalur "${b.label}" dobel.`);
        seen[k] = 1;
        if (!b.act && n.branches.length > 1) w.push(`${tag}: jalur "${b.label}" tidak berisi tindakan, dianggap selesai.`);
        b.c = classAct(b.act);
        if (b.c.k === "loop" && !nodes.some((t, j) => !t.sub && t.ref === b.c.n && j < i)) {
          w.push(`${tag}: "kembali ke ${b.c.n}" tidak menunjuk langkah sebelumnya, dianggap selesai.`);
          b.c = { k: "end" };
        }
      });
      if (n.branches.length > 1 && i < nodes.length - 1 && n.branches.every(b => b.c.k !== "box"))
        w.push(`${tag}: semua jalur berakhir atau mengulang, sehingga langkah sesudahnya tidak punya jalur masuk.`);
    });
    return w;
  }

  /* macro=true: ringkas (sublangkah dilipat); macro=false: mikro (semua langkah) */
  function chart(text, macro, maxBoxes) {
    const { steps, warn } = parse(text);
    const nodes = macro
      ? condense(steps.filter(s => !s.sub), maxBoxes)
      : steps.map(s => ({ ref: s.ref, sub: s.sub, text: s.text, q: s.q, branches: s.branches.map(b => ({ ...b })) }));
    return { nodes, warn: warn.concat(validate(nodes)) };
  }

  /* ---------- 3) TATA LETAK + RENDER ---------- */
  function render(nodes, withTerm) {
    const FS = 15, LH = 19, PV = 10, M = 12, S = 2;
    const FONT = "Arial, Helvetica, sans-serif", F = FS + "px " + FONT;
    const dec = nodes.some(n => n.q);
    const fanN = Math.max(0, ...nodes.filter(n => n.q && n.branches.length > 1).map(n => n.branches.length));
    const hasFan = fanN > 0, hasSide = nodes.some(n => n.q && n.branches.length === 1);
    const FW = 150, FG = 14, BW = dec ? 300 : 460, TW = 150, SW = 155;

    let LW = Math.max(560, fanN * (FW + FG) - FG + 2 * M);
    if (hasFan && hasSide) LW = Math.max(LW, 760);
    const CX = hasFan ? LW / 2 : dec ? 185 : 280;
    const SX = CX + BW / 2 + 60;

    const m = document.createElement("canvas").getContext("2d");
    m.font = F;
    const wrap = (txt, max) => {
      const out = []; let cur = "";
      for (const w of txt.split(/\s+/)) {
        const t = cur ? cur + " " + w : w;
        if (cur && m.measureText(t).width > max) { out.push(cur); cur = w; } else cur = t;
      }
      if (cur) out.push(cur);
      return out.length ? out : [""];
    };

    /* daftar simbol */
    const items = [];
    if (withTerm) items.push({ k: "t", text: "Mulai" });
    let si = 0;
    nodes.forEach(n => {
      si = n.sub ? si + 1 : 0;
      items.push({ k: n.q ? "d" : "p", text: n.text, n: n.ref, sub: n.sub, lab: n.sub ? n.ref + "." + si : String(n.ref), br: n.branches });
    });
    if (withTerm) items.push({ k: "t", text: "Selesai" });

    /* ukuran */
    items.forEach(it => {
      if (it.k === "t") { it.w = TW; it.l = [it.text]; it.h = 34; }
      else if (it.k === "p") { it.w = BW; it.l = wrap(it.text, BW - 24); it.h = it.l.length * LH + PV * 2; }
      else {
        it.w = BW; it.l = wrap(it.text, BW * 0.56);
        it.h = Math.max(it.l.length * LH + 40, 70);
        if (it.br.length === 1) {
          it.act = it.br[0].c;
          if (it.act.k === "box") {
            it.sl = wrap(it.act.t, SW - 18);
            it.sh = it.sl.length * LH + PV * 2;
            it.h = Math.max(it.h, it.sh + 14);
          }
        } else {
          it.fan = true;
          it.br.forEach(b => {
            if (b.c.k === "end") { b.l = ["Selesai"]; b.h = 30; }
            else {
              b.l = wrap(b.c.k === "loop" ? "Kembali ke langkah " + b.c.n : b.c.t, FW - 16);
              b.h = b.l.length * LH + PV * 2;
            }
          });
          it.hasJoin = it.br.some(b => b.c.k === "box");
          it.rh = Math.max(...it.br.map(b => b.h));
          it.ext = 40 + it.rh + 16;
        }
      }
    });

    const gapOf = it => (it.k === "d" ? 32 : 24);
    let y = M;
    items.forEach(it => {
      it.top = y; it.cy = y + it.h / 2; it.bot = y + it.h;
      it.end = it.bot + (it.ext || 0);
      y = it.end + gapOf(it);
    });
    const H = items[items.length - 1].end + M;

    /* gambar */
    const c = document.createElement("canvas");
    c.width = LW * S; c.height = Math.round(H * S);
    const g = c.getContext("2d");
    g.scale(S, S);
    g.fillStyle = "#fff"; g.fillRect(0, 0, LW, H);
    g.font = F; g.textBaseline = "middle"; g.lineWidth = 1.4; g.strokeStyle = "#111";

    const ln = (x1, y1, x2, y2) => { g.beginPath(); g.moveTo(x1, y1); g.lineTo(x2, y2); g.stroke(); };
    const ah = (x, yy, d) => {
      g.beginPath(); g.moveTo(x, yy);
      if (d === "d") { g.lineTo(x - 5, yy - 10); g.lineTo(x + 5, yy - 10); }
      else if (d === "l") { g.lineTo(x + 10, yy - 5); g.lineTo(x + 10, yy + 5); }
      else { g.lineTo(x - 10, yy - 5); g.lineTo(x - 10, yy + 5); }
      g.closePath(); g.fillStyle = "#111"; g.fill();
    };
    const txt = (t, x, yy, al, font, col) => {
      g.save(); g.font = font || F; g.fillStyle = col || "#111"; g.textAlign = al || "center";
      g.fillText(t, x, yy); g.restore();
    };
    const lab = (t, x, yy, al) => txt(t, x, yy, al, "bold 12px " + FONT);
    const pill = (x, yy, w, h) => {
      g.fillStyle = "#dcefe9"; g.beginPath(); g.roundRect(x, yy, w, h, h / 2); g.fill(); g.stroke();
    };
    const box = (x, yy, w, h, fill, dash) => {
      g.fillStyle = fill; g.fillRect(x, yy, w, h);
      g.save(); if (dash) g.setLineDash([5, 3]); g.strokeRect(x, yy, w, h); g.restore();
    };
    const lines = (arr, x, cy) => arr.forEach((s, k) => txt(s, x, cy - arr.length * LH / 2 + LH / 2 + k * LH));

    let loops = 0;
    items.forEach((it, i) => {
      const last = i === items.length - 1;
      const x0 = CX - it.w / 2, rx = CX + it.w / 2;

      if (it.k === "t") pill(x0, it.top, it.w, it.h);
      else if (it.k === "p") box(x0, it.top, it.w, it.h, it.sub ? "#e8eff5" : "#f3f6f9");
      else {
        g.fillStyle = "#fff4d6"; g.beginPath();
        g.moveTo(CX, it.top); g.lineTo(rx, it.cy); g.lineTo(CX, it.bot); g.lineTo(x0, it.cy);
        g.closePath(); g.fill(); g.stroke();
      }
      lines(it.l, CX, it.cy);
      if (it.lab) txt(it.lab, x0 - 6, it.cy, "right", "12px " + FONT, "#666");

      /* panah turun ke simbol berikutnya */
      if (!last && (!it.fan || it.hasJoin)) {
        const gp = gapOf(it), y1 = it.end;
        ln(CX, y1, CX, y1 + gp - 1);
        ah(CX, y1 + gp, "d");
        if (it.k === "d" && !it.fan) {
          const sl = it.br[0].label.toLowerCase();
          lab(sl === "tidak" ? "Ya" : sl === "ya" ? "Tidak" : "Lanjut", CX - 14, it.bot + 16);
        }
      }

      /* keputusan dua jalur: jalur samping */
      if (it.k === "d" && !it.fan) {
        const a = it.act, blab = it.br[0].label, yJoin = it.bot + 18;
        if (a.k === "loop") {
          const tg = items.find((t, j) => j < i && t.k !== "t" && !t.sub && t.n === a.n);
          const lx = rx + 15 + Math.min(loops++, 3) * 12;
          ln(rx, it.cy, lx, it.cy); ln(lx, it.cy, lx, tg.cy); ln(lx, tg.cy, rx + 1, tg.cy);
          ah(rx, tg.cy, "l");
          lab(blab, lx + 4, it.cy - 8, "left");
        } else {
          const sy = it.cy - (a.k === "box" ? it.sh / 2 : 15);
          if (a.k === "box") {
            box(SX, sy, SW, it.sh, "#f3f6f9");
            lines(it.sl, SX + SW / 2, it.cy);
          } else {
            pill(SX, sy, 100, 30); txt("Selesai", SX + 50, it.cy);
          }
          ln(rx, it.cy, SX, it.cy); ah(SX, it.cy, "r");
          lab(blab, (rx + SX) / 2, it.cy - 8);
          if (a.k === "box" && !last) {
            ln(SX + SW / 2, sy + it.sh, SX + SW / 2, yJoin);
            ln(SX + SW / 2, yJoin, CX + 1, yJoin);
            ah(CX + 1, yJoin, "l");
          }
        }
      }

      /* keputusan banyak jalur: cabang sebanyak jalur yang ditulis */
      if (it.fan) {
        const n = it.br.length, step = FW + FG, xa = CX - (n - 1) * step / 2;
        const busY = it.bot + 14, bTop = it.bot + 40, joinY = it.end;
        const joins = [];
        ln(CX, it.bot, CX, busY);
        ln(xa, busY, xa + (n - 1) * step, busY);
        it.br.forEach((b, k) => {
          const bx = xa + k * step;
          ln(bx, busY, bx, bTop - 1); ah(bx, bTop, "d");
          lab(b.label, bx + 6, (busY + bTop) / 2 + 2, "left");
          if (b.c.k === "end") { pill(bx - 55, bTop, 110, 30); txt("Selesai", bx, bTop + 15); }
          else {
            box(bx - FW / 2, bTop, FW, b.h, b.c.k === "loop" ? "#fff" : "#f3f6f9", b.c.k === "loop");
            lines(b.l, bx, bTop + b.h / 2);
          }
          if (b.c.k === "box") { joins.push(bx); if (!last) ln(bx, bTop + b.h, bx, joinY); }
        });
        if (joins.length && !last) {
          const xs = [...joins, CX];
          ln(Math.min(...xs), joinY, Math.max(...xs), joinY);
        }
      }
    });

    return { canvas: c, w: LW, h: H };
  }

  return { parse, chart, render, docLines, classAct };
})();
