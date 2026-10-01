const TEMPLATE_URL = "assets/template-spo.docx";

async function loadTemplateZip() {
  const response = await fetch(TEMPLATE_URL);
  if (!response.ok) throw new Error(`Template SPO tidak dapat dimuat (${response.status}).`);
  return JSZip.loadAsync(await response.arrayBuffer());
}

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const $ = id => document.getElementById(id);
const TEXT_IDS = ["unit","judul","pengertian","tujuan","kebijakan","prosedur","bagan"];
const NUM_IDS = ["n_tujuan","n_kebijakan","n_prosedur"];
let baganMode = "auto", baganFile = null;

/* ---------- daftar dinamis ---------- */
function addRow(listId, value = "", ph = "") {
  const r = document.createElement("div"); r.className = "row";
  const i = document.createElement("input"); i.type = "text"; i.value = value; i.placeholder = ph;
  const x = document.createElement("button"); x.type = "button"; x.className = "x"; x.textContent = "✕"; x.title = "Hapus";
  x.onclick = () => { if ($(listId).children.length > 1) r.remove(); else i.value = ""; save(); };
  i.oninput = save; r.append(i, x); $(listId).append(r);
}
const listVals = id => [...$(id).querySelectorAll("input")].map(i => i.value.trim()).filter(Boolean);
$("addUnit").onclick = () => addRow("unitList", "", "mis. Instalasi Rawat Inap");
$("addDoc").onclick = () => addRow("docList", "", "mis. Formulir Serah Terima Pasien");

/* ---------- simpan draf (opsional, gagal senyap) ---------- */
function save() {
  try {
    const d = {t:{}, n:{}, u:listVals("unitList"), d:listVals("docList")};
    TEXT_IDS.forEach(k => d.t[k] = $(k).value); NUM_IDS.forEach(k => d.n[k] = $(k).checked);
    localStorage.setItem("spo_draft", JSON.stringify(d));
  } catch (e) {}
}
function load() {
  let d = null;
  try { d = JSON.parse(localStorage.getItem("spo_draft")); } catch (e) {}
  if (d) {
    TEXT_IDS.forEach(k => $(k).value = d.t?.[k] || "");
    NUM_IDS.forEach(k => {
      if (d.n && k in d.n) $(k).checked = d.n[k];
    });
    (d.u?.length ? d.u : [""]).forEach(v =>
      addRow("unitList", v, "mis. Instalasi Rawat Inap")
    );
    (d.d?.length ? d.d : [""]).forEach(v =>
      addRow("docList", v, "mis. Formulir Serah Terima Pasien")
    );
  } else {
    addRow("unitList");
    addRow("docList");
  }
}

[...TEXT_IDS, ...NUM_IDS].forEach(k =>
  $(k).addEventListener("input", save)
);

$("reset").onclick = () => {
  if (!confirm("Kosongkan semua isian?")) return;

  TEXT_IDS.forEach(k => $(k).value = "");
  $("unitList").innerHTML = "";
  $("docList").innerHTML = "";

  addRow("unitList");
  addRow("docList");

  baganFile = null;
  $("file").value = "";
  $("imgprev").style.display = "none";
  save();
};

/* ---------- bagan alir ---------- */
function setMode(m) {
  baganMode = m;

  $("tabAuto").setAttribute("aria-pressed", m === "auto");
  $("tabImg").setAttribute("aria-pressed", m === "img");
  $("tabTxt").setAttribute("aria-pressed", m === "txt");

  $("paneAuto").hidden = m !== "auto";
  $("paneImg").hidden = m !== "img";
  $("paneTxt").hidden = m !== "txt";

  if (m === "auto") updPrev();
}

$("tabAuto").onclick = () => setMode("auto");
$("tabImg").onclick = () => setMode("img");
$("tabTxt").onclick = () => setMode("txt");

$("file").onchange = e => {
  baganFile = e.target.files[0] || null;
  const p = $("imgprev");

  if (baganFile) {
    p.src = URL.createObjectURL(baganFile);
    p.style.display = "block";
  } else {
    p.style.display = "none";
  }
};

async function toPng(file) {
  const bmp = await createImageBitmap(file);
  const s = Math.min(1, 1800 / bmp.width);

  const c = document.createElement("canvas");
  c.width = Math.round(bmp.width * s);
  c.height = Math.round(bmp.height * s);

  const g = c.getContext("2d");
  g.fillStyle = "#fff";
  g.fillRect(0, 0, c.width, c.height);
  g.drawImage(bmp, 0, 0, c.width, c.height);

  const blob = await new Promise(r =>
    c.toBlob(r, "image/png")
  );

  return {
    buf: await blob.arrayBuffer(),
    w: c.width,
    h: c.height
  };
}

/* ---------- diagram alir otomatis ---------- */
function parseLine(l) {
  let m = l.match(/^(\d+[.)]\s*)?(.+?)\?\s*\|\s*tidak\s*:\s*(.+)$/i);

  if (m)
    return {
      pre: m[1] || "",
      text: m[2].trim(),
      q: true,
      act: m[3].trim()
    };

  m = l.match(/^(\d+[.)]\s*)?(.+\?)$/);

  if (m)
    return {
      pre: m[1] || "",
      text: m[2].trim(),
      q: true,
      act: ""
    };

  m = l.match(/^(\d+[.)]\s*)?(.+)$/);

  return {
    pre: m[1] || "",
    text: m[2].trim(),
    q: false,
    act: ""
  };
}

const stepsOf = () =>
  lines($("prosedur").value)
    .map(parseLine)
    .filter(s => s.text);

function classAct(a) {
  if (!a || /^(selesai|berhenti|stop)\.?$/i.test(a))
    return { k: "end" };

  const m = a.match(
    /^kembali\s+ke\s+(?:langkah\s+)?(\d+)\.?$/i
  );

  return m
    ? { k: "loop", n: +m[1] }
    : { k: "box", t: a };
}

function dispLine(s) {
  if (!s.act) return s.pre + s.text;

  const c = classAct(s.act);

  return s.pre + s.text +
    " Jika tidak: " +
    (
      c.k === "end"
        ? "selesai."
        : c.k === "loop"
          ? "kembali ke langkah " + c.n + "."
          : s.act
    );
}

function flowCanvas(steps, withTerm) {
  const dec = steps.some(s => s.q);

  const LW = 560;
  const FS = 15;
  const LH = 19;
  const PV = 10;
  const M = 12;
  const FONT = "Arial, Helvetica, sans-serif";
  const F = FS + "px " + FONT;

  const BW = dec ? 300 : 460;
  const CX = dec ? 185 : 280;
  const TW = 150;
  const SX = 395;
  const SW = 155;

  const m = document.createElement("canvas").getContext("2d");
  m.font = F;

  const wrap = (txt, max) => {
    const out = [];
    let cur = "";

    for (const w of txt.split(/\s+/)) {
      const t = cur ? cur + " " + w : w;

      if (
        cur &&
        m.measureText(t).width > max
      ) {
        out.push(cur);
        cur = w;
      } else {
        cur = t;
      }
    }

    if (cur) out.push(cur);

    return out;
  };

  const items = [];

  if (withTerm)
    items.push({
      k: "t",
      text: "Mulai"
    });

  steps.forEach((s, i) =>
    items.push({
      k: s.q ? "d" : "p",
      text: s.text,
      n: i + 1,
      act: s.q ? classAct(s.act) : null
    })
  );

  if (withTerm)
    items.push({
      k: "t",
      text: "Selesai"
    });

  items.forEach(it => {
    if (it.k === "t") {
      it.w = TW;
      it.l = [it.text];
      it.h = 34;
    }

    else if (it.k === "p") {
      it.w = BW;
      it.l = wrap(it.text, BW - 24);
      it.h = it.l.length * LH + PV * 2;
    }

    else {
      it.w = BW;
      it.l = wrap(it.text, BW * 0.56);
      it.h = Math.max(
        it.l.length * LH + 40,
        70
      );

      if (
        it.act.k === "loop" &&
        !items.some(
          t => t.n === it.act.n && t.n < it.n
        )
      ) {
        it.act = { k: "end" };
      }

      if (it.act.k === "box") {
        it.sl = wrap(
          it.act.t,
          SW - 18
        );

        it.sh = it.sl.length * LH + PV * 2;

        it.h = Math.max(
          it.h,
          it.sh + 14
        );
      }
    }
  });

  const gapOf = it =>
    it.k === "d" ? 32 : 24;

  let y = M;

  items.forEach(it => {
    it.top = y;
    it.cy = y + it.h / 2;
    it.bot = y + it.h;
    y = it.bot + gapOf(it);
  });

  const H =
    items[items.length - 1].bot + M;

  const S = 2;

  const c = document.createElement("canvas");
  c.width = LW * S;
  c.height = Math.round(H * S);

  const g = c.getContext("2d");

  g.scale(S, S);
  g.fillStyle = "#fff";
  g.fillRect(0, 0, LW, H);

  g.font = F;
  g.textBaseline = "middle";
  g.lineWidth = 1.4;
  g.strokeStyle = "#111";

  const ln = (x1, y1, x2, y2) => {
    g.beginPath();
    g.moveTo(x1, y1);
    g.lineTo(x2, y2);
    g.stroke();
  };

  const ah = (x, y, d) => {
    g.beginPath();
    g.moveTo(x, y);

    if (d === "d") {
      g.lineTo(x - 5, y - 10);
      g.lineTo(x + 5, y - 10);
    }

    else if (d === "l") {
      g.lineTo(x + 10, y - 5);
      g.lineTo(x + 10, y + 5);
    }

    else {
      g.lineTo(x - 10, y - 5);
      g.lineTo(x - 10, y + 5);
    }

    g.closePath();
    g.fillStyle = "#111";
    g.fill();
  };

  const txt = (
    t,
    x,
    y,
    al,
    font,
    col
  ) => {
    g.save();

    g.font = font || F;
    g.fillStyle = col || "#111";
    g.textAlign = al || "center";

    g.fillText(t, x, y);

    g.restore();
  };

  const pill = (
    x,
    yy,
    w,
    h,
    fill
  ) => {
    g.fillStyle = fill;
    g.beginPath();
    g.roundRect(
      x,
      yy,
      w,
      h,
      h / 2
    );
    g.fill();
    g.stroke();
  };

  const lab = (
    t,
    x,
    yy,
    al
  ) =>
    txt(
      t,
      x,
      yy,
      al,
      "bold 12px " + FONT
    );

  let loops = 0;

  items.forEach((it, i) => {
    const last = i === items.length - 1;
    const x0 = CX - it.w / 2;

    if (it.k === "t")
      pill(
        x0,
        it.top,
        it.w,
        it.h,
        "#dcefe9"
      );

    else if (it.k === "p") {
      g.fillStyle = "#f3f6f9";
      g.fillRect(
        x0,
        it.top,
        it.w,
        it.h
      );
      g.strokeRect(
        x0,
        it.top,
        it.w,
        it.h
      );
    }

    else {
      g.fillStyle = "#fff4d6";

      g.beginPath();
      g.moveTo(CX, it.top);
      g.lineTo(
        CX + it.w / 2,
        it.cy
      );
      g.lineTo(CX, it.bot);
      g.lineTo(
        x0,
        it.cy
      );
      g.closePath();

      g.fill();
      g.stroke();
    }

    const y0 =
      it.cy -
      it.l.length * LH / 2 +
      LH / 2;

    it.l.forEach((s, k) =>
      txt(
        s,
        CX,
        y0 + k * LH
      )
    );

    if (it.n)
      txt(
        String(it.n),
        x0 - 6,
        it.cy,
        "right",
        "12px " + FONT,
        "#666"
      );

    if (!last) {
      const gp = gapOf(it);

      ln(
        CX,
        it.bot,
        CX,
        it.bot + gp - 1
      );

      ah(
        CX,
        it.bot + gp,
        "d"
      );

      if (it.k === "d")
        lab(
          "Ya",
          CX - 14,
          it.bot + 16
        );
    }

    if (it.k === "d") {
      const rx = CX + it.w / 2;
      const a = it.act;
      const yJoin = it.bot + 18;

      if (a.k === "loop") {
        const tg = items.find(
          t =>
            t.n === a.n &&
            t.n < it.n
        );

        const lx =
          350 +
          Math.min(loops++, 3) * 12;

        ln(
          rx,
          it.cy,
          lx,
          it.cy
        );

        ln(
          lx,
          it.cy,
          lx,
          tg.cy
        );

        ln(
          lx,
          tg.cy,
          rx + 1,
          tg.cy
        );

        ah(
          rx,
          tg.cy,
          "l"
        );

        lab(
          "Tidak",
          lx + 4,
          it.cy - 8,
          "left"
        );
      }

      else if (a.k === "box") {
        const sy =
          it.cy -
          it.sh / 2;

        g.fillStyle = "#f3f6f9";

        g.fillRect(
          SX,
          sy,
          SW,
          it.sh
        );

        g.strokeRect(
          SX,
          sy,
          SW,
          it.sh
        );

        const s0 =
          it.cy -
          it.sl.length * LH / 2 +
          LH / 2;

        it.sl.forEach((s, k) =>
          txt(
            s,
            SX + SW / 2,
            s0 + k * LH
          )
        );

        ln(
          rx,
          it.cy,
          SX,
          it.cy
        );

        ah(
          SX,
          it.cy,
          "r"
        );

        lab(
          "Tidak",
          (rx + SX) / 2,
          it.cy - 8
        );

        if (!last) {
          ln(
            SX + SW / 2,
            sy + it.sh,
            SX + SW / 2,
            yJoin
          );

          ln(
            SX + SW / 2,
            yJoin,
            CX + 1,
            yJoin
          );

          ah(
            CX + 1,
            yJoin,
            "l"
          );
        }
      }

      else {
        pill(
          SX,
          it.cy - 15,
          100,
          30,
          "#dcefe9"
        );

        txt(
          "Selesai",
          SX + 50,
          it.cy
        );

        ln(
          rx,
          it.cy,
          SX,
          it.cy
        );

        ah(
          SX,
          it.cy,
          "r"
        );

        lab(
          "Tidak",
          (rx + SX) / 2,
          it.cy - 8
        );
      }
    }
  });

  return {
    canvas: c,
    w: LW,
    h: H
  };
}

function updPrev() {
  const st = stepsOf();
  const p = $("autoprev");

  if (!st.length) {
    p.style.display = "none";
    return;
  }

  p.src =
    flowCanvas(
      st,
      $("term").checked
    ).canvas.toDataURL("image/png");

  p.style.display = "block";
}

$("prosedur").addEventListener(
  "input",
  () => {
    if (baganMode === "auto")
      updPrev();
  }
);

$("term").onchange = updPrev;

/* ---------- pengolah XML docx ---------- */
const pText = p =>
  [...p.getElementsByTagNameNS(W, "t")]
    .map(t => t.textContent)
    .join("");

const kids = (n, name) =>
  [...n.childNodes]
    .filter(c => c.localName === name);

const lines = s =>
  s.split(/\r?\n/)
    .map(x => x.trim())
    .filter(Boolean);

const PPR_AFTER_IND = [
  "contextualSpacing",
  "mirrorIndents",
  "suppressOverlap",
  "jc",
  "textDirection",
  "textAlignment",
  "textboxTightWrap",
  "outlineLvl",
  "divId",
  "cnfStyle",
  "rPr",
  "sectPr",
  "pPrChange"
];

function baseRPr(p) {
  let best = null;
  let len = -1;

  for (
    const r of p.getElementsByTagNameNS(W, "r")
  ) {
    const l =
      [...r.getElementsByTagNameNS(W, "t")]
        .reduce(
          (a, t) =>
            a + t.textContent.length,
          0
        );

    if (l > len) {
      len = l;
      best = r;
    }
  }

  return best
    ? kids(best, "rPr")[0]?.cloneNode(true)
    : null;
}

function mkRun(
  doc,
  rPr,
  text,
  tab
) {
  const r =
    doc.createElementNS(W, "w:r");

  if (rPr)
    r.appendChild(
      rPr.cloneNode(true)
    );

  if (tab)
    r.appendChild(
      doc.createElementNS(W, "w:tab")
    );

  if (text != null) {
    const t =
      doc.createElementNS(W, "w:t");

    t.setAttribute(
      "xml:space",
      "preserve"
    );

    t.textContent = text;

    r.appendChild(t);
  }

  return r;
}

function mkPara(
  doc,
  tmpl,
  rPr,
  text,
  num
) {
  const np = tmpl.cloneNode(true);

  [...np.attributes].forEach(a => {
    if (a.name.startsWith("w14:"))
      np.removeAttribute(a.name);
  });

  [...np.childNodes].forEach(c => {
    if (c.localName !== "pPr")
      np.removeChild(c);
  });

  if (num != null) {
    let pPr = kids(np, "pPr")[0];

    if (!pPr) {
      pPr =
        doc.createElementNS(W, "w:pPr");

      np.insertBefore(
        pPr,
        np.firstChild
      );
    }

    kids(pPr, "ind").forEach(
      n => pPr.removeChild(n)
    );

    const ind =
      doc.createElementNS(W, "w:ind");

    ind.setAttribute(
      "w:left",
      "360"
    );

    ind.setAttribute(
      "w:hanging",
      "360"
    );

    const ref =
      [...pPr.childNodes].find(
        c =>
          PPR_AFTER_IND.includes(
            c.localName
          )
      );

    pPr.insertBefore(
      ind,
      ref || null
    );

    np.appendChild(
      mkRun(
        doc,
        rPr,
        num + "."
      )
    );

    np.appendChild(
      mkRun(
        doc,
        rPr,
        text,
        true
      )
    );
  }

  else {
    np.appendChild(
      mkRun(
        doc,
        rPr,
        text
      )
    );
  }

  return np;
}

function fillLines(
  doc,
  p,
  arr,
  numbered
) {
  if (!arr.length) {
    arr = ["-"];
    numbered = false;
  }

  const rPr = baseRPr(p);
  const parent = p.parentNode;

  arr.forEach((ln, i) => {
    const t = numbered
      ? ln.replace(
          /^\d+[.)]\s*/,
          ""
        )
      : ln;

    const np =
      mkPara(
        doc,
        p,
        rPr,
        t,
        numbered
          ? i + 1
          : null
      );

    if (i < arr.length - 1) {
      const pPr =
        kids(np, "pPr")[0];

      const sp =
        pPr &&
        kids(
          pPr,
          "spacing"
        )[0];

      if (
        sp &&
        sp.hasAttributeNS(
          W,
          "after"
        )
      ) {
        sp.setAttributeNS(
          W,
          "w:after",
          "0"
        );
      }
    }

    parent.insertBefore(
      np,
      p
    );
  });

  parent.removeChild(p);
}

function cellWidthEmu(p) {
  for (
    let n = p.parentNode;
    n;
    n = n.parentNode
  ) {
    if (n.localName === "tc") {
      const w =
        kids(n, "tcPr")[0] &&
        kids(
          kids(n, "tcPr")[0],
          "tcW"
        )[0];

      const dxa = w
        ? parseInt(
            w.getAttributeNS(
              W,
              "w"
            )
          )
        : 0;

      if (dxa > 800)
        return Math.round(
          (dxa - 240) /
          1440 *
          914400
        );
    }
  }

  return Math.round(
    5.5 * 914400
  );
}

function fillImage(
  doc,
  p,
  img,
  maxW
) {
  let cx =
    Math.round(
      img.w * 9525
    );

  let cy =
    Math.round(
      img.h * 9525
    );

  const k =
    Math.min(
      1,
      maxW / cx,
      (6 * 914400) / cy
    );

  cx =
    Math.round(cx * k);

  cy =
    Math.round(cy * k);

  const frag = `
<w:drawing xmlns:w="${W}"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">

<wp:inline distT="0" distB="0" distL="0" distR="0">

<wp:extent cx="${cx}" cy="${cy}"/>

<wp:docPr
 id="9001"
 name="Bagan alir"/>

<wp:cNvGraphicFramePr>
<a:graphicFrameLocks
 noChangeAspect="1"/>
</wp:cNvGraphicFramePr>

<a:graphic>
<a:graphicData
 uri="http://schemas.openxmlformats.org/drawingml/2006/picture">

<pic:pic>

<pic:nvPicPr>
<pic:cNvPr
 id="9001"
 name="bagan_alir.png"/>
<pic:cNvPicPr/>
</pic:nvPicPr>

<pic:blipFill>
<a:blip
 r:embed="rIdBaganAlir"/>
<a:stretch>
<a:fillRect/>
</a:stretch>
</pic:blipFill>

<pic:spPr>
<a:xfrm>
<a:off x="0" y="0"/>
<a:ext cx="${cx}" cy="${cy}"/>
</a:xfrm>

<a:prstGeom prst="rect">
<a:avLst/>
</a:prstGeom>

</pic:spPr>

</pic:pic>

</a:graphicData>
</a:graphic>

</wp:inline>
</w:drawing>`;

  const d =
    doc.importNode(
      new DOMParser()
        .parseFromString(
          frag,
          "application/xml"
        )
        .documentElement,
      true
    );

  const np =
    mkPara(
      doc,
      p,
      null,
      null,
      null
    );

  const r =
    doc.createElementNS(
      W,
      "w:r"
    );

  r.appendChild(d);
  np.appendChild(r);

  p.parentNode.replaceChild(
    np,
    p
  );
}

async function buildDocx(v) {
  const zip =
    await loadTemplateZip();

  const doc =
    new DOMParser().parseFromString(
      await zip
        .file("word/document.xml")
        .async("string"),
      "application/xml"
    );

  let usedImg = false;

  for (
    const p of [
      ...doc.getElementsByTagNameNS(
        W,
        "p"
      )
    ]
  ) {
    const m =
      pText(p)
        .replace(/\$+/g, "$")
        .trim()
        .match(/^\$(.+)\$$/);

    if (!m) continue;

    const key =
      m[1]
        .trim()
        .toLowerCase();

    if (
      key === "unit" ||
      key === "judul spo"
    ) {
      fillLines(
        doc,
        p,
        [
          v[
            key === "unit"
              ? "unit"
              : "judul"
          ]
        ],
        false
      );

      continue;
    }

    if (
      [
        "pengertian",
        "tujuan",
        "kebijakan",
        "prosedur"
      ].includes(key)
    ) {
      fillLines(
        doc,
        p,
        lines(v[key]),
        v.n[key]
      );

      continue;
    }

    if (key === "bagan alir") {
      if (v.img) {
        fillImage(
          doc,
          p,
          v.img,
          cellWidthEmu(p)
        );

        usedImg = true;
      }

      else {
        fillLines(
          doc,
          p,
          lines(v.bagan),
          false
        );
      }

      continue;
    }

    const li =
      key.match(
        /^(unit|dokumen) terkait (\d+)$/
      );

    if (li) {
      const arr =
        li[1] === "unit"
          ? v.units
          : v.docs;

      if (li[2] === "1")
        fillLines(
          doc,
          p,
          arr,
          false
        );
      else
        p.parentNode.removeChild(p);
    }
  }

  const out =
    new XMLSerializer()
      .serializeToString(doc)
      .replace(
        /^<\?xml[^>]*\?>\s*/,
        ""
      );

  zip.file(
    "word/document.xml",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    out
  );

  if (usedImg) {
    zip.file(
      "word/media/bagan_alir.png",
      v.img.buf
    );

    const rp =
      "word/_rels/document.xml.rels";

    let rels =
      await zip
        .file(rp)
        .async("string");

    rels =
      rels.replace(
        "</Relationships>",
        '<Relationship Id="rIdBaganAlir" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/bagan_alir.png"/></Relationships>'
      );

    zip.file(
      rp,
      rels
    );

    let ct =
      await zip
        .file("[Content_Types].xml")
        .async("string");

    if (!/Extension="png"/i.test(ct)) {
      ct =
        ct.replace(
          "<Override",
          '<Default Extension="png" ContentType="image/png"/><Override'
        );
    }

    zip.file(
      "[Content_Types].xml",
      ct
    );
  }

  return zip.generateAsync({
    type: "blob",
    compression: "DEFLATE",
    mimeType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  });
}

/* ---------- unduh ---------- */
function msg(t, err) {
  $("msg").textContent = t;
  $("msg").className =
    err ? "err" : "";
}

$("go").onclick = async () => {
  const unit =
    $("unit").value.trim();

  const judul =
    $("judul").value.trim();

  if (!unit || !judul) {
    msg(
      "Unit dan Judul SPO wajib diisi.",
      true
    );

    (
      unit
        ? $("judul")
        : $("unit")
    ).focus();

    return;
  }

  if (
    typeof JSZip === "undefined"
  ) {
    msg(
      "Pustaka JSZip gagal dimuat. Periksa koneksi internet.",
      true
    );

    return;
  }

  $("go").disabled = true;

  msg(
    "Membuat dokumen…"
  );

  try {
    const v = {
      unit,
      judul,
      n: {},
      units: listVals("unitList"),
      docs: listVals("docList"),
      bagan: $("bagan").value
    };

    [
      "pengertian",
      "tujuan",
      "kebijakan",
      "prosedur"
    ].forEach(k => {
      v[k] = $(k).value;

      v.n[k] =
        k === "pengertian"
          ? false
          : $("n_" + k).checked;
    });

    v.prosedur =
      lines(v.prosedur)
        .map(l =>
          dispLine(
            parseLine(l)
          )
        )
        .join("\n");

    v.img = null;

    if (
      baganMode === "img" &&
      baganFile
    ) {
      v.img =
        await toPng(
          baganFile
        );
    }

    if (
      baganMode === "auto" &&
      stepsOf().length
    ) {
      const f =
        flowCanvas(
          stepsOf(),
          $("term").checked
        );

      const blob =
        await new Promise(r =>
          f.canvas.toBlob(
            r,
            "image/png"
          )
        );

      v.img = {
        buf:
          await blob.arrayBuffer(),
        w: f.w,
        h: f.h
      };
    }

    if (
      !v.img &&
      baganMode !== "txt"
    ) {
      v.bagan = "";
    }

    const blob =
      await buildDocx(v);

    const a =
      document.createElement("a");

    a.href =
      URL.createObjectURL(
        blob
      );

    a.download =
      (
        "SPO - " +
        judul
      )
        .replace(
          /[\\/:*?"<>|]+/g,
          " "
        )
        .replace(
          /\s+/g,
          " "
        )
        .trim() +
      ".docx";

    document.body.appendChild(a);
    a.click();
    a.remove();

    setTimeout(
      () =>
        URL.revokeObjectURL(
          a.href
        ),
      4000
    );

    msg(
      "Selesai – file terunduh."
    );
  }

  catch (e) {
    console.error(e);

    msg(
      "Gagal membuat dokumen: " +
      e.message,
      true
    );
  }

  $("go").disabled = false;
};

load();
updPrev();
