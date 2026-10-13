const TEMPLATE_URL = "assets/template-spo.docx";

async function loadTemplateZip() {
  const response = await fetch(TEMPLATE_URL);
  if (!response.ok) throw new Error(`Template SPO tidak dapat dimuat (${response.status}).`);
  return JSZip.loadAsync(await response.arrayBuffer());
}

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const $ = id => document.getElementById(id);
const TEXT_IDS = ["unit","judul","pengertian","tujuan","nomorsk","jenissk","prosedur","bagan"];
const NUM_IDS = ["n_tujuan","n_prosedur","ringkas"];
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

/* ---------- diagram alir otomatis (mesin: js/flow.js) ---------- */
const chartOf = () =>
  FLOW.chart($("prosedur").value, $("ringkas").checked, +$("maxbox").value || 8);

function updPrev() {
  const c = chartOf(), p = $("autoprev"), w = $("flowwarn");
  w.innerHTML = "";
  c.warn.forEach(t => { const li = document.createElement("li"); li.textContent = t; w.append(li); });
  w.hidden = !c.warn.length;

  if (!c.nodes.length) { p.style.display = "none"; return; }
  p.src = FLOW.render(c.nodes, $("term").checked).canvas.toDataURL("image/png");
  p.style.display = "block";
}

$("prosedur").addEventListener("input", () => { if (baganMode === "auto") updPrev(); });
$("term").onchange = updPrev;
$("ringkas").onchange = updPrev;
$("maxbox").onchange = updPrev;

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

function setInd(doc, np, left, hanging) {
  let pPr = kids(np, "pPr")[0];
  if (!pPr) { pPr = doc.createElementNS(W, "w:pPr"); np.insertBefore(pPr, np.firstChild); }
  kids(pPr, "ind").forEach(n => pPr.removeChild(n));
  const ind = doc.createElementNS(W, "w:ind");
  ind.setAttribute("w:left", left);
  if (hanging) ind.setAttribute("w:hanging", hanging);
  const ref = [...pPr.childNodes].find(c => PPR_AFTER_IND.includes(c.localName));
  pPr.insertBefore(ind, ref || null);
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
    setInd(doc, np, "360", "360");

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

  let n = 0;

  arr.forEach((ln0, i) => {
    const sub = ln0.charAt(0) === "\u0001";
    const ln = sub ? ln0.slice(1) : ln0;
    const t = numbered && !sub ? ln.replace(/^\d+[.)]\s*/, "") : ln;

    const np = mkPara(doc, p, rPr, sub ? "– " + t : t, numbered && !sub ? ++n : null);

    if (sub) setInd(doc, np, "720", "0");

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

/* Ganti placeholder inline ($Nomor SK$, $Jenis SK$) di dalam satu paragraf.
   Teks digabung dari semua <w:t> lalu ditaruh di <w:t> pertama, sehingga
   placeholder yang terpecah antar-run tetap terganti. */
function replaceInline(p, map) {
  const ts = [...p.getElementsByTagNameNS(W, "t")];
  if (!ts.length) return false;
  const full = ts.map(t => t.textContent).join("");
  let hit = false;
  const out = full.replace(/\$([^$]+)\$/g, (m, k) => {
    const key = k.trim().toLowerCase();
    if (key in map) { hit = true; return map[key] || "-"; }
    return m;
  });
  if (!hit) return false;
  ts[0].textContent = out;
  ts[0].setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve");
  ts.slice(1).forEach(t => (t.textContent = ""));
  return true;
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

  /* Template menandai bagian yang berubah dengan stabilo kuning;
     hapus agar dokumen hasil generate bersih. */
  [...doc.getElementsByTagNameNS(W, "highlight")].forEach(h =>
    h.parentNode.removeChild(h)
  );

  for (
    const p of [
      ...doc.getElementsByTagNameNS(
        W,
        "p"
      )
    ]
  ) {
    if (
      replaceInline(p, {
        "nomor sk": v.nomorsk,
        "jenis sk": v.jenissk
      })
    ) continue;

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
        /^(unit|dokumen) terkait(?: (\d+))?$/
      );

    if (li) {
      const arr =
        li[1] === "unit"
          ? v.units
          : v.docs;

      if (!li[2] || li[2] === "1")
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

    v.nomorsk = $("nomorsk").value.trim();
    v.jenissk = $("jenissk").value.trim();

    [
      "pengertian",
      "tujuan",
      "prosedur"
    ].forEach(k => {
      v[k] = $(k).value;

      v.n[k] =
        k === "pengertian"
          ? false
          : $("n_" + k).checked;
    });

    v.prosedur = FLOW.docLines(v.prosedur).join("\n");

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

    if (baganMode === "auto") {
      const c = chartOf();

      if (c.nodes.length) {
        if (c.warn.length &&
          !confirm("Bagan alir punya " + c.warn.length + " peringatan:\n\n- " +
            c.warn.join("\n- ") + "\n\nTetap unduh?")) {
          msg("Dibatalkan. Perbaiki prosedur lalu coba lagi.", true);
          $("go").disabled = false;
          return;
        }

        const f = FLOW.render(c.nodes, $("term").checked);
        const blob = await new Promise(r => f.canvas.toBlob(r, "image/png"));
        v.img = { buf: await blob.arrayBuffer(), w: f.w, h: f.h };
      }
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
