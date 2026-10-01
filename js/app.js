const TEMPLATE_URL = "assets/template-spo.docx";

async function loadTemplateZip() {
  const response = await fetch(TEMPLATE_URL);

  if (!response.ok) {
    throw new Error(
      `Template SPO tidak dapat dimuat (${response.status}).`
    );
  }

  return JSZip.loadAsync(
    await response.arrayBuffer()
  );
}


/* =========================================================
   KONFIGURASI
========================================================= */

const W =
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const $ = id =>
  document.getElementById(id);

const TEXT_IDS = [
  "unit",
  "judul",
  "pengertian",
  "tujuan",
  "kebijakan",
  "prosedur",
  "bagan"
];

const NUM_IDS = [
  "n_tujuan",
  "n_kebijakan",
  "n_prosedur"
];

let baganMode = "auto";
let baganFile = null;


/* =========================================================
   KAPITAL OTOMATIS UNIT & JUDUL
========================================================= */

function upperField(id) {

  const el = $(id);

  if (!el) return;

  el.value =
    el.value.toLocaleUpperCase("id-ID");
}


["unit", "judul"].forEach(id => {

  $(id).addEventListener(
    "input",
    () => {

      const el = $(id);

      const pos =
        el.selectionStart;

      upperField(id);

      try {

        el.setSelectionRange(
          pos,
          pos
        );

      } catch (e) {}

      save();

    }
  );

});


/* =========================================================
   DAFTAR DINAMIS
========================================================= */

function addRow(
  listId,
  value = "",
  ph = ""
) {

  const r =
    document.createElement("div");

  r.className = "row";


  const i =
    document.createElement("input");

  i.type = "text";

  i.value = value;

  i.placeholder = ph;


  const x =
    document.createElement("button");

  x.type = "button";

  x.className = "x";

  x.textContent = "✕";

  x.title = "Hapus";


  x.onclick = () => {

    if (
      $(listId).children.length > 1
    ) {

      r.remove();

    } else {

      i.value = "";

    }

    save();

  };


  i.oninput = save;


  r.append(
    i,
    x
  );


  $(listId).append(r);
}


const listVals = id =>
  [
    ...$(id)
      .querySelectorAll("input")
  ]
    .map(i => i.value.trim())
    .filter(Boolean);


$("addUnit").onclick =
  () =>
    addRow(
      "unitList",
      "",
      "mis. Instalasi Rawat Inap"
    );


$("addDoc").onclick =
  () =>
    addRow(
      "docList",
      "",
      "mis. Formulir Serah Terima Pasien"
    );


/* =========================================================
   SIMPAN DRAFT
========================================================= */

function save() {

  try {

    const d = {
      t: {},
      n: {},
      u: listVals("unitList"),
      d: listVals("docList")
    };


    TEXT_IDS.forEach(
      k =>
        d.t[k] =
          $(k).value
    );


    NUM_IDS.forEach(
      k =>
        d.n[k] =
          $(k).checked
    );


    localStorage.setItem(
      "spo_draft",
      JSON.stringify(d)
    );

  } catch (e) {}

}


/* =========================================================
   LOAD DRAFT
========================================================= */

function load() {

  let d = null;


  try {

    d =
      JSON.parse(
        localStorage.getItem(
          "spo_draft"
        )
      );

  } catch (e) {}


  if (d) {

    TEXT_IDS.forEach(
      k =>
        $(k).value =
          d.t?.[k] || ""
    );


    NUM_IDS.forEach(
      k => {

        if (
          d.n &&
          k in d.n
        ) {

          $(k).checked =
            d.n[k];

        }

      }
    );


    /*
     * Pastikan data lama
     * juga menjadi kapital.
     */
    upperField("unit");
    upperField("judul");


    (
      d.u?.length
        ? d.u
        : [""]
    ).forEach(
      v =>
        addRow(
          "unitList",
          v,
          "mis. Instalasi Rawat Inap"
        )
    );


    (
      d.d?.length
        ? d.d
        : [""]
    ).forEach(
      v =>
        addRow(
          "docList",
          v,
          "mis. Formulir Serah Terima Pasien"
        )
    );

  }

  else {

    addRow("unitList");

    addRow("docList");

  }

}


/* =========================================================
   AUTO SAVE
========================================================= */

[
  ...TEXT_IDS,
  ...NUM_IDS
].forEach(
  k =>
    $(k).addEventListener(
      "input",
      save
    )
);


/* =========================================================
   RESET
========================================================= */

$("reset").onclick =
  () => {

    if (
      !confirm(
        "Kosongkan semua isian?"
      )
    ) return;


    TEXT_IDS.forEach(
      k =>
        $(k).value = ""
    );


    $("unitList").innerHTML =
      "";

    $("docList").innerHTML =
      "";


    addRow("unitList");

    addRow("docList");


    baganFile = null;

    $("file").value = "";

    $("imgprev").style.display =
      "none";


    /*
     * Kembali ke mode
     * bagan otomatis.
     */
    setMode("auto");


    save();

    updPrev();

  };


/* =========================================================
   BAGAN ALIR - MODE
========================================================= */

function setMode(m) {

  baganMode = m;


  $("tabAuto")
    .setAttribute(
      "aria-pressed",
      m === "auto"
    );


  $("tabImg")
    .setAttribute(
      "aria-pressed",
      m === "img"
    );


  $("tabTxt")
    .setAttribute(
      "aria-pressed",
      m === "txt"
    );


  $("paneAuto").hidden =
    m !== "auto";


  $("paneImg").hidden =
    m !== "img";


  $("paneTxt").hidden =
    m !== "txt";


  if (m === "auto") {

    updPrev();

  }

}


$("tabAuto").onclick =
  () =>
    setMode("auto");


$("tabImg").onclick =
  () =>
    setMode("img");


$("tabTxt").onclick =
  () =>
    setMode("txt");


/* =========================================================
   UPLOAD BAGAN
========================================================= */

$("file").onchange =
  e => {

    baganFile =
      e.target.files[0] ||
      null;


    const p =
      $("imgprev");


    if (baganFile) {

      p.src =
        URL.createObjectURL(
          baganFile
        );

      p.style.display =
        "block";

    }

    else {

      p.style.display =
        "none";

    }

  };


/* =========================================================
   KONVERSI GAMBAR
========================================================= */

async function toPng(file) {

  const bmp =
    await createImageBitmap(file);


  const s =
    Math.min(
      1,
      1800 / bmp.width
    );


  const c =
    document.createElement(
      "canvas"
    );


  c.width =
    Math.round(
      bmp.width * s
    );


  c.height =
    Math.round(
      bmp.height * s
    );


  const g =
    c.getContext("2d");


  g.fillStyle =
    "#fff";


  g.fillRect(
    0,
    0,
    c.width,
    c.height
  );


  g.drawImage(
    bmp,
    0,
    0,
    c.width,
    c.height
  );


  const blob =
    await new Promise(
      r =>
        c.toBlob(
          r,
          "image/png"
        )
    );


  return {

    buf:
      await blob.arrayBuffer(),

    w:
      c.width,

    h:
      c.height

  };

}


/* =========================================================
   PARSE PROSEDUR
========================================================= */

function parseLine(l) {

  let m =
    l.match(
      /^(\d+[.)]\s*)?(.+?)\?\s*\|\s*tidak\s*:\s*(.+)$/i
    );


  if (m) {

    return {

      pre:
        m[1] || "",

      text:
        m[2].trim(),

      q:true,

      act:
        m[3].trim()

    };

  }


  m =
    l.match(
      /^(\d+[.)]\s*)?(.+\?)$/
    );


  if (m) {

    return {

      pre:
        m[1] || "",

      text:
        m[2].trim(),

      q:true,

      act:""

    };

  }


  m =
    l.match(
      /^(\d+[.)]\s*)?(.+)$/
    );


  return {

    pre:
      m[1] || "",

    text:
      m[2].trim(),

    q:false,

    act:""

  };

}


const stepsOf =
  () =>
    lines(
      $("prosedur").value
    )
      .map(parseLine)
      .filter(
        s => s.text
      );


/* =========================================================
   KLASIFIKASI JALUR
========================================================= */

function classAct(a) {

  if (
    !a ||
    /^(selesai|berhenti|stop)\.?$/i.test(a)
  ) {

    return {
      k:"end"
    };

  }


  const m =
    a.match(
      /^kembali\s+ke\s+(?:langkah\s+)?(\d+)\.?$/i
    );


  return m

    ? {
        k:"loop",
        n:+m[1]
      }

    : {
        k:"box",
        t:a
      };

}


/* =========================================================
   TAMPILAN PROSEDUR DI DOCX
========================================================= */

function dispLine(s) {

  if (!s.act)
    return s.pre + s.text;


  const c =
    classAct(s.act);


  return (
    s.pre +
    s.text +
    " Jika tidak: " +
    (
      c.k === "end"
        ? "selesai."
        : c.k === "loop"
          ? "kembali ke langkah " +
            c.n +
            "."
          : s.act
    )
  );

}


/* =========================================================
   RINGKAS TEKS BAGAN
========================================================= */

function compactFlowText(
  text,
  max = 68
) {

  let t =
    String(text || "")
      .replace(
        /\s+/g,
        " "
      )
      .trim();


  /*
   * Hilangkan nomor
   * dari tampilan bagan.
   */
  t =
    t.replace(
      /^\d+[.)]\s*/,
      ""
    );


  /*
   * Kata pembuka yang
   * tidak terlalu penting
   * untuk bagan.
   */
  const replacements = [

    [
      /^kemudian\s+/i,
      ""
    ],

    [
      /^selanjutnya\s+/i,
      ""
    ],

    [
      /^setelah itu\s+/i,
      ""
    ],

    [
      /^lalu\s+/i,
      ""
    ],

    [
      /^selanjutnya lakukan\s+/i,
      ""
    ],

    [
      /\bmelakukan pemeriksaan\b/gi,
      "periksa"
    ],

    [
      /\bmelakukan pengecekan\b/gi,
      "cek"
    ],

    [
      /\bdilakukan pemeriksaan\b/gi,
      "periksa"
    ],

    [
      /\bmemastikan bahwa\b/gi,
      "pastikan"
    ],

    [
      /\bsegera dilakukan\b/gi,
      "segera"
    ],

    [
      /\bkemudian dilakukan\b/gi,
      "lakukan"
    ],

    [
      /\bdan kemudian\b/gi,
      "lalu"
    ],

    [
      /\bselanjutnya\b/gi,
      ""
    ]

  ];


  replacements.forEach(
    ([pattern, replacement]) => {

      t =
        t.replace(
          pattern,
          replacement
        );

    }
  );


  t =
    t.replace(
      /\s+/g,
      " "
    )
    .trim();


  /*
   * Batasi panjang teks
   * agar kotak bagan tidak
   * terlalu besar.
   */
  if (
    t.length > max
  ) {

    const cut =
      t.slice(
        0,
        max
      );


    const lastSpace =
      cut.lastIndexOf(
        " "
      );


    t =
      (
        lastSpace > 30
          ? cut.slice(
              0,
              lastSpace
            )
          : cut
      ).trim() +
      "…";

  }


  return t;

}


/* =========================================================
   CANVAS BAGAN ALIR
========================================================= */

function flowCanvas(
  steps,
  withTerm
) {

  const dec =
    steps.some(
      s => s.q
    );


  const LW = 560;

  const FS = 15;

  const LH = 19;

  const PV = 10;

  const M = 12;

  const FONT =
    "Arial, Helvetica, sans-serif";

  const F =
    FS +
    "px " +
    FONT;


  const BW =
    dec
      ? 300
      : 460;


  const CX =
    dec
      ? 185
      : 280;


  const TW = 150;

  const SX = 395;

  const SW = 155;


  const m =
    document
      .createElement("canvas")
      .getContext("2d");


  m.font = F;


  /* ---------- wrap text ---------- */

  const wrap =
    (
      txt,
      max
    ) => {

      const out = [];

      let cur = "";


      for (
        const w
        of txt.split(/\s+/)
      ) {

        const t =
          cur
            ? cur +
              " " +
              w
            : w;


        if (
          cur &&
          m.measureText(t).width >
            max
        ) {

          out.push(cur);

          cur = w;

        }

        else {

          cur = t;

        }

      }


      if (cur)
        out.push(cur);


      return out;

    };


  const items = [];


  if (withTerm) {

    items.push({
      k:"t",
      text:"Mulai"
    });

  }


  /*
   * Prosedur diringkas hanya
   * untuk BAGAN ALIR.
   *
   * Prosedur asli tetap utuh
   * di dokumen.
   */
  steps.forEach(
    (s, i) => {

      const act =
        s.q
          ? classAct(s.act)
          : null;


      if (
        act &&
        act.k === "box"
      ) {

        act.t =
          compactFlowText(
            act.t,
            45
          );

      }


      items.push({

        k:
          s.q
            ? "d"
            : "p",

        text:
          compactFlowText(
            s.text,
            68
          ),

        n:
          i + 1,

        act

      });

    }
  );


  if (withTerm) {

    items.push({

      k:"t",

      text:"Selesai"

    });

  }


  /* ---------- ukuran setiap item ---------- */

  items.forEach(
    it => {

      if (
        it.k === "t"
      ) {

        it.w = TW;

        it.l = [
          it.text
        ];

        it.h = 34;

      }


      else if (
        it.k === "p"
      ) {

        it.w = BW;

        it.l =
          wrap(
            it.text,
            BW - 24
          );

        it.h =
          it.l.length *
            LH +
          PV * 2;

      }


      else {

        it.w = BW;

        it.l =
          wrap(
            it.text,
            BW * .56
          );

        it.h =
          Math.max(
            it.l.length *
              LH +
              40,
            70
          );


        /*
         * Validasi loop.
         */
        if (
          it.act.k === "loop" &&
          !items.some(
            t =>
              t.n ===
                it.act.n &&
              t.n <
                it.n
          )
        ) {

          it.act = {
            k:"end"
          };

        }


        if (
          it.act.k === "box"
        ) {

          it.sl =
            wrap(
              it.act.t,
              SW - 18
            );


          it.sh =
            it.sl.length *
              LH +
            PV * 2;


          it.h =
            Math.max(
              it.h,
              it.sh + 14
            );

        }

      }

    }
  );


  const gapOf =
    it =>
      it.k === "d"
        ? 32
        : 24;


  let y = M;


  items.forEach(
    it => {

      it.top = y;

      it.cy =
        y +
        it.h / 2;

      it.bot =
        y +
        it.h;

      y =
        it.bot +
        gapOf(it);

    }
  );


  const H =
    items[
      items.length - 1
    ].bot +
    M;


  const S = 2;


  const c =
    document.createElement(
      "canvas"
    );


  c.width =
    LW * S;

  c.height =
    Math.round(
      H * S
    );


  const g =
    c.getContext("2d");


  g.scale(
    S,
    S
  );


  g.fillStyle =
    "#fff";

  g.fillRect(
    0,
    0,
    LW,
    H
  );


  g.font = F;

  g.textBaseline =
    "middle";

  g.lineWidth = 1.4;

  g.strokeStyle =
    "#111";


  /* ---------- line ---------- */

  const ln =
    (
      x1,
      y1,
      x2,
      y2
    ) => {

      g.beginPath();

      g.moveTo(
        x1,
        y1
      );

      g.lineTo(
        x2,
        y2
      );

      g.stroke();

    };


  /* ---------- arrow ---------- */

  const ah =
    (
      x,
      y,
      d
    ) => {

      g.beginPath();

      g.moveTo(
        x,
        y
      );


      if (
        d === "d"
      ) {

        g.lineTo(
          x - 5,
          y - 10
        );

        g.lineTo(
          x + 5,
          y - 10
        );

      }


      else if (
        d === "l"
      ) {

        g.lineTo(
          x + 10,
          y - 5
        );

        g.lineTo(
          x + 10,
          y + 5
        );

      }


      else {

        g.lineTo(
          x - 10,
          y - 5
        );

        g.lineTo(
          x - 10,
          y + 5
        );

      }


      g.closePath();

      g.fillStyle =
        "#111";

      g.fill();

    };


  /* ---------- text ---------- */

  const txt =
    (
      t,
      x,
      y,
      al,
      font,
      col
    ) => {

      g.save();

      g.font =
        font || F;

      g.fillStyle =
        col || "#111";

      g.textAlign =
        al || "center";

      g.fillText(
        t,
        x,
        y
      );

      g.restore();

    };


  /* ---------- pill ---------- */

  const pill =
    (
      x,
      yy,
      w,
      h,
      fill
    ) => {

      g.fillStyle =
        fill;

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


  const lab =
    (
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
        "bold 12px " +
          FONT
      );


  let loops = 0;


  /* ---------- gambar item ---------- */

  items.forEach(
    (it, i) => {

      const last =
        i ===
        items.length - 1;


      const x0 =
        CX -
        it.w / 2;


      /* TERMINAL */

      if (
        it.k === "t"
      ) {

        pill(
          x0,
          it.top,
          it.w,
          it.h,
          "#dcefe9"
        );

      }


      /* PROSES */

      else if (
        it.k === "p"
      ) {

        g.fillStyle =
          "#f3f6f9";

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


      /* KEPUTUSAN */

      else {

        g.fillStyle =
          "#fff4d6";


        g.beginPath();


        g.moveTo(
          CX,
          it.top
        );


        g.lineTo(
          CX +
            it.w / 2,
          it.cy
        );


        g.lineTo(
          CX,
          it.bot
        );


        g.lineTo(
          x0,
          it.cy
        );


        g.closePath();

        g.fill();

        g.stroke();

      }


      /* ---------- teks ---------- */

      const y0 =
        it.cy -
        it.l.length *
          LH / 2 +
        LH / 2;


      it.l.forEach(
        (s, k) =>
          txt(
            s,
            CX,
            y0 +
              k * LH
          )
      );


      /* ---------- nomor ---------- */

      if (it.n) {

        txt(
          String(it.n),
          x0 - 6,
          it.cy,
          "right",
          "12px " +
            FONT,
          "#666"
        );

      }


      /* ---------- alur utama ---------- */

      if (!last) {

        const gp =
          gapOf(it);


        ln(
          CX,
          it.bot,
          CX,
          it.bot +
            gp -
            1
        );


        ah(
          CX,
          it.bot +
            gp,
          "d"
        );


        if (
          it.k === "d"
        ) {

          lab(
            "Ya",
            CX - 14,
            it.bot + 16
          );

        }

      }


      /* ---------- keputusan ---------- */

      if (
        it.k === "d"
      ) {

        const rx =
          CX +
          it.w / 2;


        const a =
          it.act;


        const yJoin =
          it.bot + 18;


        /* LOOP */

        if (
          a.k === "loop"
        ) {

          const tg =
            items.find(
              t =>
                t.n ===
                  a.n &&
                t.n <
                  it.n
            );


          if (tg) {

            const lx =
              350 +
              Math.min(
                loops++,
                3
              ) *
              12;


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

        }


        /* ACTION BOX */

        else if (
          a.k === "box"
        ) {

          const sy =
            it.cy -
            it.sh / 2;


          g.fillStyle =
            "#f3f6f9";


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
            it.sl.length *
              LH / 2 +
            LH / 2;


          it.sl.forEach(
            (s, k) =>
              txt(
                s,
                SX +
                  SW / 2,
                s0 +
                  k * LH
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
              SX +
                SW / 2,
              sy +
                it.sh,
              SX +
                SW / 2,
              yJoin
            );


            ln(
              SX +
                SW / 2,
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


        /* END */

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

    }
  );


  return {

    canvas:c,

    w:LW,

    h:H

  };

}


/* =========================================================
   PREVIEW BAGAN
========================================================= */

function updPrev() {

  const st =
    stepsOf();


  const p =
    $("autoprev");


  if (!st.length) {

    p.style.display =
      "none";

    return;

  }


  p.src =
    flowCanvas(
      st,
      $("term").checked
    )
      .canvas
      .toDataURL(
        "image/png"
      );


  p.style.display =
    "block";

}


$("prosedur").addEventListener(
  "input",
  () => {

    if (
      baganMode === "auto"
    ) {

      updPrev();

    }

  }
);


$("term").onchange =
  updPrev;


/* =========================================================
   PENGOLAH XML DOCX
========================================================= */

const pText =
  p =>
    [
      ...p.getElementsByTagNameNS(
        W,
        "t"
      )
    ]
      .map(
        t =>
          t.textContent
      )
      .join("");


const kids =
  (
    n,
    name
  ) =>
    [
      ...n.childNodes
    ]
      .filter(
        c =>
          c.localName ===
          name
      );


const lines =
  s =>
    String(s || "")
      .split(/\r?\n/)
      .map(
        x =>
          x.trim()
      )
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


/* =========================================================
   BASE RUN
========================================================= */

function baseRPr(p) {

  let best = null;

  let len = -1;


  for (
    const r of
      p.getElementsByTagNameNS(
        W,
        "r"
      )
  ) {

    const l =
      [
        ...r.getElementsByTagNameNS(
          W,
          "t"
        )
      ]
        .reduce(
          (
            a,
            t
          ) =>
            a +
            t.textContent.length,
          0
        );


    if (
      l > len
    ) {

      len = l;

      best = r;

    }

  }


  return best
    ? kids(
        best,
        "rPr"
      )[0]?.cloneNode(true)
    : null;

}


/* =========================================================
   RUN
========================================================= */

function mkRun(
  doc,
  rPr,
  text,
  tab
) {

  const r =
    doc.createElementNS(
      W,
      "w:r"
    );


  if (rPr) {

    r.appendChild(
      rPr.cloneNode(true)
    );

  }


  if (tab) {

    r.appendChild(
      doc.createElementNS(
        W,
        "w:tab"
      )
    );

  }


  if (text != null) {

    const t =
      doc.createElementNS(
        W,
        "w:t"
      );


    t.setAttribute(
      "xml:space",
      "preserve"
    );


    t.textContent =
      text;


    r.appendChild(t);

  }


  return r;

}


/* =========================================================
   PARAGRAPH
========================================================= */

function mkPara(
  doc,
  tmpl,
  rPr,
  text,
  num
) {

  const np =
    tmpl.cloneNode(true);


  [...np.attributes]
    .forEach(
      a => {

        if (
          a.name.startsWith(
            "w14:"
          )
        ) {

          np.removeAttribute(
            a.name
          );

        }

      }
    );


  [...np.childNodes]
    .forEach(
      c => {

        if (
          c.localName !==
          "pPr"
        ) {

          np.removeChild(c);

        }

      }
    );


  if (
    num != null
  ) {

    let pPr =
      kids(
        np,
        "pPr"
      )[0];


    if (!pPr) {

      pPr =
        doc.createElementNS(
          W,
          "w:pPr"
        );


      np.insertBefore(
        pPr,
        np.firstChild
      );

    }


    kids(
      pPr,
      "ind"
    )
      .forEach(
        n =>
          pPr.removeChild(n)
      );


    const ind =
      doc.createElementNS(
        W,
        "w:ind"
      );


    ind.setAttribute(
      "w:left",
      "360"
    );


    ind.setAttribute(
      "w:hanging",
      "360"
    );


    const ref =
      [
        ...pPr.childNodes
      ]
        .find(
          c =>
            PPR_AFTER_IND
              .includes(
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


/* =========================================================
   FILL LINES
========================================================= */

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


  const rPr =
    baseRPr(p);


  const parent =
    p.parentNode;


  arr.forEach(
    (
      ln,
      i
    ) => {

      const t =
        numbered

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


      if (
        i <
        arr.length - 1
      ) {

        const pPr =
          kids(
            np,
            "pPr"
          )[0];


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

    }
  );


  parent.removeChild(p);

}


/* =========================================================
   CELL WIDTH
========================================================= */

function cellWidthEmu(p) {

  for (
    let n =
      p.parentNode;
    n;
    n =
      n.parentNode
  ) {

    if (
      n.localName ===
      "tc"
    ) {

      const w =
        kids(
          n,
          "tcPr"
        )[0] &&
        kids(
          kids(
            n,
            "tcPr"
          )[0],
          "tcW"
        )[0];


      const dxa =
        w
          ? parseInt(
              w.getAttributeNS(
                W,
                "w"
              )
            )
          : 0;


      if (
        dxa > 800
      ) {

        return Math.round(
          (
            dxa - 240
          ) /
          1440 *
          914400
        );

      }

    }

  }


  return Math.round(
    5.5 *
    914400
  );

}


/* =========================================================
   IMAGE DOCX
========================================================= */

function fillImage(
  doc,
  p,
  img,
  maxW
) {

  let cx =
    Math.round(
      img.w *
      9525
    );


  let cy =
    Math.round(
      img.h *
      9525
    );


  const k =
    Math.min(
      1,
      maxW / cx,
      (
        6 *
        914400
      ) /
      cy
    );


  cx =
    Math.round(
      cx * k
    );


  cy =
    Math.round(
      cy * k
    );


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

<a:off
 x="0"
 y="0"/>

<a:ext
 cx="${cx}"
 cy="${cy}"/>

</a:xfrm>

<a:prstGeom
 prst="rect">

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


/* =========================================================
   BUILD DOCX
========================================================= */

async function buildDocx(v) {

  const zip =
    await loadTemplateZip();


  const doc =
    new DOMParser()
      .parseFromString(
        await zip
          .file(
            "word/document.xml"
          )
          .async(
            "string"
          ),
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
        .replace(
          /\$+/g,
          "$"
        )
        .trim()
        .match(
          /^\$(.+)\$$/
        );


    if (!m)
      continue;


    const key =
      m[1]
        .trim()
        .toLowerCase();


    /* ---------- unit / judul ---------- */

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


    /* ---------- isi ---------- */

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
        lines(
          v[key]
        ),
        v.n[key]
      );


      continue;

    }


    /* ---------- bagan ---------- */

    if (
      key ===
      "bagan alir"
    ) {

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
          lines(
            v.bagan
          ),
          false
        );

      }


      continue;

    }


    /* ---------- unit/dokumen terkait ---------- */

    const li =
      key.match(
        /^(unit|dokumen) terkait (\d+)$/
      );


    if (li) {

      const arr =
        li[1] === "unit"
          ? v.units
          : v.docs;


      if (
        li[2] === "1"
      ) {

        fillLines(
          doc,
          p,
          arr,
          false
        );

      }

      else {

        p.parentNode
          .removeChild(p);

      }

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


  /* ---------- gambar ---------- */

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
        .file(
          "[Content_Types].xml"
        )
        .async("string");


    if (
      !/Extension="png"/i.test(
        ct
      )
    ) {

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

    type:"blob",

    compression:"DEFLATE",

    mimeType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

  });

}


/* =========================================================
   PESAN
========================================================= */

function msg(
  t,
  err
) {

  $("msg").textContent =
    t;


  $("msg").className =
    err
      ? "err"
      : "";

}


/* =========================================================
   DATA UNTUK DOCX & PDF
========================================================= */

async function collectData() {

  const unit =
    $("unit")
      .value
      .trim()
      .toLocaleUpperCase(
        "id-ID"
      );


  const judul =
    $("judul")
      .value
      .trim()
      .toLocaleUpperCase(
        "id-ID"
      );


  if (
    !unit ||
    !judul
  ) {

    msg(
      "Unit dan Judul SPO wajib diisi.",
      true
    );


    (
      unit
        ? $("judul")
        : $("unit")
    ).focus();


    return null;

  }


  const v = {

    unit,

    judul,

    n:{},

    units:
      listVals(
        "unitList"
      ),

    docs:
      listVals(
        "docList"
      ),

    bagan:
      $("bagan").value

  };


  [
    "pengertian",
    "tujuan",
    "kebijakan",
    "prosedur"
  ].forEach(
    k => {

      v[k] =
        $(k).value;


      v.n[k] =
        k === "pengertian"
          ? false
          : $("n_" + k)
              .checked;

    }
  );


  return v;

}


/* =========================================================
   SIAPKAN GAMBAR BAGAN
========================================================= */

async function prepareFlowImage(v) {

  v.img = null;


  /* ---------- gambar upload ---------- */

  if (
    baganMode === "img" &&
    baganFile
  ) {

    v.img =
      await toPng(
        baganFile
      );

    return;

  }


  /* ---------- otomatis ---------- */

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
      await new Promise(
        resolve =>
          f.canvas.toBlob(
            resolve,
            "image/png"
          )
      );


    v.img = {

      buf:
        await blob.arrayBuffer(),

      w:
        f.w,

      h:
        f.h

    };

  }


  if (
    !v.img &&
    baganMode !== "txt"
  ) {

    v.bagan = "";

  }

}


/* =========================================================
   UNDUH DOCX
========================================================= */

$("go").onclick =
  async () => {

    if (
      typeof JSZip ===
      "undefined"
    ) {

      msg(
        "Pustaka JSZip gagal dimuat. Periksa koneksi internet.",
        true
      );

      return;

    }


    $("go").disabled =
      true;


    msg(
      "Membuat dokumen Word…"
    );


    try {

      const v =
        await collectData();


      if (!v) {

        $("go").disabled =
          false;

        return;

      }


      /*
       * DOCX mempertahankan
       * teks prosedur lengkap.
       */
      v.prosedur =
        lines(
          v.prosedur
        )
          .map(
            l =>
              dispLine(
                parseLine(l)
              )
          )
          .join("\n");


      await prepareFlowImage(
        v
      );


      const blob =
        await buildDocx(
          v
        );


      const a =
        document.createElement(
          "a"
        );


      const url =
        URL.createObjectURL(
          blob
        );


      a.href =
        url;


      a.download =
        (
          "SPO - " +
          v.judul
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


      document.body.appendChild(
        a
      );


      a.click();

      a.remove();


      setTimeout(
        () =>
          URL.revokeObjectURL(
            url
          ),
        4000
      );


      msg(
        "Selesai – file Word terunduh."
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


    $("go").disabled =
      false;

  };


/* =========================================================
   PDF
========================================================= */

function pdfSafeName(
  name
) {

  return String(
    name || "SPO"
  )
    .replace(
      /[\\/:*?"<>|]+/g,
      " "
    )
    .replace(
      /\s+/g,
      " "
    )
    .trim();

}


/* =========================================================
   PDF TEXT
========================================================= */

function pdfLines(
  doc,
  text,
  x,
  y,
  width,
  options = {}
) {

  const fontSize =
    options.fontSize ||
    10;


  const lineHeight =
    options.lineHeight ||
    5.2;


  const bold =
    options.bold ||
    false;


  doc.setFont(
    "helvetica",
    bold
      ? "bold"
      : "normal"
  );


  doc.setFontSize(
    fontSize
  );


  const lines =
    doc.splitTextToSize(
      String(
        text || "-"
      ),
      width
    );


  lines.forEach(
    line => {

      if (
        y > 280
      ) {

        doc.addPage();

        y = 18;

      }


      doc.text(
        line,
        x,
        y
      );


      y +=
        lineHeight;

    }
  );


  return y;

}


/* =========================================================
   PDF SECTION
========================================================= */

function pdfSection(
  doc,
  title,
  y
) {

  if (
    y > 275
  ) {

    doc.addPage();

    y = 18;

  }


  doc.setFont(
    "helvetica",
    "bold"
  );


  doc.setFontSize(
    11
  );


  doc.setTextColor(
    0,
    63,
    125
  );


  doc.text(
    title.toUpperCase(),
    18,
    y
  );


  y += 6;


  doc.setDrawColor(
    7,
    148,
    71
  );


  doc.line(
    18,
    y,
    192,
    y
  );


  doc.setTextColor(
    16,
    34,
    56
  );


  return y + 6;

}


/* =========================================================
   PDF BUILD
========================================================= */

async function buildPdf(
  v
) {

  if (
    typeof window.jspdf ===
    "undefined" ||
    !window.jspdf.jsPDF
  ) {

    throw new Error(
      "Pustaka PDF gagal dimuat. Periksa koneksi internet."
    );

  }


  const {
    jsPDF
  } =
    window.jspdf;


  const doc =
    new jsPDF({

      orientation:
        "portrait",

      unit:
        "mm",

      format:
        "a4"

    });


  const pageWidth =
    210;


  const margin =
    18;


  const contentWidth =
    pageWidth -
    margin * 2;


  let y = 18;


  /* =====================================================
     HEADER
  ===================================================== */

  doc.setFont(
    "helvetica",
    "bold"
  );


  doc.setFontSize(
    15
  );


  doc.setTextColor(
    0,
    63,
    125
  );


  doc.text(
    "STANDAR PROSEDUR OPERASIONAL",
    pageWidth / 2,
    y,
    {
      align:"center"
    }
  );


  y += 8;


  doc.setFontSize(
    11
  );


  doc.text(
    v.judul,
    pageWidth / 2,
    y,
    {
      align:"center",
      maxWidth:
        contentWidth
    }
  );


  y += 8;


  doc.setDrawColor(
    0,
    63,
    125
  );


  doc.setLineWidth(
    .6
  );


  doc.line(
    margin,
    y,
    pageWidth -
      margin,
    y
  );


  y += 8;


  /* =====================================================
     IDENTITAS
  ===================================================== */

  doc.setFontSize(
    10
  );


  doc.setFont(
    "helvetica",
    "bold"
  );


  doc.text(
    "UNIT",
    margin,
    y
  );


  doc.setFont(
    "helvetica",
    "normal"
  );


  doc.text(
    ": " + v.unit,
    margin + 28,
    y
  );


  y += 6;


  doc.setFont(
    "helvetica",
    "bold"
  );


  doc.text(
    "JUDUL SPO",
    margin,
    y
  );


  doc.setFont(
    "helvetica",
    "normal"
  );


  const judulLines =
    doc.splitTextToSize(
      v.judul,
      contentWidth - 28
    );


  doc.text(
    ": " +
      judulLines[0],
    margin + 28,
    y
  );


  for (
    let i = 1;
    i <
    judulLines.length;
    i++
  ) {

    y += 5;


    doc.text(
      judulLines[i],
      margin + 30,
      y
    );

  }


  y += 9;


  /* =====================================================
     PENGERTIAN
  ===================================================== */

  y =
    pdfSection(
      doc,
      "Pengertian",
      y
    );


  y =
    pdfLines(
      doc,
      lines(
        v.pengertian
      ).join("\n") ||
        "-",
      margin,
      y,
      contentWidth
    );


  y += 5;


  /* =====================================================
     TUJUAN
  ===================================================== */

  y =
    pdfSection(
      doc,
      "Tujuan",
      y
    );


  const tujuan =
    lines(
      v.tujuan
    );


  tujuan.forEach(
    (
      item,
      i
    ) => {

      const prefix =
        v.n.tujuan
          ? `${i + 1}. `
          : "";


      y =
        pdfLines(
          doc,
          prefix +
            item.replace(
              /^\d+[.)]\s*/,
              ""
            ),
          margin,
          y,
          contentWidth
        );

    }
  );


  y += 5;


  /* =====================================================
     KEBIJAKAN
  ===================================================== */

  y =
    pdfSection(
      doc,
      "Kebijakan",
      y
    );


  const kebijakan =
    lines(
      v.kebijakan
    );


  kebijakan.forEach(
    (
      item,
      i
    ) => {

      const prefix =
        v.n.kebijakan
          ? `${i + 1}. `
          : "";


      y =
        pdfLines(
          doc,
          prefix +
            item.replace(
              /^\d+[.)]\s*/,
              ""
            ),
          margin,
          y,
          contentWidth
        );

    }
  );


  y += 5;


  /* =====================================================
     PROSEDUR
  ===================================================== */

  y =
    pdfSection(
      doc,
      "Prosedur / Langkah-langkah",
      y
    );


  const prosedur =
    lines(
      v.prosedur
    );


  prosedur.forEach(
    (
      item,
      i
    ) => {

      const clean =
        item.replace(
          /^\d+[.)]\s*/,
          ""
        );


      const prefix =
        v.n.prosedur
          ? `${i + 1}. `
          : "";


      y =
        pdfLines(
          doc,
          prefix +
            clean,
          margin,
          y,
          contentWidth
        );


      y += 1;

    }
  );


  /* =====================================================
     BAGAN ALIR
  ===================================================== */

  let flowImage =
    null;


  if (v.img) {

    const bytes =
      new Uint8Array(
        v.img.buf
      );


    let binary = "";


    const chunk =
      0x8000;


    for (
      let i = 0;
      i <
      bytes.length;
      i += chunk
    ) {

      binary +=
        String.fromCharCode(
          ...bytes.subarray(
            i,
            i + chunk
          )
        );

    }


    flowImage =
      "data:image/png;base64," +
      btoa(binary);

  }


  if (flowImage) {

    y += 5;


    y =
      pdfSection(
        doc,
        "Bagan Alir",
        y
      );


    const maxW =
      contentWidth;


    const maxH =
      115;


    let iw =
      v.img.w;


    let ih =
      v.img.h;


    const scale =
      Math.min(
        maxW / iw,
        maxH / ih
      );


    iw *= scale;

    ih *= scale;


    if (
      y + ih >
      280
    ) {

      doc.addPage();

      y = 18;

    }


    doc.addImage(
      flowImage,
      "PNG",
      (
        pageWidth -
        iw
      ) / 2,
      y,
      iw,
      ih
    );


    y +=
      ih + 8;

  }

  else if (
    v.bagan.trim()
  ) {

    y += 5;


    y =
      pdfSection(
        doc,
        "Bagan Alir",
        y
      );


    y =
      pdfLines(
        doc,
        lines(
          v.bagan
        ).join("\n"),
        margin,
        y,
        contentWidth
      );


    y += 5;

  }


  /* =====================================================
     UNIT TERKAIT
  ===================================================== */

  if (
    v.units.length
  ) {

    y =
      pdfSection(
        doc,
        "Unit Terkait",
        y
      );


    v.units.forEach(
      (
        item,
        i
      ) => {

        y =
          pdfLines(
            doc,
            `${i + 1}. ${item}`,
            margin,
            y,
            contentWidth
          );

      }
    );


    y += 5;

  }


  /* =====================================================
     DOKUMEN TERKAIT
  ===================================================== */

  if (
    v.docs.length
  ) {

    y =
      pdfSection(
        doc,
        "Dokumen Terkait",
        y
      );


    v.docs.forEach(
      (
        item,
        i
      ) => {

        y =
          pdfLines(
            doc,
            `${i + 1}. ${item}`,
            margin,
            y,
            contentWidth
          );

      }
    );

  }


  /* =====================================================
     FOOTER SEMUA HALAMAN
  ===================================================== */

  const totalPages =
    doc.internal
      .getNumberOfPages();


  for (
    let i = 1;
    i <= totalPages;
    i++
  ) {

    doc.setPage(i);


    doc.setDrawColor(
      210,
      220,
      230
    );


    doc.setLineWidth(
      .3
    );


    doc.line(
      margin,
      287,
      pageWidth -
        margin,
      287
    );


    doc.setFont(
      "helvetica",
      "normal"
    );


    doc.setFontSize(
      8
    );


    doc.setTextColor(
      100,
      114,
      135
    );


    doc.text(
      "Cincop by AmegoVerse",
      margin,
      293
    );


    doc.text(
      `Halaman ${i} dari ${totalPages}`,
      pageWidth -
        margin,
      293,
      {
        align:"right"
      }
    );

  }


  doc.setTextColor(
    16,
    34,
    56
  );


  return doc;

}


/* =========================================================
   UNDUH PDF
========================================================= */

$("pdf").onclick =
  async () => {

    if (
      typeof window.jspdf ===
      "undefined"
    ) {

      msg(
        "Pustaka PDF gagal dimuat. Periksa koneksi internet.",
        true
      );

      return;

    }


    $("pdf").disabled =
      true;


    msg(
      "Membuat PDF…"
    );


    try {

      const v =
        await collectData();


      if (!v) {

        $("pdf").disabled =
          false;

        return;

      }


      /*
       * Prosedur tetap lengkap.
       */
      v.prosedur =
        lines(
          v.prosedur
        ).join("\n");


      await prepareFlowImage(
        v
      );


      const pdf =
        await buildPdf(
          v
        );


      pdf.save(
        "SPO - " +
        pdfSafeName(
          v.judul
        ) +
        ".pdf"
      );


      msg(
        "Selesai – PDF berhasil dibuat."
      );

    }

    catch (e) {

      console.error(e);


      msg(
        "Gagal membuat PDF: " +
        e.message,
        true
      );

    }


    $("pdf").disabled =
      false;

  };


/* =========================================================
   INISIALISASI
========================================================= */

load();

updPrev();
