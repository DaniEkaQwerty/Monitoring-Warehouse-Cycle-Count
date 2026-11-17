/* ===================================================================
       KONFIGURASI INPUT OTOMATIS
       Ubah dua baris ini kalau nama folder / pola file di share berbeda.
       - TALLY_DIR_NAME : folder Tally Sheet di network share
       - TALLY_FILE_RE  : pola nama file Tally Sheet (Tally_Sheet_WE_... / WEEK ...)
       =================================================================== */
const TALLY_DIR_NAME = "TALLYSHEET";
const TALLY_FILE_RE = /^(?:tally[\s_-]*sheet|week)(?![a-z])/i;

const $ = (id) => document.getElementById(id);
const esc = (s) =>
  String(s == null ? "" : s).replace(
    /[&<>"]/g,
    (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c],
  );
const pad = (n) => String(n).padStart(2, "0");
const nf = (n) => Number(n).toLocaleString("id-ID");
const RM =
  window.matchMedia &&
  matchMedia("(prefers-reduced-motion: reduce)").matches;
const ymd = (d) =>
  d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate());
let tabIn = { set: new Set(), n: 0, err: "" };
let tallyWB = null,
  tallySheets = [],
  tallyByName = new Map(),
  files = [],
  fid = 0,
  res = null,
  tab = "all",
  q = "",
  sf = "all",
  autoMsg = "",
  manualSheet = "",
  lastTallyErr = "",
  scanCatalog = [],
  scanDir = null,
  scanBusy = false,
  scanSkipped = 0,
  filterPart = "",
  filterStatus = "",
  activePopCol = null;

function showErr(m) {
  const e = $("err");
  e.style.display = m ? "block" : "none";
  e.textContent = m || "";
}

/* ---------- animasi ---------- */
function reveal(el, on) {
  clearTimeout(el._t);
  if (on) {
    if (el.style.display !== "block") {
      el.style.display = "block";
      void el.offsetWidth;
    }
    requestAnimationFrame(() => el.classList.add("in"));
  } else {
    el.classList.remove("in");
    el._t = setTimeout(
      () => {
        el.style.display = "none";
      },
      RM ? 0 : 320,
    );
  }
}
function tween(el, to) {
  const from = el._v || 0;
  el._v = to;
  cancelAnimationFrame(el._r);
  if (RM || from === to) {
    el.textContent = nf(to);
    return;
  }
  const t0 = performance.now(),
    D = 550;
  const step = (n) => {
    const p = Math.min(1, (n - t0) / D),
      e = 1 - Math.pow(1 - p, 3);
    el.textContent = nf(Math.round(from + (to - from) * e));
    if (p < 1) el._r = requestAnimationFrame(step);
  };
  el._r = requestAnimationFrame(step);
}

/* ---------- baca file ---------- */
function parseTime(v) {
  if (v == null || v === "") return null;
  if (typeof v === "number")
    return Math.round((v - 25569) * 86400) * 1000;
  if (v instanceof Date)
    return Date.UTC(
      v.getFullYear(),
      v.getMonth(),
      v.getDate(),
      v.getHours(),
      v.getMinutes(),
      v.getSeconds(),
    );
  const s = String(v).trim();
  let m = s.match(
    /^(\d{4})-(\d{2})-(\d{2})[ T](\d{1,2}):(\d{2})(?::(\d{2}))?/,
  );
  if (m)
    return Date.UTC(+m[1], m[2] - 1, +m[3], +m[4], +m[5], +(m[6] || 0));
  m = s.match(
    /^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})[ T](\d{1,2}):(\d{2})(?::(\d{2}))?/,
  );
  if (m)
    return Date.UTC(+m[3], m[2] - 1, +m[1], +m[4], +m[5], +(m[6] || 0));
  return null;
}
const fmt = (ms) =>
  new Date(ms).toISOString().slice(0, 19).replace("T", " ");
const normPart = (v) =>
  String(v == null ? "" : v)
    .trim()
    .toUpperCase();
/* Exception nomor MD: abaikan 4 digit angka setelah spasi di akhir part number file Trx (contoh: "HLW03-JA11 0112" -> "HLW03-JA11") */
const normPartTrx = (v) => normPart(v).replace(/\s+\d{4}$/, "").trim();
/* THROW = dihitung sebagai OUT, tapi jumlahnya ditampilkan terpisah di keterangan */
const isThrow = (r) => /throw/i.test(r.ket || "");
/* OUT PREPARE = barang persiapan keluar; dibaca dan ditampilkan di detail tapi TIDAK dihitung sebagai OUT / tidak mengurangi Tally */
const isOutPrepare = (r) =>
  /out\s*prepare/i.test(r.loc || "") ||
  /out\s*prepare/i.test(r.ket || "") ||
  (/^prepare$/i.test((r.loc || "").trim())) ||
  (/^prepare$/i.test((r.ket || "").trim())) ||
  (r.loc.toUpperCase() === "OUT" && /prepare/i.test(r.ket || ""));
const isOutRow = (r) => !isOutPrepare(r) && (r.loc.toUpperCase() === "OUT" || isThrow(r));
/* RE = keterangan "RE" (kata utuh, huruf besar/kecil bebas); "RELABEL"/"RETURN" tidak ikut */
const isRelabel = (r) => /(^|[^a-z])re([^a-z]|$)/i.test(r.ket || "");

function getTrxStatus(r) {
  if (isOutPrepare(r)) return "OUT PREPARE";
  if (isRelabel(r)) return "RE";
  if (isThrow(r)) return "THROW";
  return "OUT";
}
const STATUS_BADGE = {
  OUT: '<span class="b b-out">OUT</span>',
  THROW: '<span class="b b-throw">THROW</span>',
  RE: '<span class="b b-re">RE</span>',
  "OUT PREPARE": '<span class="b b-prep">OUT PREPARE</span>',
};
function toQty(v) {
  if (typeof v === "number") return v;
  if (v == null) return null;
  const n = Number(String(v).replace(/[,\s]/g, ""));
  return isNaN(n) || String(v).trim() === "" ? null : n;
}

function parseTrx(wb) {
  const need = ["part", "qty", "unique", "locator", "time"];
  let pick = null,
    found = [];
  for (const name of wb.SheetNames) {
    const aoa = XLSX.utils.sheet_to_json(wb.Sheets[name], {
      header: 1,
      raw: true,
      defval: null,
    });
    if (!aoa.length) continue;
    const hdr = aoa[0].map((h) =>
      String(h == null ? "" : h)
        .trim()
        .toLowerCase(),
    );
    found = hdr;
    if (need.every((k) => hdr.includes(k))) {
      pick = { aoa, hdr };
      break;
    }
  }
  if (!pick)
    throw new Error(
      "Kolom harus ada: Part, Qty, Unique, Locator, Time. Ditemukan: " +
      (found.filter(Boolean).join(", ") || "kosong"),
    );
  const ix = {};
  need.forEach((k) => (ix[k] = pick.hdr.indexOf(k)));
  ix.ket = pick.hdr.indexOf("keterangan");
  const rows = [];
  let bad = 0,
    minT = Infinity,
    maxT = -Infinity;
  for (let i = 1; i < pick.aoa.length; i++) {
    const r = pick.aoa[i];
    const u = r[ix.unique];
    if (u == null || String(u).trim() === "") continue;
    const t = parseTime(r[ix.time]);
    if (t == null) {
      bad++;
      continue;
    }
    const loc = String(r[ix.locator] == null ? "" : r[ix.locator]).trim();
    const ket =
      ix.ket < 0 || r[ix.ket] == null ? "" : String(r[ix.ket]).trim();
    rows.push({
      part: normPartTrx(r[ix.part]),
      qty: toQty(r[ix.qty]),
      unique: String(u).trim(),
      loc,
      t,
      ket,
    });
    if (t < minT) minT = t;
    if (t > maxT) maxT = t;
  }
  return { rows, bad, minT, maxT };
}
function guessShift(name, rows) {
  const m = name.match(/shift\s*[-_ ]?([123])/i);
  if (m) return +m[1];
  if (!rows.length) return 1;
  const ts = rows.map((r) => r.t).sort((a, b) => a - b),
    mid = new Date(ts[ts.length >> 1]);
  const h = mid.getUTCHours() + mid.getUTCMinutes() / 60;
  return h >= 7 && h < 15 ? 1 : h >= 15 && h < 23 ? 2 : 3;
}
function fileFullDate(name, minT) {
  const m = name.match(/(\d{1,2})[-_.](\d{1,2})[-_.](\d{2,4})/);
  if (m) {
    let y = +m[3];
    if (y < 100) y += 2000;
    return new Date(y, +m[2] - 1, +m[1]);
  }
  const x = new Date(minT);
  return new Date(x.getUTCFullYear(), x.getUTCMonth(), x.getUTCDate());
}
const fileDateLabel = (f) => {
  const d = f.fdate;
  return pad(d.getDate()) + "-" + pad(d.getMonth() + 1) + "-" + d.getFullYear();
};
/* Tally Sheet = persiapan: scan tanggal D dicek ke sheet tanggal D-1 (shift sama) */
const sheetDateOf = (d) => new Date(d.getFullYear(), d.getMonth(), d.getDate() - 1);
const sheetDateLabel = (f) => {
  const d = sheetDateOf(f.fdate);
  return pad(d.getDate()) + "-" + pad(d.getMonth() + 1) + "-" + d.getFullYear();
};

/* ---------- Tally Sheet: nama sheet + dua blok ---------- */
const HARI_OFF = { SABTU: 0, JUMAT: 1, KAMIS: 2, RABU: 3, SELASA: 4, SENIN: 5, MINGGU: 6 };
const SHIFT_OF = { PAGI: 1, MALAM: 3 };
const TALLY_COLS = [
  { pi: 3, qi: 7, ki: 8, label: "Kiri" },
  { pi: 13, qi: 17, ki: 18, label: "Kanan" },
];

// Sheet = [HARI] [PAGI|MALAM] [DD] [MM] [YYYY], tanggal = Week Ending Sabtu.
// Tanggal sebenarnya = tanggal WE - offset hari (SABTU=0 ... MINGGU=6).
function tallyMeta(name) {
  const s = String(name == null ? "" : name).trim().replace(/\s+/g, " ").toUpperCase();
  const m = s.match(
    /^(MINGGU|SENIN|SELASA|RABU|KAMIS|JUMAT|SABTU)\s+(PAGI|MALAM)\s+(\d{1,2})\s*[._\-\/ ]?\s*(\d{1,2})\s*[._\-\/ ]?\s*(\d{4})$/,
  );
  if (!m) return null;
  const dd = +m[3],
    mm = +m[4],
    yyyy = +m[5];
  const we = new Date(yyyy, mm - 1, dd);
  if (we.getFullYear() !== yyyy || we.getMonth() !== mm - 1 || we.getDate() !== dd)
    return null;
  const date = new Date(yyyy, mm - 1, dd - HARI_OFF[m[1]]);
  return {
    name,
    day: m[1],
    shift: SHIFT_OF[m[2]],
    we,
    date,
    key: ymd(date),
    weKey: ymd(we),
  };
}
const tallySheetsOf = (wb) =>
  wb.SheetNames.map(tallyMeta)
    .filter(Boolean)
    .sort((a, b) => a.date - b.date || a.shift - b.shift);

function parseTally(wb, name) {
  const ws = wb.Sheets[name];
  if (!ws || !ws["!ref"]) return { map: new Map(), rows: [], noQty: 0 };
  const rg = XLSX.utils.decode_range(ws["!ref"]);
  const val = (r, c) => {
    const x = ws[XLSX.utils.encode_cell({ r, c })];
    return !x || x.t === "e" ? null : x.v;
  };
  const isHdr = (p) => p === "PART NUMBER" || p === "PART NO";
  // nomor halaman cetak: naik setiap ketemu baris judul "Part Number"
  const pageAt = [];
  let pg = 0;
  for (let r = rg.s.r; r <= rg.e.r; r++) {
    for (const b of TALLY_COLS) {
      const raw = val(r, b.pi);
      if (raw != null && isHdr(normPart(raw))) {
        pg++;
        break;
      }
    }
    pageAt[r] = pg;
  }
  const map = new Map(),
    rows = [];
  let noQty = 0;
  TALLY_COLS.forEach((b, bi) => {
    for (let r = rg.s.r; r <= rg.e.r; r++) {
      const raw = val(r, b.pi);
      if (raw == null) continue;
      const part = normPart(raw);
      if (!part || isHdr(part)) continue;
      const qty = toQty(val(r, b.qi));
      if (qty == null) {
        noQty++;
        continue;
      }
      const k = val(r, b.ki),
        ket = k == null ? "" : String(k).trim();
      const l = val(r, b.pi - 2),
        po = l == null ? "" : String(l).trim();
      const rec = { blok: b.label, part, qty, ket, po, page: pageAt[r], bi, r };
      rows.push(rec);
      if (!map.has(part)) map.set(part, { part, qty: 0, det: [] });
      const o = map.get(part);
      o.qty += qty;
      o.det.push(rec);
    }
  });
  // urutan baca: halaman, lalu blok kiri sebelum kanan, lalu baris (urutan pengurangan)
  rows.sort((a, b) => a.page - b.page || a.bi - b.bi || a.r - b.r);
  return { map, rows, noQty };
}

/* ---------- OUT: dedupe Unique global + RELABEL ---------- */
function analyzeOut(rows) {
  // Satu Unique = satu baris terakhir (Time paling besar) di SEMUA file/shift.
  const byU = new Map();
  for (const r of rows) {
    const o = byU.get(r.unique);
    if (!o) byU.set(r.unique, { ...r });
    else if (r.t >= o.t) Object.assign(o, r);
  }
  const finals = [...byU.values()];
  // RELABEL = barang sisa proses OUT; dikurangkan dari total OUT Part yang sama.
  const relabelByShift = new Map();
  for (const r of rows) {
    if (!isRelabel(r)) continue;
    if (!relabelByShift.has(r.shift)) relabelByShift.set(r.shift, new Map());
    const m = relabelByShift.get(r.shift);
    if (!m.has(r.part)) m.set(r.part, { part: r.part, qty: 0, rows: [] });
    const o = m.get(r.part);
    o.qty += r.qty || 0;
    o.rows.push(r);
  }
  return { finals, relabelByShift, unique: byU.size, scans: rows.length };
}

/* Pool OUT per Part (per shift) = jumlah semua Qty OUT (termasuk THROW) - RELABEL.
   Baris Tally TIDAK digabung: dibaca berurutan, tiap baris mengurangi pool Part yang sama.
   Baris terpenuhi penuh = Sudah Out, selain itu = Belum Out (Kurang N).
   Qty THROW dihitung sebagai OUT, tapi jumlahnya dicatat terpisah di keterangan. */
function compareShift(tallyRows, finals, relabelMap, skip, shift) {
  const outMap = new Map();
  for (const r of finals) {
    if (r.shift !== shift) continue;
    if (!isOutRow(r)) continue;
    if (!outMap.has(r.part))
      outMap.set(r.part, { part: r.part, qty: 0, thr: 0, rows: [], thrRows: [] });
    const o = outMap.get(r.part);
    o.qty += r.qty || 0;
    o.rows.push(r);
    if (isThrow(r)) {
      o.thr += r.qty || 0;
      o.thrRows.push(r);
    }
  }
  const prepMap = new Map();
  for (const r of finals) {
    if (r.shift !== shift) continue;
    if (!isOutPrepare(r)) continue;
    if (!prepMap.has(r.part))
      prepMap.set(r.part, { part: r.part, qty: 0, rows: [] });
    const p = prepMap.get(r.part);
    p.qty += r.qty || 0;
    p.rows.push(r);
  }
  const pools = new Map();
  const getPool = (part) => {
    let p = pools.get(part);
    if (!p) {
      const o = outMap.get(part) || { qty: 0, thr: 0, rows: [], thrRows: [] };
      const rl = relabelMap.get(part) || { qty: 0, rows: [] };
      const pr = prepMap.get(part) || { qty: 0, rows: [] };
      p = {
        part,
        out: o.qty,
        outRows: o.rows,
        throw: o.thr,
        throwRows: o.thrRows,
        relabel: rl.qty,
        relabelRows: rl.rows,
        prepare: pr.qty,
        prepareRows: pr.rows,
        start: Math.max(0, o.qty - rl.qty),
      };
      p.left = p.start;
      pools.set(part, p);
    }
    return p;
  };
  for (const part of outMap.keys()) getPool(part);
  for (const part of prepMap.keys()) getPool(part);
  const thrNote = (p) => (p.throw ? " · Throw " + nf(p.throw) : "");
  const tallyParts = new Set(tallyRows.map((r) => r.part));
  const items = [],
    extra = [],
    excluded = [];
  for (const t of tallyRows) {
    // BARU: Part yang ada di TAB_IN tidak dibandingkan ke Scan OUT (dikecualikan)
    if (skip && skip.has(t.part)) {
      const pe = getPool(t.part);
      excluded.push({
        part: t.part,
        shift,
        po: t.po,
        tally: t.qty,
        tallyDet: [t],
        out: pe.out,
        outRows: pe.outRows,
        prepare: pe.prepare,
        prepareRows: pe.prepareRows,
        throw: pe.throw,
        throwRows: pe.throwRows,
        relabel: pe.relabel,
        relabelRows: pe.relabelRows,
        net: pe.start,
        status: "tin",
        k: "Dikecualikan (TAB_IN)",
      });
      continue;
    }
    const p = getPool(t.part);
    const before = p.left,
      take = Math.min(t.qty, before),
      short = t.qty - take;
    p.left -= take;
    items.push({
      part: t.part,
      shift,
      po: t.po,
      tally: t.qty,
      tallyDet: [t],
      out: p.out,
      outRows: p.outRows,
      prepare: p.prepare,
      prepareRows: p.prepareRows,
      throw: p.throw,
      throwRows: p.throwRows,
      relabel: p.relabel,
      relabelRows: p.relabelRows,
      net: p.start,
      before,
      take,
      after: p.left,
      short,
      selisih: take - t.qty,
      status: short === 0 ? "out" : "belum",
      k:
        short === 0
          ? p.throw
            ? "Termasuk Throw " + nf(p.throw)
            : ""
          : (p.out === 0 && !p.relabel ? "Belum ada scan OUT" : "Kurang " + nf(short)) + thrNote(p),
    });
  }
  let throwQty = 0,
    prepareQty = 0;
  for (const [part, p] of pools) {
    throwQty += p.throw;
    prepareQty += p.prepare;
    if (p.left <= 0) continue;
    if (skip && skip.has(part)) {
      // Part TAB_IN tanpa baris Tally tetap ditampilkan sebagai dikecualikan
      if (!tallyParts.has(part))
        excluded.push({
          part,
          shift,
          po: "",
          tally: null,
          tallyDet: [],
          out: p.out,
          outRows: p.outRows,
          prepare: p.prepare,
          prepareRows: p.prepareRows,
          throw: p.throw,
          throwRows: p.throwRows,
          relabel: p.relabel,
          relabelRows: p.relabelRows,
          net: p.left,
          status: "tin",
          k: "Dikecualikan (TAB_IN)",
        });
      continue;
    }
    extra.push({
      part,
      shift,
      tally: null,
      tallyDet: [],
      out: p.out,
      outRows: p.outRows,
      prepare: p.prepare,
      prepareRows: p.prepareRows,
      throw: p.throw,
      throwRows: p.throwRows,
      relabel: p.relabel,
      relabelRows: p.relabelRows,
      net: p.left,
      status: "extra",
      k:
        (tallyParts.has(part) ? "Sisa OUT (Tally sudah terpenuhi)" : "Tidak ada di Tally") +
        thrNote(p),
    });
  }
  return { items, extra, excluded, throwQty, prepareQty };
}

function showOut(on) {
  reveal($("out"), on);
  reveal($("res"), on);
}
function niceTop(max) {
  if (max <= 0) return { step: 1, top: 4 };
  const raw = max / 4,
    p = Math.pow(10, Math.floor(Math.log10(raw))),
    m = raw / p;
  const step = Math.max(1, (m <= 1 ? 1 : m <= 2 ? 2 : m <= 5 ? 5 : 10) * p);
  return { step, top: step * Math.ceil(max / step) };
}
function renderChart(anim) {
  const el = $("chart");
  if (!res || !el) return;
  const S = res.shifts.map((s) => {
    const p = res.per[s];
    return {
      s,
      has: !!p,
      tag: p ? p.tag : 0,
      out: p ? p.out : 0,
      miss: p ? p.miss : 0,
    };
  });
  const { step, top } = niceTop(Math.max(1, ...S.map((x) => x.tag)));
  const n = Math.round(top / step);
  let grid = "";
  for (let i = 0; i <= n; i++)
    grid +=
      '<i style="bottom:' + (i * step / top) * 100 + '%"><b>' + nf(i * step) + "</b></i>";
  const bar = (cls, v, name, s) =>
    '<div class="br ' + cls + '" data-h="' + (v / top) * 100 +
    '" title="Shift ' + s + " · " + name + ": " + nf(v) + '"><span>' + nf(v) + "</span></div>";
  const cols = S.map((x, i) => {
    const dim = (sf !== "all" && +sf !== x.s) || !x.has;
    const pct = x.has && x.tag ? Math.round((x.out / x.tag) * 100) + "% sudah out" : x.has ? "-" : "belum ada data";
    return (
      '<div class="grp' + (dim ? " dim" : "") + '"><div class="bars" style="--d:' + i * 80 + 'ms">' +
      bar("br2", x.out, "Sudah Out", x.s) +
      bar("br3", x.miss, "Belum Out", x.s) +
      '</div><div class="lb">Shift ' + x.s + "<small>" + pct + "</small></div></div>"
    );
  }).join("");
  const seg =
    res.shifts.length > 1
      ? '<div class="bar" style="margin:0">' +
      [["all", "Semua"]]
        .concat(res.shifts.map((x) => [String(x), "Shift " + x]))
        .map(([k, l]) => '<button class="tab' + (String(sf) === k ? " on" : "") + '" data-s="' + k + '">' + l + "</button>")
        .join("") + "</div>"
      : "";
  el.innerHTML =
    '<div class="head"><h2 style="margin:0">Ringkasan per shift</h2>' + seg + '</div><div class="legend" style="margin-bottom:8px">' +
    '<span><i style="background:var(--c2)"></i>Sudah Out</span>' +
    '<span><i style="background:var(--c3)"></i>Belum Out</span></div>' +
    '<div class="plot"><div class="gl">' + grid + '</div><div class="cols">' + cols + "</div></div>" +
    '<div class="note" id="cnote"></div>';
  const parts = [];
  if (res.extra.length) parts.push(nf(res.extra.length) + " Part punya sisa OUT / tanpa Tally");
  if (res.tabin.length) parts.push(nf(res.tabin.length) + " Part dikecualikan (TAB_IN)");
  if (res.relabelQty) parts.push(nf(res.relabelQty) + " qty RELABEL dikurangkan");
  if (res.throwQty) parts.push(nf(res.throwQty) + " qty THROW dihitung sebagai OUT");
  if (res.prepareQty) parts.push(nf(res.prepareQty) + " qty OUT PREPARE (belum keluar)");
  el.querySelector("#cnote").textContent = parts.length
    ? parts.join(" · ")
    : "Semua Part OUT ada di Tally.";
  const go = () =>
    el.querySelectorAll(".br").forEach((b) => (b.style.height = b.dataset.h + "%"));
  if (anim && !RM) requestAnimationFrame(() => requestAnimationFrame(go));
  else {
    el.querySelectorAll(".br").forEach((b) => (b.style.transition = "none"));
    go();
  }
}

/* ---------- alur ---------- */
function autoSheetFor(f) {
  if (!f.fdate) return null;
  const key = ymd(sheetDateOf(f.fdate));
  return tallySheets.find((m) => m.key === key && m.shift === f.shift) || null;
}
function sheetForFile(f) {
  if (manualSheet && tallyByName.has(manualSheet)) {
    const m = tallyByName.get(manualSheet);
    if (m.shift === f.shift) return m;
  }
  if (f.sheet && tallyByName.has(f.sheet)) return tallyByName.get(f.sheet);
  return null;
}
function autoSelect() {
  autoMsg = "";
  manualSheet = "";
  if (!tallyWB || !files.length) return;
  const usable = files.filter((f) => f.shift === 1 || f.shift === 3);
  if (!usable.length) return;
  let first = null,
    missing = null;
  for (const f of usable) {
    f.sheet = null;
    const m = autoSheetFor(f);
    if (m) {
      f.sheet = m.name;
      if (!first) first = m;
    } else if (!missing) missing = f;
  }
  if (first) $("sheet").value = first.name;
  if (missing)
    autoMsg = "Sheet " + sheetDateLabel(missing) + " Shift " + missing.shift + " tidak ada, pilih manual";
  else autoMsg = "Otomatis dari tanggal + shift file";
}

function run() {
  showErr("");
  $("info").style.display = "none";
  reveal($("sel"), !!tallyWB);
  $("auto").textContent = autoMsg;
  const usable = files.filter((f) => f.shift === 1 || f.shift === 3);
  if (!tallyWB || !usable.length || !tallySheets.length) {
    res = null;
    showOut(false);
    if (tallyWB && files.length && !usable.length) {
      $("info").style.display = "block";
      $("info").textContent = "Hanya ada file Scan Shift 2. Shift 2 tidak dipakai.";
    }
    return;
  }
  const allRows = [];
  for (const f of usable)
    for (const r of f.rows) allRows.push({ ...r, shift: f.shift, file: f.name });
  const a = analyzeOut(allRows);
  const shifts = [...new Set(usable.map((f) => f.shift))].sort((x, y) => x - y);
  const per = {},
    list = [],
    extra = [],
    tabinAll = [],
    msgs = [],
    usedSheets = [];
  let noQty = 0,
    bad = 0,
    relabelQty = 0,
    throwQty = 0,
    prepareQty = 0;
  for (const s of shifts) {
    const fs = usable
      .filter((f) => f.shift === s)
      .sort((x, y) => x.name.localeCompare(y.name));
    let meta = null;
    for (const f of fs) {
      const m = sheetForFile(f);
      if (m) {
        meta = m;
        break;
      }
    }
    if (!meta) {
      const f0 = fs[0];
      msgs.push("Sheet " + sheetDateLabel(f0) + " Shift " + s + " tidak ada, pilih manual.");
      continue;
    }
    const t = parseTally(tallyWB, meta.name);
    if (!usedSheets.includes(meta.name)) usedSheets.push(meta.name);
    noQty += t.noQty;
    if (!t.rows.length)
      msgs.push("Sheet " + meta.name + " tidak punya baris Part + Jumlah.");
    const relMap = a.relabelByShift.get(s) || new Map();
    relMap.forEach((v) => (relabelQty += v.qty));
    const cmp = compareShift(t.rows, a.finals, relMap, tabIn.set, s);
    throwQty += cmp.throwQty;
    prepareQty += cmp.prepareQty;
    cmp.items.forEach((x) => list.push(x));
    cmp.extra.forEach((x) => extra.push(x));
    cmp.excluded.forEach((x) => tabinAll.push(x));
    const scanS = a.finals.filter((r) => r.shift === s);
    per[s] = {
      tag: cmp.items.length,
      out: cmp.items.filter((x) => x.status === "out").length,
      miss: cmp.items.filter((x) => x.status === "belum").length,
      extra: cmp.extra.length,
      tin: cmp.excluded.length,
      scans: scanS.length,
      unique: new Set(scanS.map((r) => r.unique)).size,
      rawOut: allRows.filter((r) => r.shift === s && isOutRow(r)).length,
      finalOut: scanS.filter(isOutRow).length,
      rawPrep: allRows.filter((r) => r.shift === s && isOutPrepare(r)).length,
      finalPrep: scanS.filter(isOutPrepare).length,
    };
  }
  files.forEach((f) => (bad += f.bad));
  if (!Object.keys(per).length) {
    if (msgs.length) {
      $("info").style.display = "block";
      $("info").innerHTML = msgs.map(esc).join("<br>");
    }
    res = null;
    showOut(false);
    return;
  }
  if (noQty) msgs.push(noQty + " baris Tally tanpa Qty dilewati.");
  if (bad) msgs.push(bad + " baris scan dilewati (Time tidak terbaca).");
  if (!tabIn.n)
    msgs.push(
      (tabIn.err
        ? "TAB_IN belum terbaca: " + tabIn.err
        : "TAB_IN belum di-upload") +
      ". Part TAB_IN masih ikut dibandingkan.",
    );
  if (msgs.length) {
    $("info").style.display = "block";
    $("info").innerHTML = msgs.map(esc).join("<br>");
  }
  res = {
    shifts: Object.keys(per).map(Number).sort((x, y) => x - y),
    per,
    list,
    extra,
    tabin: tabinAll,
    sheets: usedSheets,
    scans: a.scans,
    unique: a.unique,
    rawOut: allRows.filter(isOutRow).length,
    finalOut: a.finals.filter(isOutRow).length,
    rawPrep: allRows.filter(isOutPrepare).length,
    finalPrep: a.finals.filter(isOutPrepare).length,
    relabelQty,
    throwQty,
    prepareQty,
    finals: a.finals,
  };
  if (sf !== "all" && !res.shifts.includes(+sf)) sf = "all";
  if (tab === "tin" && !res.tabin.length) tab = "all";
  showOut(true);
  render(true);
}

const BADGE = {
  out: '<span class="b ok">Sudah Out</span>',
  belum: '<span class="b miss">Belum Out</span>',
  extra: '<span class="b dbl">Sisa OUT</span>',
  tin: '<span class="b dbl">Dikecualikan</span>',
};
const inS = (x) => sf === "all" || x.shift === +sf;

function detailHTML(x) {
  const t = x.tally == null
    ? '<span class="muted">-</span>'
    : x.tallyDet && x.tallyDet.length
      ? x.tallyDet
        .map(
          (d) =>
            '<span class="seg">' + (d.po ? "Line " + esc(d.po) + " · " : "") + esc(d.blok) + " " + nf(d.qty) +
            (d.ket ? ' <span class="muted">(' + esc(d.ket) + ")</span>" : "") +
            "</span>",
        )
        .join("")
      : nf(x.tally);
  const o = x.outRows && x.outRows.length
    ? x.outRows
      .slice()
      .sort((a, b) => a.t - b.t)
      .map(
        (r) =>
          '<span class="seg">' + esc(r.unique) + " " + nf(r.qty || 0) +
          ' <span class="muted">' + esc(r.loc) + (isThrow(r) ? " · THROW" : "") + "</span></span>",
      )
      .join("")
    : '<span class="muted">belum ada scan OUT</span>';
  const pr = x.prepareRows && x.prepareRows.length
    ? x.prepareRows
      .slice()
      .sort((a, b) => a.t - b.t)
      .map(
        (r) =>
          '<span class="seg">' + esc(r.unique) + " " + nf(r.qty || 0) +
          ' <span class="muted">' + esc(r.loc || "PREPARE") + (r.ket ? " · " + esc(r.ket) : "") + "</span></span>",
      )
      .join("") +
      '<span class="seg muted">total ' + nf(x.prepare || 0) + " (tidak mengurangi Tally)</span>"
    : '<span class="muted">-</span>';
  const rl = x.relabelRows && x.relabelRows.length
    ? x.relabelRows
      .map(
        (r) =>
          '<span class="seg">' + esc(r.unique) + " " + nf(r.qty || 0) +
          (r.ket ? ' <span class="muted">(' + esc(r.ket) + ")</span>" : "") +
          "</span>",
      )
      .join("")
    : '<span class="muted">-</span>';
  const th = x.throwRows && x.throwRows.length
    ? x.throwRows
      .map((r) => '<span class="seg">' + esc(r.unique) + " " + nf(r.qty || 0) + "</span>")
      .join("") +
    '<span class="seg muted">total ' + nf(x.throw) + " (sudah termasuk OUT)</span>"
    : '<span class="muted">-</span>';
  return (
    '<div class="dgrid"><div><b>Tally</b>' + t + "</div><div><b>OUT</b>" + o +
    "</div><div><b>OUT PREPARE</b>" + pr +
    "</div><div><b>RELABEL</b>" + rl + "</div><div><b>THROW</b>" + th + "</div>" +
    (x.take == null
      ? ""
      : '<div><b>Pool OUT</b><span class="seg">OUT ' + nf(x.out) + " - RELABEL " + nf(x.relabel) + " = " + nf(x.net) +
      '</span><span class="seg">sebelum baris ini ' + nf(x.before) +
      '</span><span class="seg">dipakai ' + nf(x.take) +
      '</span><span class="seg">sisa ' + nf(x.after) + "</span></div>") +
    "</div>"
  );
}

function render(anim) {
  if (!res) return;
  const sh = res.shifts.filter((s) => sf === "all" || s === +sf);
  const sum = (k) => sh.reduce((a, s) => a + (res.per[s] ? res.per[s][k] : 0), 0);
  const L = res.list.filter(inS),
    E = res.extra.filter(inS),
    K = res.tabin.filter(inS),
    SO = (res.finals || []).filter(inS);
  const v = {
    tag: sum("tag"),
    out: sum("out"),
    miss: sum("miss"),
  };
  document
    .querySelectorAll(".st b")
    .forEach((b) => tween(b, v[b.dataset.k]));
  $("rt").textContent =
    "Sheet " + res.sheets.join(", ") + " · " +
    res.shifts.map((s) => "Shift " + s).join(", ");
  const T = [
    ["all", "Semua", L.length],
    ["out", "Sudah Out", v.out],
    ["belum", "Belum Out", v.miss],
    ["extra", "Sisa OUT", E.length],
  ];
  if (K.length) T.push(["tin", "Dikecualikan", K.length]);
  T.push(["shift_out", "Shift OUT", SO.length]);
  $("tabs").innerHTML = T.map(
    ([k, l, n]) =>
      '<button class="tab' +
      (tab === k ? " on" : "") +
      '" data-t="' +
      k +
      '">' +
      esc(l) +
      " " +
      n +
      "</button>",
  ).join("");

  $("foot").textContent =
    nf(res.scans) +
    " baris scan · " +
    nf(res.unique) +
    " Unique · " +
    nf(res.rawOut) +
    " baris OUT" +
    (res.rawOut !== res.finalOut ? " (" + nf(res.finalOut) + " final)" : "") +
    (res.prepareQty ? " · " + nf(res.prepareQty) + " qty OUT PREPARE" : "") +
    (res.throwQty ? " · " + nf(res.throwQty) + " qty THROW (termasuk OUT)" : "") +
    (K.length ? " · " + nf(K.length) + " baris Tally dikecualikan (TAB_IN)" : "") +
    " · " +
    nf(v.tag) +
    " baris Tally dibandingkan · pool OUT per Part, per shift";
  renderChart(anim);
  rows(anim);
}

function rows(anim) {
  const f = q.trim().toLowerCase(),
    multi = res.shifts.length > 1,
    off = multi ? 1 : 0;
  const keep = (a) => !f || a.join(" ").toLowerCase().includes(f);
  const L = res.list.filter(inS),
    E = res.extra.filter(inS),
    K = res.tabin.filter(inS);
  const S = (x) => (multi ? [x.shift] : []);
  const H = (a) => (multi ? ["Shift"] : []).concat(a);
  const txtOf = (x) =>
    [
      x.part,
      x.po || "",
      x.k,
      ...(x.outRows || []).map((r) => r.unique),
      ...(x.prepareRows || []).map((r) => r.unique),
      ...(x.tallyDet || []).map((d) => d.ket),
    ];
  let head,
    body,
    numCols = [];
  const trPair = (x, cells, i) => {
    const rw = anim && i < 28;
    const row =
      '<tr class="it' +
      (x.status === "belum" ? " miss" : "") +
      (rw ? " rw" : "") +
      '" title="Klik untuk rincian"' +
      (rw ? ' style="--d:' + i * 22 + 'ms"' : "") +
      ">" +
      cells
        .map(
          (c, j) =>
            "<td" + (numCols.includes(j) ? ' class="n"' : "") + ">" + c + "</td>",
        )
        .join("") +
      "</tr>";
    const det =
      '<tr class="drow" style="display:none"><td colspan="' +
      cells.length +
      '">' +
      detailHTML(x) +
      "</td></tr>";
    return row + det;
  };
  const car = (x) => '<span class="car">▸</span>' + esc(x.part);
  const tallyCell = (x) => {
    if (x.tally == null) return "";
    const d = x.tallyDet || [];
    const sub = d.map((r) => esc(r.blok) + " " + nf(r.qty)).join(" · ");
    return nf(x.tally) + (sub ? '<span class="sub">' + sub + "</span>" : "");
  };
  const outCell = (x) => {
    if (x.out == null) return "";
    const rows = (x.outRows || []).slice().sort((a, b) => a.t - b.t);
    const sub =
      rows.slice(0, 3).map((r) => esc(r.unique) + " " + nf(r.qty || 0)).join(" · ") +
      (rows.length > 3 ? " +" + (rows.length - 3) : "");
    return nf(x.out) + (sub ? '<span class="sub">' + sub + "</span>" : "");
  };
  const relCell = (x) => {
    if (!x.relabel) return "";
    const rows = x.relabelRows || [];
    const sub = rows.map((r) => esc(r.unique) + " " + nf(r.qty || 0)).join(" · ");
    return nf(x.relabel) + (sub ? '<span class="sub">' + sub + "</span>" : "");
  };
  const sgn = (x) => (x.selisih > 0 ? "+" : "") + nf(x.selisih);
  const stat = (x) => BADGE[x.status] + ' <span class="muted">' + esc(x.k) + "</span>";

  if (tab === "shift_out") {
    head = H(["Part", "Unique", "Qty", "Locator", "Waktu", "Status", "Keterangan"]);
    numCols = [2].map((i) => i + off);
    let A = (res.finals || [])
      .filter(inS)
      .slice()
      .sort((a, b) => a.shift - b.shift || a.t - b.t);
    if (filterPart) A = A.filter((r) => r.part === filterPart);
    if (filterStatus) A = A.filter((r) => getTrxStatus(r) === filterStatus);
    const txtOfTrx = (r) => [
      r.part,
      r.unique,
      String(r.qty || ""),
      r.loc || "",
      fmt(r.t),
      getTrxStatus(r),
      r.ket || "",
    ];
    body = A.filter((r) => keep(txtOfTrx(r))).map((r, i) => {
      const rw = anim && i < 28;
      const st = getTrxStatus(r);
      const cells = S(r)
        .map(esc)
        .concat([
          esc(r.part),
          esc(r.unique),
          nf(r.qty || 0),
          esc(r.loc || "-"),
          esc(fmt(r.t)),
          STATUS_BADGE[st] || esc(st),
          esc(r.ket || "-"),
        ]);
      return (
        '<tr class="it' +
        (rw ? " rw" : "") +
        '"' +
        (rw ? ' style="--d:' + i * 22 + 'ms"' : "") +
        ">" +
        cells
          .map(
            (c, j) =>
              "<td" + (numCols.includes(j) ? ' class="n"' : "") + ">" + c + "</td>",
          )
          .join("") +
        "</tr>"
      );
    });
  } else if (tab === "extra") {
    head = H(["Part", "Qty OUT", "RELABEL", "Sisa OUT", "Status"]);
    numCols = [1, 2, 3].map((i) => i + off);
    let A = E;
    if (filterPart) A = A.filter((x) => x.part === filterPart);
    body = A.filter((x) => keep(txtOf(x))).map((x, i) =>
      trPair(
        x,
        S(x)
          .map(esc)
          .concat([car(x), outCell(x), relCell(x), nf(x.net), stat(x)]),
        i,
      ),
    );
  } else if (tab === "tin") {
    head = H(["Part", "Line", "Qty Tally", "Qty OUT", "RELABEL", "Sisa OUT", "Status"]);
    numCols = [2, 3, 4, 5].map((i) => i + off);
    let A = K;
    if (filterPart) A = A.filter((x) => x.part === filterPart);
    body = A.filter((x) => keep(txtOf(x))).map((x, i) =>
      trPair(
        x,
        S(x)
          .map(esc)
          .concat([
            car(x),
            esc(x.po || ""),
            tallyCell(x),
            outCell(x),
            relCell(x),
            nf(x.net),
            stat(x),
          ]),
        i,
      ),
    );
  } else {
    head = H(["Part", "Line", "Qty Tally", "Terpenuhi", "Kurang", "Status"]);
    numCols = [2, 3, 4].map((i) => i + off);
    let A = L;
    if (tab === "out") A = A.filter((x) => x.status === "out");
    if (tab === "belum") A = A.filter((x) => x.status === "belum");
    if (filterPart) A = A.filter((x) => x.part === filterPart);
    body = A.filter((x) => keep(txtOf(x))).map((x, i) =>
      trPair(
        x,
        S(x)
          .map(esc)
          .concat([
            car(x),
            esc(x.po || ""),
            nf(x.tally),
            nf(x.take),
            x.short ? nf(x.short) : "",
            stat(x),
          ]),
        i,
      ),
    );
  }

  $("th").innerHTML =
    "<tr>" +
    head
      .map((h, i) => {
        const isNum = numCols.includes(i);
        const isPart = h.toLowerCase() === "part";
        if (isPart) {
          const isF = !!filterPart;
          return (
            '<th class="th-filter-th' + (isNum ? " n" : "") + '" title="Klik untuk filter Part Number">' +
            '<div class="th-box">' +
            '<span>' + esc(h) + '</span>' +
            '<button type="button" class="th-filter-btn' + (isF ? " on" : "") + '" title="Filter Part">' +
            (isF ? '▼' : '▾') +
            '</button>' +
            '</div>' +
            '</th>'
          );
        }
        return "<th" + (isNum ? ' class="n"' : "") + ">" + esc(h) + "</th>";
      })
      .join("") +
    "</tr>";
  updateFilterChips();

  $("tb").innerHTML = body.length
    ? body.join("")
    : '<tr><td colspan="' +
    head.length +
    '"><div class="empty">Tidak ada data</div></td></tr>';
}

function getAvailableParts() {
  if (!res) return [];
  let source = [];
  if (tab === "shift_out") {
    source = (res.finals || []).filter(inS);
  } else if (tab === "extra") {
    source = res.extra.filter(inS);
  } else if (tab === "tin") {
    source = res.tabin.filter(inS);
  } else {
    let A = res.list.filter(inS);
    if (tab === "out") A = A.filter((x) => x.status === "out");
    if (tab === "belum") A = A.filter((x) => x.status === "belum");
    source = A;
  }
  const counts = new Map();
  for (const item of source) {
    if (item.part) counts.set(item.part, (counts.get(item.part) || 0) + 1);
  }
  const sorted = [...counts.keys()].sort((a, b) => a.localeCompare(b));
  return sorted.map((p) => ({ part: p, count: counts.get(p) }));
}

function updatePartSelect() {
  const sel = $("filterPartSelect");
  if (!sel) return;
  const parts = getAvailableParts();
  const totalRows = parts.reduce((a, b) => a + b.count, 0);
  let html = '<option value="">Semua Part' + (totalRows ? " (" + totalRows + ")" : "") + "</option>";
  for (const p of parts) {
    html +=
      '<option value="' +
      esc(p.part) +
      '"' +
      (filterPart === p.part ? " selected" : "") +
      ">" +
      esc(p.part) +
      " (" +
      p.count +
      ")</option>";
  }
  sel.innerHTML = html;
  sel.value = filterPart;
  sel.classList.toggle("on", !!filterPart);
}

function updateFilterChips() {
  const el = $("fchips");
  if (!el) return;
  const chips = [];
  if (filterPart) {
    chips.push(
      '<span class="fchip">Part: <b>' +
        esc(filterPart) +
        '</b><button type="button" class="fchip-del" data-clear="part" title="Hapus filter Part">×</button></span>',
    );
  }
  el.innerHTML = chips.join("");
}

function openFilterPop(anchorEl) {
  const pop = $("filterPop");
  if (!pop || !res) return;
  if (pop.style.display !== "none") {
    closeFilterPop();
    return;
  }
  const titleEl = $("fpTitle");
  const searchEl = $("fpSearch");
  const optEl = $("fpOptions");
  searchEl.value = "";
  titleEl.textContent = "Filter Part Number";
  searchEl.placeholder = "Cari Part Number...";

  const parts = getAvailableParts();
  const totalRows = parts.reduce((a, b) => a + b.count, 0);

  let html =
    '<div class="fp-opt' +
    (!filterPart ? " sel" : "") +
    '" data-val=""><span>(Pilih Semua / Reset)</span><small>' +
    totalRows +
    "</small></div>";
  for (const p of parts) {
    html +=
      '<div class="fp-opt' +
      (filterPart === p.part ? " sel" : "") +
      '" data-val="' +
      esc(p.part) +
      '">' +
      "<span>" +
      esc(p.part) +
      "</span><small>" +
      p.count +
      "</small></div>";
  }
  optEl.innerHTML = html;

  pop.style.display = "block";
  const rect = anchorEl.getBoundingClientRect();
  pop.style.top = rect.bottom + 4 + "px";
  let left = rect.left;
  const popWidth = 260;
  if (left + popWidth > window.innerWidth - 12) left = window.innerWidth - popWidth - 12;
  if (left < 10) left = 10;
  pop.style.left = left + "px";

  setTimeout(() => searchEl.focus(), 40);
}

function closeFilterPop() {
  const pop = $("filterPop");
  if (pop) pop.style.display = "none";
}

/* ---------- TAB_IN (exception dari Excel, otomatis dari path) ---------- */
function setTabIn(parts) {
  const set = new Set();
  for (const p of parts) {
    const v = normPart(p);
    if (v) set.add(v);
  }
  tabIn = set.size
    ? { set, n: set.size, err: "" }
    : { set, n: 0, err: 'kolom A sheet "TAB_IN" kosong' };
  if (tallyWB) run();
}
function tabInParts(buf) {
  const wb = XLSX.read(buf, { type: "array" });
  let nm = wb.SheetNames.find((n) => n.trim().toLowerCase() === "tab_in");
  if (!nm && wb.SheetNames.length === 1) nm = wb.SheetNames[0];
  if (!nm)
    throw new Error(
      'sheet "TAB_IN" tidak ada di file (sheet yang ada: ' +
      wb.SheetNames.join(", ") +
      ")",
    );
  const col = XLSX.utils
    .sheet_to_json(wb.Sheets[nm], { header: 1, raw: true, defval: null })
    .map((r) => r[0]);
  if (col.length && /^part$/i.test(String(col[0] == null ? "" : col[0]).trim()))
    col.shift();
  return col;
}
async function loadTabInUp(list) {
  const file = list[0];
  if (!file) return;
  try {
    const parts = tabInParts(await file.arrayBuffer());
    setTabIn(parts);
    if (!tabIn.n) throw new Error(tabIn.err);
    $("dz3").classList.add("ok");
    $("m3").textContent = file.name + " · " + nf(tabIn.n) + " part";
    showErr("");
  } catch (e) {
    tabIn = { set: new Set(), n: 0, err: e.message };
    $("dz3").classList.remove("ok");
    $("m3").textContent = "Klik atau seret Data.xlsx";
    if (tallyWB) run();
    showErr("TAB_IN: " + e.message);
  }
}

function download() {
  if (!res) return;
  const wb = XLSX.utils.book_new(),
    st = { out: "Sudah Out", belum: "Belum Out", extra: "Sisa OUT", tin: "Dikecualikan (TAB_IN)" };
  const add = (rows, name) =>
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), name);
  const tallyTxt = (x) =>
    (x.tallyDet || [])
      .map((d) => d.blok + " " + d.qty + (d.ket ? " (" + d.ket + ")" : ""))
      .join("; ");
  const outTxt = (x) =>
    (x.outRows || [])
      .map((r) => r.unique + " " + (r.qty || 0) + " [" + r.loc + (isThrow(r) ? " THROW" : "") + "]")
      .join("; ");
  const relTxt = (x) =>
    (x.relabelRows || [])
      .map((r) => r.unique + " " + (r.qty || 0) + (r.ket ? " (" + r.ket + ")" : ""))
      .join("; ");
  const thrTxt = (x) =>
    (x.throwRows || []).map((r) => r.unique + " " + (r.qty || 0)).join("; ");
  const prepTxt = (x) =>
    (x.prepareRows || [])
      .map((r) => r.unique + " " + (r.qty || 0) + " [" + (r.loc || "PREPARE") + (r.ket ? " " + r.ket : "") + "]")
      .join("; ");
  const hasil = (x) => ({
    Shift: x.shift,
    Part: x.part,
    Line: x.po || "",
    "Qty Tally": x.tally == null ? "" : x.tally,
    Terpenuhi: x.take,
    Kurang: x.short || "",
    Status: st[x.status],
    Keterangan: x.k,
    "Qty OUT PREPARE": x.prepare || "",
    "Throw (sudah termasuk OUT)": x.throw || "",
    "Pool OUT (OUT-RELABEL)": x.net,
    "Pool sebelum baris": x.before,
    "Pool sisa": x.after,
    "Rincian Tally": tallyTxt(x),
    "Rincian OUT": outTxt(x),
    "Rincian OUT PREPARE": prepTxt(x),
    "Rincian RELABEL": relTxt(x),
    "Rincian THROW": thrTxt(x),
  });
  add(res.list.map(hasil), "Hasil");
  add(
    res.extra.map((x) => ({
      Shift: x.shift,
      Part: x.part,
      "Qty OUT": x.out,
      "Qty OUT PREPARE": x.prepare || "",
      RELABEL: x.relabel || "",
      Throw: x.throw || "",
      "Sisa OUT": x.net,
      Status: x.k,
      "Rincian OUT": outTxt(x),
      "Rincian OUT PREPARE": prepTxt(x),
      "Rincian RELABEL": relTxt(x),
      "Rincian THROW": thrTxt(x),
    })),
    "Sisa OUT",
  );
  add(
    res.tabin.map((x) => ({
      Shift: x.shift,
      Part: x.part,
      Line: x.po || "",
      "Qty Tally": x.tally == null ? "" : x.tally,
      "Qty OUT": x.out,
      "Qty OUT PREPARE": x.prepare || "",
      RELABEL: x.relabel || "",
      Throw: x.throw || "",
      "Sisa OUT": x.net,
      Status: st[x.status],
      Keterangan: x.k,
      "Rincian Tally": tallyTxt(x),
      "Rincian OUT": outTxt(x),
      "Rincian OUT PREPARE": prepTxt(x),
      "Rincian RELABEL": relTxt(x),
      "Rincian THROW": thrTxt(x),
    })),
    "TAB_IN",
  );
  add(
    res.finals
      .slice()
      .sort((a, b) => a.shift - b.shift || a.t - b.t)
      .map((r) => ({
        Shift: r.shift,
        Part: r.part,
        Qty: r.qty,
        Unique: r.unique,
        Locator: r.loc,
        Waktu: fmt(r.t),
        Status: getTrxStatus(r),
        Keterangan: r.ket,
        File: r.file,
      })),
    "Scan OUT",
  );
  XLSX.writeFile(wb, "Cek_OUT.xlsx");
}

/* ---------- daftar file scan ---------- */
function renderFiles() {
  $("dz2").classList.toggle("ok", files.length > 0);
  $("rtrx").hidden = !files.length;
  $("m2").textContent = files.length
    ? files.length + " file dimuat" +
    (files[0] && (files[0].shift === 1 || files[0].shift === 3)
      ? " · Shift " + files[0].shift
      : "")
    : scanCatalog.length
      ? scanCatalog.length + " file siap · pilih dropdown"
      : "Klik atau seret folder Scan";
}

const MON_SHORT = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"];

function scanMeta(name) {
  const s = String(name == null ? "" : name);
  const sm = s.match(/shift\s*[-_ ]?([123])(?!\d)/i);
  if (!sm) return null;
  const shift = +sm[1];
  if (shift !== 1 && shift !== 3) return null;
  let d = null,
    m = null,
    y = null;
  let r = s.match(/(?<!\d)(\d{1,2})[-_.\/](\d{1,2})[-_.\/](\d{2,4})(?!\d)/);
  if (r) {
    d = +r[1];
    m = +r[2];
    y = +r[3];
  } else {
    r = s.match(/(?<!\d)(\d{4})[-_.\/](\d{1,2})[-_.\/](\d{1,2})(?!\d)/);
    if (r) {
      y = +r[1];
      m = +r[2];
      d = +r[3];
    }
  }
  if (d == null) return null;
  if (y < 100) y += 2000;
  const dt = new Date(y, m - 1, d);
  if (dt.getFullYear() !== y || dt.getMonth() !== m - 1 || dt.getDate() !== d) return null;
  return {
    date: dt,
    dateKey: ymd(dt),
    shift,
    label:
      pad(d) + " " + MON_SHORT[m - 1] + " " + y + " " +
      (shift === 1 ? "Pagi" : "Malam") + "/shift " + shift,
  };
}

async function collectScanHandles(dir, depth, out, cap, lvl) {
  for await (const [n, h] of dir.entries()) {
    if (out.length >= cap) return;
    if (h.kind === "file") {
      if (/\.(xlsx|xlsm|xls|csv)$/i.test(n) && !n.startsWith("~$")) out.push({ n, h, lvl });
    } else if (depth > 0) {
      await collectScanHandles(h, depth - 1, out, cap, lvl + 1);
    }
  }
}

function renderScanSelect() {
  const wrap = $("trxwrap"),
    sel = $("trxsel");
  if (!wrap || !sel) return;
  wrap.style.display = scanCatalog.length ? "block" : "none";
  sel.innerHTML =
    '<option value="">Pilih file scan…</option>' +
    scanCatalog
      .map((x, i) => '<option value="' + i + '">' + esc(x.label) + "</option>")
      .join("");
  sel.value = "";
  sel.disabled = !scanCatalog.length;
  $("trxmsg").textContent = scanCatalog.length
    ? scanCatalog.length + " file" +
    (scanSkipped ? " · " + scanSkipped + " dilewati" : " · pilih satu file")
    : "";
}

async function loadScanDir(dir) {
  if (scanBusy) return;
  scanBusy = true;
  scanDir = dir;
  $("m2").textContent = "Membaca folder…";
  showErr("");
  try {
    const found = [];
    await collectScanHandles(dir, 4, found, 5000, 0);
    const byKey = new Map();
    let skipped = 0;
    for (const f of found) {
      const meta = scanMeta(f.n);
      if (!meta) {
        skipped++;
        continue;
      }
      const key = meta.dateKey + "-" + meta.shift;
      const prev = byKey.get(key);
      if (!prev || f.lvl < prev.lvl) {
        byKey.set(key, {
          name: f.n,
          label: meta.label,
          date: meta.date,
          dateKey: meta.dateKey,
          shift: meta.shift,
          handle: f.h,
          lvl: f.lvl,
        });
      }
    }
    scanCatalog = [...byKey.values()].sort(
      (a, b) => b.date - a.date || a.shift - b.shift,
    );
    scanSkipped = skipped;
    renderScanSelect();
    renderFiles();
    if (!scanCatalog.length) {
      showErr(
        'Tidak ada file scan berformat tanggal + "Shift1"/"Shift3" di folder "' +
        dir.name +
        '".' +
        (skipped ? " " + skipped + " file dilewati." : ""),
      );
    }
  } catch (e) {
    showErr("Folder Scan: " + (e.message || String(e)));
  } finally {
    scanBusy = false;
    if (!scanCatalog.length) $("m2").textContent = "Klik atau seret folder Scan";
  }
}

async function pickScanFolder() {
  if (!window.showDirectoryPicker) {
    showErr("Browser ini tidak mendukung akses folder. Buka file ini di Chrome atau Edge.");
    return;
  }
  if (scanBusy) return;
  try {
    const dir = await window.showDirectoryPicker({ id: "scan", mode: "read" });
    await loadScanDir(dir);
  } catch (e) {
    if (e && e.name === "AbortError") {
      if (!scanCatalog.length) $("m2").textContent = "Klik atau seret folder Scan";
    } else {
      showErr("Folder Scan: " + (e.message || String(e)));
    }
  }
}

async function dropScanFolder(e) {
  e.preventDefault();
  e.stopPropagation();
  $("dz2").classList.remove("drag");
  if (scanBusy) return;
  const items = e.dataTransfer && e.dataTransfer.items;
  if (!items || !items.length) return;
  const ps = [];
  for (const it of items) {
    if (it.kind !== "file" || !it.getAsFileSystemHandle) continue;
    try {
      ps.push(it.getAsFileSystemHandle().catch(() => null));
    } catch (err) {
      ps.push(Promise.resolve(null));
    }
  }
  const handles = await Promise.all(ps);
  const dir = handles.find((h) => h && h.kind === "directory") || null;
  if (!dir) {
    showErr("Seret sebuah folder ke kotak Scan Transaksi, bukan file.");
    return;
  }
  await loadScanDir(dir);
}

async function applyScanSelection() {
  const sel = $("trxsel");
  if (!sel || scanBusy) return;
  const i = sel.value === "" ? -1 : +sel.value;
  const entry = i >= 0 ? scanCatalog[i] : null;
  if (!entry) return;
  scanBusy = true;
  sel.disabled = true;
  showErr("");
  try {
    const file = await entry.handle.getFile();
    files = [];
    await loadTrx([file]);
  } catch (e) {
    showErr("Scan Transaksi: " + (e.message || String(e)));
  } finally {
    scanBusy = false;
    sel.disabled = false;
  }
}

/* ---------- upload ---------- */
async function loadTally(list) {
  const file = list[0];
  if (!file) return;
  try {
    const wb = XLSX.read(await file.arrayBuffer(), { type: "array" });
    const metas = tallySheetsOf(wb);
    if (!metas.length)
      throw new Error(
        'File "' + file.name + '": tidak ada sheet Tally berformat "HARI PAGI/MALAM DD MM YYYY" ' +
        '(contoh: "SENIN PAGI 07 11 2026"). Sheet ada: ' + wb.SheetNames.join(", "),
      );
    tallyWB = wb;
    tallySheets = metas;
    tallyByName = new Map(metas.map((m) => [m.name, m]));
    lastTallyErr = "";
    $("sheet").innerHTML = metas
      .map(
        (m) =>
          '<option value="' + esc(m.name) + '">' + esc(m.name) + " · " + ymd(m.date) + "</option>",
      )
      .join("");
    $("sheet").disabled = false;
    $("dz1").classList.add("ok");
    $("m1").textContent = file.name + " · " + metas.length + " sheet";
    autoSelect();
    showErr("");
    run();
  } catch (e) {
    tallyWB = null;
    tallySheets = [];
    tallyByName = new Map();
    lastTallyErr = e.message;
    $("dz1").classList.remove("ok");
    $("m1").textContent = "Klik atau seret file Tally";
    $("sheet").innerHTML = "<option>-</option>";
    $("sheet").disabled = true;
    run();
    showErr(e.message);
  }
}
async function loadTrx(list) {
  const errs = [];
  for (const file of list) {
    try {
      const t = parseTrx(
        XLSX.read(await file.arrayBuffer(), { type: "array" }),
      );
      if (!t.rows.length) throw new Error("tidak ada data");
      const old = files.findIndex((f) => f.name === file.name);
      if (old >= 0) files.splice(old, 1);
      files.push({
        id: ++fid,
        name: file.name,
        rows: t.rows,
        bad: t.bad,
        minT: t.minT,
        maxT: t.maxT,
        shift: guessShift(file.name, t.rows),
        fdate: fileFullDate(file.name, t.minT),
        sheet: null,
      });
    } catch (e) {
      errs.push(file.name + ": " + e.message);
    }
  }
  files.sort((a, b) => a.shift - b.shift || a.name.localeCompare(b.name));
  renderFiles();
  autoSelect();
  showErr(errs.join(" | "));
  if (errs.length) $("err").style.display = "block";
  runKeepErr(errs);
}
function runKeepErr(errs) {
  const m = errs.join(" | ");
  run();
  if (m) showErr(m);
}

function resetTrx() {
  files = [];
  res = null;
  autoMsg = "";
  manualSheet = "";
  renderFiles();
  renderScanSelect();
  run();
}
function resetAll() {
  tallyWB = null;
  tallySheets = [];
  tallyByName = new Map();
  files = [];
  res = null;
  tab = "all";
  q = "";
  sf = "all";
  autoMsg = "";
  manualSheet = "";
  lastTallyErr = "";
  scanCatalog = [];
  scanDir = null;
  scanBusy = false;
  scanSkipped = 0;
  filterPart = "";
  filterStatus = "";
  closeFilterPop();
  $("q").value = "";
  $("dz1").classList.remove("ok");
  $("m1").textContent = "Klik atau seret file Tally";
  $("sheet").innerHTML = "<option>-</option>";
  $("sheet").disabled = true;
  $("trxmsg").textContent = "";
  renderFiles();
  renderScanSelect();
  showErr("");
  run();
}

function bindDrop(zone, input, fn) {
  input.onchange = () => {
    if (input.files.length) fn([...input.files]);
    input.value = "";
  };
  ["dragenter", "dragover"].forEach((ev) =>
    zone.addEventListener(ev, (e) => {
      e.preventDefault();
      zone.classList.add("drag");
    }),
  );
  ["dragleave", "drop"].forEach((ev) =>
    zone.addEventListener(ev, (e) => {
      e.preventDefault();
      zone.classList.remove("drag");
    }),
  );
  zone.addEventListener("drop", (e) => {
    if (e.dataTransfer.files.length) fn([...e.dataTransfer.files]);
  });
}
bindDrop($("dz1"), $("f1"), loadTally);
bindDrop($("dz3"), $("f3"), loadTabInUp);
window.addEventListener("dragover", (e) => e.preventDefault());
window.addEventListener("drop", (e) => e.preventDefault());

$("dz2").onclick = (e) => {
  if (e.target.closest("#rtrx")) return;
  pickScanFolder();
};
$("dz2").onkeydown = (e) => {
  if (e.target.closest("#rtrx")) return;
  if (e.key === "Enter" || e.key === " ") {
    e.preventDefault();
    pickScanFolder();
  }
};
["dragenter", "dragover"].forEach((ev) =>
  $("dz2").addEventListener(ev, (e) => {
    e.preventDefault();
    e.stopPropagation();
    $("dz2").classList.add("drag");
  }),
);
$("dz2").addEventListener("dragleave", (e) => {
  e.preventDefault();
  e.stopPropagation();
  $("dz2").classList.remove("drag");
});
$("dz2").addEventListener("drop", dropScanFolder);
$("trxsel").onchange = applyScanSelection;

$("sheet").onchange = () => {
  manualSheet = $("sheet").value;
  autoMsg = "";
  run();
};
$("rtrx").onclick = (e) => {
  e.preventDefault();
  e.stopPropagation();
  resetTrx();
};
$("home").onclick = resetAll;
$("home").onkeydown = (e) => {
  if (e.key === "Enter" || e.key === " ") {
    e.preventDefault();
    resetAll();
  }
};
$("tabs").onclick = (e) => {
  const b = e.target.closest("[data-t]");
  if (!b) return;
  tab = b.dataset.t;
  closeFilterPop();
  render(true);
};
$("th").onclick = (e) => {
  const th = e.target.closest(".th-filter-th");
  if (!th) return;
  e.stopPropagation();
  openFilterPop(th);
};
$("fold").onclick = (e) => {
  const btn = e.target.closest(".fchip-del");
  if (!btn) return;
  filterPart = "";
  closeFilterPop();
  rows(false);
};
const partSel = $("filterPartSelect");
if (partSel) {
  partSel.onchange = (e) => {
    filterPart = e.target.value;
    closeFilterPop();
    rows(false);
  };
}
const fpop = $("filterPop");
if (fpop) {
  fpop.onclick = (e) => {
    const opt = e.target.closest(".fp-opt");
    if (!opt) return;
    filterPart = opt.dataset.val || "";
    closeFilterPop();
    rows(false);
  };
}
const fsearch = $("fpSearch");
if (fsearch) {
  fsearch.oninput = (e) => {
    const v = e.target.value.trim().toLowerCase();
    const opts = $("fpOptions").querySelectorAll(".fp-opt");
    opts.forEach((o) => {
      const text = o.textContent.toLowerCase();
      o.style.display = !v || text.includes(v) ? "flex" : "none";
    });
  };
}
document.addEventListener("click", (e) => {
  const pop = $("filterPop");
  if (!pop || pop.style.display === "none") return;
  if (e.target.closest("#filterPop") || e.target.closest(".th-filter-th")) return;
  closeFilterPop();
});
$("tb").onclick = (e) => {
  const tr = e.target.closest("tr.it");
  if (!tr) return;
  const d = tr.nextElementSibling;
  if (!d || !d.classList.contains("drow")) return;
  const show = d.style.display === "none";
  d.style.display = show ? "table-row" : "none";
  const c = tr.querySelector(".car");
  if (c) c.textContent = show ? "▾" : "▸";
};
$("chart").onclick = (e) => {
  const b = e.target.closest("[data-s]");
  if (!b) return;
  sf = b.dataset.s;
  render(true);
};
$("q").oninput = (e) => {
  q = e.target.value;
  rows(false);
};
$("dl").onclick = download;
$("tg").onclick = () => {
  const s = $("fold").classList.toggle("shut");
  $("fold").inert = s;
  $("tg").textContent = s ? "Buka ▾" : "Tutup ▴";
  $("tg").setAttribute("aria-expanded", !s);
};

/* ---------- INPUT OTOMATIS (File System Access API) ---------- */
(() => {
  const MON_EN = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  const MON_ID = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  const monAlt = (m) => [...new Set([MON_EN[m], MON_ID[m], MON_EN[m].slice(0, 3), MON_ID[m].slice(0, 3)])].join("|");
  const monRe = (m) => new RegExp("(^|[^a-z])(" + monAlt(m) + ")|^\\s*0?" + (m + 1) + "(?!\\d)", "i");
  const dstr = (d) => pad(d.getDate()) + "-" + pad(d.getMonth() + 1) + "-" + d.getFullYear();
  const st = () => { };

  /* simpan handle folder supaya cukup pilih sekali */
  const dbOpen = () => new Promise((res, rej) => {
    const r = indexedDB.open("tally_auto", 1);
    r.onupgradeneeded = () => r.result.createObjectStore("h");
    r.onsuccess = () => res(r.result);
    r.onerror = () => rej(r.error);
  });
  const dbDo = async (mode, fn) => {
    const d = await dbOpen();
    return new Promise((res, rej) => {
      const q = fn(d.transaction("h", mode).objectStore("h"));
      q.onsuccess = () => res(q.result);
      q.onerror = () => rej(q.error);
    });
  };

  /* helper folder */
  const sub = async (d, n) => { try { return await d.getDirectoryHandle(n); } catch (e) { return null; } };
  const fil = async (d, n) => { try { return await d.getFileHandle(n); } catch (e) { return null; } };
  const ls = async (d) => { const o = []; for await (const [n, h] of d.entries()) o.push({ n, h }); return o; };
  async function collect(d, depth, out) {
    for (const e of await ls(d)) {
      if (e.h.kind === "file") {
        if (/\.xls[xm]?$/i.test(e.n) && !e.n.startsWith("~$")) out.push(e);
      } else if (depth > 0) await collect(e.h, depth - 1, out);
    }
  }

  /* hari/shift sekarang -> tanggal Sabtu (Week Ending), sama dengan aturan di VBA */
  function weSat(d, s) {
    const w = d.getDay();
    const add = w === 0 ? (s >= 2 ? 6 : -1) : 6 - w;
    return new Date(d.getFullYear(), d.getMonth(), d.getDate() + add);
  }

  /* 2 = tanggal lengkap cocok, 1 = cocok "WE <tgl>", 0 = tidak cocok */
  function matchDate(name, dt) {
    const D = dt.getDate(), M = dt.getMonth() + 1, Y = dt.getFullYear();
    const s = name.replace(/\.[^.]+$/, "");
    let m, re = /(?<!\d)(?=(\d{1,2})\s*[._\-\/ ]\s*(\d{1,2})\s*[._\-\/ ]\s*(\d{2,4}))/g;
    while ((m = re.exec(s))) {
      let y = +m[3]; if (y < 100) y += 2000;
      if (+m[1] === D && +m[2] === M && y === Y) return 2;
      re.lastIndex++;
    }
    re = /(?<!\d)(?=(\d{4})\s*[._\-\/]\s*(\d{1,2})\s*[._\-\/]\s*(\d{1,2}))/g;
    while ((m = re.exec(s))) {
      if (+m[1] === Y && +m[2] === M && +m[3] === D) return 2;
      re.lastIndex++;
    }
    const c = s.match(/(^|\D)(\d{2})(\d{2})(\d{4})(\D|$)/);
    if (c && +c[2] === D && +c[3] === M && +c[4] === Y) return 2;
    if (new RegExp("(^|\\D)0?" + D + "\\s*[-_.\\s]?\\s*(" + monAlt(M - 1) + ")", "i").test(s)) return 2;
    if (new RegExp("(^|[^a-z])(we|week\\s*ending)[\\s._\\-]*0?" + D + "(?!\\d)", "i").test(s)) return 1;
    return 0;
  }

  async function getRoot(key, label) {
    let h = null;
    try { h = await dbDo("readonly", (o) => o.get(key)); } catch (e) { }
    if (h) {
      try {
        let p = await h.queryPermission({ mode: "read" });
        if (p !== "granted") p = await h.requestPermission({ mode: "read" });
        if (p === "granted") return h;
      } catch (e) { }
    }
    st("Pilih folder " + label + " (cukup sekali)…");
    h = await window.showDirectoryPicker({ id: key, mode: "read" });
    try { await dbDo("readwrite", (o) => o.put(h, key)); } catch (e) { }
    return h;
  }
  const forget = async (key) => { try { await dbDo("readwrite", (o) => o.delete(key)); } catch (e) { } };
  async function resolveDir(h, name) {
    if (h.name.toLowerCase() === name.toLowerCase()) return h;
    return sub(h, name);
  }
  async function findDirDeep(h, name, depth) {
    if (h.name && h.name.toLowerCase() === name.toLowerCase()) return h;
    if (depth <= 0) return null;
    const d = await sub(h, name);
    if (d) return d;
    for (const e of await ls(h)) {
      if (e.h.kind === "directory") {
        const r = await findDirDeep(e.h, name, depth - 1);
        if (r) return r;
      }
    }
    return null;
  }

  /* Tally Sheet: folder TALLYSHEET > folder bulan > WEEK ... / Tally_Sheet_WE_... */
  const tallyPriority = (n) => (/^tally[\s_-]*sheet(?![a-z])/i.test(n) ? 2 : 1);
  async function findTally(root, sat) {
    const cands = [];
    await collect(root, 3, cands);
    const hits = cands
      .filter((c) => TALLY_FILE_RE.test(c.n))
      .map((c) => ({ c, s: matchDate(c.n, sat), p: tallyPriority(c.n) }))
      .filter((x) => x.s > 0);
    if (!hits.length) {
      if (!cands.some((c) => TALLY_FILE_RE.test(c.n))) await forget("tally"); /* salah folder: pilih ulang di klik berikutnya */
      throw new Error("File Tally Sheet untuk Week Ending " + dstr(sat) + " tidak ditemukan. File yang terlihat: " +
        (cands.filter((c) => TALLY_FILE_RE.test(c.n)).slice(0, 8).map((c) => c.n).join(", ") || "tidak ada"));
    }
    const top = Math.max(...hits.map((x) => x.s));
    const topHits = hits.filter((x) => x.s === top);
    const bestP = Math.max(...topHits.map((x) => x.p));
    let best = null;
    for (const x of topHits.filter((x) => x.p === bestP)) {
      const f = await x.c.h.getFile();
      if (!best || f.lastModified > best.lastModified) best = f;
    }
    return best;
  }

  /* Tab In.xlsx di Monitoring */
  async function findTabIn(mon) {
    const h = await fil(mon, "Tab In.xlsx");
    if (h) return h.getFile();
    const e = (await ls(mon)).find((x) => x.h.kind === "file" && /^tab[\s_-]*in\.xls[xm]?$/i.test(x.n));
    return e ? e.h.getFile() : null;
  }

  /* Scan Transaksi: satu shift per kalkulasi. Week Ending dihitung sesuai shift file. */
  async function findTrx(mon, sd, s) {
    const data = await sub(mon, "Data");
    if (!data) throw new Error('Folder "Data" tidak ditemukan di Monitoring.');
    const dd = pad(sd.getDate()), mm = pad(sd.getMonth() + 1), yyyy = String(sd.getFullYear()), yy = yyyy.slice(2);
    const cache = {};
    const memo = async (k, get) => (k in cache ? cache[k] : (cache[k] = await get()));
    const names = [
      "Trx_Current_" + dd + "-" + mm + "-" + yyyy + "_Shift" + s + ".xlsx",
      "Trx_Current_" + dd + "-" + mm + "-" + yy + "_Shift" + s + ".xlsx",
    ];
    const sat = weSat(sd, s);
    const weName = "WE_ " + pad(sat.getDate()) + "_" + pad(sat.getMonth() + 1) + "_" + sat.getFullYear();
    const where = [
      () => memo("we" + weName, () => sub(data, weName)),
      async () => data,
      () => memo("ar" + weName, async () => {
        const y = await sub(data, String(sat.getFullYear()));
        const mo = y && (await sub(y, MON_EN[sat.getMonth()]));
        return (mo && (await sub(mo, weName))) || null;
      }),
    ];
    for (const w of where) {
      const d = await w();
      if (!d) continue;
      for (const n of names) {
        const h = await fil(d, n);
        if (h) return [{ shift: s, file: await h.getFile() }];
      }
    }
    return [];
  }

  /* Folder Scan = folder WE_ langsung di bawah Monitoring\Data. Dimuat sebagai FOLDER
     (sama seperti seret folder manual: dropdown terisi), lalu file tanggal+shift dipilih. */
  async function loadScanFolderFor(data, d, s) {
    const sat = weSat(d, s);
    const dirs = (await ls(data))
      .filter((e) => e.h.kind === "directory" && matchDate(e.n, sat) === 2)
      .sort((a, b) => b.n.length - a.n.length);
    const key = ymd(d);
    for (const e of dirs) {
      await loadScanDir(e.h);
      const i = scanCatalog.findIndex((x) => x.dateKey === key && x.shift === s);
      if (i < 0) continue;
      $("trxsel").value = String(i);
      await applyScanSelection();
      if (files.length) return e.n;
    }
    return null;
  }

  async function autoInput() {
    const btn = $("btnAuto");
    if (!window.showDirectoryPicker) {
      showErr("Browser ini tidak mendukung akses folder. Buka file ini di Chrome atau Edge.");
      return;
    }
    btn.disabled = true;
    showErr("");
    st("");
    try {
      const now = new Date(), h = now.getHours();
      // 07-15 = Shift 1; 15-23 tetap pakai file Shift 1 hari itu; 23-07 = Shift 3.
      const sh = h >= 7 && h < 15 ? 1 : h >= 15 && h < 23 ? 1 : 3;
      const sd = new Date(now.getFullYear(), now.getMonth(), now.getDate() - (h < 7 ? 1 : 0));
      // Tally Sheet memakai minggu Minggu-Sabtu, jadi selalu pakai varian "Sabtu berikutnya".
      const tgt = sheetDateOf(sd);
      const sat = weSat(tgt, 3);

      /* lokasi 1: folder Tally Sheet di network share */
      /* folder dipilih sendiri (nama bebas), disimpan, lalu file dicari di dalamnya */
      const tallyRoot = await getRoot("tally", "Tally Sheet");

      /* lokasi 2: \\192.168.1.2\mis\qr_logistic\Monitoring (folder Data) */
      const mHandle = await getRoot("monitoring", "Monitoring (qr_logistic\\Monitoring)");
      /* terima folder Monitoring, atau folder induknya; yang penting ada subfolder Data */
      let mon = (await sub(mHandle, "Data")) ? mHandle : await findDirDeep(mHandle, "Monitoring", 2);
      if (!mon || !(await sub(mon, "Data"))) {
        await forget("monitoring");
        throw new Error("Folder yang dipilih tidak punya subfolder Data. Klik lagi dan pilih ulang.");
      }
      const data = await sub(mon, "Data");

      st("Mencari file…");
      const [tallyFile, tabFile] = await Promise.all([
        findTally(tallyRoot, sat),
        findTabIn(mon),
      ]);
      let useD = sd, fb = false;

      st("Memuat data…");
      resetTrx();
      await loadTally([tallyFile]);
      if (!tallyWB)
        throw new Error(lastTallyErr || "Tally Sheet gagal dibaca.");
      if (tabFile) await loadTabInUp([tabFile]);

      /* Scan: muat folder WE_ di dalam Data, pilih file tanggal + shift */
      let scanFolder = await loadScanFolderFor(data, sd, sh);

      /* Kalau scan hari ini belum ada: pakai hari terakhir yang punya sheet di Tally Sheet */
      if (!scanFolder && tallyWB) {
        st("Scan hari ini belum ada, mencari hari terakhir di Tally Sheet…");
        const cand = [...new Set(
          tallySheets
            .map((m) => m.date)
            .filter((d) => d < tgt)
            .sort((a, b) => b - a)
            .map(ymd),
        )].map((s) => {
          const [y, mo, da] = s.split("-").map(Number);
          return new Date(y, mo - 1, da + 1);
        });
        for (const d of cand.slice(0, 14)) {
          scanFolder = await loadScanFolderFor(data, d, sh);
          if (scanFolder) { useD = d; fb = true; break; }
        }
      }
      if (!scanFolder) throw new Error("File scan Shift " + sh + " untuk tanggal " + dstr(sd) + " dan hari-hari sebelumnya di Tally Sheet belum ada di folder WE_ dalam Data.");

      st(
        "Tally: " + tallyFile.name + "\n" +
        (tabFile ? "Tab In: OK" : "Tab In.xlsx tidak ditemukan") + "\n" +
        "Folder Scan: " + scanFolder + "\n" +
        "Scan " + dstr(useD) + (fb ? " (hari terakhir di Tally Sheet, hari ini belum ada)" : "") + ": Shift " + sh
      );
    } catch (e) {
      if (e && e.name === "AbortError") st("Dibatalkan.");
      else { st(""); showErr(e.message || String(e)); }
    } finally {
      btn.disabled = false;
    }
  }
  $("btnAuto").onclick = autoInput;
})();