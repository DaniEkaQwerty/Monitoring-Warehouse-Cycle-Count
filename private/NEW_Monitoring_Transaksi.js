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
    const TAGCOLS = { 1: [0, 1, 2], 2: [6, 7, 8], 3: [11, 12, 13] };
    let tabIn = { set: new Set(), n: 0, err: "" };
    let tagWB = null,
      files = [],
      fid = 0,
      res = null,
      tab = "all",
      filterItem = "",
      filterLoc = "",
      filterMd = "",
      activePopCol = null,
      q = "",
      sf = "all",
      uOpen = false,
      kOpen = false;

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
    function toQty(v) {
      if (typeof v === "number") return v;
      if (v == null) return null;
      const n = Number(String(v).replace(/[,\s]/g, ""));
      return isNaN(n) || String(v).trim() === "" ? null : n;
    }
    const isNumSheet = (n) => /^\s*\d+\s*$/.test(n);
    const isMC = (k) => /^MC(?![A-Z0-9])/i.test(k);
    // BARU: Move dari kolom Keterangan (file Trx), Throw/Out dari kolom Locator
    const trxCats = (ket, loc) => {
      const c = [];
      if (/^MOVE(?![A-Z])/i.test(ket)) c.push("move");
      const l = String(loc).trim().toUpperCase();
      if (l === "THROW") c.push("throw");
      else if (l === "OUT") c.push("out");
      return c;
    };

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
      ix.ket = pick.hdr.indexOf("keterangan"); // BARU (opsional)
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
        const pRaw = normPart(r[ix.part]);
        const p = normPartTrx(r[ix.part]);
        const mm = pRaw.match(/\s+(\d{4})$/);
        rows.push({
          part: p,
          md: mm ? mm[1] : "",
          hasMd: !!mm,
          qty: toQty(r[ix.qty]),
          unique: String(u).trim(),
          loc,
          t,
          ket,
          cats: trxCats(ket, loc),
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
    function fileDate(name, minT) {
      const m = name.match(/(\d{1,2})[-_.](\d{1,2})[-_.](\d{2,4})/);
      if (m) return [+m[1], +m[2]];
      const x = new Date(minT);
      return [x.getUTCDate(), x.getUTCMonth() + 1];
    }

    function parseTag(wb, name, sh, skip) {
      const ws = wb.Sheets[name];
      if (!ws || !ws["!ref"]) return { rows: [], noQty: 0, mc: 0, tabin: [] };
      const rg = XLSX.utils.decode_range(ws["!ref"]);
      const [ci, cq, ck] = TAGCOLS[sh];
      const val = (r, c) => {
        const x = ws[XLSX.utils.encode_cell({ r, c })];
        return !x || x.t === "e" ? null : x.v;
      };
      const rows = [];
      let noQty = 0,
        mc = 0;
      const tabin = [];
      for (let r = rg.s.r; r <= rg.e.r; r++) {
        const it = val(r, ci);
        if (it == null) continue;
        const item = normPart(it);
        if (!item || item === "ITEM") continue;
        const k = val(r, ck),
          ket = k == null ? "" : String(k).trim();
        if (isMC(ket)) {
          mc++;
          continue;
        }
        const qty = toQty(val(r, cq));
        // BARU: item yang ada di TAB_IN (Access) tidak dicocokkan ke Scan
        if (skip && skip.has(item)) {
          tabin.push({ row: r + 1, item, qty, ket });
          continue;
        }
        if (qty == null) noQty++;
        rows.push({ row: r + 1, item, qty, ket });
      }
      return { rows, noQty, mc, tabin };
    }

    /* ---------- pencocokan ---------- */
    // Cari kombinasi scan yang jumlah Qty-nya pas dengan Qty tag
    function pickSum(pool, target) {
      const arr = pool.filter(
        (r) => r.qty != null && r.qty > 0 && r.qty <= target,
      );
      let best = null,
        iter = 0;
      const dfs = (i, sum, ch) => {
        if (best || iter++ > 30000) return;
        if (sum === target) {
          best = ch.slice();
          return;
        }
        if (sum > target || i >= arr.length) return;
        ch.push(arr[i]);
        dfs(i + 1, sum + arr[i].qty, ch);
        ch.pop();
        dfs(i + 1, sum, ch);
      };
      dfs(0, 0, []);
      return best;
    }

    // trxRows = scan SATU shift saja (shift lain tidak ikut dicocokkan)
    function analyze(trxRows, tagRows) {
      // 1) buang double scan: Unique + Locator + Waktu sama dihitung satu
      const seen = new Set(),
        byU = new Map();
      let scans = 0;
      for (const r of trxRows) {
        const k = r.unique + "\u0001" + r.loc + "\u0001" + r.t;
        if (seen.has(k)) continue;
        seen.add(k);
        scans++;
        const o = byU.get(r.unique);
        if (!o)
          byU.set(r.unique, {
            unique: r.unique,
            part: r.part,
            qty: r.qty,
            loc: r.loc,
            t: r.t,
            t0: r.t,
            used: false,
          });
        else {
          o.t0 = Math.min(o.t0, r.t);
          if (r.t >= o.t) {
            o.part = r.part;
            o.qty = r.qty;
            o.loc = r.loc;
            o.t = r.t;
          }
        }
      }
      const recs = [...byU.values()].sort((a, b) => a.t0 - b.t0);
      // 2) cocokkan ITEM + QTY persis
      const queues = new Map();
      for (const r of recs) {
        const k = r.part + "|" + r.qty;
        (queues.get(k) || queues.set(k, []).get(k)).push(r);
      }
      const list = tagRows.map((t) => {
        const qq = t.qty == null ? null : queues.get(t.item + "|" + t.qty);
        const u = qq && qq.length ? qq.shift() : null;
        if (u) u.used = true;
        return { ...t, us: u ? [u] : [], status: u ? "ok" : "miss" };
      });
      // 3) sisanya: ITEM sama di shift yang sama tapi QTY / QR beda -> Beda QR (tetap sudah masuk)
      for (const x of list) {
        if (x.status !== "miss" || x.qty == null) continue;
        const pool = recs.filter((r) => !r.used && r.part === x.item);
        if (!pool.length) continue;
        let ch = pickSum(pool, x.qty);
        if (!ch) {
          ch = [];
          let acc = 0;
          for (const r of pool) {
            if (acc >= x.qty) break;
            ch.push(r);
            acc += r.qty || 0;
          }
        }
        ch.forEach((r) => (r.used = true));
        x.us = ch;
        x.status = "bqr";
      }
      list.forEach(
        (x) => (x.sumQ = x.us.reduce((a, r) => a + (r.qty || 0), 0)),
      );
      // scan ITEM yang barisnya Beda QR (sudah masuk) tidak dianggap "Tidak di ID Tag"
      const bqrItems = new Set(
        list.filter((x) => x.status === "bqr").map((x) => x.item),
      );
      return {
        list,
        extra: recs.filter((r) => !r.used && !bqrItems.has(r.part)),
        scans,
        unique: recs.length,
      };
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
      const S = [1, 2, 3].map((s) => {
        const p = res.per[s];
        return p
          ? { s, has: true, tag: p.tag, in: p.ok + p.bqr, miss: p.miss }
          : { s, has: false, tag: 0, in: 0, miss: 0 };
      });
      const { step, top } = niceTop(
        Math.max(...S.map((x) => x.in + x.miss)),
      );
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
        const pct = x.has && x.tag ? Math.round((x.in / x.tag) * 100) + "% masuk" : x.has ? "-" : "belum ada data";
        return (
          '<div class="grp' + (dim ? " dim" : "") + '"><div class="bars" style="--d:' + i * 80 + 'ms">' +
          bar("br2", x.in, "Sudah masuk", x.s) +
          bar("br3", x.miss, "Belum masuk", x.s) +
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
        '<span><i style="background:var(--c2)"></i>Sudah masuk</span>' +
        '<span><i style="background:var(--c3)"></i>Belum masuk</span></div>' +
        '<div class="plot"><div class="gl">' + grid + '</div><div class="cols">' + cols + "</div></div>";
      const go = () =>
        el.querySelectorAll(".br").forEach((b) => (b.style.height = b.dataset.h + "%"));
      if (anim && !RM) requestAnimationFrame(() => requestAnimationFrame(go));
      else {
        el.querySelectorAll(".br").forEach((b) => (b.style.transition = "none"));
        go();
      }
    }

    /* ---------- alur ---------- */
    let autoMsg = "";
    function autoSelect() {
      autoMsg = "";
      if (!tagWB || !files.length) return;
      const f = files[0],
        [d, m] = f.fdate,
        cand = pad(d) + pad(m);
      const hit = [...$("sheet").options]
        .map((o) => o.value)
        .find((v) => v.trim() === cand);
      if (hit) {
        $("sheet").value = hit;
        autoMsg = "Otomatis dari tanggal file";
      } else autoMsg = 'Sheet "' + cand + '" tidak ada, pilih manual';
    }

    function run() {
      showErr("");
      $("info").style.display = "none";
      reveal($("sel"), !!tagWB);
      $("auto").textContent = autoMsg;
      if (!tagWB || !files.length || !$("sheet").value) {
        res = null;
        showOut(false);
        return;
      }
      const shifts = [...new Set(files.map((f) => f.shift))].sort();
      const per = {},
        list = [],
        extra = [],
        skip = [],
        tabinAll = [],
        mdAll = [],
        msgs = [];
      let noQty = 0,
        bad = 0;
      const seenMd = new Set();
      for (const s of shifts) {
        const tag = parseTag(tagWB, $("sheet").value, s, tabIn.set);
        tag.tabin.forEach((x) => tabinAll.push({ ...x, shift: s }));
        const rows = [];
        files
          .filter((f) => f.shift === s)
          .forEach((f) => {
            for (const r of f.rows) {
              if (r.hasMd) {
                const mk = r.unique + "\u0001" + r.loc + "\u0001" + r.t;
                if (!seenMd.has(mk)) {
                  seenMd.add(mk);
                  mdAll.push({
                    shift: s,
                    part: r.part,
                    md: r.md,
                    qty: r.qty,
                    unique: r.unique,
                    loc: r.loc,
                    t: r.t,
                  });
                }
              }
              // Move / Throw / Out tidak ikut dicocokkan ke ID Tag
              if (r.cats.length) skip.push({ ...r, shift: s });
              else rows.push(r);
            }
          });
        const a = analyze(rows, tag.rows);
        a.list.forEach((x) => {
          x.shift = s;
          list.push(x);
        });
        a.extra.forEach((x) => {
          x.shift = s;
          extra.push(x);
        });
        per[s] = {
          tag: a.list.length,
          ok: a.list.filter((x) => x.status === "ok").length,
          bqr: a.list.filter((x) => x.status === "bqr").length,
          miss: a.list.filter((x) => x.status === "miss").length,
          scans: a.scans,
          unique: a.unique,
          mc: tag.mc,
          tin: tag.tabin.length,
        };
        noQty += tag.noQty;
        if (!tag.rows.length)
          msgs.push(
            "Shift " +
            s +
            ": tidak ada ITEM di sheet " +
            $("sheet").value +
            ".",
          );
      }
      files.forEach((f) => (bad += f.bad));
      if (noQty)
        msgs.push(noQty + " baris tag tanpa QTY dihitung belum masuk.");
      if (bad) msgs.push(bad + " baris scan dilewati (Time tidak terbaca).");
      if (!tabIn.n)
        msgs.push(
          (tabIn.err
            ? "TAB_IN belum terbaca: " + tabIn.err
            : "TAB_IN belum di-upload") +
          ". Item TAB_IN masih ikut dicocokkan.",
        );
      if (msgs.length) {
        $("info").style.display = "block";
        $("info").innerHTML = msgs.map(esc).join("<br>");
      }
      res = {
        shifts,
        per,
        list,
        extra,
        skip,
        tabin: tabinAll,
        md: mdAll,
        sheet: $("sheet").value,
      };
      if (sf !== "all" && !shifts.includes(+sf)) sf = "all";
      showOut(true);
      render(true);
    }

    const BADGE = {
      ok: '<span class="b ok">OK</span>',
      bqr: '<span class="b ok">OK</span>',
      miss: '<span class="b miss">Belum masuk</span>',
    };
    // Unique / Locator / Waktu untuk baris tag (bisa lebih dari satu QR)
    const locsOf = (x) => [...new Set(x.us.map((r) => r.loc))];
    const lTxt = (x) => locsOf(x).join(", ");
    const tTxt = (x) =>
      x.us.length ? fmt(Math.max(...x.us.map((r) => r.t))) : "";
    const uTxt = (x) =>
      x.us.length
        ? x.us[0].unique + (x.us.length > 1 ? " +" + (x.us.length - 1) : "")
        : "";
    const inS = (x) => sf === "all" || x.shift === +sf;
    const isItemTab = () =>
      ["bqr", "tabin", "all", "ok", "miss"].includes(tab);
    const partOf = (x) => (isItemTab() ? x.item : x.part);
    const locsOfX = (x) =>
      ["bqr", "all", "ok", "miss"].includes(tab) ? locsOf(x) : [x.loc];
    function baseRows() {
      if (!res) return [];
      if (tab === "bqr")
        return res.list.filter(inS).filter((x) => x.status === "bqr");
      if (["move", "throw", "out"].includes(tab))
        return res.skip
          .filter((x) => inS(x) && x.cats.includes(tab))
          .sort((a, b) => a.t - b.t);
      if (tab === "extra") return res.extra.filter(inS);
      if (tab === "tabin") return res.tabin.filter(inS);
      if (tab === "md")
        return res.md.filter(inS).sort((a, b) => a.t - b.t);
      let A = res.list.filter(inS);
      if (tab === "ok") A = A.filter((x) => x.status !== "miss");
      if (tab === "miss") A = A.filter((x) => x.status === "miss");
      return A;
    }
    function filteredRows(skipCol) {
      let A = baseRows();
      if (skipCol !== "part" && filterItem)
        A = A.filter((x) => partOf(x) === filterItem);
      if (skipCol !== "md" && filterMd)
        A = A.filter((x) => x.md === filterMd);
      if (skipCol !== "loc" && filterLoc)
        A = A.filter((x) => (locsOfX(x) || []).includes(filterLoc));
      return A;
    }
    function getFilterValues(col) {
      const A = filteredRows(col);
      const counts = new Map();
      for (const x of A) {
        const vals =
          col === "part" ? [partOf(x)] : col === "md" ? [x.md] : locsOfX(x);
        for (const v of vals || []) {
          if (v) counts.set(v, (counts.get(v) || 0) + 1);
        }
      }
      return {
        total: A.length,
        values: [...counts.keys()]
          .sort((a, b) => a.localeCompare(b))
          .map((v) => ({ val: v, count: counts.get(v) })),
      };
    }

    function render(anim) {
      if (!res) return;
      const sh = res.shifts.filter((s) => sf === "all" || s === +sf);
      const sum = (k) => sh.reduce((a, s) => a + res.per[s][k], 0);
      const L = res.list.filter(inS),
        E = res.extra.filter(inS),
        K = res.skip.filter(inS);
      const v = {
        tag: sum("tag"),
        in: sum("ok") + sum("bqr"),
        bqr: sum("bqr"),
        miss: sum("miss"),
        md: res.md.filter(inS).length,
      };
      document
        .querySelectorAll(".st b")
        .forEach((b) => tween(b, v[b.dataset.k]));
      $("rt").textContent =
        "Sheet " +
        res.sheet +
        " · " +
        res.shifts.map((s) => "Shift " + s).join(", ");
      const T = [
        ["all", "Semua", L.length],
        ["ok", "Sudah masuk", v.in],
        ["miss", "Belum masuk", v.miss],
        ["md", "ITEM MD", v.md],
      ];
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
        nf(sum("scans")) +
        " scan · " +
        nf(sum("unique")) +
        " Unique · " +
        nf(sum("mc")) +
        " baris MC disembunyikan · " +
        nf(v.md) +
        " item MD · dicocokkan lewat ITEM + QTY, per shift";
      renderChart(anim);
      rows(anim);
    }

    function rows(anim, opened) {
      const f = q.trim().toLowerCase(),
        multi = res.shifts.length > 1,
        off = multi ? 1 : 0;
      const keep = (a) => !f || a.join(" ").toLowerCase().includes(f);
      const L = res.list.filter(inS),
        E = res.extra.filter(inS);
      const S = (x) => (multi ? [x.shift] : []);
      let head,
        body,
        numCols,
        uCol = -1,
        kCol = -1;
      const tcls = (j) => {
        const c =
          (numCols.includes(j) ? "n " : "") +
          (j === uCol ? "uc " + (opened === "u" ? "slide" : "") : "") +
          (j === kCol ? "kc " + (opened === "k" ? "slide" : "") : "");
        return c.trim() ? ' class="' + c.trim() + '"' : "";
      };
      const tog = (k, lbl, o) =>
        '<button type="button" class="ut" data-c="' +
        k +
        '" title="' +
        (o ? "Tutup" : "Buka") +
        " kolom " +
        lbl +
        '" aria-expanded="' +
        o +
        '">' +
        lbl +
        (o ? " ‹" : " ›") +
        "</button>";
      const tr = (cells, cls, i) => {
        const rw = anim && i < 28;
        return (
          "<tr" +
          (cls || rw
            ? ' class="' + (cls || "") + (rw ? " rw" : "") + '"'
            : "") +
          (rw ? ' style="--d:' + i * 22 + 'ms"' : "") +
          ">" +
          cells
            .map(
              (c, j) =>
                "<td" +
                tcls(j) +
                ">" +
                ((j === uCol && !uOpen) || (j === kCol && !kOpen) ? "" : c) +
                "</td>",
            )
            .join("") +
          "</tr>"
        );
      };
      // baris polos (semua sel di-escape)
      const plain = (raw) =>
        raw.filter(keep).map((c, i) => tr(c.map(esc), "", i));
      const H = (a) => (multi ? ["Shift"] : []).concat(a);
      if (tab === "bqr") {
        head = H(["ITEM", "QTY", "KET", "QR", "Qty scan", "Unique", "Locator", "Waktu"]);
        numCols = [1, 3, 4].map((i) => i + off);
        uCol = 5 + off;
        kCol = 2 + off;
        let A = filteredRows();
        body = plain(
          A.map((x) =>
            S(x).concat([
              x.item,
              x.qty == null ? "" : x.qty,
              x.ket,
              x.us.length,
              nf(x.sumQ),
              x.us.map((r) => r.unique).join(", "),
              lTxt(x),
              tTxt(x),
            ]),
          ),
        );
      } else if (["move", "throw", "out"].includes(tab)) {
        head = H(["Unique", "Part", "Qty", "Locator", "Waktu", "Keterangan"]);
        numCols = [2 + off];
        uCol = off;
        let A = filteredRows();
        body = plain(
          A.map((x) =>
            S(x).concat([
              x.unique,
              x.part,
              x.qty == null ? "" : x.qty,
              x.loc,
              fmt(x.t),
              x.ket,
            ]),
          ),
        );
      } else if (tab === "extra") {
        head = H(["Unique", "Part", "Qty", "Locator", "Waktu"]);
        numCols = [2 + off];
        uCol = off;
        let A = filteredRows();
        body = plain(
          A.map((x) => S(x).concat([x.unique, x.part, x.qty, x.loc, fmt(x.t)])),
        );
      } else if (tab === "tabin") {
        head = H(["ITEM", "QTY", "KET"]);
        numCols = [1 + off];
        kCol = 2 + off;
        let A = filteredRows();
        body = plain(
          A.map((x) => S(x).concat([x.item, x.qty == null ? "" : x.qty, x.ket])),
        );
      } else if (tab === "md") {
        head = H(["Part", "No. MD", "Qty", "Unique", "Locator", "Waktu"]);
        numCols = [2 + off];
        uCol = 3 + off;
        body = plain(
          filteredRows().map((x) =>
            S(x).concat([
              x.part,
              x.md,
              x.qty == null ? "" : x.qty,
              x.unique,
              x.loc,
              fmt(x.t),
            ]),
          ),
        );
      } else {
        head = H(["ITEM", "QTY", "KET", "Status", "Unique", "Locator", "Waktu"]);
        numCols = [1 + off];
        uCol = 4 + off;
        kCol = 2 + off;
        const A = filteredRows();
        body = A.filter((x) =>
          keep([
            x.item,
            x.qty,
            x.ket,
            x.us.map((r) => r.unique).join(" "),
            lTxt(x),
            x.status === "miss" ? "belum masuk" : "ok",
          ]),
        ).map((x, i) =>
          tr(
            S(x)
              .map(esc)
              .concat([
                esc(x.item),
                esc(x.qty == null ? "" : x.qty),
                esc(x.ket),
                BADGE[x.status],
                esc(uTxt(x)),
                esc(lTxt(x)),
                tTxt(x),
              ]),
            x.status === "miss" ? "miss" : "",
            i,
          ),
        );
      }
      $("th").parentNode.classList.toggle("uoff", !uOpen);
      $("th").parentNode.classList.toggle("koff", !kOpen);
      const partIdx = head.findIndex(
        (h) => h.toUpperCase() === "ITEM" || h.toUpperCase() === "PART",
      );
      const mdIdx = tab === "md" ? head.indexOf("No. MD") : -1;
      const locIdx =
        tab === "miss" || tab === "tabin"
          ? -1
          : head.findIndex((h) => h.toUpperCase() === "LOCATOR");
      $("th").innerHTML =
        "<tr>" +
        head
          .map((h, i) => {
            const fcol =
              i === partIdx ? "part" : i === mdIdx ? "md" : i === locIdx ? "loc" : null;
            if (fcol) {
              const active =
                fcol === "part" ? filterItem : fcol === "md" ? filterMd : filterLoc;
              return (
                '<th class="th-filter-th' + (numCols.includes(i) ? " n" : "") + '" title="Klik untuk filter ' + esc(h) + '">' +
                '<div class="th-box">' +
                '<span>' + esc(h) + '</span>' +
                '<button type="button" class="th-filter-btn' + (active ? " on" : "") + '" data-fcol="' + fcol + '" title="Filter ' + esc(h) + '">' +
                (active ? "▼" : "▾") +
                '</button>' +
                '</div>' +
                '</th>'
              );
            }
            return (
              "<th" +
              tcls(i) +
              ">" +
              (i === uCol
                ? tog("u", "Unique", uOpen)
                : i === kCol
                  ? tog("k", "KET", kOpen)
                  : esc(h)) +
              "</th>"
            );
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

    function updateFilterChips() {
      const el = $("fchips");
      if (!el) return;
      const chips = [];
      if (filterItem) {
        chips.push(
          '<span class="fchip">' +
            (isItemTab() ? "ITEM" : "PART") +
            ': <b>' +
            esc(filterItem) +
            '</b><button type="button" class="fchip-del" data-clear="item" title="Hapus filter Part">×</button></span>',
        );
      }
      if (filterMd) {
        chips.push(
          '<span class="fchip">NO. MD: <b>' +
            esc(filterMd) +
            '</b><button type="button" class="fchip-del" data-clear="md" title="Hapus filter No. MD">×</button></span>',
        );
      }
      if (filterLoc) {
        chips.push(
          '<span class="fchip">LOKATOR: <b>' +
            esc(filterLoc) +
            '</b><button type="button" class="fchip-del" data-clear="loc" title="Hapus filter Locator">×</button></span>',
        );
      }
      el.innerHTML = chips.join("");
    }

    function openFilterPop(anchorEl, col) {
      const pop = $("filterPop");
      if (!pop || !res) return;
      if (pop.style.display !== "none" && activePopCol === col) {
        closeFilterPop();
        return;
      }
      activePopCol = col;
      const titleEl = $("fpTitle");
      const searchEl = $("fpSearch");
      const optEl = $("fpOptions");
      searchEl.value = "";
      const label =
        col === "part"
          ? isItemTab() ? "ITEM" : "Part Number"
          : col === "md" ? "No. MD" : "Locator";
      titleEl.textContent = "Filter " + label;
      searchEl.placeholder = "Cari " + label + "...";

      const { total, values } = getFilterValues(col);
      const cur =
        col === "part" ? filterItem : col === "md" ? filterMd : filterLoc;
      let html =
        '<div class="fp-opt' +
        (!cur ? " sel" : "") +
        '" data-val=""><span>(Pilih Semua / Reset)</span><small>' +
        total +
        "</small></div>";
      for (const it of values) {
        html +=
          '<div class="fp-opt' +
          (cur === it.val ? " sel" : "") +
          '" data-val="' +
          esc(it.val) +
          '">' +
          "<span>" +
          esc(it.val) +
          "</span><small>" +
          it.count +
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
      activePopCol = null;
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
      if (tagWB) run();
    }
    // Baca kolom A dari sheet bernama TAB_IN (kalau file cuma punya 1 sheet, sheet itu dipakai)
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
    // Upload Data.xlsx (kolom A sheet TAB_IN). Tetap terpakai sampai halaman ditutup / diganti file baru.
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
        if (tagWB) run();
        showErr("TAB_IN: " + e.message);
      }
    }

    function download() {
      if (!res) return;
      const wb = XLSX.utils.book_new(),
        st = { ok: "OK", bqr: "Beda QR", miss: "Belum masuk" };
      const add = (rows, name) =>
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), name);
      const hasil = (x) => ({
        Shift: x.shift,
        ITEM: x.item,
        QTY: x.qty,
        KET: x.ket,
        Status: st[x.status],
        QR: x.us.length || "",
        "Qty scan": x.us.length ? x.sumQ : "",
        Unique: x.us.map((r) => r.unique).join(", "),
        Locator: lTxt(x),
        Waktu: tTxt(x),
      });
      add(res.list.map(hasil), "Hasil");
      add(res.list.filter((x) => x.status === "bqr").map(hasil), "Beda QR");
      add(
        res.extra.map((x) => ({
          Shift: x.shift,
          Unique: x.unique,
          Part: x.part,
          Qty: x.qty,
          Locator: x.loc,
          Waktu: fmt(x.t),
        })),
        "Tidak di ID Tag",
      );
      add(
        res.skip.map((x) => ({
          Shift: x.shift,
          Unique: x.unique,
          Part: x.part,
          Qty: x.qty,
          Locator: x.loc,
          Waktu: fmt(x.t),
          Keterangan: x.ket,
          Jenis: x.cats.map((c) => c.toUpperCase()).join("+"),
        })),
        "Move-Throw-Out",
      );
      add(
        res.tabin.map((x) => ({
          Shift: x.shift,
          ITEM: x.item,
          QTY: x.qty,
          KET: x.ket,
        })),
        "TAB_IN",
      );
      add(
        res.md.map((x) => ({
          Shift: x.shift,
          Part: x.part,
          "No. MD": x.md,
          Qty: x.qty,
          Unique: x.unique,
          Locator: x.loc,
          Waktu: fmt(x.t),
        })),
        "ITEM MD",
      );
      XLSX.writeFile(wb, "Cek_" + res.sheet + ".xlsx");
    }

    /* ---------- daftar file scan ---------- */
    function renderFiles() {
      const box = $("fl");
      box.innerHTML = files
        .map(
          (f) =>
            '<div class="fc" data-id="' +
            f.id +
            '"><span class="fn" title="' +
            esc(f.name) +
            '">' +
            esc(f.name) +
            '</span><span class="note" style="margin:0">' +
            nf(f.rows.length) +
            " baris</span>" +
            '<select data-sh="' +
            f.id +
            '" aria-label="Shift">' +
            [1, 2, 3]
              .map(
                (n) =>
                  '<option value="' +
                  n +
                  '"' +
                  (f.shift === n ? " selected" : "") +
                  ">Shift " +
                  n +
                  "</option>",
              )
              .join("") +
            '</select><button class="x" data-rm="' +
            f.id +
            '" aria-label="Hapus">×</button></div>',
        )
        .join("");
      $("dz2").classList.toggle("ok", files.length > 0);
      $("rtrx").hidden = !files.length;
      $("m2").textContent = files.length
        ? files.length + " file dimuat · klik untuk tambah"
        : "Klik atau seret file (boleh lebih dari satu)";
    }

    /* ---------- upload ---------- */
    async function loadTag(list) {
      const file = list[0];
      if (!file) return;
      try {
        const wb = XLSX.read(await file.arrayBuffer(), { type: "array" });
        const nums = wb.SheetNames.filter(isNumSheet);
        if (!nums.length)
          throw new Error(
            "Tidak ada sheet yang namanya angka saja. Sheet ada: " +
            wb.SheetNames.join(", "),
          );
        tagWB = wb;
        $("sheet").innerHTML = nums
          .map(
            (n) => '<option value="' + esc(n) + '">' + esc(n) + "</option>",
          )
          .join("");
        $("sheet").disabled = false;
        $("dz1").classList.add("ok");
        $("m1").textContent = file.name + " · " + nums.length + " sheet";
        autoSelect();
        showErr("");
        run();
      } catch (e) {
        tagWB = null;
        $("dz1").classList.remove("ok");
        $("sheet").disabled = true;
        showErr(e.message);
        run();
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
            fdate: fileDate(file.name, t.minT),
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
      filterItem = "";
      filterLoc = "";
      filterMd = "";
      closeFilterPop();
      autoMsg = "";
      renderFiles();
      run();
    }
    function resetAll() {
      tagWB = null;
      files = [];
      res = null;
      tab = "all";
      filterItem = "";
      filterLoc = "";
      filterMd = "";
      closeFilterPop();
      q = "";
      sf = "all";
      autoMsg = "";
      $("q").value = "";
      $("dz1").classList.remove("ok");
      $("m1").textContent = "Klik atau seret file";
      $("sheet").innerHTML = "<option>-</option>";
      $("sheet").disabled = true;
      renderFiles();
      showErr("");
      run();
    }
    function removeFile(id) {
      const el = document.querySelector('.fc[data-id="' + id + '"]');
      const done = () => {
        files = files.filter((f) => f.id !== id);
        renderFiles();
        autoSelect();
        run();
      };
      if (el && !RM) {
        el.classList.add("out");
        setTimeout(done, 200);
      } else done();
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
    bindDrop($("dz1"), $("f1"), loadTag);
    bindDrop($("dz2"), $("f2"), loadTrx);
    bindDrop($("dz3"), $("f3"), loadTabInUp);
    window.addEventListener("dragover", (e) => e.preventDefault());
    window.addEventListener("drop", (e) => e.preventDefault());

    $("sheet").onchange = () => {
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
    $("fl").onclick = (e) => {
      const b = e.target.closest("[data-rm]");
      if (b) removeFile(+b.dataset.rm);
    };
    $("fl").onchange = (e) => {
      const s = e.target.closest("[data-sh]");
      if (!s) return;
      const f = files.find((x) => x.id === +s.dataset.sh);
      if (f) {
        f.shift = +s.value;
        files.sort(
          (a, b) => a.shift - b.shift || a.name.localeCompare(b.name),
        );
        renderFiles();
        run();
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
      const fb = th && th.querySelector("[data-fcol]");
      if (fb) {
        e.stopPropagation();
        openFilterPop(th, fb.dataset.fcol);
        return;
      }
      const b = e.target.closest("[data-c]");
      if (!b) return;
      const k = b.dataset.c;
      if (k === "u") uOpen = !uOpen;
      else kOpen = !kOpen;
      rows(false, (k === "u" ? uOpen : kOpen) ? k : null);
    };
    $("fold").onclick = (e) => {
      const btn = e.target.closest(".fchip-del");
      if (!btn) return;
      const c = btn.dataset.clear;
      if (c === "loc") filterLoc = "";
      else if (c === "md") filterMd = "";
      else filterItem = "";
      closeFilterPop();
      rows(false);
    };
    const fpop = $("filterPop");
    if (fpop) {
      fpop.onclick = (e) => {
        const opt = e.target.closest(".fp-opt");
        if (!opt) return;
        const v = opt.dataset.val || "";
        if (activePopCol === "loc") filterLoc = v;
        else if (activePopCol === "md") filterMd = v;
        else filterItem = v;
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
      const st = () => { };  /* tanpa keterangan status (estetika); error tetap lewat showErr */

      /* simpan handle folder supaya cukup pilih sekali */
      const dbOpen = () => new Promise((res, rej) => {
        const r = indexedDB.open("idtag_auto", 1);
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

      /* helper folder: pakai lookup langsung (cepat), listing hanya kalau nama tidak pasti */
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
        const w = d.getDay(); // 0 = Minggu
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

      /* satu handle tersimpan per lokasi (key), karena ID Tag dan Monitoring ada di share yang berbeda */
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
      /* terima folder yang dipilih langsung, atau folder induknya (nama dicari 1 level ke bawah) */
      async function resolveDir(h, name) {
        if (h.name.toLowerCase() === name.toLowerCase()) return h;
        return sub(h, name);
      }

      /* ID Tag: 01 REKAP SJ 480 DATA ENTRY FG > folder tahun > ID TAG > folder bulan > file Week Ending */
      async function findIdTag(rekap, sat, now) {
        const yrs = [...new Set([sat.getFullYear(), now.getFullYear()])];
        const top = await ls(rekap);
        let yd = null;
        for (const y of yrs) {
          const e = top.find((x) => x.h.kind === "directory" && x.n.includes(String(y)));
          if (e) { yd = e.h; break; }
        }
        if (!yd) throw new Error("Folder tahun " + yrs[0] + " tidak ditemukan di 01 REKAP SJ 480 DATA ENTRY FG.");
        let idd = await sub(yd, "ID TAG");
        if (!idd) {
          const e = (await ls(yd)).find((x) => x.h.kind === "directory" && /id\s*[-_]?\s*tag/i.test(x.n));
          idd = e && e.h;
        }
        if (!idd) throw new Error("Folder ID TAG tidak ditemukan di " + yd.name + ".");

        const months = [...new Set([sat.getMonth(), now.getMonth()])];
        const ents = await ls(idd);
        const cands = [];
        for (const m of months) {
          const re = monRe(m);
          for (const e of ents) if (e.h.kind === "directory" && re.test(e.n)) await collect(e.h, 2, cands);
        }
        for (const e of ents) if (e.h.kind === "file" && /\.xls[xm]?$/i.test(e.n) && !e.n.startsWith("~$")) cands.push(e);

        const hits = cands.map((c) => ({ c, s: matchDate(c.n, sat) })).filter((x) => x.s > 0);
        if (!hits.length) {
          throw new Error("File ID Tag untuk Week Ending " + dstr(sat) + " tidak ditemukan. File yang terlihat: " +
            (cands.slice(0, 8).map((c) => c.n).join(", ") || "tidak ada"));
        }
        const top2 = Math.max(...hits.map((x) => x.s));
        let best = null;
        for (const x of hits.filter((x) => x.s === top2)) {
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

      /* Scan Transaksi: Data > WE_ ... (belum lengkap) | Data (lepas) | tahun\bulan\WE_ ... (arsip). Lookup langsung per shift.
         Week Ending dihitung per shift (Minggu Shift 1 masuk WE sebelumnya), sama dengan aturan di VBA. */
      async function findTrx(mon, sd) {
        const data = await sub(mon, "Data");
        if (!data) throw new Error('Folder "Data" tidak ditemukan di Monitoring.');
        const dd = pad(sd.getDate()), mm = pad(sd.getMonth() + 1), yyyy = String(sd.getFullYear()), yy = yyyy.slice(2);
        const cache = {};
        const memo = async (k, get) => (k in cache ? cache[k] : (cache[k] = await get()));
        const names = (s) => [
          "Trx_Current_" + dd + "-" + mm + "-" + yyyy + "_Shift" + s + ".xlsx",
          "Trx_Current_" + dd + "-" + mm + "-" + yy + "_Shift" + s + ".xlsx",
        ];
        const one = async (s) => {
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
            for (const n of names(s)) { const h = await fil(d, n); if (h) return h; }
          }
          return null;
        };
        const hs = await Promise.all([1, 2, 3].map(one));
        const out = [];
        for (let i = 0; i < 3; i++) if (hs[i]) out.push({ shift: i + 1, file: await hs[i].getFile() });
        return out;
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
          const sh = h >= 7 && h < 15 ? 1 : h >= 15 && h < 23 ? 2 : 3;
          const sd = new Date(now.getFullYear(), now.getMonth(), now.getDate() - (h < 7 ? 1 : 0)); // tanggal shift
          const sat = weSat(sd, sh);

          /* lokasi 1: \\192.168.1.2\data\logistik\01 REKAP SJ 480 DATA ENTRY FG (ID Tag) */
          const rHandle = await getRoot("rekap", "01 REKAP SJ 480 DATA ENTRY FG");
          const rekap = await resolveDir(rHandle, "01 REKAP SJ 480 DATA ENTRY FG");
          if (!rekap) {
            await forget("rekap");
            throw new Error("Folder yang dipilih bukan 01 REKAP SJ 480 DATA ENTRY FG. Klik lagi dan pilih ulang.");
          }
          /* lokasi 2: \\192.168.1.2\mis\qr_logistic\Monitoring (Tab In.xlsx dan Data) */
          const mHandle = await getRoot("monitoring", "Monitoring (qr_logistic\\Monitoring)");
          const mon = await resolveDir(mHandle, "Monitoring");
          if (!mon) {
            await forget("monitoring");
            throw new Error("Folder yang dipilih bukan Monitoring. Klik lagi dan pilih ulang.");
          }

          st("Mencari file…");
          const [tagFile, tabFile, trx0] = await Promise.all([
            findIdTag(rekap, sat, now),
            findTabIn(mon),
            findTrx(mon, sd),
          ]);
          let trx = trx0, useD = sd, fb = false;

          st("Memuat data…");
          resetTrx();
          await loadTag([tagFile]);
          if (tabFile) await loadTabInUp([tabFile]);

          /* Kalau scan hari ini belum ada: pakai hari terakhir yang punya sheet di file ID Tag (sheet DDMM, sebelum hari ini) */
          if (!trx.length && tagWB) {
            st("Scan hari ini belum ada, mencari hari terakhir di ID Tag…");
            const cand = tagWB.SheetNames
              .map((n) => n.trim())
              .filter((n) => /^\d{3,4}$/.test(n))
              .map((n) => {
                n = n.padStart(4, "0");
                const D = +n.slice(0, 2), M = +n.slice(2);
                if (D < 1 || D > 31 || M < 1 || M > 12) return null;
                let d = new Date(sd.getFullYear(), M - 1, D);
                if (d > sd) d = new Date(sd.getFullYear() - 1, M - 1, D);
                return d;
              })
              .filter((d) => d && d < sd)
              .sort((a, b) => b - a);
            for (const d of cand.slice(0, 14)) {
              const r = await findTrx(mon, d);
              if (r.length) { trx = r; useD = d; fb = true; break; }
            }
          }
          if (!trx.length) throw new Error("File scan untuk tanggal " + dstr(sd) + " dan hari-hari sebelumnya di ID Tag belum ada di folder Data.");

          await loadTrx(trx.map((x) => x.file));

          const got = trx.map((x) => x.shift);
          const lost = [1, 2, 3].filter((s) => !got.includes(s));
          st(
            "ID Tag: " + tagFile.name + "\n" +
            (tabFile ? "Tab In: OK" : "Tab In.xlsx tidak ditemukan") + "\n" +
            "Scan " + dstr(useD) + (fb ? " (hari terakhir di ID Tag, hari ini belum ada)" : "") + ": Shift " + got.join(", ") +
            (lost.length ? " (Shift " + lost.join(", ") + " belum ada)" : "")
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
