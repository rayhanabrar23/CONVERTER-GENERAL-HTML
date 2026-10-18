/* PEI Core — port dari "Daily RP and LR Calculation.py" (Streamlit).
 * Murni logika (tanpa DOM/File API) supaya bisa dites di Node dan dipakai browser.
 * Semua input sudah berupa teks / array of objects (hasil parse XLSX di lapisan UI). */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.PEICore = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const RATIO_THRESHOLD = 0.65;
  const AUTO_ADJUST_TARGET = 0.63;
  const LR1 = 'LR-1. Volume Buy ≤ Available Qty';
  const LR3 = 'LR-3. Nilai Beli ≤ Max LR Final';
  const LR4 = 'LR-4. LR Final ≤ Max LR Final';

  const dict = () => Object.create(null);
  const copyDict = (o) => Object.assign(dict(), o);

  // ── util gaya Python ──
  const FLOAT_RE = /^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/;
  function pyFloat(s) { // float() Python; NaN jika gagal
    if (typeof s === 'number') return s;
    const t = String(s == null ? '' : s).trim();
    return FLOAT_RE.test(t) ? Number(t) : NaN;
  }
  const f0 = (s) => { const v = pyFloat(s); return Number.isNaN(v) ? 0 : v; };
  const toNumeric = (s) => pyFloat(String(s == null ? '' : s).replace(/,/g, ''));
  const splitLines = (c) => String(c).trim().split(/\r\n|\r|\n/);
  const str = (v) => (v == null ? '' : String(v)).trim();

  const fmtRp = (v) => 'Rp ' + Math.round(Number(v)).toLocaleString('en-US').replace(/,/g, '.');
  const fmtPct = (v) => (Number(v) * 100).toFixed(2) + '%';
  const toInt = (v) => Math.floor(Number(v) + 1e-6); // buang sisa floating point (7845750.000000001 → 7845750)

  /** Baca nilai Rupiah dari ketikan / sel Excel. Rupiah harus bilangan bulat ≥ 0.
   * Terima: 7845750 | 7.845.750 | 7,845,750 | Rp 7.845.750 | 7.845.750,00
   * Tolak : 7.845.750,50 | 7845750.5 (desimal), -5, teks lain. Kosong → {value:null}. */
  function parseRupiah(raw) {
    if (raw == null) return { value: null };
    if (typeof raw === 'number') {
      if (!Number.isFinite(raw)) return { error: 'bukan angka' };
      if (raw < 0) return { error: 'tidak boleh negatif' };
      const r = Math.round(raw);
      return Math.abs(raw - r) < 1e-6 ? { value: r } : { error: 'ada desimal — Rupiah harus bilangan bulat' };
    }
    let s = String(raw).trim(); if (!s) return { value: null };
    s = s.replace(/^rp\.?\s*/i, '').replace(/\s+/g, '');
    if (s.startsWith('-')) return { error: 'tidak boleh negatif' };
    if (/^\d+$/.test(s)) return { value: Number(s) };
    if (/^\d{1,3}(\.\d{3})+$/.test(s)) return { value: Number(s.replace(/\./g, '')) };
    if (/^\d{1,3}(,\d{3})+$/.test(s)) return { value: Number(s.replace(/,/g, '')) };
    const m = s.match(/^(\d{1,3}(?:([.,])\d{3})+|\d+)([.,])(\d+)$/);
    if (m && m[3] !== m[2]) {
      if (/^0+$/.test(m[4])) return { value: Number(m[1].replace(/[.,]/g, '')) };
      return { error: 'ada desimal — Rupiah harus bilangan bulat (contoh: 7845750 atau 7.845.750)' };
    }
    return { error: 'format angka tidak dikenali (contoh: 7845750 atau 7.845.750)' };
  }

  // ── rename kolom (find_and_rename) ──
  const ALIASES = {
    stock_key: ['no_share', 'no_shares', 'Stock Code', 'Stockcode', 'Stock', 'SYMBOL', 'StockCode'],
    sid_key: ['SID', 'SID_No', 'Client_SID'],
    cid_key: ['no_cust', 'CID', 'Client_ID', 'Account_No'],
    avail_risk: ['Available Quantity', 'availablequantity', 'Available Qty', 'AvailableQuantity'],
    name_key: ['Name', 'Client_Name', 'Nama'],
    haircut_key: ['Haircut', 'haircut', 'HC'],
    margin_flag: ['Margin', 'margin', 'MARGIN', 'flag_margin', 'MarginFlag'],
  };
  function headersOf(rows) {
    const h = []; const seen = new Set();
    for (const r of rows) for (const k of Object.keys(r)) if (!seen.has(k)) { seen.add(k); h.push(k); }
    return h;
  }
  function findAndRename(rows) {
    const headers = headersOf(rows); const map = {};
    for (const [official, aliases] of Object.entries(ALIASES)) {
      for (const c of headers) if (aliases.includes(String(c).trim())) { map[c] = official; break; }
    }
    const out = rows.map((r) => { const n = {}; for (const [k, v] of Object.entries(r)) n[map[k] || k] = v; return n; });
    return { rows: out, headers: headers.map((h) => map[h] || h) };
  }

  // ── perhitungan ──
  function calcCollateral(stocks, cp, hc) {
    let total = 0; const detail = [];
    for (const stock of Object.keys(stocks)) {
      const qty = stocks[stock];
      const price = cp.has(stock) ? cp.get(stock) : 0;
      const h = hc.has(stock) ? hc.get(stock) : 0.05;
      const coll = qty * price * (1 - h);
      total += coll; detail.push({ stock, qty, cp: price, hc: h, collateral: coll });
    }
    return [total, detail];
  }
  function solveRpMin(loanEx, accrued, collAfterRp) {
    const rp = collAfterRp > 0 ? (loanEx + accrued) - collAfterRp * RATIO_THRESHOLD : loanEx + accrued;
    return Math.min(Math.max(rp, 0), loanEx + accrued); // v5: batas = RP Max (accrued ikut terbayar)
  }
  const solveRpMax = (loanEx, accrued) => loanEx + accrued;
  function solveLr(collAfterLr, loanAfterRp, accrued, availLim, totalRp, totalBuyVal) {
    const lr63 = collAfterLr > 0 ? Math.max(collAfterLr * AUTO_ADJUST_TARGET - (loanAfterRp + accrued), 0) : 0;
    const availEf = availLim + totalRp;
    const capBuy = totalBuyVal * 1.1;
    const cands = [['Rasio 63%', lr63], ['Available Limit', availEf], ['Nilai Beli x110%', capBuy]];
    let best = cands[0];
    for (const c of cands) if (c[1] < best[1]) best = c; // min() Python: elemen pertama bila sama
    return [lr63, availEf, best[1], best[0]];
  }
  function getLoanStatus(availQty, volBuy) {
    if (volBuy <= 0) return '';
    if (availQty < 0) return '';
    return availQty > volBuy ? 'LOAN PEI' : 'LOAN PARTIAL';
  }

  // ── parser file ──
  function loadClosingPrice(rows) {
    const out = new Map();
    for (const r of rows) {
      const code = str(r['no_share']).toUpperCase();
      const price = toNumeric(r['kurs_now']);
      if (!Number.isNaN(price) && code && code !== 'NAN') out.set(code, price);
    }
    return out;
  }
  function loadRiskParameter(content) {
    const hc = new Map(), avq = new Map();
    for (let line of splitLines(content)) {
      line = line.trim();
      if (!line || line.startsWith('StockCode')) continue;
      const p = line.split('|');
      if (p.length < 4) continue;
      const code = p[0].trim().toUpperCase();
      const h = pyFloat(p[2]); const a = pyFloat(p[3]);
      hc.set(code, Number.isNaN(h) ? 0 : h / 100); avq.set(code, Number.isNaN(a) ? 0 : a);
    }
    return { hc, avq };
  }
  function parseOpFile(content) {
    const result = new Map(); const opBrokers = new Set(); result.brokers = opBrokers; // kode partisipan di header OP
    for (let line of splitLines(content)) {
      line = line.trim(); if (!line) continue;
      const p = line.split('|');
      if (p[0] === '0') {
        if (p.length < 7) continue;
        const sid = p[3].trim(); if (p[2] && p[2].trim()) opBrokers.add(p[2].trim().toUpperCase());
        result.set(sid, {
          loan_existing: f0(p[5]), accrued_interest: f0(p[6]),
          available_limit: p.length > 7 ? f0(p[7]) : 0,
          name: p.length > 4 ? p[4].trim() : sid, stocks: dict(),
        });
      } else if (p[0] === '1') {
        if (p.length < 5) continue;
        const sid = p[2].trim(); const stock = p[3].trim().toUpperCase(); const vol = f0(p[4]);
        if (result.has(sid) && stock && vol > 0) {
          const s = result.get(sid).stocks; s[stock] = (s[stock] || 0) + vol;
        }
      }
    }
    return result;
  }
  function parseCreditLimitFile(content) {
    const result = new Map(); let valueDate = null; const brokers = new Set();
    splitLines(content).forEach((raw, i) => {
      const line = raw.trim(); if (!line) return;
      const p = line.split('|');
      if (i === 0 && p[0].trim().toLowerCase() === 'value date') return;
      if (p.length < 7) return;
      const sid = p[2].trim();
      const v = pyFloat(p[6].replace(/,/g, ''));
      if (valueDate === null) valueDate = p[0].trim();
      if (p[1].trim()) brokers.add(p[1].trim().toUpperCase());
      result.set(sid, { available_limit: Number.isNaN(v) ? 0 : v, name: p[3].trim(), value_date: p[0].trim() });
    });
    return { data: result, valueDate, brokers };
  }
  function loadParticipantStockPosition(content) {
    const result = new Map();
    splitLines(content).forEach((raw, i) => {
      const line = raw.trim(); if (!line) return;
      if (i === 0 && line.toLowerCase().startsWith('date')) return;
      const p = line.split('|'); if (p.length < 9) return;
      const sid = p[2].trim(); const stock = p[4].trim().toUpperCase();
      const receive = f0(p[7]), deliver = f0(p[8]); const delta = receive - deliver;
      if (delta === 0) return;
      if (!result.has(sid)) result.set(sid, dict());
      result.get(sid)[stock] = { receive, deliver, delta };
    });
    return result;
  }
  function loadLrBelumSettle(content) {
    const result = new Map();
    for (let line of splitLines(content)) {
      line = line.trim(); if (!line) continue;
      const p = line.split('|');
      if (p[0] !== '0' || p.length < 8) continue;
      const sid = p[2].trim();
      result.set(sid, (result.get(sid) || 0) + f0(p[p.length - 1]));
    }
    return result;
  }
  function cloneOp(op) {
    const out = new Map();
    for (const [sid, d] of op) out.set(sid, Object.assign({}, d, { stocks: copyDict(d.stocks) }));
    return out;
  }
  function applyLrAdjustmentToOp(op, lr) {
    const out = cloneOp(op);
    for (const [sid, v] of lr) if (out.has(sid)) out.get(sid).loan_existing += v;
    return out;
  }
  function applyPositionAdjustmentToOp(op, pos) {
    const out = cloneOp(op);
    for (const [sid, adj] of pos) {
      if (!out.has(sid)) continue;
      const stocks = out.get(sid).stocks;
      for (const stock of Object.keys(adj)) {
        const nq = (stocks[stock] || 0) + adj[stock].delta;
        if (nq <= 0) delete stocks[stock]; else stocks[stock] = nq;
      }
    }
    return out;
  }

  // Netting Invoice. rows = array of objects (header → nilai). Kembalikan juga jejak untuk diagnostik.
  function parseNettingInvoice(rowsIn) {
    const { rows, headers } = findAndRename(rowsIn);
    for (const req of ['stock_key', 'sid_key', 'cid_key', 'bors', 'tot_vol'])
      if (!headers.includes(req)) throw new Error(`Kolom '${req}' tidak ditemukan di file Netting Invoice.`);
    const hasBoard = headers.includes('board'), hasLorf = headers.includes('lorf');
    const valCol = headers.includes('amt_done') ? 'amt_done' : (headers.includes('amt_pay') ? 'amt_pay' : null);
    if (!valCol) throw new Error('Kolom nilai transaksi (amt_done/amt_pay) tidak ditemukan di file Netting Invoice.');
    const hasName = headers.includes('name');

    const netting = new Map(), cidToSid = new Map(), sidToName = new Map();
    const trace = { rawRows: new Map(), droppedBoard: new Map(), droppedLorf: new Map(), rawSids: new Set() };
    const inc = (m, k) => m.set(k, (m.get(k) || 0) + 1);
    for (const r of rows) {
      const sid = str(r.sid_key), cid = str(r.cid_key), stock = str(r.stock_key).toUpperCase();
      const side = str(r.bors).toUpperCase();
      if (sid) { trace.rawSids.add(sid); inc(trace.rawRows, sid); }
      if (hasBoard && str(r.board).toUpperCase() !== 'RG') { if (sid) inc(trace.droppedBoard, sid); continue; }
      if (hasLorf && str(r.lorf).toUpperCase() !== 'D') { if (sid) inc(trace.droppedLorf, sid); continue; }
      if (!sid || !stock) continue;
      const vol = f0(String(r.tot_vol == null ? '' : r.tot_vol).replace(/,/g, ''));
      const val = f0(String(r[valCol] == null ? '' : r[valCol]).replace(/,/g, ''));
      if (cid) cidToSid.set(cid, sid);
      if (hasName && !sidToName.has(sid)) { const nm = str(r.name); if (nm && nm.toLowerCase() !== 'nan') sidToName.set(sid, nm); }
      if (!netting.has(sid)) netting.set(sid, dict());
      const m = netting.get(sid);
      const d = m[stock] || (m[stock] = { buy_lot: 0, sell_lot: 0, buy_value: 0, sell_value: 0 });
      if (side === 'B') { d.buy_lot += vol; d.buy_value += val; }
      else if (side === 'S') { d.sell_lot += vol; d.sell_value += val; }
    }
    for (const m of netting.values()) for (const d of Object.values(m)) {
      d.net_lot = d.buy_lot - d.sell_lot; d.net_value = d.buy_value - d.sell_value;
    }
    return { netting, cidToSid, sidToName, trace };
  }
  function splitNetting(netting, cp) {
    const netBuy = new Map(), netSell = new Map();
    for (const [sid, stocks] of netting) {
      for (const stock of Object.keys(stocks)) {
        const d = stocks[stock]; const netLot = d.buy_lot - d.sell_lot;
        const price = cp.has(stock) ? cp.get(stock) : 0;
        if (netLot > 0) { if (!netBuy.has(sid)) netBuy.set(sid, dict()); netBuy.get(sid)[stock] = { lot: netLot, value: d.buy_lot * price }; }
        else if (netLot < 0) { if (!netSell.has(sid)) netSell.set(sid, dict()); netSell.get(sid)[stock] = { lot: -netLot, value: d.sell_lot * price }; }
      }
    }
    return { netBuy, netSell };
  }

  // SID Client (xlsx rows) → peta PEI
  function parseSidClient(rowsIn) {
    const { rows, headers } = findAndRename(rowsIn);
    for (const req of ['sid_key', 'cid_key'])
      if (!headers.includes(req)) throw new Error(`Kolom '${req}' tidak ditemukan di file SID Client.`);
    const sidToCid = new Map(), cidToName = new Map(), pei = new Set(), seen = new Map();
    const hasName = headers.includes('name_key');
    for (const r of rows) {
      const sid = str(r.sid_key), cid = str(r.cid_key);
      sidToCid.set(sid, cid); pei.add(sid);
      if (hasName && !cidToName.has(cid)) cidToName.set(cid, str(r.name_key));
      if (!seen.has(sid)) seen.set(sid, []); seen.get(sid).push({ cid, name: hasName ? str(r.name_key) : '' });
    }
    // SID ganda: conflict = CID berbeda (berbahaya), selain itu baris dobel biasa
    const sidDupes = [];
    for (const [sid, list] of seen) if (sid && list.length > 1) {
      const cids = [...new Set(list.map((x) => x.cid))];
      sidDupes.push({ sid, rows: list, conflict: cids.length > 1, used: sidToCid.get(sid) });
    }
    return { sidToCid, cidToName, peiSids: pei, sidDupes };
  }

  /** Bangun konteks dari input mentah.
   * inputs: {closingRows, riskText, opText, nettingRows, clText, sidClientRows, posText?, lrSettleText?, hasilBuyRows?} */
  function buildContext(inputs) {
    const cp = loadClosingPrice(inputs.closingRows);
    const risk = loadRiskParameter(inputs.riskText);
    const opRaw = parseOpFile(inputs.opText);
    let op = opRaw; let pos = new Map(), lrs = new Map();
    if (inputs.posText) { pos = loadParticipantStockPosition(inputs.posText); op = applyPositionAdjustmentToOp(op, pos); }
    if (inputs.lrSettleText) { lrs = loadLrBelumSettle(inputs.lrSettleText); op = applyLrAdjustmentToOp(op, lrs); }
    const net = parseNettingInvoice(inputs.nettingRows);
    const { netBuy, netSell } = splitNetting(net.netting, cp);
    const cl = parseCreditLimitFile(inputs.clText);
    const sc = parseSidClient(inputs.sidClientRows);
    return {
      cp, hc: risk.hc, avq: risk.avq, opRaw, op, pos, lrs, netting: net.netting, trace: net.trace,
      cidToSidInvoice: net.cidToSid, sidToNameInvoice: net.sidToName, netBuy, netSell,
      cl: cl.data, clValueDate: cl.valueDate, brokerCodes: brokerCodesOf(cl.brokers, opRaw.brokers || new Set()),
      sidToCid: sc.sidToCid, cidToName: sc.cidToName, peiSids: sc.peiSids, sidDupes: sc.sidDupes, hasPos: !!inputs.posText,
      hasilBuyRows: inputs.hasilBuyRows || null, // Map sid → array of arrays (sheet LR versi partisipan, ≥14 kolom), opsional
    };
  }

  /** Hitung satu SID.
   * opts.mode: 'pei' (tab TRX PEI Python) | 'val' (tab Validasi Python)
   * opts.rpOverride / opts.lrOverride: nilai RP Final / LR Final hasil edit (sumber kebenaran tunggal)
   * opts.adjusted: Auto-Adjust LR sudah diterapkan untuk SID ini */
  function computeSid(sid, ctx, opts) {
    opts = opts || {}; const mode = opts.mode || 'val';
    const op = ctx.op.get(sid) || { loan_existing: 0, accrued_interest: 0, name: null, stocks: dict() };
    const cl = ctx.cl.get(sid) || { available_limit: 0, name: null };
    const loanEx = op.loan_existing, accrued = op.accrued_interest, availLim = cl.available_limit;
    const cid = ctx.sidToCid.get(sid) || '';
    const name = op.name || ctx.cidToName.get(cid) || cl.name || sid;
    const stocksOp = op.stocks;
    const sellStocks = ctx.netSell.get(sid) || dict();
    const buyStocks = ctx.netBuy.get(sid) || dict();
    const checks = []; const add = (label, passed, detail) => checks.push({ label, passed, detail: detail || '' });

    const rpDetail = []; const stocksAfterRp = copyDict(stocksOp); let totalRpMin = 0;
    const hasRp = Object.keys(sellStocks).length > 0;
    const rpActive = hasRp && loanEx > 0;
    const subtract = (stock, lot) => { stocksAfterRp[stock] = (stocksAfterRp[stock] || 0) - lot; if (stocksAfterRp[stock] <= 0) delete stocksAfterRp[stock]; };

    if (mode === 'pei') {
      for (const stock of Object.keys(sellStocks)) {
        const lotSell = sellStocks[stock].lot, lotOp = stocksOp[stock] || 0;
        const lotKeluar = Math.min(lotSell, lotOp); const price = ctx.cp.has(stock) ? ctx.cp.get(stock) : 0;
        rpDetail.push({ stock, lot_sell: lotSell, lot_op: lotOp, lot_keluar: lotKeluar, price, rp_min: lotKeluar * price, ada_di_op: lotOp > 0 });
      }
      for (const d of rpDetail) { if (d.ada_di_op) totalRpMin += d.rp_min; }
      for (const d of rpDetail) if (d.ada_di_op && d.lot_keluar > 0) subtract(d.stock, d.lot_keluar);
    } else if (rpActive) {
      for (const stock of Object.keys(sellStocks)) {
        const lotSell = sellStocks[stock].lot, lotOp = stocksOp[stock] || 0;
        const lotKeluar = Math.min(lotSell, lotOp); const price = ctx.cp.has(stock) ? ctx.cp.get(stock) : 0;
        if (lotOp > 0) {
          totalRpMin += lotKeluar * price;
          rpDetail.push({ stock, lot_sell: lotSell, lot_op: lotOp, lot_keluar: lotKeluar, price, rp_min: lotKeluar * price, ada_di_op: true });
          subtract(stock, lotKeluar);
        }
      }
    }

    const [collAfterRp] = calcCollateral(stocksAfterRp, ctx.cp, ctx.hc);
    const totalRpMinRatio = (mode === 'pei' || rpActive) ? solveRpMin(loanEx, accrued, collAfterRp) : 0;
    const totalRpMaks = rpActive ? solveRpMax(loanEx, accrued) : 0;
    const rpFinal = (opts.rpOverride != null) ? opts.rpOverride : totalRpMaks;
    const rpCap = Math.round(totalRpMaks);
    const rpOver = opts.rpOverride != null && opts.rpOverride > rpCap;
    // v5: RP Final di atas RP Max tidak boleh menggelembungkan perhitungan (dan gagal RP-4 → tidak masuk iFast)
    const rpCalc = Math.min(rpFinal, totalRpMaks);
    // v5: alokasi RP → pokok dulu, sisanya ke accrued. Accrued yang sudah terbayar tidak dihitung lagi.
    const principalPaid = Math.min(rpCalc, loanEx);
    const loanAfterRp = Math.max(loanEx - rpCalc, 0);
    const accruedAfterRp = Math.max(accrued - Math.max(rpCalc - loanEx, 0), 0);
    const numRp = loanAfterRp + accruedAfterRp;
    const rasioRp = collAfterRp > 0 ? numRp / collAfterRp : null;

    // checks RP
    if (!rpActive) {
      const msg = loanEx <= 0 ? '⏭ Dilewati — Loan Existing = 0' : 'Tidak ada transaksi jual';
      add('RP-1. Pengecekan Saham Jual', true, msg); add('RP-2. Lot Sell ≤ Lot di OP', true, msg); add('RP-3. Rasio After RP < 65%', true, msg);
    } else {
      const rp1 = Object.keys(sellStocks).filter((s) => (stocksOp[s] || 0) === 0);
      add('RP-1. Pengecekan Saham Jual', true, rp1.length ? `ℹ️ Saham reguler (bukan OP): ${rp1.join(', ')}` : '✅ Semua saham jual ada di OP');
      const rp2 = rpDetail.filter((d) => d.lot_sell > d.lot_op && d.ada_di_op).map((d) => `${d.stock}: ${d.lot_sell} → ${d.lot_op}`);
      add('RP-2. Lot Sell ≤ Lot di OP', true, rp2.length ? '⚠️ Auto-adjusted: ' + rp2.join('; ')
        : `Total RP Max (rekomendasi): ${fmtRp(totalRpMaks)} | RP Min (ref): ${fmtRp(totalRpMinRatio)}`);
      if (rasioRp !== null) add('RP-3. Rasio After RP < 65%', rasioRp < RATIO_THRESHOLD, `Rasio: ${fmtPct(rasioRp)} | Numerator: ${fmtRp(numRp)} | Coll: ${fmtRp(collAfterRp)}`);
      else if (numRp <= 0) add('RP-3. Rasio After RP < 65%', true, 'Numerator ≤ 0');
      else add('RP-3. Rasio After RP < 65%', false, 'Collateral = 0');
    }
    if (opts.rpOverride != null && (rpActive || opts.rpOverride > 0)) // hanya muncul bila RP diedit
      add('RP-4. RP Final ≤ RP Max', !rpOver, rpOver ? `❌ RP Final ${fmtRp(rpFinal)} melebihi RP Max ${fmtRp(totalRpMaks)} (Loan + Accrued) → tidak masuk iFast` : `RP Final ${fmtRp(rpFinal)} ≤ RP Max ${fmtRp(totalRpMaks)}`);

    // LR
    const hasLr = Object.keys(buyStocks).length > 0;
    const stocksAfterLr = copyDict(stocksAfterRp); const buyStatus = [];
    for (const stock of Object.keys(buyStocks)) {
      stocksAfterLr[stock] = (stocksAfterLr[stock] || 0) + buyStocks[stock].lot;
      const a = ctx.avq.has(stock) ? ctx.avq.get(stock) : 0;
      buyStatus.push({ stock, lot_beli: buyStocks[stock].lot, available_qty: a, status: getLoanStatus(a, buyStocks[stock].lot) });
    }
    const [collAfterLr] = calcCollateral(stocksAfterLr, ctx.cp, ctx.hc);
    let totalBuyVal = 0; for (const s of Object.keys(buyStocks)) totalBuyVal += buyStocks[s].value;
    // v5: limit yang kembali = pokok yang dibayar (Used Limit di Credit Limit = pokok loan, tanpa accrued)
    const [maxLr63, availEf, maxLrFinal, binding] = solveLr(collAfterLr, loanAfterRp, accruedAfterRp, availLim, principalPaid, totalBuyVal);
    const lrCap = toInt(maxLrFinal);
    const lrFinal = (opts.lrOverride != null) ? opts.lrOverride : lrCap; // v5: default dibulatkan ke bawah (Rupiah bulat)
    const lrOver = opts.lrOverride != null && opts.lrOverride > lrCap;
    const rasioLr = collAfterLr > 0 ? (loanAfterRp + accruedAfterRp + Math.min(totalBuyVal, lrFinal)) / collAfterLr : null;

    if (!hasLr) {
      add(LR1, true, 'Tidak ada Loan Request');
      add('LR-2. Kapasitas LR (63% / Avail Limit / Nilai Beli x110%)', true, 'Tidak ada Loan Request');
      add(LR3, true, 'Tidak ada Loan Request');
    } else {
      let pass1 = true; const det1 = [];
      if (ctx.hasilBuyRows) { // sumber asli: sheet LR versi partisipan (kolom 1=saham, 4=avail, 13=vol)
        for (const row of (ctx.hasilBuyRows.get(sid) || [])) {
          const vol = f0(row[13]), avq = f0(row[4]), stk = String(row[1] == null ? '' : row[1]);
          if (avq === 0) { pass1 = false; det1.push(`${stk}: DIBATALKAN (Available=0)`); }
          else if (vol > avq) { pass1 = false; det1.push(`${stk}: Vol ${vol} > Avail ${avq}`); }
        }
      } else { // tanpa file versi partisipan: pakai Available Qty dari Risk Parameter
        for (const b of buyStatus) {
          if (b.available_qty === 0) { pass1 = false; det1.push(`${b.stock}: DIBATALKAN (Available=0)`); }
          else if (b.lot_beli > b.available_qty) { pass1 = false; det1.push(`${b.stock}: Vol ${b.lot_beli} > Avail ${b.available_qty}`); }
        }
      }
      add(LR1, pass1, det1.length ? det1.join('; ') : `Total Nilai Beli: ${fmtRp(totalBuyVal)}`);
      add('LR-2. Kapasitas LR (63% / Avail Limit / Nilai Beli x110%)', true,
        `LR @63%: ${fmtRp(maxLr63)} | Avail Efektif: ${fmtRp(availEf)} | Nilai Beli x110%: ${fmtRp(totalBuyVal * 1.1)} | Max LR Final: ${fmtRp(lrFinal)} (binding: ${binding})`);
      let pass3 = totalBuyVal <= lrFinal;
      let det3 = `Nilai Beli: ${fmtRp(totalBuyVal)} vs Max LR Final: ${fmtRp(lrFinal)} | Rasio jika full: ${rasioLr !== null ? fmtPct(rasioLr) : 'N/A'}`;
      if (!pass3) {
        if (opts.adjusted) { pass3 = true; det3 = `✂️ Di-adjust ke ${fmtRp(lrFinal)} | Rasio: ${rasioLr !== null ? fmtPct(rasioLr) : 'N/A'}`; }
        else det3 += ` || ⚠️ Potong Loan Request ke: ${fmtRp(lrFinal)}`;
      }
      add(LR3, pass3, det3);
      if (opts.lrOverride != null) // hanya muncul bila LR diedit
        add(LR4, !lrOver, lrOver ? `❌ LR Final ${fmtRp(lrFinal)} melebihi Max LR Final ${fmtRp(lrCap)} (binding: ${binding}) → tidak masuk iFast` : `LR Final ${fmtRp(lrFinal)} ≤ Max LR Final ${fmtRp(lrCap)}`);
    }

    return {
      sid, name, cid: ctx.sidToCid.get(sid) || sid, checks, has_rp: hasRp, has_lr: hasLr, rp_skipped: loanEx <= 0,
      loan_existing: loanEx, accrued, avail_limit: availLim,
      stocks_op: stocksOp, stocks_after_rp: stocksAfterRp, stocks_after_lr: stocksAfterLr, buy_stocks: buyStocks,
      rp_detail: rpDetail, total_rp_maks: totalRpMaks, total_rp_min: totalRpMin, total_rp_min_ratio: totalRpMinRatio,
      rp_final: rpFinal, rp_edited: opts.rpOverride != null, rp_cap: rpCap, rp_over: rpOver,
      loan_after_rp: loanAfterRp, accrued_after_rp: accruedAfterRp, principal_paid: principalPaid, rasio_rp: rasioRp, lr_cap: lrCap, lr_over: lrOver,
      coll_before_rp: calcCollateral(stocksOp, ctx.cp, ctx.hc)[0], coll_after_rp: collAfterRp, coll_after_lr: collAfterLr,
      max_lr_63: maxLr63, avail_efektif: availEf, max_lr_final: maxLrFinal, lr_binding: binding,
      lr_final: lrFinal, lr_edited: opts.lrOverride != null, total_buy_val: totalBuyVal, rasio_lr: rasioLr,
      buy_status_detail: buyStatus, adjusted: !!opts.adjusted,
    };
  }

  const lolosRp = (d) => !d.rp_skipped && d.has_rp && d.checks.filter((c) => c.label.startsWith('RP-')).every((c) => c.passed);
  const lolosLr = (d) => d.has_lr && d.checks.filter((c) => c.label.startsWith('LR-')).every((c) => c.passed);

  // ── state = SATU sumber kebenaran ──
  function newState() { return { rp: new Map(), lr: new Map(), adjusted: new Set(), audit: [] }; }
  function setOverride(state, kind, sid, value, user, note) {
    const prev = state[kind].get(sid);
    if (value == null || Number.isNaN(value)) state[kind].delete(sid); else state[kind].set(sid, value);
    state.audit.push({ ts: new Date().toISOString(), user: user || '-', kind: kind.toUpperCase(), sid, from: prev == null ? null : prev, to: value == null ? null : value, note: note || '' });
  }
  function optsFor(state, sid, mode) {
    return { mode, rpOverride: state.rp.has(sid) ? state.rp.get(sid) : null, lrOverride: state.lr.has(sid) ? state.lr.get(sid) : null, adjusted: state.adjusted.has(sid) };
  }
  function runPei(ctx, state) {
    const out = new Map();
    for (const sid of ctx.peiSids) out.set(sid, computeSid(sid, ctx, optsFor(state, sid, 'pei')));
    return out;
  }
  function runValidasi(ctx, state) {
    const all = new Set([...ctx.op.keys(), ...ctx.cl.keys(), ...ctx.netBuy.keys(), ...ctx.netSell.keys()]);
    const out = new Map();
    for (const sid of [...all].sort()) {
      const r = computeSid(sid, ctx, optsFor(state, sid, 'val')); r.is_pei = ctx.peiSids.has(sid); out.set(sid, r);
    }
    return out;
  }
  /** Kode partisipan yang tertulis di berkas: {codes:[...], fromCl:[...], fromOp:[...]} */
  function brokerCodesOf(fromCl, fromOp) { return { codes: [...new Set([...fromCl, ...fromOp])].sort(), fromCl: [...fromCl].sort(), fromOp: [...fromOp].sort() }; }
  /** clPartisipan: Credit Limit partisipan di PEI (Rupiah). Tanpa nilai → check global tidak bisa dinilai. */
  function globalResult(results, clPartisipan) {
    let rp = 0, lr = 0;
    for (const d of results.values()) { if (lolosLr(d)) lr += d.lr_final; if (lolosRp(d)) rp += d.rp_final; }
    if (clPartisipan == null || Number.isNaN(clPartisipan)) return { passed: null, total_rp: rp, total_lr: lr, detail: `CL Partisipan belum diisi | RP: ${fmtRp(rp)} | LR: ${fmtRp(lr)}` };
    return { passed: clPartisipan + rp > lr, total_rp: rp, total_lr: lr,
      detail: `CL Partisipan: ${fmtRp(clPartisipan)} + RP: ${fmtRp(rp)} = ${fmtRp(clPartisipan + rp)} | LR: ${fmtRp(lr)}` };
  }
  function applyAutoAdjust(state, results, user) {
    const touched = [];
    for (const [sid, d] of results) {
      if (d.has_lr && d.is_pei && d.checks.some((c) => c.label === LR3 && !c.passed)) { state.adjusted.add(sid); touched.push(sid);
        state.audit.push({ ts: new Date().toISOString(), user: user || '-', kind: 'AUTO-ADJUST', sid, from: null, to: d.lr_final, note: 'LR-3 di-adjust ke Max LR Final' }); }
    }
    return touched;
  }

  function cancelAutoAdjust(state, user, sids) {
    const list = sids ? [...sids].filter((s) => state.adjusted.has(s)) : [...state.adjusted];
    for (const sid of list) { state.adjusted.delete(sid); state.audit.push({ ts: new Date().toISOString(), user: user || '-', kind: 'BATAL AUTO-ADJUST', sid, from: null, to: null, note: 'Auto-Adjust dibatalkan' }); }
    return list;
  }
  const stateSig = (s) => JSON.stringify([[...s.rp].sort(), [...s.lr].sort(), [...s.adjusted].sort()]);
  // ── v5: terapkan RP/LR Final dari file (RP_Final di langkah 3, Hasil_TRX_PEI di langkah 4) ──
  function cloneState(s) { return { rp: new Map(s.rp), lr: new Map(s.lr), adjusted: new Set(s.adjusted), audit: s.audit.slice() }; }
  /** Baca kolom nilai dari baris file. Kembalikan {map: sid→Rupiah, errors:[...]} — nilai salah format TIDAK diam-diam jadi 0. */
  function readFinalColumn(rows, col, sheetLabel) {
    const map = new Map(), errors = [];
    for (const r of rows) {
      const sid = str(r.SID); if (!sid) continue;
      const p = parseRupiah(r[col]);
      if (p.error || p.value == null) { errors.push(`${sheetLabel}${sid}: ${col} "${r[col] == null ? '' : r[col]}" ${p.error || 'kosong (isi 0 bila tidak ada)'}`); continue; }
      if (map.has(sid) && map.get(sid) !== p.value) { errors.push(`${sheetLabel}${sid}: ${col} berbeda antar baris (${fmtRp(map.get(sid))} vs ${fmtRp(p.value)})`); continue; }
      map.set(sid, p.value);
    }
    return { map, errors };
  }
  /** vals.rp / vals.lr: Map sid→nilai (SID nasabah PEI yang tidak ada di map dianggap 0).
   * Semua-atau-tidak-sama-sekali: bila ada nilai melebihi RP Max / Max LR Final, state TIDAK diubah. */
  function applyFinals(ctx, state, vals, user, note) {
    const st = cloneState(state); const errors = [], changed = [];
    let pei = runPei(ctx, st);
    for (const kind of ['rp', 'lr']) {
      const m = vals[kind]; if (!m) continue;
      for (const [sid, d] of pei) {
        const nv = m.has(sid) ? m.get(sid) : 0;
        const cap = kind === 'rp' ? d.rp_cap : d.lr_cap;
        if (nv > cap) { errors.push(`${sid} ${d.name}: ${kind.toUpperCase()} Final ${fmtRp(nv)} melebihi ${kind === 'rp' ? 'RP Max' : 'Max LR Final'} ${fmtRp(cap)}`); continue; }
        if (Math.round(d[kind + '_final']) !== Math.round(nv)) { setOverride(st, kind, sid, nv, user, note); changed.push(`${kind.toUpperCase()} ${sid} ${d.name}: ${fmtRp(d[kind + '_final'])} → ${fmtRp(nv)}`); }
      }
      if (kind === 'rp') pei = runPei(ctx, st); // kapasitas LR bergantung pada RP Final
    }
    if (vals.adjusted && !errors.length) { // v5.1: kolom Auto-Adjust di Hasil_TRX_PEI
      for (const [sid, d] of pei) {
        const want = vals.adjusted.has(sid) && d.has_lr, have = st.adjusted.has(sid);
        if (want === have) continue;
        if (want) { st.adjusted.add(sid); st.audit.push({ ts: new Date().toISOString(), user: user || '-', kind: 'AUTO-ADJUST', sid, from: null, to: null, note }); }
        else cancelAutoAdjust(st, user, [sid]);
        changed.push(`Auto-Adjust ${sid} ${d.name}: ${have ? 'Ya' : 'Tidak'} → ${want ? 'Ya' : 'Tidak'}`);
      }
    }
    const notPei = new Set(); for (const k of ['rp', 'lr']) if (vals[k]) for (const sid of vals[k].keys()) if (!ctx.peiSids.has(sid)) notPei.add(sid);
    return { ok: !errors.length, state: errors.length ? state : st, errors, changed, notPei: [...notPei] };
  }
  /** sheets: {nama sheet: array of objects}. Hasil_TRX_PEI buatan aplikasi ini → nilai RP/LR Final per SID.
   * Summary = daftar lengkap; sheet Repayment (RP) & Loan Request (LR) harus sama dengan Summary. */
  function parseHasilTrx(sheets) {
    const sum = sheets['Summary'];
    if (!sum || !sum.length || !['SID', 'RP Final', 'LR Final'].every((c) => c in sum[0]))
      return { ok: false, errors: ['Sheet "Summary" (kolom SID, RP Final, LR Final) tidak ditemukan — ini bukan file Hasil_TRX_PEI dari aplikasi ini.'] };
    const a = readFinalColumn(sum, 'RP Final', 'Summary '), b = readFinalColumn(sum, 'LR Final', 'Summary ');
    const errors = [...a.errors, ...b.errors];
    const seen = new Set(); for (const r of sum) { const sid = str(r.SID); if (seen.has(sid)) errors.push(`Summary: SID ${sid} muncul lebih dari sekali`); seen.add(sid); }
    for (const [sheet, col, ref] of [['Repayment (RP)', 'RP Final', a.map], ['Loan Request (LR)', 'LR Final', b.map]]) {
      const rows = sheets[sheet] || []; const c = readFinalColumn(rows, col, sheet + ' '); errors.push(...c.errors);
      for (const [sid, v] of c.map) {
        if (!ref.has(sid)) errors.push(`${sheet} ${sid}: tidak ada di Summary`);
        else if (ref.get(sid) !== v) errors.push(`${sid}: ${col} di sheet ${sheet} ${fmtRp(v)} ≠ Summary ${fmtRp(ref.get(sid))} — samakan dulu`);
      }
      if (sheet === 'Loan Request (LR)') for (const [sid, v] of ref) if (v > 0 && !c.map.has(sid)) errors.push(`${sid}: LR Final ${fmtRp(v)} di Summary tapi tidak ada di sheet Loan Request (LR)`);
    }
    let adjusted = null;
    if ('Auto-Adjust' in sum[0]) {
      adjusted = new Set(); const yes = (v) => /^(ya|y|yes|true|1)$/i.test(str(v)), no = (v) => /^(tidak|t|no|n|false|0|-|)$/i.test(str(v));
      for (const r of sum) { const sid = str(r.SID); if (!sid) continue; const v = r['Auto-Adjust'];
        if (yes(v)) adjusted.add(sid); else if (!no(v)) errors.push(`Summary ${sid}: Auto-Adjust "${v}" tidak dikenali (isi Ya/Tidak)`); }
      for (const r of sheets['Loan Request (LR)'] || []) { const sid = str(r.SID);
        if ('Auto-Adjust' in r && yes(r['Auto-Adjust']) !== adjusted.has(sid)) errors.push(`${sid}: Auto-Adjust di sheet Loan Request (LR) ≠ Summary — samakan dulu`); }
    }
    return { ok: !errors.length, errors: [...new Set(errors)], rp: a.map, lr: b.map, adjusted, sids: seen };
  }
  /** Terapkan Hasil_TRX_PEI ke state. File harus berasal dari data yang sedang diproses (SID sama). */
  function applyHasilTrx(ctx, state, parsed, user) {
    if (!parsed.ok) return { ok: false, state, errors: parsed.errors, changed: [] };
    const missing = [...ctx.peiSids].filter((s) => !parsed.sids.has(s)), extra = [...parsed.sids].filter((s) => !ctx.peiSids.has(s));
    if (missing.length || extra.length) return { ok: false, state, changed: [], errors: [`File Hasil_TRX_PEI tidak cocok dengan data yang sedang diproses: ${missing.length} nasabah PEI tidak ada di file, ${extra.length} SID di file bukan nasabah PEI hari ini${extra.length ? ' (mis. ' + extra.slice(0, 3).join(', ') + ')' : ''}. Pastikan file berasal dari proses hari ini.`] };
    return applyFinals(ctx, state, { rp: parsed.rp, lr: parsed.lr, adjusted: parsed.adjusted }, user, 'dari Hasil TRX');
  }

  // ── export: SEMUA membaca rp_final / lr_final (nilai hasil edit) ──
  const needCode = (code) => { const c = str(code).toUpperCase(); if (!c) throw new Error('Kode Partisipan belum diisi — file iFast tidak bisa dibuat.'); return c; };
  function ifastRepayment(results, participantCode) {
    const pc = needCode(participantCode); const loan = [], stock = [];
    for (const [sid, d] of results) {
      if (!lolosRp(d)) continue; if (d.rp_final <= 0) continue;
      loan.push({ 'Participant Code': pc, 'SID Client': sid, 'Repayment Value': d.rp_final });
      for (const rd of d.rp_detail) if (rd.lot_keluar > 0) stock.push({ 'SID Client': sid, 'Stock Code': rd.stock, 'Quantity': Math.trunc(rd.lot_keluar) });
    }
    return { 'Repayment Direct': loan, 'Stock': stock };
  }
  function ifastLoan(results, participantCode) {
    const pc = needCode(participantCode); const loan = [], stock = [];
    for (const [sid, d] of results) {
      if (!lolosLr(d)) continue; if (d.lr_final <= 0) continue;
      loan.push({ 'Participant Code': pc, 'SID Client': sid, 'Loan Value': toInt(d.lr_final) });
      for (const s of Object.keys(d.stocks_after_lr)) {
        const lotBeli = d.stocks_after_lr[s] - (d.stocks_after_rp[s] || 0);
        if (lotBeli > 0) stock.push({ 'SID Client': sid, 'Stock Code': s, 'Quantity': Math.trunc(lotBeli) });
      }
    }
    return { 'Loan': loan, 'Stock': stock };
  }
  const pct = (v) => (v !== null && v !== undefined ? fmtPct(v) : 'N/A');
  const failLabels = (d, pre) => d.checks.filter((c) => c.label.startsWith(pre) && !c.passed).map((c) => c.label.split('.')[0]).join(', ');
  function statusRp(d) {
    if (d.rp_skipped) return 'Tidak ada RP (Loan Existing = 0)';
    if (lolosRp(d)) return d.rp_final > 0 ? 'Masuk iFast' : 'RP Final = 0 (tidak dikirim)';
    return 'TIDAK masuk iFast — gagal ' + failLabels(d, 'RP-');
  }
  function statusLr(d) {
    if (d.lr_final <= 0) return 'LR Final = 0 (tidak dikirim)';
    if (lolosLr(d)) return d.adjusted && d.total_buy_val > d.lr_final ? `Masuk iFast — di-Auto-Adjust (Nilai Beli ${fmtRp(d.total_buy_val)} dipotong ke ${fmtRp(d.lr_final)})` : 'Masuk iFast';
    return 'TIDAK masuk iFast — gagal ' + failLabels(d, 'LR-') + (d.checks.some((c) => c.label === LR3 && !c.passed) ? ' (perlu Auto-Adjust)' : '');
  }
  function hasilTrxPei(res) { // pei-mode results → 3 sheet seperti Hasil_TRX_PEI.xlsx
    const rp = [], lr = [], sum = [];
    for (const [sid, d] of res) {
      for (const rd of d.rp_detail) { if (!rd.ada_di_op) continue;
        rp.push({ SID: sid, Nama: d.name, Saham: rd.stock, 'Lot Jual': Math.trunc(rd.lot_sell), 'Lot OP': Math.trunc(rd.lot_op), 'Lot Keluar': Math.trunc(rd.lot_keluar), Harga: rd.price,
          'Nilai Jual (ref)': rd.rp_min, 'RP Min (ref)': d.total_rp_min_ratio, 'RP Final': d.rp_final, 'RP Max (rekomendasi)': d.total_rp_maks,
          'Loan After RP': d.loan_after_rp, 'Rasio After RP': pct(d.rasio_rp), 'Status RP': statusRp(d) }); }
      if (d.lr_final > 0) lr.push({ SID: sid, Nama: d.name, 'RP Final': d.rp_final, 'Avail Limit': d.avail_limit, 'Avail Efektif': d.avail_efektif, 'LR @63%': toInt(d.max_lr_63),
        'Loan After RP': d.loan_after_rp, 'Coll After LR': d.coll_after_lr, 'Rasio LR': pct(d.rasio_lr), 'Max LR Final': toInt(d.max_lr_final), 'LR Final': d.lr_final, Binding: d.lr_binding,
        'Nilai Beli': d.total_buy_val, 'Auto-Adjust': d.adjusted ? 'Ya' : 'Tidak', 'Status LR': statusLr(d) });
      sum.push({ SID: sid, Nama: d.name, 'Loan Existing': d.loan_existing, 'Coll Awal': d.coll_before_rp,
        'RP Max (rekomendasi)': d.total_rp_maks, 'RP Final': d.rp_final, 'Loan After RP': d.loan_after_rp, 'Rasio RP': pct(d.rasio_rp),
        'Avail Efektif': d.avail_efektif, 'Max LR Final': toInt(d.max_lr_final), 'LR Final': d.lr_final, Binding: d.lr_binding, 'Rasio LR': pct(d.rasio_lr),
        'Nilai Beli': d.total_buy_val, 'Auto-Adjust': d.adjusted ? 'Ya' : 'Tidak', 'Status RP': d.has_rp ? statusRp(d) : '-', 'Status LR': d.has_lr ? statusLr(d) : '-' });
    }
    return { 'Repayment (RP)': rp, 'Loan Request (LR)': lr, Summary: sum };
  }
  function toCsv(rows, columns) {
    const cols = columns || (rows.length ? Object.keys(rows[0]) : []);
    const esc = (v) => { const s = v == null ? '' : String(v); return /[",\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; };
    return [cols.map(esc).join(',')].concat(rows.map((r) => cols.map((c) => esc(r[c])).join(','))).join('\r\n');
  }

  // ── Diagnostik "Kenapa tidak muncul?" per SID/CID/nama ──
  const normId = (x) => String(x || '').toUpperCase().replace(/[^A-Z0-9]/g, '').replace(/^0+/, '');
  function resolveSid(query, ctx) {
    const q = str(query); if (!q) return { sid: null, how: 'kosong', variants: [] };
    const pool = new Set([...ctx.peiSids, ...ctx.op.keys(), ...ctx.cl.keys(), ...ctx.trace.rawSids, ...ctx.netting.keys()]);
    if (pool.has(q)) return { sid: q, how: 'SID persis', variants: [] };
    for (const [sid, cid] of ctx.sidToCid) if (cid === q) return { sid, how: 'dicari lewat CID', variants: [] };
    const n = normId(q); const variants = [...pool].filter((x) => normId(x) === n);
    if (variants.length) return { sid: variants[0], how: 'SID cocok setelah dinormalisasi (format beda antar file)', variants };
    const names = []; const ql = q.toLowerCase();
    for (const [sid, d] of ctx.op) if (String(d.name).toLowerCase().includes(ql)) names.push(sid);
    for (const [sid, d] of ctx.cl) if (String(d.name).toLowerCase().includes(ql) && !names.includes(sid)) names.push(sid);
    for (const [cid, nm] of ctx.cidToName) if (String(nm).toLowerCase().includes(ql)) { for (const [sid, c] of ctx.sidToCid) if (c === cid && !names.includes(sid)) names.push(sid); }
    if (names.length) return { sid: names[0], how: `dicari lewat nama (${names.length} kecocokan)`, variants: names };
    return { sid: null, how: 'tidak ditemukan di file mana pun', variants: [] };
  }
  /** Telusuri tahap per tahap. vres = hasil runValidasi (opsional). */
  function diagnoseSid(query, ctx, vres) {
    const r = resolveSid(query, ctx); const st = []; const push = (stage, status, msg) => st.push({ stage, status, msg });
    if (!r.sid) { push('Pencarian', 'drop', `"${query}" ${r.how}. Cek penulisan SID/CID/nama.`); return { sid: null, resolved: r, stages: st, verdict: 'Tidak ditemukan' }; }
    const sid = r.sid; push('Pencarian', r.variants.length > 1 ? 'warn' : 'ok', `${sid} — ${r.how}${r.variants.length > 1 ? ': ' + r.variants.join(', ') : ''}`);
    const inFile = { 'SID Client': ctx.peiSids.has(sid), 'OP': ctx.op.has(sid), 'Credit Limit': ctx.cl.has(sid), 'Netting Invoice (mentah)': ctx.trace.rawSids.has(sid) };
    const other = [...new Set([...ctx.peiSids, ...ctx.op.keys(), ...ctx.cl.keys(), ...ctx.trace.rawSids])].filter((x) => x !== sid && normId(x) === normId(sid));
    if (other.length) push('Format SID', 'warn', `Ada SID mirip dengan format berbeda di file: ${other.join(', ')} → file tidak akan saling cocok.`);

    const dup = (ctx.sidDupes || []).find((x) => x.sid === sid);
    if (dup) push('SID Client ganda', dup.conflict ? 'warn' : 'ok', `SID muncul ${dup.rows.length}× di SID Client: ${dup.rows.map((x) => x.cid + ' ' + x.name).join(' | ')}${dup.conflict ? ` → CID BERBEDA; yang dipakai ${dup.used}. Perbaiki file SID Client.` : ' (baris dobel, isi sama)'}`);
    if (!inFile['SID Client']) push('SID Client', 'drop', 'SID tidak ada di file SID Client → bukan nasabah PEI, tidak diproses di TRX PEI.'); else push('SID Client', 'ok', 'Terdaftar sebagai nasabah PEI.');

    const raw = ctx.trace.rawRows.get(sid) || 0, db = ctx.trace.droppedBoard.get(sid) || 0, dl = ctx.trace.droppedLorf.get(sid) || 0;
    if (!raw) push('Netting Invoice', 'drop', 'Tidak ada baris transaksi untuk SID ini di Netting Invoice.');
    else {
      push('Netting Invoice', (db || dl) ? 'warn' : 'ok', `${raw} baris mentah; dibuang karena board≠RG: ${db}; karena lorf≠D: ${dl}; terpakai: ${raw - db - dl}.`);
      const net = ctx.netting.get(sid) || dict();
      for (const stock of Object.keys(net)) {
        const d = net[stock]; const nl = d.buy_lot - d.sell_lot;
        const out = nl > 0 ? `NET BELI ${nl} → masuk LR` : (nl < 0 ? `NET JUAL ${-nl} → masuk RP` : 'NETTING NOL (beli = jual) → DIABAIKAN, tidak ada RP/LR');
        push('Netting ' + stock, nl === 0 ? 'drop' : 'ok', `beli ${d.buy_lot} / jual ${d.sell_lot} → ${out}`);
      }
      if (!Object.keys(net).length) push('Netting Invoice', 'drop', 'Semua baris SID ini terbuang oleh filter board/lorf.');
    }
    const stocksUsed = new Set([...Object.keys(ctx.netBuy.get(sid) || {}), ...Object.keys(ctx.netSell.get(sid) || {})]);
    for (const s of stocksUsed) {
      if (!ctx.cp.has(s)) push('Closing Price ' + s, 'drop', 'Harga penutupan tidak ada di file Closing Price → nilai beli/jual = 0 (LR jadi 0 / RP referensi 0).');
      if (!ctx.hc.has(s)) push('Risk Parameter ' + s, 'warn', 'Saham tidak ada di Risk Parameter → haircut default 5% dipakai.');
    }
    const op = ctx.op.get(sid);
    if (!op) push('OP', 'warn', 'SID tidak ada di Outstanding Position → Loan Existing 0, tidak ada RP.');
    else push('OP', op.loan_existing > 0 ? 'ok' : 'warn', `Loan Existing ${fmtRp(op.loan_existing)}, accrued ${fmtRp(op.accrued_interest)}${op.loan_existing > 0 ? '' : ' → RP TIDAK muncul (hanya muncul jika Loan Existing > 0)'}.`);
    const cl = ctx.cl.get(sid);
    push('Credit Limit', cl ? (cl.available_limit > 0 ? 'ok' : 'warn') : 'warn', cl ? `Available Limit ${fmtRp(cl.available_limit)}${cl.available_limit > 0 ? '' : ' → kapasitas LR 0 kecuali ada RP'}.` : 'SID tidak ada di Credit Limit → Available Limit 0.');

    const d = vres && vres.get(sid);
    let verdict = 'OK';
    if (d) {
      if (d.has_rp) {
        if (d.rp_skipped) push('RP', 'drop', 'Ada jual bersih tapi Loan Existing = 0 → tidak ada RP.');
        else { const ok = lolosRp(d); push('RP', ok ? 'ok' : 'drop', `RP Final ${fmtRp(d.rp_final)}${d.rp_edited ? ' (diedit)' : ''}. ${ok ? 'Masuk file iFast RP.' : 'GAGAL check: ' + d.checks.filter((c) => c.label.startsWith('RP-') && !c.passed).map((c) => c.label + ' — ' + c.detail).join('; ')}`); if (!ok) verdict = 'RP gugur di validasi'; }
      } else push('RP', 'warn', 'Tidak ada jual bersih → tidak ada RP.');
      if (d.has_lr) {
        const ok = lolosLr(d); const failing = d.checks.filter((c) => c.label.startsWith('LR-') && !c.passed);
        if (d.lr_final <= 0) { push('LR', 'drop', `LR Final = 0 (binding: ${d.lr_binding}; LR@63% ${fmtRp(d.max_lr_63)}, Avail Efektif ${fmtRp(d.avail_efektif)}, Nilai Beli ×110% ${fmtRp(d.total_buy_val * 1.1)}) → tidak masuk file iFast.`); verdict = 'LR = 0'; }
        else if (!ok) { push('LR', 'drop', `LR Final ${fmtRp(d.lr_final)} tapi GAGAL check → tidak masuk file iFast: ${failing.map((c) => c.label + ' — ' + c.detail).join('; ')}${failing.some((c) => c.label === LR3) ? ' (gunakan Auto-Adjust LR untuk memotong ke Max LR Final)' : ''}`); verdict = 'LR gugur di validasi'; }
        else push('LR', 'ok', `LR Final ${fmtRp(d.lr_final)}${d.lr_edited ? ' (diedit)' : ''}, binding ${d.lr_binding}. Masuk file iFast LR.`);
      } else push('LR', 'warn', 'Tidak ada beli bersih → tidak ada LR.');
      if (!d.has_rp && !d.has_lr) verdict = 'Tidak ada transaksi bersih (RP/LR)';
    } else push('Hasil', 'warn', 'Validasi belum dijalankan.');
    if (!inFile['SID Client'] && verdict === 'OK') verdict = 'Bukan nasabah PEI di SID Client';
    return { sid, resolved: r, stages: st, verdict };
  }

  // ── deteksi berkas wajib dari nama (untuk upload folder) ──
  const SLOT_PATTERNS = [
    ['closing', /closing\s*price/i, ['.xlsx']],
    ['risk', /risk\s*param/i, ['.txt']],
    ['op', /^op[\s_\-]/i, ['.txt']],
    ['netting', /(list\s*(of\s*)?invoice|netting)/i, ['.csv', '.xls', '.xlsx']],
    ['cl', /credit\s*limit/i, ['.txt']],
    ['sidclient', /sid\s*client/i, ['.xlsx']],
    ['pos', /stock\s*position/i, ['.txt']], // opsional, tapi ikut dideteksi otomatis (v5)
  ];
  function detectSlot(filename) {
    const name = String(filename || '').split(/[\\/]/).pop();
    const ext = (name.match(/\.[^.]+$/) || [''])[0].toLowerCase();
    for (const [id, re, exts] of SLOT_PATTERNS) if (re.test(name) && exts.includes(ext)) return id;
    return null;
  }
  // files: [{name, lastModified}] → { picked: {slotId: file}, dupes: {slotId: jumlah}, ignored: [nama] }
  function matchFolderFiles(files) {
    const by = {}; const ignored = [];
    for (const f of files) { const id = detectSlot(f.name); if (!id) { ignored.push(f.name); continue; } (by[id] = by[id] || []).push(f); }
    const picked = {}, dupes = {};
    for (const [id, arr] of Object.entries(by)) { arr.sort((a, b) => (b.lastModified || 0) - (a.lastModified || 0)); picked[id] = arr[0]; if (arr.length > 1) dupes[id] = arr.length; }
    return { picked, dupes, ignored };
  }

  return {
    detectSlot, matchFolderFiles, RATIO_THRESHOLD, AUTO_ADJUST_TARGET, LR3, LR4, brokerCodesOf, parseRupiah, toInt,
    cloneState, readFinalColumn, applyFinals, parseHasilTrx, applyHasilTrx, cancelAutoAdjust, stateSig, statusRp, statusLr,
    pyFloat, findAndRename, calcCollateral, solveRpMin, solveRpMax, solveLr, getLoanStatus, fmtRp, fmtPct,
    loadClosingPrice, loadRiskParameter, parseOpFile, parseCreditLimitFile, loadParticipantStockPosition, loadLrBelumSettle,
    applyLrAdjustmentToOp, applyPositionAdjustmentToOp, parseNettingInvoice, splitNetting, parseSidClient, buildContext,
    computeSid, lolosRp, lolosLr, newState, setOverride, optsFor, runPei, runValidasi, globalResult, applyAutoAdjust,
    ifastRepayment, ifastLoan, hasilTrxPei, toCsv, resolveSid, diagnoseSid,
  };
});
