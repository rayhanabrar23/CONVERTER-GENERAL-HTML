// Tes perbaikan v5 (logika inti). Jalankan: node tests/fixes_v5.test.js  (butuh tests/data dari gen_data.py)
const fs = require('fs'), path = require('path'), XLSX = require('xlsx'), assert = require('assert');
const C = (() => { const C0 = require('../public/core.js'); const CODE = JSON.parse(require('fs').readFileSync(require('path').join(__dirname, 'data', 'meta.json'), 'utf8')).code;
  return Object.assign({}, C0, { CODE, ifastRepayment: (r, c) => C0.ifastRepayment(r, c || CODE), ifastLoan: (r, c) => C0.ifastLoan(r, c || CODE) }); })(); const D = path.join(__dirname, 'data');
const xrows = (f) => { const w = XLSX.readFile(D + '/' + f); return XLSX.utils.sheet_to_json(w.Sheets[w.SheetNames[0]]); };
const txt = (f) => fs.readFileSync(D + '/' + f, 'utf8');
const base = { closingRows: xrows('closing.xlsx'), riskText: txt('risk.txt'), opText: txt('op.txt'), nettingRows: xrows('netting.xlsx'), clText: txt('cl.txt'), sidClientRows: xrows('sidclient.xlsx'), posText: txt('pos.txt'), lrSettleText: txt('lrsettle.txt') };
const ctx = C.buildContext(base);
let pass = 0; const t = (name, fn) => { fn(); pass++; console.log('  ✓', name); };
const val0 = C.runValidasi(ctx, C.newState());
const lrSid = [...val0].find(([, d]) => C.lolosLr(d) && d.lr_final > 0)[0];
const rpSid = [...val0].find(([, d]) => C.lolosRp(d) && d.rp_final > 0)[0];
const sheetRows = (o) => { const out = {}; for (const [k, v] of Object.entries(o)) out[k] = JSON.parse(JSON.stringify(v)); return out; };

t('#2 parseRupiah: format Indonesia/Inggris diterima, desimal & teks ditolak', () => {
  for (const s of ['7845750', '7.845.750', '7,845,750', 'Rp 7.845.750', 'Rp. 7.845.750', '7.845.750,00', ' 7845750 ']) assert.strictEqual(C.parseRupiah(s).value, 7845750, s);
  for (const s of ['7.845.750,50', '7845750.5', '-5', 'abc', '1.234.5']) assert.ok(C.parseRupiah(s).error, s);
  assert.strictEqual(C.parseRupiah(7845750.000000001).value, 7845750); assert.ok(C.parseRupiah(7845750.5).error); assert.strictEqual(C.parseRupiah('').value, null);
});
t('#1 LR Final > Max LR Final → LR-4 gagal, tidak masuk iFast', () => {
  const st = C.newState(); const cap = val0.get(lrSid).lr_cap; C.setOverride(st, 'lr', lrSid, cap * 100, 't');
  const r = C.runValidasi(ctx, st); const d = r.get(lrSid);
  assert.ok(d.lr_over); assert.ok(!C.lolosLr(d)); assert.ok(d.checks.some((c) => c.label === C.LR4 && !c.passed));
  assert.ok(!C.ifastLoan(r).Loan.some((x) => x['SID Client'] === lrSid));
  C.setOverride(st, 'lr', lrSid, cap, 't'); assert.ok(C.lolosLr(C.runValidasi(ctx, st).get(lrSid)), 'tepat di batas tetap lolos');
});
t('#3 RP Final > RP Max → RP-4 gagal, tidak masuk iFast, kapasitas LR tidak menggelembung', () => {
  const st = C.newState(); const b = val0.get(rpSid); C.setOverride(st, 'rp', rpSid, b.rp_cap + 5e9, 't');
  const r = C.runValidasi(ctx, st); const d = r.get(rpSid);
  assert.ok(d.rp_over && !C.lolosRp(d)); assert.ok(!C.ifastRepayment(r)['Repayment Direct'].some((x) => x['SID Client'] === rpSid));
  assert.strictEqual(d.avail_efektif, b.avail_efektif); assert.strictEqual(d.max_lr_final, b.max_lr_final);
  // SID tanpa RP (RP Max 0) diberi RP Final → tetap tidak menambah Avail Efektif
  const noRp = [...val0].find(([, x]) => x.total_rp_maks === 0 && x.has_lr)[1]; const s2 = C.newState(); C.setOverride(s2, 'rp', noRp.sid, 1e9, 't');
  assert.strictEqual(C.runValidasi(ctx, s2).get(noRp.sid).avail_efektif, noRp.avail_efektif);
});
t('#5 Loan Value & LR Final default selalu Rupiah bulat', () => {
  for (const r of C.ifastLoan(val0).Loan) assert.ok(Number.isInteger(r['Loan Value']), r['Loan Value']);
  for (const d of val0.values()) assert.ok(Number.isInteger(d.lr_final));
  const h = C.hasilTrxPei(C.runPei(ctx, C.newState())); for (const r of h['Loan Request (LR)']) for (const k of ['LR Final', 'Max LR Final', 'LR @63%']) assert.ok(Number.isInteger(r[k]), k);
});
t('#10 RP lunas (pokok + accrued) → accrued tidak dihitung lagi; limit kembali sebesar pokok', () => {
  const d = [...val0.values()].find((x) => C.lolosRp(x) && x.accrued > 0 && x.rp_final >= x.loan_existing + x.accrued);
  assert.ok(d, 'perlu SID dengan RP penuh & accrued > 0');
  assert.strictEqual(d.loan_after_rp, 0); assert.strictEqual(d.accrued_after_rp, 0);
  if (d.coll_after_rp > 0) assert.strictEqual(d.rasio_rp, 0);
  assert.strictEqual(d.avail_efektif, d.avail_limit + d.loan_existing);
  // RP sebagian: sisa pokok berkurang dulu, accrued utuh
  const st = C.newState(); C.setOverride(st, 'rp', d.sid, Math.round(d.loan_existing / 2), 't'); const p = C.runValidasi(ctx, st).get(d.sid);
  assert.strictEqual(p.accrued_after_rp, d.accrued); assert.strictEqual(p.loan_after_rp, d.loan_existing - Math.round(d.loan_existing / 2));
});
t('#7 SID ganda di SID Client terdeteksi (CID beda = conflict)', () => {
  const rows = base.sidClientRows.slice(); rows.push(Object.assign({}, rows[0], { no_cust: 'C99999', Name: 'ORANG LAIN' })); rows.push(Object.assign({}, rows[1]));
  const c2 = C.buildContext(Object.assign({}, base, { sidClientRows: rows }));
  const a = c2.sidDupes.find((x) => x.sid === rows[0].SID), b = c2.sidDupes.find((x) => x.sid === rows[1].SID);
  assert.ok(a.conflict && a.rows.length === 2); assert.ok(b && !b.conflict);
  assert.ok(C.diagnoseSid(rows[0].SID, c2, C.runValidasi(c2, C.newState())).stages.some((s) => s.stage === 'SID Client ganda' && s.status === 'warn'));
  assert.strictEqual(ctx.sidDupes.length, 0);
});
t('#6 StockPosition dikenali dari nama berkas; tanpa berkas itu ctx.hasPos = false', () => {
  assert.strictEqual(C.detectSlot('StockPosition_20261013.txt'), 'pos'); assert.strictEqual(C.detectSlot('Stock Position 13 Okt.txt'), 'pos'); assert.strictEqual(C.detectSlot('BUYMARGIN_ZZ_1.txt'), null);
  assert.strictEqual(C.buildContext(Object.assign({}, base, { posText: undefined })).hasPos, false); assert.strictEqual(ctx.hasPos, true);
});
// ── #4 Hasil_TRX_PEI benar-benar dipakai di Langkah 4 ──
const st0 = C.newState(); const hasil0 = sheetRows(C.hasilTrxPei(C.runPei(ctx, st0)));
t('#4 Hasil_TRX_PEI tanpa perubahan → diterima, tidak ada nilai berubah', () => {
  const p = C.parseHasilTrx(hasil0); assert.ok(p.ok, p.errors.join()); const r = C.applyHasilTrx(ctx, st0, p, 't'); assert.ok(r.ok); assert.strictEqual(r.changed.length, 0);
});
t('#4 LR Final diedit di Excel (sheet LR + Summary) → dipakai di file iFast', () => {
  const h = sheetRows(hasil0); const row = h['Loan Request (LR)'].find((x) => x.SID === lrSid); const nv = Math.floor(row['LR Final'] / 2);
  row['LR Final'] = nv; h.Summary.find((x) => x.SID === lrSid)['LR Final'] = nv;
  const r = C.applyHasilTrx(ctx, st0, C.parseHasilTrx(h), 't'); assert.ok(r.ok, r.errors.join()); assert.strictEqual(r.changed.length, 1);
  assert.strictEqual(r.state.lr.get(lrSid), nv); assert.strictEqual(st0.lr.has(lrSid), false, 'state lama tidak tersentuh');
});
t('#4 sheet LR ≠ Summary → ditolak; nilai desimal → ditolak; LR > Max → ditolak', () => {
  const h = sheetRows(hasil0); h['Loan Request (LR)'].find((x) => x.SID === lrSid)['LR Final'] = 1;
  assert.ok(!C.parseHasilTrx(h).ok);
  const h2 = sheetRows(hasil0); h2.Summary.find((x) => x.SID === lrSid)['LR Final'] = '7.845.750,50'; assert.ok(!C.parseHasilTrx(h2).ok);
  const h3 = sheetRows(hasil0); const big = h3.Summary.find((x) => x.SID === lrSid)['LR Final'] * 10; h3.Summary.find((x) => x.SID === lrSid)['LR Final'] = big; h3['Loan Request (LR)'].find((x) => x.SID === lrSid)['LR Final'] = big;
  const r = C.applyHasilTrx(ctx, st0, C.parseHasilTrx(h3), 't'); assert.ok(!r.ok && /melebihi Max LR Final/.test(r.errors.join())); assert.strictEqual(r.state, st0);
});
t('#4 file sampah / file dari hari lain → ditolak', () => {
  assert.ok(!C.parseHasilTrx({ 'Loan Request (LR)': [{ x: 1 }] }).ok);
  const h = sheetRows(hasil0); h.Summary = h.Summary.slice(1);
  assert.ok(!C.applyHasilTrx(ctx, st0, C.parseHasilTrx(h), 't').ok);
});
t('applyFinals: RP dari file melebihi RP Max → semua ditolak, state tidak berubah', () => {
  const m = new Map([[rpSid, val0.get(rpSid).rp_cap + 1]]); const r = C.applyFinals(ctx, st0, { rp: m }, 't', 'x');
  assert.ok(!r.ok); assert.strictEqual(r.state, st0); assert.strictEqual(st0.rp.size, 0);
});

// ── v5.1: Auto-Adjust tercatat di Hasil_TRX_PEI & di Catatan ──
// konteks dengan Available Qty cukup untuk semua saham → yang diuji murni LR-3 / Auto-Adjust
const ctxA = Object.assign({}, ctx, { avq: new Map([...ctx.avq.keys(), ...[...ctx.netBuy.values()].flatMap((m) => Object.keys(m))].map((k) => [k, 9e12])) });
const valA = C.runValidasi(ctxA, C.newState());
const adjCand = [...valA].filter(([, d]) => d.is_pei && d.has_lr && d.lr_final > 0 && d.checks.filter((c) => c.label.startsWith('LR-') && !c.passed).every((c) => c.label === C.LR3) && d.checks.some((c) => c.label === C.LR3 && !c.passed)).map(([s]) => s);
t('v5.1 Hasil_TRX_PEI: kolom Auto-Adjust & Status LR berubah setelah Auto-Adjust', () => {
  assert.ok(adjCand.length, 'perlu SID yang hanya gagal LR-3'); const sid = adjCand[0];
  const st = C.newState(); const before = C.hasilTrxPei(C.runPei(ctxA, st))['Loan Request (LR)'].find((r) => r.SID === sid);
  assert.strictEqual(before['Auto-Adjust'], 'Tidak'); assert.match(before['Status LR'], /TIDAK masuk iFast — gagal LR-3 \(perlu Auto-Adjust\)/);
  C.applyAutoAdjust(st, C.runValidasi(ctxA, st), 't'); const h = C.hasilTrxPei(C.runPei(ctxA, st));
  const after = h['Loan Request (LR)'].find((r) => r.SID === sid); assert.strictEqual(after['Auto-Adjust'], 'Ya'); assert.match(after['Status LR'], /^Masuk iFast — di-Auto-Adjust/);
  assert.strictEqual(h.Summary.find((r) => r.SID === sid)['Auto-Adjust'], 'Ya');
  // status di Hasil = isi file iFast
  const inIfast = new Set(C.ifastLoan(C.runValidasi(ctxA, st)).Loan.map((r) => r['SID Client']));
  for (const r of h['Loan Request (LR)']) assert.strictEqual(/^Masuk iFast/.test(r['Status LR']), inIfast.has(r.SID), r.SID + ' ' + r['Status LR']);
});
t('v5.1 unggah Hasil (Auto-Adjust = Ya) di sesi baru → status Auto-Adjust ikut, SID masuk iFast', () => {
  const st = C.newState(); C.applyAutoAdjust(st, C.runValidasi(ctxA, st), 't'); const h = sheetRows(C.hasilTrxPei(C.runPei(ctxA, st)));
  const r = C.applyHasilTrx(ctxA, C.newState(), C.parseHasilTrx(h), 'u'); assert.ok(r.ok, r.errors.join());
  for (const sid of adjCand) assert.ok(r.state.adjusted.has(sid), sid);
  const loan = new Set(C.ifastLoan(C.runValidasi(ctxA, r.state)).Loan.map((x) => x['SID Client'])); for (const sid of adjCand) assert.ok(loan.has(sid));
  assert.ok(r.state.audit.some((a) => a.kind === 'AUTO-ADJUST' && a.note === 'dari Hasil TRX'));
});
t('v5.1 Auto-Adjust diubah jadi Tidak di Excel → dibatalkan & tercatat; LR ≠ Summary → ditolak', () => {
  const st = C.newState(); C.applyAutoAdjust(st, C.runValidasi(ctxA, st), 't'); const h = sheetRows(C.hasilTrxPei(C.runPei(ctxA, st))); const sid = adjCand[0];
  h.Summary.find((x) => x.SID === sid)['Auto-Adjust'] = 'Tidak'; assert.ok(!C.parseHasilTrx(h).ok, 'LR sheet masih Ya → ditolak');
  h['Loan Request (LR)'].find((x) => x.SID === sid)['Auto-Adjust'] = 'Tidak';
  const r = C.applyHasilTrx(ctxA, st, C.parseHasilTrx(h), 'u'); assert.ok(r.ok, r.errors.join()); assert.ok(!r.state.adjusted.has(sid));
  assert.ok(r.state.audit.some((a) => a.kind === 'BATAL AUTO-ADJUST' && a.sid === sid));
  assert.ok(!C.ifastLoan(C.runValidasi(ctxA, r.state)).Loan.some((x) => x['SID Client'] === sid));
});
t('v5.1 Hasil lama tanpa kolom Auto-Adjust → status Auto-Adjust di sesi tidak disentuh', () => {
  const st = C.newState(); C.applyAutoAdjust(st, C.runValidasi(ctxA, st), 't'); const h = sheetRows(C.hasilTrxPei(C.runPei(ctxA, st)));
  for (const k of ['Summary', 'Loan Request (LR)']) for (const r of h[k]) delete r['Auto-Adjust'];
  const r = C.applyHasilTrx(ctxA, st, C.parseHasilTrx(h), 'u'); assert.ok(r.ok); assert.strictEqual(r.state.adjusted.size, st.adjusted.size);
});
t('v5.1 Batalkan Auto-Adjust tercatat di Catatan', () => {
  const st = C.newState(); C.applyAutoAdjust(st, C.runValidasi(ctxA, st), 't'); const n0 = st.adjusted.size; const sig = C.stateSig(st);
  const gone = C.cancelAutoAdjust(st, 'tester'); assert.strictEqual(gone.length, n0); assert.strictEqual(st.adjusted.size, 0);
  assert.strictEqual(st.audit.filter((a) => a.kind === 'BATAL AUTO-ADJUST' && a.user === 'tester').length, n0); assert.notStrictEqual(C.stateSig(st), sig);
});
console.log(`\n${pass} tes perbaikan v5/v5.1 lulus`);
