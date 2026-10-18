// Tes perilaku baru: edit RP/LR masuk ke iFast (Bug 1) + diagnostik (Bug 2). Jalankan: node behavior.js
const fs = require('fs'), path = require('path'), XLSX = require('xlsx'), assert = require('assert');
const C = (() => { const C0 = require('../public/core.js'); const CODE = JSON.parse(require('fs').readFileSync(require('path').join(__dirname, 'data', 'meta.json'), 'utf8')).code;
  return Object.assign({}, C0, { CODE, ifastRepayment: (r, c) => C0.ifastRepayment(r, c || CODE), ifastLoan: (r, c) => C0.ifastLoan(r, c || CODE) }); })(); const D = path.join(__dirname, 'data');
const xrows = (f) => { const w = XLSX.readFile(D + '/' + f); return XLSX.utils.sheet_to_json(w.Sheets[w.SheetNames[0]]); };
const txt = (f) => fs.readFileSync(D + '/' + f, 'utf8');
const ctx = C.buildContext({ closingRows: xrows('closing.xlsx'), riskText: txt('risk.txt'), opText: txt('op.txt'), nettingRows: xrows('netting.xlsx'),
  clText: txt('cl.txt'), sidClientRows: xrows('sidclient.xlsx'), posText: txt('pos.txt'), lrSettleText: txt('lrsettle.txt') });
let pass = 0; const t = (name, fn) => { fn(); pass++; console.log('  ✓', name); };
const state = C.newState();

let vres = C.runValidasi(ctx, state);
const rpSid = [...vres].find(([, d]) => C.lolosRp(d) && d.coll_after_rp > 0)[0];
const lrSid = [...vres].find(([, d]) => C.lolosLr(d) && d.lr_final > 0)[0];

t('Bug 1: edit RP Final masuk ke file iFast RP', () => {
  const before = C.ifastRepayment(vres)['Repayment Direct'].find((r) => r['SID Client'] === rpSid)['Repayment Value'];
  const edited = Math.round(before * 0.99);
  C.setOverride(state, 'rp', rpSid, edited, 'tester');
  vres = C.runValidasi(ctx, state);                         // "rerun" halaman: editan tidak boleh hilang
  const after = C.ifastRepayment(vres)['Repayment Direct'].find((r) => r['SID Client'] === rpSid)['Repayment Value'];
  assert.strictEqual(after, edited); assert.notStrictEqual(after, before);
});
t('Bug 1: edit LR Final masuk ke file iFast LR (+ Hasil_TRX_PEI memakai angka edit)', () => {
  const edited = 12345678;
  C.setOverride(state, 'lr', lrSid, edited, 'tester');
  vres = C.runValidasi(ctx, state); const pei = C.runPei(ctx, state);
  const row = C.ifastLoan(vres)['Loan'].find((r) => r['SID Client'] === lrSid);
  if (row) assert.strictEqual(row['Loan Value'], edited);   // hanya jika lolos check (LR-3: nilai beli <= LR final)
  const hasil = C.hasilTrxPei(pei); const s = hasil.Summary.find((r) => r.SID === lrSid);
  assert.strictEqual(s['LR Final'], edited);
});
t('RP dipotong besar → RP-3 gagal → SID keluar dari file iFast & diagnostik menjelaskan', () => {
  const s4 = C.newState(); const b = C.runValidasi(ctx, s4).get(rpSid); C.setOverride(s4, 'rp', rpSid, Math.round(b.rp_final * 0.2), 't');
  const r = C.runValidasi(ctx, s4); const d = r.get(rpSid);
  if (d.rasio_rp !== null && d.rasio_rp >= 0.65) {
    assert.ok(!C.ifastRepayment(r)['Repayment Direct'].some((x) => x['SID Client'] === rpSid));
    assert.ok(C.diagnoseSid(rpSid, ctx, r).stages.some((s) => s.stage === 'RP' && s.status === 'drop' && /RP-3/.test(s.msg)));
  } else console.log('    (rasio tetap < 65% pada data ini, dilewati)');
});
t('RP Final yang diedit ikut mengubah Loan After RP & kapasitas LR (rantai RP→LR)', () => {
  const s2 = C.newState(); const base = C.runValidasi(ctx, s2);
  const sid = [...base].find(([, d]) => d.has_rp && d.has_lr && !d.rp_skipped)?.[0]; assert.ok(sid, 'perlu SID dengan RP+LR');
  const b = base.get(sid); C.setOverride(s2, 'rp', sid, 0, 't'); const a = C.runValidasi(ctx, s2).get(sid);
  assert.strictEqual(a.loan_after_rp, b.loan_existing); assert.strictEqual(a.avail_efektif, b.avail_limit);
});
t('Editan bertahan setelah run ulang & tercatat di audit log', () => {
  assert.strictEqual(state.rp.get(rpSid) > 0, true); assert.ok(state.audit.filter((a) => a.kind === 'RP' && a.sid === rpSid).length >= 1);
  const a = state.audit.find((x) => x.sid === lrSid && x.kind === 'LR'); assert.strictEqual(a.user, 'tester'); assert.strictEqual(a.to, 12345678);
});
t('Auto-Adjust LR: nasabah gagal LR-3 jadi lolos & nilai iFast = Max LR Final', () => {
  const s3 = C.newState(); const r0 = C.runValidasi(ctx, s3);
  const bad = [...r0].filter(([, d]) => d.has_lr && d.is_pei && d.checks.some((c) => c.label === C.LR3 && !c.passed)).map(([s]) => s);
  if (!bad.length) return console.log('    (tidak ada kasus LR-3 gagal di data sintetis, dilewati)');
  C.applyAutoAdjust(s3, r0, 't'); const r1 = C.runValidasi(ctx, s3);
  for (const s of bad) { assert.ok(r1.get(s).checks.find((c) => c.label === C.LR3).passed); }
});
t('Diagnostik: netting nol (beli = jual) dijelaskan', () => {
  const sids = JSON.parse(txt('meta.json')).sids; const dg = C.diagnoseSid(sids[5], ctx, vres);
  assert.ok(dg.stages.some((s) => s.stage === 'Netting BBCA' && /NETTING NOL/.test(s.msg)), JSON.stringify(dg.stages.map((s) => s.stage + ':' + s.msg)));
});
t('Diagnostik: SID dengan format beda (spasi / huruf kecil / 0 di depan) tetap ketemu', () => {
  const sid = JSON.parse(txt('meta.json')).sids[2]; const r = C.resolveSid('  ' + sid.toLowerCase() + ' ', ctx); assert.strictEqual(r.sid, sid);
});
t('Diagnostik: cari lewat CID & nama; tidak ketemu → pesan jelas', () => {
  assert.ok(C.resolveSid('C00002', ctx).sid); assert.ok(C.resolveSid('nasabah 3', ctx).sid); assert.strictEqual(C.resolveSid('ZZZ999', ctx).sid, null);
});
t('Diagnostik: nasabah bukan PEI, nasabah tanpa OP, dan harga saham hilang terdeteksi', () => {
  const sids = JSON.parse(txt('meta.json')).sids;
  assert.ok(C.diagnoseSid(sids[4], ctx, vres).stages.some((s) => s.stage === 'SID Client' && s.status === 'drop'));
  assert.ok(C.diagnoseSid(sids[37], ctx, vres).stages.some((s) => s.stage === 'OP' && s.status === 'warn'));
  const withMdka = sids.find((s) => (ctx.netBuy.get(s) || {}).MDKA || (ctx.netSell.get(s) || {}).MDKA);
  if (withMdka) assert.ok(C.diagnoseSid(withMdka, ctx, vres).stages.some((s) => s.stage === 'Closing Price MDKA' && s.status === 'drop'));
});
t('CSV = XLSX (isi sama) untuk file iFast', () => {
  for (const [name, sheets] of [['RP', C.ifastRepayment(vres)], ['LR', C.ifastLoan(vres)]]) for (const [sh, rows] of Object.entries(sheets)) {
    if (!rows.length) continue;
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), sh);
    const fromX = XLSX.utils.sheet_to_json(XLSX.read(XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' })).Sheets[sh]);
    const csvWb = XLSX.read(C.toCsv(rows), { type: 'string' }); const fromC = XLSX.utils.sheet_to_json(csvWb.Sheets[csvWb.SheetNames[0]]);
    assert.deepStrictEqual(fromC, fromX, name + ' ' + sh);
  }
});
console.log(`\n${pass} tes perilaku lulus`);
