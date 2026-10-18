// Tes UI perbaikan v5 (jsdom). Jalankan: node tests/ui_v5.test.js
const fs = require('fs'), path = require('path'), assert = require('assert'); const { JSDOM } = require('jsdom');
const P = path.join(__dirname, '..', 'public'), D = path.join(__dirname, 'data');
const wait = (ms = 60) => new Promise((r) => setTimeout(r, ms));
async function boot() {
  const dom = new JSDOM(fs.readFileSync(P + '/index.html', 'utf8').replace(/<script[^>]*><\/script>/g, ''), { url: 'https://pei.test/', runScripts: 'outside-only', pretendToBeVisual: true });
  const w = dom.window; const calls = { xlsx: [] };
  w.URL.createObjectURL = () => 'blob:x'; w.HTMLAnchorElement.prototype.click = function () {};
  for (const f of ['vendor/xlsx.full.min.js', 'config.js', 'core.js']) w.eval(fs.readFileSync(path.join(P, f), 'utf8'));
  w.XLSX.writeFile = (wb, name) => calls.xlsx.push({ wb, name });
  w.eval(fs.readFileSync(P + '/app.js', 'utf8')); await wait();
  return { w, d: w.document, calls };
}
const NAMES = { 'closing price 121026.xlsx': 'closing.xlsx', 'RiskParameter_20261013.txt': 'risk.txt', 'OP_20261013_ZZ.txt': 'op.txt', 'list invoice 121026.csv': 'netting.csv',
  'CreditLimit_20261013_ZZ.txt': 'cl.txt', 'SID Client.xlsx': 'sidclient.xlsx', 'StockPosition_20261013.txt': 'pos.txt' };
async function folder(f, { noPos, over = {} } = {}) {
  const inp = f.d.querySelector('input[aria-label="Pilih beberapa berkas"]');
  const files = Object.entries(NAMES).filter(([n]) => !(noPos && /StockPosition/.test(n))).map(([n, src]) => new f.w.File([over[src] || fs.readFileSync(D + '/' + src)], n));
  Object.defineProperty(inp, 'files', { value: files, configurable: true }); inp.dispatchEvent(new f.w.Event('change')); await wait(400);
  const cl = f.d.querySelector('input[aria-label="Credit Limit Partisipan"]'); if (cl && !cl.value) { cl.value = '160.000.000.000'; cl.dispatchEvent(new f.w.Event('change')); await wait(); }
}
const main = (d) => d.getElementById('main').textContent;
const btn = (d, txt) => [...d.querySelectorAll('button')].find((b) => b.textContent.includes(txt));
const setFile = async (w, d, label, file) => { const inp = d.querySelector(`input[type=file][aria-label="${label}"]`); assert.ok(inp, label); Object.defineProperty(inp, 'files', { value: [file], configurable: true }); inp.dispatchEvent(new w.Event('change')); await wait(150); };
const toFile = (w, wb, name) => new w.File([w.XLSX.write(wb, { type: 'array', bookType: 'xlsx' })], name);
const typeIn = async (w, inp, v) => { inp.value = v; inp.dispatchEvent(new w.Event('change')); await wait(); };
// SID yang lolos semua check LR (mis. LR-1 Available Qty) di data sintetis — dihitung langsung dari core
const OK_LR = (() => { const C = require('../public/core.js'), XL = require('xlsx'); const xr = (f) => { const wb = XL.readFile(D + '/' + f); return XL.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]); };
  const tx = (f) => fs.readFileSync(D + '/' + f, 'utf8');
  const ctx = C.buildContext({ closingRows: xr('closing.xlsx'), riskText: tx('risk.txt'), opText: tx('op.txt'), nettingRows: xr('netting.xlsx'), clText: tx('cl.txt'), sidClientRows: xr('sidclient.xlsx'), posText: tx('pos.txt') });
  return new Set([...C.runValidasi(ctx, C.newState())].filter(([, x]) => x.is_pei && C.lolosLr(x) && x.lr_final > 1000).map(([k]) => k)); })();
let n = 0; const t = async (name, fn) => { await fn(); n++; console.log('  ✓', name); };

(async () => {
  const f = await boot(); const { w, d, calls } = f;
  await t('#6 folder: StockPosition ikut terdeteksi otomatis', async () => {
    await folder(f); assert.match(main(d), /Participant Stock Position ikut terdeteksi: StockPosition_20261013\.txt/);
    btn(d, 'Proses').click(); await wait(200); assert.doesNotMatch(main(d), /Participant Stock Position tidak diunggah/);
  });
  const rpIn = () => d.querySelector('input.cell[aria-label^="RP Final"]');
  await t('#2/#3 RP Final: desimal ditolak, melebihi RP Max ditolak, format 1.234.567 diterima', async () => {
    const inp = rpIn(); const sid = inp.getAttribute('aria-label').split(' ').pop(); const v0 = Number(inp.value);
    await typeIn(w, inp, '1.000,50'); assert.match(main(d), /tidak disimpan: "1\.000,50" ada desimal/); assert.strictEqual(Number(rpIn().value), v0);
    await typeIn(w, rpIn(), String(v0 + 1)); assert.match(main(d), /melebihi RP Max/); assert.strictEqual(Number(rpIn().value), v0);
    const half = Math.round(v0 / 2); await typeIn(w, rpIn(), half.toLocaleString('id-ID')); assert.strictEqual(Number(rpIn().value), half); assert.doesNotMatch(main(d), /tidak disimpan/);
    globalThis.__rpSid = sid; globalThis.__rpHalf = half;
  });
  let rpWb;
  await t('#3 upload RP Final: teks format Indonesia diterima; nilai > RP Max → file ditolak', async () => {
    btn(d, 'Simpan RP Final').click(); await wait(); rpWb = calls.xlsx.pop().wb; btn(d, 'Lanjut ke Langkah 3').click(); await wait();
    const rows = w.XLSX.utils.sheet_to_json(rpWb.Sheets['RP Final']);
    const bad = rows.map((r) => Object.assign({}, r, r.SID === globalThis.__rpSid ? { 'RP Final': r['RP Max (rekomendasi)'] + 1000 } : {}));
    let wb = w.XLSX.utils.book_new(); w.XLSX.utils.book_append_sheet(wb, w.XLSX.utils.json_to_sheet(bad), 'RP Final');
    await setFile(w, d, 'Upload RP Final', toFile(w, wb, 'rp_bad.xlsx')); assert.match(main(d), /melebihi RP Max .*file ditolak/); assert.match(main(d), /Upload file RP Final untuk menghitung LR/);
    const txt = rows.map((r) => Object.assign({}, r, { 'RP Final': Math.round(r['RP Final']).toLocaleString('id-ID') }));
    wb = w.XLSX.utils.book_new(); w.XLSX.utils.book_append_sheet(wb, w.XLSX.utils.json_to_sheet(txt), 'RP Final');
    await setFile(w, d, 'Upload RP Final', toFile(w, wb, 'rp_txt.xlsx')); assert.match(main(d), /RP Final dimuat/);
    const junk = rows.map((r, i) => Object.assign({}, r, i === 0 ? { 'RP Final': 'abc' } : {}));
    wb = w.XLSX.utils.book_new(); w.XLSX.utils.book_append_sheet(wb, w.XLSX.utils.json_to_sheet(junk), 'RP Final');
    await setFile(w, d, 'Upload RP Final', toFile(w, wb, 'rp_junk.xlsx')); assert.match(main(d), /tidak bisa dibaca/);
    wb = w.XLSX.utils.book_new(); w.XLSX.utils.book_append_sheet(wb, w.XLSX.utils.json_to_sheet(rows), 'RP Final'); await setFile(w, d, 'Upload RP Final', toFile(w, wb, 'rp_ok.xlsx'));
  });
  let lrSid, lrCap;
  await t('#1/#2 LR Final: typo desimal & di atas Max LR Final ditolak', async () => {
    const inp = [...d.querySelectorAll('input.cell[aria-label^="LR Final"]')].find((i) => Number(i.value) > 1000 && OK_LR.has(i.getAttribute('aria-label').split(' ').pop())); lrSid = inp.getAttribute('aria-label').split(' ').pop(); lrCap = Number(inp.value);
    const get = () => [...d.querySelectorAll('input.cell')].find((i) => i.getAttribute('aria-label') === `LR Final ${lrSid}`);
    await typeIn(w, inp, lrCap.toLocaleString('id-ID') + ',50'); assert.match(main(d), /ada desimal/); assert.strictEqual(Number(get().value), lrCap);
    await typeIn(w, get(), String(lrCap * 100)); assert.match(main(d), /melebihi Max LR Final/); assert.strictEqual(Number(get().value), lrCap);
  });
  await t('v5.1 Langkah 3: Auto-Adjust setelah Final & Unduh → Hasil usang, wajib unduh ulang; Batalkan tercatat', async () => {
    btn(d, 'Final & Unduh').click(); await wait(); const h0 = calls.xlsx.pop().wb; assert.ok(!btn(d, 'Lanjut ke Langkah 4').disabled);
    const ad = btn(d, 'Auto-Adjust LR'); assert.ok(!ad.disabled, 'ada nasabah gagal LR-3'); ad.click(); await wait();
    assert.match(main(d), /Ada perubahan .* setelah Hasil_TRX_PEI terakhir dibuat/); assert.ok(btn(d, 'Lanjut ke Langkah 4').disabled);
    btn(d, 'Final & Unduh ulang').click(); await wait(); const h1 = calls.xlsx.pop().wb; assert.ok(!btn(d, 'Lanjut ke Langkah 4').disabled);
    const ya = (wb) => w.XLSX.utils.sheet_to_json(wb.Sheets.Summary).filter((r) => r['Auto-Adjust'] === 'Ya').length;
    assert.strictEqual(ya(h0), 0); assert.ok(ya(h1) > 0, 'Hasil baru mencatat Auto-Adjust = Ya');
    btn(d, 'Batalkan Auto-Adjust').click(); await wait(); assert.ok(btn(d, 'Lanjut ke Langkah 4').disabled, 'batal = beda dengan Hasil terakhir → unduh ulang');
    btn(d, 'Final & Unduh ulang').click(); await wait(); assert.strictEqual(ya(calls.xlsx.pop().wb), 0); assert.ok(!btn(d, 'Lanjut ke Langkah 4').disabled);
    [...d.querySelectorAll('#tabs button')].find((x) => x.textContent.includes('Catatan')).click(); await wait();
    assert.match(main(d), /BATAL AUTO-ADJUST/); assert.match(main(d), /Unduh catatan/);
    [...d.querySelectorAll('#tabs button')].find((x) => x.textContent.includes('3. LR')).click(); await wait();
  });
  let hasilWb;
  await t('#4 Langkah 4: LR Final yang diedit di Hasil_TRX_PEI (sheet LR + Summary) dipakai untuk iFast', async () => {
    btn(d, 'Final & Unduh').click(); await wait(); hasilWb = calls.xlsx.pop().wb; btn(d, 'Lanjut ke Langkah 4').click(); await wait();
    const X = w.XLSX; const nv = Math.floor(lrCap / 3);
    const edit = (sheet) => { const rows = X.utils.sheet_to_json(hasilWb.Sheets[sheet]); rows.find((r) => r.SID === lrSid)['LR Final'] = nv; hasilWb.Sheets[sheet] = X.utils.json_to_sheet(rows); };
    edit('Loan Request (LR)');
    await setFile(w, d, 'Upload Hasil TRX PEI', toFile(w, hasilWb, 'Hasil_half.xlsx')); assert.match(main(d), /ditolak: .*≠ Summary/); assert.ok(btn(d, 'Jalankan Validasi').disabled);
    edit('Summary');
    await setFile(w, d, 'Upload Hasil TRX PEI', toFile(w, hasilWb, 'Hasil_TRX_PEI.xlsx')); assert.match(main(d), /1 perubahan dibanding kondisi sekarang/);
    btn(d, 'Jalankan Validasi').click(); await wait(200); assert.match(main(d), /diambil dari file/);
    const ad = btn(d, 'Auto-Adjust LR'); if (ad && !ad.disabled) { ad.click(); await wait(200); }
    const ex = [...d.querySelectorAll('button')].filter((b) => b.textContent.includes('Excel (.xlsx)'))[1]; ex.click(); await wait();
    const row = X.utils.sheet_to_json(calls.xlsx.pop().wb.Sheets.Loan).find((r) => r['SID Client'] === lrSid);
    assert.ok(row, 'SID ada di Loan iFast'); assert.strictEqual(row['Loan Value'], nv);
  });
  await t('#4 Langkah 4: file bukan Hasil_TRX_PEI ditolak', async () => {
    const wb = w.XLSX.utils.book_new(); w.XLSX.utils.book_append_sheet(wb, w.XLSX.utils.aoa_to_sheet([['x'], [1]]), 'Loan Request (LR)');
    await setFile(w, d, 'Upload Hasil TRX PEI', toFile(w, wb, 'junk.xlsx')); assert.match(main(d), /bukan file Hasil_TRX_PEI/); assert.ok(btn(d, 'Jalankan Validasi').disabled);
  });
  await t('#5 sel RP/LR Final di layar selalu Rupiah bulat (nilai iFast dicek di fixes_v5)', async () => {
    await (async () => { [...d.querySelectorAll('#tabs button')].find((x) => x.textContent.includes('3. LR')).click(); await wait(); })();
    const cells = [...d.querySelectorAll('input.cell')]; assert.ok(cells.length); assert.ok(cells.every((i) => /^\d+$/.test(i.value)));
  });
  await t('#6/#7 tanpa StockPosition & dengan SID ganda → peringatan tampil di langkah RP', async () => {
    const g = await boot(); const X = require('xlsx'); const w0 = X.readFile(D + '/sidclient.xlsx'); const sc = X.utils.sheet_to_json(w0.Sheets[w0.SheetNames[0]]);
    sc.push(Object.assign({}, sc[0], { no_cust: 'C99999', Name: 'ORANG LAIN' }));
    const wb = X.utils.book_new(); X.utils.book_append_sheet(wb, X.utils.json_to_sheet(sc), 'Sheet1');
    await folder(g, { noPos: true, over: { 'sidclient.xlsx': X.write(wb, { type: 'buffer', bookType: 'xlsx' }) } });
    assert.match(main(g.d), /Participant Stock Position \(StockPosition_\*\.txt\) tidak ada di folder/);
    btn(g.d, 'Proses').click(); await wait(200);
    assert.match(main(g.d), /Participant Stock Position tidak diunggah/); assert.match(main(g.d), new RegExp(`SID ganda di SID Client dengan CID berbeda: ${sc[0].SID}`));
  });
  console.log(`\n${n} tes UI v5 lulus`); process.exit(0);
})().catch((e) => { console.error('GAGAL:', e.message); process.exit(1); });
