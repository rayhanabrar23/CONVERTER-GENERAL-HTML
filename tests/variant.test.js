// Tes identitas partisipan sesuai public/config.js (versi terkunci / versi umum). Jalankan: node tests/variant.test.js
const fs = require('fs'), path = require('path'), assert = require('assert'); const { JSDOM } = require('jsdom');
const P = path.join(__dirname, '..', 'public'), D = path.join(__dirname, 'data');
const META = JSON.parse(fs.readFileSync(D + '/meta.json', 'utf8'));
const wait = (ms = 60) => new Promise((r) => setTimeout(r, ms));
async function boot() {
  const dom = new JSDOM(fs.readFileSync(P + '/index.html', 'utf8').replace(/<script[^>]*><\/script>/g, ''), { url: 'https://pei.test/', runScripts: 'outside-only', pretendToBeVisual: true });
  const w = dom.window; const calls = { xlsx: [] };
  w.URL.createObjectURL = () => 'blob:x'; w.HTMLAnchorElement.prototype.click = function () {};
  for (const f of ['vendor/xlsx.full.min.js', 'config.js', 'core.js']) w.eval(fs.readFileSync(path.join(P, f), 'utf8'));
  w.XLSX.writeFile = (wb, name) => calls.xlsx.push({ wb, name });
  w.eval(fs.readFileSync(P + '/app.js', 'utf8')); await wait();
  return { w, d: w.document, calls, CFG: w.PEI_CONFIG };
}
const NAMES = { 'closing price 1.xlsx': 'closing.xlsx', 'RiskParameter_1.txt': 'risk.txt', 'OP_1.txt': 'op.txt', 'list invoice 1.csv': 'netting.csv', 'CreditLimit_1.txt': 'cl.txt', 'SID Client.xlsx': 'sidclient.xlsx', 'StockPosition_1.txt': 'pos.txt' };
// ganti kode partisipan di berkas OP (kolom 3 header) dan/atau Credit Limit (kolom 2)
const recode = (txt, from, to, col) => txt.split('\n').map((l) => { const p = l.split('|'); if (p[col] === from) p[col] = to; return p.join('|'); }).join('\n');
async function load(f, over = {}) {
  const inp = f.d.querySelector('input[aria-label="Pilih beberapa berkas"]');
  Object.defineProperty(inp, 'files', { value: Object.entries(NAMES).map(([n, s]) => new f.w.File([over[s] || fs.readFileSync(D + '/' + s)], n)), configurable: true });
  inp.dispatchEvent(new f.w.Event('change')); await wait(400);
}
const main = (d) => d.getElementById('main').textContent;
const btn = (d, t) => [...d.querySelectorAll('button')].find((b) => b.textContent.includes(t));
const setIn = async (f, label, v) => { const i = f.d.querySelector(`input[aria-label="${label}"]`); i.value = v; i.dispatchEvent(new f.w.Event('change')); await wait(); };
let n = 0; const t = async (name, fn) => { await fn(); n++; console.log('  ✓', name); };

(async () => {
  const f0 = await boot(); const CFG = f0.CFG; const LOCKED = !!CFG.participantCode; const CODE = META.code;
  console.log(`  (versi: ${CFG.variant}, kode data uji: ${CODE})`);
  await t('judul halaman & header mengikuti config.js', async () => {
    assert.strictEqual(f0.d.querySelector('.hero h1').textContent, CFG.title); assert.ok(f0.d.title.startsWith(CFG.title));
  });
  await t('Credit Limit Partisipan: wajib ada sebelum Jalankan Proses; format Rupiah ketat', async () => {
    await load(f0); const b = btn(f0.d, 'Jalankan Proses');
    if (CFG.clPartisipan == null) { assert.ok(b.disabled); assert.match(main(f0.d), /Isi Credit Limit Partisipan dulu/); }
    await setIn(f0, 'Credit Limit Partisipan', '1.000,50'); assert.match(main(f0.d), /ada desimal/);
    await setIn(f0, 'Credit Limit Partisipan', '160.000.000.000'); assert.ok(!btn(f0.d, 'Jalankan Proses').disabled);
  });
  await t('Kode Partisipan terbaca dari berkas & dipakai di file iFast', async () => {
    btn(f0.d, 'Jalankan Proses').click(); await wait(200);
    const expect = LOCKED ? CFG.participantCode : CODE; assert.match(main(f0.d), new RegExp(`Kode Partisipan: ${expect}`));
    // jalankan sampai Langkah 4
    btn(f0.d, 'Simpan RP Final').click(); await wait(); const rp = f0.calls.xlsx.pop().wb; btn(f0.d, 'Lanjut ke Langkah 3').click(); await wait();
    const toFile = (wb, nm) => new f0.w.File([f0.w.XLSX.write(wb, { type: 'array', bookType: 'xlsx' })], nm);
    const up = async (label, file) => { const i = f0.d.querySelector(`input[aria-label="${label}"]`); Object.defineProperty(i, 'files', { value: [file], configurable: true }); i.dispatchEvent(new f0.w.Event('change')); await wait(150); };
    await up('Upload RP Final', toFile(rp, 'rp.xlsx')); btn(f0.d, 'Final & Unduh').click(); await wait(); const hasil = f0.calls.xlsx.pop().wb;
    btn(f0.d, 'Lanjut ke Langkah 4').click(); await wait(); await up('Upload Hasil TRX PEI', toFile(hasil, 'Hasil_TRX_PEI.xlsx'));
    btn(f0.d, 'Jalankan Validasi').click(); await wait(200);
    assert.match(main(f0.d), /Limit Partisipan (LOLOS|GAGAL) — CL Partisipan: Rp 160\.000\.000\.000/);
    const ex = [...f0.d.querySelectorAll('button')].filter((b) => b.textContent.includes('Excel (.xlsx)'));
    let rows = 0;
    for (const [i, sheet] of [[0, 'Repayment Direct'], [1, 'Loan']]) {
      if (ex[i].disabled) continue; ex[i].click(); await wait(); const o = f0.calls.xlsx.pop();
      for (const r of f0.w.XLSX.utils.sheet_to_json(o.wb.Sheets[sheet])) { assert.strictEqual(r['Participant Code'], expect); rows++; }
      assert.ok(!o.name.includes(expect), 'nama file iFast tidak memuat kode partisipan');
    }
    assert.ok(rows > 0, 'ada baris iFast');
  });
  if (LOCKED) {
    await t('versi terkunci: berkas partisipan lain ditolak', async () => {
      const f = await boot(); await load(f, { 'cl.txt': recode(fs.readFileSync(D + '/cl.txt', 'utf8'), CODE, 'QQ', 1), 'op.txt': recode(fs.readFileSync(D + '/op.txt', 'utf8'), CODE, 'QQ', 2) });
      btn(f.d, 'Jalankan Proses').click(); await wait(200); assert.match(main(f.d), new RegExp(`milik partisipan QQ, bukan ${CFG.participantCode}`)); assert.ok(f.d.querySelector('#tabs button:nth-child(2)').disabled);
    });
  } else {
    await t('versi umum: Kode Partisipan otomatis ikut data partisipan lain (QQ)', async () => {
      const f = await boot(); await load(f, { 'cl.txt': recode(fs.readFileSync(D + '/cl.txt', 'utf8'), CODE, 'QQ', 1), 'op.txt': recode(fs.readFileSync(D + '/op.txt', 'utf8'), CODE, 'QQ', 2) });
      await setIn(f, 'Credit Limit Partisipan', '5.000.000.000'); btn(f.d, 'Jalankan Proses').click(); await wait(200); assert.match(main(f.d), /Kode Partisipan: QQ \(dari berkas\)/);
    });
    await t('versi umum: berkas dengan dua Broker Code berbeda ditolak', async () => {
      const f = await boot(); await load(f, { 'op.txt': recode(fs.readFileSync(D + '/op.txt', 'utf8'), CODE, 'QQ', 2) });
      await setIn(f, 'Credit Limit Partisipan', '5.000.000.000'); btn(f.d, 'Jalankan Proses').click(); await wait(200); assert.match(main(f.d), /lebih dari satu Broker Code/);
    });
    await t('versi umum: Kode Partisipan isian manual yang beda dengan berkas ditolak; isian valid dipakai', async () => {
      const f = await boot(); await load(f); await setIn(f, 'Credit Limit Partisipan', '5.000.000.000');
      await setIn(f, 'Kode Partisipan', 'qq'); btn(f.d, 'Jalankan Proses').click(); await wait(200); assert.match(main(f.d), /\(QQ\) berbeda dengan Broker Code di berkas/);
      await setIn(f, 'Kode Partisipan', 'Q-1'); assert.match(main(f.d), /tidak valid/);
      await setIn(f, 'Kode Partisipan', ''); btn(f.d, 'Jalankan Proses').click(); await wait(200); assert.match(main(f.d), new RegExp(`Kode Partisipan: ${CODE} \\(dari berkas\\)`));
    });
  }
  console.log(`\n${n} tes versi lulus`); process.exit(0);
})().catch((e) => { console.error('GAGAL:', e.message); process.exit(1); });
