const fs = require('fs'), path = require('path'), assert = require('assert'); const { JSDOM } = require('jsdom');
const P = path.join(__dirname, '..', 'public'), D = path.join(__dirname, 'data');
const wait = (ms = 60) => new Promise((r) => setTimeout(r, ms));
async function boot(role) {
  const dom = new JSDOM(fs.readFileSync(P + '/index.html', 'utf8').replace(/<script[^>]*><\/script>/g, ''), { url: 'https://pei.test/', runScripts: 'outside-only', pretendToBeVisual: true });
  const w = dom.window; const calls = { audit: [], xlsx: [], csv: [] };
  w.fetch = async () => { calls.net = (calls.net || 0) + 1; return { ok: false }; };
  w.URL.createObjectURL = (b) => { calls.csv.push(b); return 'blob:x'; }; w.HTMLAnchorElement.prototype.click = function () {};
  for (const f of ['vendor/xlsx.full.min.js', 'config.js', 'core.js']) w.eval(fs.readFileSync(path.join(P, f), 'utf8'));
  w.XLSX.writeFile = (wb, name) => calls.xlsx.push({ wb, name });
  w.eval(fs.readFileSync(P + '/app.js', 'utf8')); await wait();
  return { w, d: w.document, calls };
}
async function upload(w, d, over = {}) {
  const map = { '1. Closing Price': 'closing.xlsx', '2. Risk Parameter': 'risk.txt', '3. Outstanding Position': 'op.txt', '4. List of Invoice': 'netting.csv', '5. Credit Limit': 'cl.txt', '6. SID Client (daftar nasabah PEI)': 'sidclient.xlsx', '7. Participant Stock Position': 'pos.txt', '8. Data Transaksi LR': 'lrsettle.txt' };
  for (const [label, file] of Object.entries(map)) {
    const inp = d.querySelector(`input[type=file][aria-label="${label}"]`); assert.ok(inp, label);
    Object.defineProperty(inp, 'files', { value: [new w.File([over[file] || fs.readFileSync(D + '/' + file)], file)], configurable: true });
    inp.dispatchEvent(new w.Event('change')); await wait(80);
  }
  await fillCl(w, d);
}
const fillCl = async (w, d) => { const i = d.querySelector('input[aria-label="Credit Limit Partisipan"]'); if (i && !i.value) { i.value = '160.000.000.000'; i.dispatchEvent(new w.Event('change')); await wait(); } };
const tab = async (d, name) => { [...d.querySelectorAll('#tabs button')].find((b) => b.textContent.includes(name)).click(); await wait(); };
const btnByText = (d, txt) => [...d.querySelectorAll('button')].find((b) => b.textContent.includes(txt));
const setFile = async (w, d, label, file) => { const inp = d.querySelector(`input[type=file][aria-label="${label}"]`); assert.ok(inp, label); Object.defineProperty(inp, 'files', { value: [file], configurable: true }); inp.dispatchEvent(new w.Event('change')); await wait(150); };
const toFile = (w, wb, name) => new w.File([w.XLSX.write(wb, { type: 'array', bookType: 'xlsx' })], name);
let n = 0; const t = async (name, fn) => { await fn(); n++; console.log('  ✓', name); };
const tabsOn = (d) => [...d.querySelectorAll('#tabs button')].filter((b) => !b.disabled).map((b) => b.textContent.replace('✓ ', ''));
(async () => {
  const { w, d, calls } = await boot('PEI');
  await t('sebelum diproses: hanya Langkah 1 yang terbuka', async () => assert.deepStrictEqual(tabsOn(d), ['1. Upload']));
  await upload(w, d);
  await t('Langkah 1: tombol Proses aktif setelah 6 berkas wajib', async () => { const b = btnByText(d, 'Proses'); assert.ok(!b.disabled); b.click(); await wait(200); });
  await t('setelah proses: masuk Langkah 2 (RP); Langkah 3 & 4 masih terkunci', async () => {
    assert.ok(d.querySelector('#main table')); assert.match(d.getElementById('main').textContent, /Total Nasabah PEI Terdaftar/);
    assert.ok(!tabsOn(d).includes('3. LR') && !tabsOn(d).some((x) => x.startsWith('4.')));
    assert.ok(d.querySelector('#main .on, #tabs .on').textContent.includes('2. RP'));
  });
  await t('Langkah 2: tombol Lanjut mati sebelum RP Final diunduh', async () => assert.ok(btnByText(d, 'Lanjut ke Langkah 3').disabled));
  const rpInput = () => d.querySelector('input.cell[aria-label^="RP Final"]');
  let sidEdited, before;
  await t('Langkah 2: edit RP Final → ditandai & tercatat di Catatan', async () => {
    const inp = rpInput(); sidEdited = inp.getAttribute('aria-label').split(' ').pop(); before = Number(inp.value);
    inp.value = String(Math.round(before * 0.99)); inp.dispatchEvent(new w.Event('change')); await wait();
    const again = [...d.querySelectorAll('input.cell')].find((i) => i.getAttribute('aria-label') === `RP Final ${sidEdited}`);
    assert.strictEqual(Number(again.value), Math.round(before * 0.99)); assert.ok(again.className.includes('ed'));
  });
  let rpWb;
  await t('Langkah 2: unduh RP_Final (sheet "RP Final", kolom SID & RP Final) → Lanjut terbuka', async () => {
    btnByText(d, 'Simpan RP Final').click(); await wait();
    const o = calls.xlsx.pop(); assert.match(o.name, /^RP_Final_\d{8}\.xlsx$/); rpWb = o.wb;
    const rows = w.XLSX.utils.sheet_to_json(rpWb.Sheets['RP Final']); const r = rows.find((x) => x.SID === sidEdited);
    assert.ok(r); assert.strictEqual(r['RP Final'], Math.round(before * 0.99)); assert.ok('RP Max (rekomendasi)' in r);
    const nx = btnByText(d, 'Lanjut ke Langkah 3'); assert.ok(!nx.disabled); nx.click(); await wait();
  });
  await t('Langkah 3: sebelum upload RP Final tidak ada tabel LR', async () => {
    assert.match(d.getElementById('main').textContent, /Upload file RP Final untuk menghitung LR/); assert.strictEqual(d.querySelectorAll('input.cell[aria-label^="LR Final"]').length, 0);
  });
  await t('Langkah 3: file salah (tanpa kolom RP Final) ditolak dengan pesan jelas', async () => {
    const wb = w.XLSX.utils.book_new(); w.XLSX.utils.book_append_sheet(wb, w.XLSX.utils.json_to_sheet([{ A: 1 }]), 'X');
    await setFile(w, d, 'Upload RP Final', toFile(w, wb, 'salah.xlsx')); assert.match(d.getElementById('main').textContent, /Kolom "SID" dan "RP Final" tidak ditemukan/);
  });
  // file RP Final: buang satu SID lain supaya terbaca RP=0
  const rows0 = w.XLSX.utils.sheet_to_json(rpWb.Sheets['RP Final']); const sidDropped = rows0.find((r) => r.SID !== sidEdited).SID;
  await t('Langkah 3: upload RP Final → LR dihitung; SID yang tidak ada di file dianggap RP Final = 0', async () => {
    const wb = w.XLSX.utils.book_new(); w.XLSX.utils.book_append_sheet(wb, w.XLSX.utils.json_to_sheet(rows0.filter((r) => r.SID !== sidDropped)), 'RP Final');
    await setFile(w, d, 'Upload RP Final', toFile(w, wb, 'RP_Final_x.xlsx'));
    assert.match(d.getElementById('main').textContent, /RP Final dimuat/); assert.ok(d.querySelectorAll('input.cell[aria-label^="LR Final"]').length > 0);
    assert.match(d.getElementById('main').textContent, /Cek LR-3/);
  });
  await t('Langkah 3: Hasil_TRX_PEI belum bisa lanjut sebelum dibuat; setelah Final & Unduh → file berisi RP Final hasil upload', async () => {
    assert.ok(btnByText(d, 'Lanjut ke Langkah 4').disabled);
    btnByText(d, 'Final & Unduh').click(); await wait(); const o = calls.xlsx.pop(); assert.strictEqual(o.name, 'Hasil_TRX_PEI.xlsx');
    assert.deepStrictEqual(Object.keys(o.wb.Sheets), ['Repayment (RP)', 'Loan Request (LR)', 'Summary']);
    const sum = w.XLSX.utils.sheet_to_json(o.wb.Sheets.Summary);
    assert.strictEqual(sum.find((x) => x.SID === sidEdited)['RP Final'], Math.round(before * 0.99), 'edit RP ikut');
    assert.strictEqual(sum.find((x) => x.SID === sidDropped)['RP Final'], 0, 'SID tidak di file → 0');
    globalThis.__hasilWb = o.wb; assert.ok(!btnByText(d, 'Lanjut ke Langkah 4').disabled);
  });
  await t('Langkah 3: edit LR Final tersimpan & tercatat; Auto-Adjust tersedia', async () => {
    const lrIn = d.querySelector('input.cell[aria-label^="LR Final"]'); const sid = lrIn.getAttribute('aria-label').split(' ').pop(); const v = Number(lrIn.value);
    lrIn.value = String(Math.round(v * 0.5)); lrIn.dispatchEvent(new w.Event('change')); await wait();
    const again = [...d.querySelectorAll('input.cell')].find((i) => i.getAttribute('aria-label') === `LR Final ${sid}`); assert.strictEqual(Number(again.value), Math.round(v * 0.5));
    assert.ok(btnByText(d, 'Auto-Adjust LR')); assert.ok(btnByText(d, 'Simpan LR Final'));
    btnByText(d, 'Final & Unduh').click(); await wait(); globalThis.__hasilWb = calls.xlsx.pop().wb; // finalkan ulang dengan edit LR
  });
  await t('Langkah 4: terbuka setelah Hasil TRX dibuat; validasi baru jalan setelah upload', async () => {
    btnByText(d, 'Lanjut ke Langkah 4').click(); await wait(); assert.ok(btnByText(d, 'Jalankan Validasi').disabled);
    await setFile(w, d, 'Upload Hasil TRX PEI', toFile(w, globalThis.__hasilWb, 'Hasil_TRX_PEI.xlsx'));
    const b = btnByText(d, 'Jalankan Validasi'); assert.ok(!b.disabled); b.click(); await wait(200);
    assert.match(d.getElementById('main').textContent, /Limit Partisipan (LOLOS|GAGAL)/); assert.match(d.getElementById('main').textContent, /Export ke sistem \(iFast\)/);
    assert.match(d.getElementById('main').textContent, /LR-1 memakai Available Qty dari Risk Parameter/);
  });
  await t('Langkah 4: file iFast Repayment memakai RP Final hasil edit', async () => {
    btnByText(d, 'Excel (.xlsx)').click(); await wait(); const wb = calls.xlsx.pop().wb;
    const r = w.XLSX.utils.sheet_to_json(wb.Sheets['Repayment Direct']).find((x) => x['SID Client'] === sidEdited); assert.ok(r, 'SID hasil edit ada'); assert.strictEqual(r['Repayment Value'], Math.round(before * 0.99));
  });
  await t('Diagnostik: netting nol dijelaskan lewat UI', async () => {
    await tab(d, 'Kenapa tidak muncul?'); const sid = JSON.parse(fs.readFileSync(D + '/meta.json')).sids[5];
    const inp = d.querySelector('input[aria-label="Cari nasabah"]'); inp.value = sid; btnByText(d, 'Telusuri').click(); await wait(); assert.match(d.getElementById('main').textContent, /NETTING NOL/);
  });
  await t('Catatan: perubahan RP/LR tercatat (user lokal)', async () => { await tab(d, 'Catatan'); const m = d.getElementById('main').textContent; assert.match(m, /Lokal/); assert.match(m, new RegExp(sidEdited)); assert.match(m, /dari file RP Final/); });
  await t('Proses ulang di Langkah 1 mereset alur (langkah 3 & 4 terkunci lagi)', async () => {
    await tab(d, '1. Upload'); btnByText(d, 'Proses').click(); await wait(200); assert.ok(!tabsOn(d).includes('3. LR')); assert.ok(d.getElementById('tabs').textContent.includes('2. RP'));
  });
  await t('XSS: nama nasabah berisi HTML (diunggah lewat OP) tampil sebagai teks, bukan elemen', async () => {
    const evil = await boot('PEI'); const op = fs.readFileSync(D + '/op.txt', 'utf8').replace(/NASABAH \d+/g, '<img src=x onerror=alert(1)>');
    await upload(evil.w, evil.d, { 'op.txt': Buffer.from(op) }); btnByText(evil.d, 'Proses').click(); await wait(200);
    assert.ok(evil.d.querySelectorAll('#main table').length > 0); assert.strictEqual(evil.d.querySelectorAll('img').length, 0);
    assert.ok(evil.d.getElementById('main').textContent.includes('<img src=x onerror=alert(1)>'));
  });
  await t('upload folder: 6 berkas wajib dikenali otomatis dari nama, berkas lain diabaikan', async () => {
    const f = await boot('PEI'); const rd = (n) => fs.readFileSync(D + '/' + n);
    const names = { 'closing price 121026.xlsx': 'closing.xlsx', 'RiskParameter_20261013.txt': 'risk.txt', 'OP_20261013_ZZ.txt': 'op.txt', 'list invoice 121026.csv': 'netting.csv',
      'CreditLimit_20261013_ZZ.txt': 'cl.txt', 'SID Client.xlsx': 'sidclient.xlsx', 'Hasil_TRX_PEI (53).xlsx': 'closing.xlsx', 'RP_Final_20261013.xlsx': 'closing.xlsx', 'BUYMARGIN_ZZ_20261013084136.txt': 'pos.txt' };
    const inp = f.d.querySelector('input[aria-label="Pilih beberapa berkas"]'); assert.ok(inp);
    Object.defineProperty(inp, 'files', { value: Object.entries(names).map(([n, src]) => new f.w.File([rd(src)], n)), configurable: true });
    inp.dispatchEvent(new f.w.Event('change')); await wait(400); await fillCl(f.w, f.d);
    assert.match(f.d.getElementById('main').textContent, /6 dari 6 berkas wajib terdeteksi/);
    const b = btnByText(f.d, 'Proses'); assert.ok(!b.disabled);
    b.click(); await wait(200); assert.ok(f.d.querySelector('#main table'));
  });
  await t('upload folder: sebagian berkas kurang → diberi tahu mana yang belum ada, tombol Proses tetap mati', async () => {
    const f = await boot('PEI'); const inp = f.d.querySelector('input[aria-label="Pilih folder"]'); assert.ok(inp);
    Object.defineProperty(inp, 'files', { value: [new f.w.File([fs.readFileSync(D + '/risk.txt')], 'RiskParameter_20261013.txt'), new f.w.File(['x'], 'catatan.txt')], configurable: true });
    inp.dispatchEvent(new f.w.Event('change')); await wait(200);
    const txt = f.d.getElementById('main').textContent; assert.match(txt, /1 dari 6 berkas wajib terdeteksi/); assert.match(txt, /Belum ada: .*Closing Price/);
    assert.ok(btnByText(f.d, 'Proses').disabled);
  });
  await t('upload folder: nama tidak dikenali sama sekali → pesan jelas', async () => {
    const f = await boot('PEI'); const inp = f.d.querySelector('input[aria-label="Pilih beberapa berkas"]');
    Object.defineProperty(inp, 'files', { value: [new f.w.File(['x'], 'foo.txt')], configurable: true }); inp.dispatchEvent(new f.w.Event('change')); await wait(100);
    assert.match(f.d.getElementById('main').textContent, /Tidak ada berkas wajib yang dikenali/);
  });
  await t('tanpa login: tidak ada permintaan jaringan sama sekali', async () => assert.ok(!calls.net));
  console.log(`\n${n} tes UI lulus`); process.exit(0);
})().catch((e) => { console.error('GAGAL:', e.message); process.exit(1); });
