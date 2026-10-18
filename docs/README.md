# PEI Generator — versi HTML tanpa login

Port dari `pages/Daily RP and LR Calculation.py` (Streamlit). Aplikasi statis murni: semua perhitungan jalan di browser, tidak ada server, tidak ada login, **berkas nasabah tidak pernah keluar dari komputer**.

> **Versi umum** — dapat dipakai semua partisipan. Kode Partisipan dibaca otomatis dari berkas; Credit Limit Partisipan diisi sendiri.

## Struktur
```
public/    index.html, config.js (pengaturan versi), app.js, core.js, style.css, vendor/xlsx
tests/     tes perilaku, tes perbaikan, tes UI
```

## Cara pakai
**Paling cepat (lokal):** klik dua kali `public/index.html`. Tidak perlu instal apa pun.

**Dihosting (Cloudflare Pages):** push ke branch baru (mis. `html-version`, bukan `main`), Create → Pages → Connect to Git, Framework *None*, Build command kosong, Output directory `public`. Tidak ada secret/variabel yang perlu diisi.

> Tanpa login, siapa pun yang tahu alamatnya bisa membuka aplikasinya. Data tetap aman karena tidak diunggah, tetapi untuk hosting internal sebaiknya batasi akses (mis. Cloudflare Access) atau pakai lokal saja.

## Alur 4 langkah (berurutan, langkah berikutnya terkunci sampai langkah sebelumnya selesai)
1. **Upload** — 6 berkas wajib (bisa per folder) + berkas opsional (Participant Stock Position, Data Transaksi LR), lalu cek **Identitas partisipan** (Kode Partisipan & Credit Limit Partisipan). Klik *Jalankan Proses*.
2. **RP** — tabel nasabah PEI dengan transaksi Sell. RP Final default = RP Max; ubah bila perlu. Unduh `RP_Final_YYYYMMDD.xlsx`.
3. **LR** — unggah file RP Final dari langkah 2. LR dihitung ulang dari RP Final itu (nasabah yang tidak ada di file dianggap RP Final = 0). LR Final bisa diedit atau di-*Auto-Adjust* bila Nilai Beli melebihi kapasitas. Klik *Final & Unduh* untuk membuat `Hasil_TRX_PEI.xlsx` (opsional: `LR_Final_YYYYMMDD.xlsx`).
4. **Validasi & Export** — unggah `Hasil_TRX_PEI.xlsx`, jalankan validasi, Auto-Adjust bila perlu, lalu unduh file iFast (Repayment Proceed dan Loan Request).

Tab *Kenapa tidak muncul?* (penelusuran per nasabah) dan *Catatan* (log perubahan) bisa dibuka kapan saja setelah langkah 1.

Catatan LR-1: bila sheet LR pada file yang diunggah di langkah 4 adalah versi partisipan yang memuat kolom Available/Volume (≥14 kolom), itu yang dipakai. Jika tidak (mis. Hasil_TRX_PEI buatan aplikasi ini), LR-1 memakai Available Qty dari Risk Parameter.

## Identitas partisipan
- **Kode Partisipan** ditulis di kolom `Participant Code` pada file iFast. Kosongkan untuk dibaca otomatis dari kolom *Broker Code* berkas Credit Limit/OP, atau isi manual. Bila berkas memuat lebih dari satu kode, atau isian berbeda dengan berkas, proses ditolak.
- **Credit Limit Partisipan** (limit partisipan di PEI) dipakai untuk check limit partisipan di Langkah 4: CL Partisipan + total RP > total LR. Wajib diisi di Langkah 1 (tersimpan di browser untuk pemakaian berikutnya).

## Untuk apa file Hasil_TRX_PEI?
Hasil_TRX_PEI.xlsx adalah **berkas keputusan final** RP & LR hari itu (3 sheet: Repayment per saham, Loan Request per nasabah, Summary semua nasabah PEI). Fungsinya:
1. **Arsip** — bukti angka RP Final & LR Final yang diputuskan beserta dasar hitungnya (RP Max, Avail Efektif, LR @63%, binding, rasio, Auto-Adjust, status masuk iFast).
2. **Bahan review/approval** — bisa dikirim ke atasan untuk dicek sebelum upload ke iFast.
3. **Acuan file iFast** — di Langkah 4, RP Final, LR Final & status Auto-Adjust *di dalam file ini* yang dipakai untuk membuat file iFast. Jadi bila reviewer mengoreksi angka di Excel, koreksinya ikut. Syarat: sheet Summary dan sheet RP/LR harus sama, nilai tidak boleh melebihi RP Max / Max LR Final, dan SID-nya harus sama dengan data yang sedang diproses — kalau tidak, file ditolak dengan pesan yang jelas.

## Upload cepat per folder
Kumpulkan berkas harian dalam satu folder, lalu klik **Pilih Folder** (atau seret folder ke kotak). 6 berkas wajib dikenali dari **nama**; berkas lain di folder diabaikan:

| Berkas wajib | Nama harus memuat | Format |
|---|---|---|
| Closing Price | `closing price` | .xlsx |
| Risk Parameter | `RiskParameter` | .txt |
| Outstanding Position | diawali `OP_` | .txt |
| List of Invoice | `list invoice` / `netting` | .csv/.xls/.xlsx |
| Credit Limit | `CreditLimit` | .txt |
| SID Client | `SID Client` | .xlsx |

Jika ada lebih dari satu berkas yang cocok, dipakai yang terbaru (tanggal modifikasi). **Participant Stock Position** (nama memuat `StockPosition`) juga dideteksi otomatis; bila tidak ada, aplikasi memberi peringatan. Data Transaksi LR tetap diunggah manual. Pola nama ada di `SLOT_PATTERNS` pada `core.js`.

## Menjalankan tes
```
npm install
python3 tests/gen_data.py      # data sintetis (kode partisipan mengikuti public/config.js)
npm test
```

## Catatan perilaku (penting)
- **Satu sumber kebenaran**: RP/LR Final yang diedit dipakai oleh Validasi, file iFast, dan Hasil_TRX_PEI. Rerun tidak mereset editan.
- RP yang dipotong terlalu kecil bisa membuat check **RP-3 (rasio < 65%)** gagal sehingga nasabah **keluar dari file iFast** (terlihat di tab Validasi dan Diagnostik).
- Dua mode hitung dipertahankan seperti Python: TRX PEI (`pei`) dan Validasi (`val`). Beda hanya untuk nasabah yang menjual tetapi Loan Existing = 0.
- Tanpa file versi partisipan (≥14 kolom), check LR-1 memakai Available Qty dari Risk Parameter.

## Riwayat perubahan
### v5 (2026-10-18)
1. **LR Final tidak bisa melebihi Max LR Final.** Edit di atas batas ditolak dengan pesan; bila batas turun setelah RP diubah, check baru `LR-4. LR Final ≤ Max LR Final` gagal → nasabah tidak masuk iFast.
2. **Input angka ketat.** Diterima: `7845750`, `7.845.750`, `7,845,750`, `Rp 7.845.750`, `7.845.750,00`. Ditolak (tidak disimpan): desimal seperti `7.845.750,50` / `7845750.5`, angka negatif, teks.
3. **RP Final tidak bisa melebihi RP Max (Loan + Accrued).** Berlaku untuk edit di layar dan file RP Final (file ditolak seluruhnya). Check baru `RP-4. RP Final ≤ RP Max`. RP berlebih tidak bisa menaikkan Avail Efektif/kapasitas LR. File RP Final dengan nilai tidak terbaca ditolak, bukan dianggap 0.
4. **Langkah 4 benar-benar memakai Hasil_TRX_PEI yang diunggah.** File sembarang/beda hari ditolak.
5. **Rupiah bulat** di LR Final default, Loan Value iFast, Max LR Final & LR @63%.
6. **Participant Stock Position dideteksi otomatis** dari folder + peringatan bila tidak diunggah.
7. **Peringatan SID ganda di SID Client**, terutama bila CID/nama berbeda.
8. **Accrued yang sudah dibayar tidak dihitung lagi.** RP dialokasikan ke pokok dulu, sisanya ke accrued. Avail Efektif = Avail Limit + **pokok** yang dibayar (Used Limit di Credit Limit = pokok). RP Min (ref) dibatasi RP Max.

### v5.1 (2026-10-18)
- **Hasil_TRX_PEI mencatat Auto-Adjust**: kolom `Nilai Beli`, `Auto-Adjust` (Ya/Tidak), `Status LR`, `Status RP`. Status di file = isi file iFast.
- **Hasil usang terdeteksi**: ada perubahan setelah Hasil dibuat → wajib *Final & Unduh ulang*; di Langkah 4 tersedia *Unduh Hasil_TRX_PEI terbaru*.
- **Status Auto-Adjust ikut terbawa** saat Hasil_TRX_PEI diunggah (termasuk sesi baru). Ya→Tidak di Excel membatalkannya.
- **Pembatalan Auto-Adjust tercatat** (`BATAL AUTO-ADJUST`); tombol Batalkan juga di Langkah 4.
- Tab **Audit** menjadi **Catatan**; unduhan `Catatan PEI YYYYMMDD.csv`.

### v5.2 (2026-10-18)
- Dua versi dari satu kode: versi khusus satu partisipan (kode dikunci) dan **versi umum**. Perbedaannya hanya `public/config.js`.
- Kode Partisipan & Credit Limit Partisipan tidak lagi ditanam di kode. Versi umum membaca Kode Partisipan dari kolom Broker Code berkas Credit Limit/OP (atau diisi manual) dan menolak berkas yang memuat lebih dari satu kode.
