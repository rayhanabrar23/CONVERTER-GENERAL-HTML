// Pengaturan versi aplikasi — satu-satunya file yang berbeda antar versi.
// Versi umum: dapat dipakai semua partisipan. Kode Partisipan dibaca otomatis dari berkas
// (kolom Broker Code di Credit Limit / OP) atau diisi manual; Credit Limit Partisipan diisi di Langkah 1.
window.PEI_CONFIG = {
  variant: 'umum',
  title: 'PEI Generator',
  chip: null,
  participantCode: null,        // null = tidak dikunci
  participantLabel: null,
  clPartisipan: null,           // wajib diisi pengguna (tersimpan di browser untuk pemakaian berikutnya)
  partnerLabel: 'partisipan',
};
