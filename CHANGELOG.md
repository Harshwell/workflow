# Changelog

Repo ini memakai changelog ringan yang fokus ke perubahan yang benar-benar relevan buat maintainer.
Formatnya sengaja sederhana: **Added / Changed / Fixed**. Tidak perlu sok formal kalau ujungnya tidak pernah dibaca.

---


## 2026-03-29


### 2026-03-29 (phase 3 consolidation pass)
- `01_Utils` menambahkan helper bersama `computeDbValueFromClaimNumber_` dan `parseClaimLastUpdatedDatetime_` untuk menghapus duplikasi logika lintas modul.
- `05b`, `05c`, `06b` sekarang mendelegasikan DB/insurance helper ke utility terpusat (`computeDbValueFromClaimNumber_`, `mapInsuranceShort_`).
- `05c` DRY_RUN guard disejajarkan ke `isDryRun_()` agar perilaku dry-run konsisten lintas modul.
- `06c` `getStatusTypeMap06c_` tidak lagi membawa hardcoded fallback map; mapping hanya dari source-of-truth config/global.
- `06b` dan `06c` parser datetime claim kini mendelegasikan ke parser terpusat `parseClaimLastUpdatedDatetime_`.

### Changed
- `03_SheetsAndValidation`: tambah `sv03_getDateAutoNumberFormatForColumn_` sebagai resolver `DATE_AUTO` yang aman (date-only fallback, datetime bila sample berisi komponen waktu).
- `05b_Pipeline_RoutingOperational`: apply highlight operational kini memakai isolasi error per-sheet agar kegagalan satu sheet tidak memutus pemrosesan sheet lainnya.
- `05c_Pipeline_OptionalSheets`: dedup EV-Bike diperketat untuk overlay `Submission` pada claim yang sudah diproses dari Raw di run yang sama.
- `06c_PostProcessAndUtils`: scan Event ID sheet `Past` dibatasi dengan jendela baris terbaru (`PAST_EVENT_SCAN_MAX_ROWS`, default 5000) demi efisiensi di workbook besar.

### Fixed
- `05a_Pipeline_RawMutate_Backup`: cache `__EXCLUDED_LAST_STATUSES` dipindah dari module-load ke lazy per-call (`__getExcludedLastStatuses05a_`) dengan runtime cache.
- `05a_Pipeline_RawMutate_Backup`: `backupOpsToRawFull_` tidak lagi me-reassign parameter `rawValues`, mengurangi risiko drift state dan side effect yang tidak eksplisit.

## 2026-03-26

### Added
- `appsscript.json` dengan `runtimeVersion: V8`, `timeZone: Asia/Jakarta`, dan `exceptionLogging: STACKDRIVER` agar konfigurasi project tidak lagi implicit. 
- `tools/static_smoke_check.js` untuk smoke-check lokal berbasis Node tanpa dependency tambahan. Script ini memuat seluruh source Apps Script ke runtime stub dan menjalankan `runSelfCheck_()` untuk menangkap load-order error sebelum deploy.

### Changed
- `00_Config` dirapikan agar `CONFIG` dibangun **setelah** seluruh konstanta yang direferensikan sudah terinisialisasi.
- Alias source-of-truth di `CONFIG` diperluas supaya caller lama tidak membaca policy yang salah atau jatuh ke fallback diam-diam.
- Helper highlight/status/date parsing di `05b`, `06b`, dan `06c` diarahkan kembali ke source-of-truth utama, bukan patch-layer terpisah.
- `06d_IntegratedMaintenance` disederhanakan menjadi entrypoint maintenance + self-check; bootstrap override ganda yang men-shadow helper utama sudah dihapus.

### Fixed
- Crash load-time `CONFIG` / TDZ (`Cannot access 'COLUMN_TYPES' before initialization`) ditutup.
- `runSelfCheck_()` sekarang bisa mendeteksi konstanta global `const` / `let` dengan benar, tidak lagi false negative hanya karena simbol tidak menjadi properti `globalThis`.
- Policy highlight operasional sekarang membaca struktur canonical `CLAIM_HIGHLIGHT_POLICY.COLORS` dan `NOTES_CANONICAL`, bukan selalu jatuh ke warna/note default.
- Resolver header di routing operasional sekarang memakai helper header matching terpusat sehingga variasi casing / spacing / alias lebih tahan banting.
- Parsing datetime lintas modul diperketat agar native `Date(string)` hanya dipakai untuk string yang memang unambiguous; fallback liar yang berpotensi menggeser timezone sudah dipersempit.
- Sort operasional yang berjalan di bawah filter sekarang memakai indeks relatif terhadap range filter, sehingga tidak lagi salah sort saat filter tidak mulai dari kolom A.

### Notes
- Repo ini masih besar, tetapi duplikasi shadow-override yang paling mengganggu sudah dipangkas dulu. Prioritas berikutnya sebaiknya fokus ke pemecahan `06a_EntryPoints` secara bertahap, bukan kosmetik folder.
