# Changelog

Repo ini memakai changelog ringan yang fokus ke perubahan yang benar-benar relevan buat maintainer.
Formatnya sengaja sederhana: **Added / Changed / Fixed**. Tidak perlu sok formal kalau ujungnya tidak pernah dibaca.

---

## 2026-03-30

### Changed
- Konsolidasi modul `06d_IntegratedMaintenance.gs`, `06e_SubHelpers.gs`, dan `06f_RuntimeAssertions.gs` ke `06c_PostProcessAndUtils.gs` agar footprint script tinggal `00` sampai `06c`.
- `static_smoke_check.js` diperbarui agar load-order source hanya memakai modul `00` sampai `06c`.
- Dokumentasi arsitektur (`README.md`, `docs/WORKFLOW_MAP.md`) disesuaikan mengikuti konsolidasi modul 06.
- Rename routing sheet `SC - Ivan` -> `SC - Meindar` di policy + flow terkait.
- SUB ingest progress kini ter-update per stage melalui `setProgressForFlow_` (search, prep, OLD/NEW, relocate, snapshot, done/fail).
- Relocation routing index kini mengabaikan bucket internal (`__*`) agar tidak jadi target pseudo-sheet.
- Log compaction default ke `LOG_COMPACTION_USE_V2_ONLY = true` untuk mengurangi duplikasi log legacy B:H.
- SUB enrichment diperluas untuk `Submission Date` alias + optional fields `Store Name`/`PA Name`/`SPA Name`.
- Optional sheets (`B2B`, `EV-Bike`) menambah self-heal kolom wajib komputasi (`Claim Number`, `Start Date`, `End Date`, `Details`) bila belum ada.

### Fixed
- Perbaikan cleanup SUB email thread: konsisten memakai variabel label `queuedLabel` (sebelumnya typo `queueLabel`).
- Main routing `Submission Date` dibuat fallback-safe ke beberapa alias header (`claim_submission_date`, `claim_submitted_datetime`, `submission_date`).
- SUB reset kolom workflow (`Update Status`, `Timestamp`, `Status`, `Remarks`) saat `Last Status` berubah untuk menghindari stale state.


## 2026-03-29

### Phase 3 consolidation pass

### 2026-03-29 (phase 3 consolidation pass)
- `01_Utils` menambahkan helper bersama `computeDbValueFromClaimNumber_` dan `parseClaimLastUpdatedDatetime_` untuk menghapus duplikasi logika lintas modul.
- `05b`, `05c`, `06b` sekarang mendelegasikan DB/insurance helper ke utility terpusat (`computeDbValueFromClaimNumber_`, `mapInsuranceShort_`).
- `05c` DRY_RUN guard disejajarkan ke `isDryRun_()` agar perilaku dry-run konsisten lintas modul.
- `06c` `getStatusTypeMap06c_` tidak lagi membawa hardcoded fallback map; mapping hanya dari source-of-truth config/global.
- `06b` dan `06c` parser datetime claim kini mendelegasikan ke parser terpusat `parseClaimLastUpdatedDatetime_`.

### Phase 4 structural pass
- Split sebagian helper SUB dari `06a_EntryPoints` ke file baru `06e_SubHelpers.gs` (implementasi append Submission + sort operational dipindahkan; `06a` menyisakan delegator untuk menjaga kompatibilitas trigger/caller).
- `static_smoke_check.js` diperbarui untuk memuat `06e_SubHelpers.gs` agar validasi load-order tetap mencakup helper baru.
- Tambah modul `06f_RuntimeAssertions.gs` dan preflight non-fatal di `runPipeline_` + `runSubEmailIngest` untuk mendeteksi simbol penting yang hilang lebih awal.

### Changed
- Phase 4B-4D incremental hardening: `enrichOperationalSheetsFromRaw06_` sekarang memakai resolver indeks raw terpusat (`__resolveEnrichRawIndexes06b_`) untuk mengecilkan kompleksitas fungsi inti.
- `06e_SubHelpers.gs` sort SUB kini mengutamakan `Submission Date` -> `Last Status Date` -> `Last Status` dan tetap mendukung `sortSpecs` custom saat diberikan.
- Header matching lintas modul mulai dikonsolidasikan melalui util bersama `findHeaderIndexByCandidates_` (dipakai oleh 05a/06c).
- Split sebagian helper SUB dari `06a_EntryPoints` ke file baru `06e_SubHelpers` (implementasi append Submission + sort operational dipindahkan; `06a` menyisakan delegator untuk menjaga kompatibilitas trigger/caller).
- `static_smoke_check.js` diperbarui untuk memuat `06e_SubHelpers` agar validasi load-order tetap mencakup helper baru.
- Tambah modul `06f_RuntimeAssertions` dan preflight non-fatal di `runPipeline_` + `runSubEmailIngest` untuk mendeteksi simbol penting yang hilang lebih awal.

### Changed
- Phase 4B-4D incremental hardening: `enrichOperationalSheetsFromRaw06_` sekarang memakai resolver indeks raw terpusat (`__resolveEnrichRawIndexes06b_`) untuk mengecilkan kompleksitas fungsi inti.
- `06e_SubHelpers` sort SUB kini mengutamakan `Submission Date` -> `Last Status Date` -> `Last Status` dan tetap mendukung `sortSpecs` custom saat diberikan.
- Header matching lintas modul mulai dikonsolidasikan melalui util bersama `findHeaderIndexByCandidates_` (dipakai oleh 05a/06c).
### 2026-03-29 (phase 4 structural pass)
- Split sebagian helper SUB dari `06a_EntryPoints` ke file baru `06e_SubHelpers` (implementasi append Submission + sort operational dipindahkan; `06a` menyisakan delegator untuk menjaga kompatibilitas trigger/caller).
- `static_smoke_check.js` diperbarui untuk memuat `06e_SubHelpers` agar validasi load-order tetap mencakup helper baru.

### Changed
- Phase 4B-4D incremental hardening: `enrichOperationalSheetsFromRaw06_` sekarang memakai resolver indeks raw terpusat (`__resolveEnrichRawIndexes06b_`) untuk mengecilkan kompleksitas fungsi inti.
- Header matching lintas modul mulai dikonsolidasikan melalui util bersama `findHeaderIndexByCandidates_` (dipakai oleh 05a/06c).
### Changed
- `03_SheetsAndValidation`: tambah `sv03_getDateAutoNumberFormatForColumn_` sebagai resolver `DATE_AUTO` yang aman (date-only fallback, datetime bila sample berisi komponen waktu).
- `05b_Pipeline_RoutingOperational`: apply highlight operational kini memakai isolasi error per-sheet agar kegagalan satu sheet tidak memutus pemrosesan sheet lainnya.
- `05c_Pipeline_OptionalSheets`: dedup EV-Bike diperketat untuk overlay `Submission` pada claim yang sudah diproses dari Raw di run yang sama.
- `06c_PostProcessAndUtils`: scan Event ID sheet `Past` dibatasi dengan jendela baris terbaru (`PAST_EVENT_SCAN_MAX_ROWS`, default 5000) demi efisiensi di workbook besar.

### Fixed
- `05a_Pipeline_RawMutate_Backup`: cache `__EXCLUDED_LAST_STATUSES` dipindah dari module-load ke lazy per-call (`__getExcludedLastStatuses05a_`) dengan runtime cache.
- `05a_Pipeline_RawMutate_Backup`: `backupOpsToRawFull_` tidak lagi me-reassign parameter `rawValues`, mengurangi risiko drift state dan side effect yang tidak eksplisit.
- `06b` status-type fallback hardcoded dihapus agar tetap konsisten ke source-of-truth (`CONFIG`/`STATUS_TYPE_BY_LAST_STATUS`).

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
