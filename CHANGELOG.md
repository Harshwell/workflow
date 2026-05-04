# Changelog

Repo ini memakai changelog ringan yang fokus ke perubahan yang benar-benar relevan buat maintainer.
Formatnya sengaja sederhana: **Added / Changed / Fixed**. Tidak perlu sok formal kalau ujungnya tidak pernah dibaca.

---

## 2026-05-04

### Fixed
- `processB2B_` tidak lagi membersihkan isi sheet sebelum ada row pengganti. Ini mencegah kondisi sheet `B2B` jadi header-only saat run tertentu tidak menghasilkan kandidat row (mis. jendela Raw sementara kosong / semua row terfilter).
- Highlight claim number untuk flag `Second-Year (Market Value)`, `First-Month Policy`, dan `Policy Remaining <= 1 Month` kini tetap mewarnai sel walau note detail memakai varian label tanpa tanda titik (normalisasi matcher note dibuat toleran).

---

## 2026-05-03

### Changed
- Normalisasi mapping SC di `Report Base` untuk menutup gap `Unknown`: keyword `MDP`, `PT DELTASINDO...` (`deltasindo`), `EzCare`/`EZ Care`, dan `B-Store` sekarang konsisten masuk jalur PIC `Meindar`.
- Autofill kolom `Branch` di SC sheets diperluas: `MDP` -> `MDP`, dan `PT DELTASINDO...` (`deltasindo`) -> `Deltafone`.

---

## 2026-04-27

### Changed
- `Submission by Month` di jalur operational + Report Base disimpan sebagai nilai **date bulanan** (tanggal 1) dengan display format `MMM yy` agar tetap terbaca bisnis namun aman untuk formula/pivot.
- `POSITION_BY_LAST_STATUS['DONE_EXPIRED']` disejajarkan ke `Exclusion` agar konsisten dengan routing sheet Exclusion.
- Dokumentasi diperbarui (`README.md`, `docs/WORKFLOW_MAP.md`) dengan ringkasan update Part 9 terbaru + maintenance quick-reference.
- Fallback source `Submission Date` di routing/enrichment diperluas untuk skenario MAIN/SUB (`claim_submitted_datetime`, `claim_submission_date`, `submission_date`).
- Mapping `Exclusion` ditambah `INSURANCE_CLAIM_WAITING_PAID` dan `CLAIM_CANCELLED` (routing + position + excluded-status set).

### Fixed
- Excluded-status filtering pada optional sheets (B2B/EV-Bike) dibuat case-insensitive dengan normalisasi uppercase sebelum cek membership.
- Mapping partner B2B diperluas untuk kebutuhan enterprise terbaru: `Bhinneka`, `PSMS`, `DIGIMAP EnE`, `Parastar`, `GSE`, `KPD`, `Tukar Ind`, `Bumilindo`.

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
- Layout Log canonical dipindah ke tabel v2 mulai kolom B (`No` .. `Severity`) sehingga tidak lagi menyisakan blok legacy B:H terpisah.
- SUB enrichment diperluas untuk `Submission Date` alias + optional fields `Store Name`/`PA Name`/`SPA Name`.
- Optional sheets (`B2B`, `EV-Bike`) menambah self-heal kolom wajib komputasi (`Claim Number`, `Start Date`, `End Date`, `Details`) bila belum ada.

### Fixed
- Perbaikan cleanup SUB email thread: konsisten memakai variabel label `queuedLabel` (sebelumnya typo `queueLabel`).
- Main routing `Submission Date` dibuat fallback-safe ke beberapa alias header (`claim_submission_date`, `claim_submitted_datetime`, `submission_date`).
- SUB reset kolom workflow (`Update Status`, `Timestamp`, `Status`, `Remarks`) saat `Last Status` berubah untuk menghindari stale state.


## 2026-03-29

### Changed
- `01_Utils` menambahkan helper bersama `computeDbValueFromClaimNumber_` dan `parseClaimLastUpdatedDatetime_` untuk mengurangi duplikasi lintas modul.
- `05b`, `05c`, `06b`, `06c` mulai konsolidasi util bersama (DB classifier, header matching, parser datetime, status-type source-of-truth).
- `05c` DRY_RUN guard disejajarkan ke `isDryRun_()` agar perilaku dry-run konsisten lintas modul.
- Split helper SUB dari `06a_EntryPoints` ke `06e_SubHelpers` (append Submission + sort operational) dengan delegator kompatibel di `06a`.
- Tambah `06f_RuntimeAssertions` + preflight non-fatal di `runPipeline_` / `runSubEmailIngest`.
- `static_smoke_check.js` diperbarui agar memuat helper SUB/runtime assertions baru.
- `enrichOperationalSheetsFromRaw06_` memakai resolver indeks raw terpusat (`__resolveEnrichRawIndexes06b_`) untuk menurunkan kompleksitas fungsi inti.
- Sort SUB di helper terpisah dipertegas: `Submission Date` -> `Last Status Date` -> `Last Status` (tetap dukung `sortSpecs` custom).
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
