# Unreleased
- Optional project Service Center Extractor, Salvage, dan Outstanding sekarang mengikuti override Farhan untuk CV Berkah/Rejeki Seluler serta EzCare Apple sejak 15 Jul 2026; Extractor otomatis membuat/menulis tab CV Berkah dan Rejeki Seluler tanpa mapping manual.
- MAIN sekarang menggunakan continuation dua tahap dengan Script Properties token + one-shot trigger untuk memisahkan ingest/backup dari clear/route/restore; Log - Main berlanjut dengan RunID sama tanpa reset pada transisi stage.
- Memisahkan audit log MAIN (`Log - Main`) dan SUB (`Log - Sub`) dengan lazy migration yang aman.
- Menambahkan backup AWB/Timestamp AWB, mapping Claim Type Reject Claim, SC CV Berkah/Rejeki Seluler, dan split EzCare Apple per 15 Jul 2026.
- Mengganti Service Type menjadi Claim Type dan menambahkan Service Center PIC di PO.

# Changelog

Repo ini memakai changelog ringan yang fokus ke perubahan yang benar-benar relevan buat maintainer.
Formatnya sengaja sederhana: **Added / Changed / Fixed**. Tidak perlu sok formal kalau ujungnya tidak pernah dibaca.

---


## 2026-07-06

### Added
- Sheet operasional `Reject Claim` untuk klaim dengan `Last Status` mengandung keyword `reject` dan `days_aging_from_last_activity` atau `last_update_datetime` masih `<= 30` hari.

### Changed
- Mapping GSI dipindahkan ke Meilani di routing SC, Service Center PIC, Service Center Extractor, dan Salvage; Rejeki Seluler/Rejeki Seluller tetap diarahkan ke Farhan.
- WebApp Movement Tracking dihapus dari runtime dan health check pipeline. Daily/Weekly Report Base tetap dipertahankan.
- Strict sync kedua setelah optional processors dibatasi ke optional sheets untuk mengurangi scan/write operational sheets berulang.
- `Submission.TAT` diformat decimal 1 digit; sheet lain tetap integer.
- Dokumentasi workflow/column contract diperbarui untuk Reject Claim, SC mapping, dan perubahan runtime.

### Fixed
- SUB relocation sekarang dapat memindahkan klaim existing ke `Reject Claim` saat status berubah menjadi reject dan masih dalam window 30 hari.
- Branch/PIC mapping mengenali Rejeki Seluler agar tidak jatuh ke unknown/unmapped.
- B2B sekarang mengecualikan status `DONE_EXPIRED`, `CLAIM_EXPIRE`, dan `CLAIM_EXPIRE_WALKIN`.
- Salvage tidak lagi menulis Last Update Timestamp ke F1; timestamp run hanya mengisi `D1:E1`.

---

## 2026-06-29

### Changed
- MAIN/SUB/FORM tidak lagi membuat atau mengisi `DB`, `Status Type`, `Update Status Asso`, `Timestamp Asso`, `Update Status Admin`, dan `Timestamp Admin` pada operational/optional sheets.
- `Aging Position` / `Aging Post.` dinormalisasi menjadi `Stage Aging`.
- `Stage Aging` kini mengambil source per sheet: `Aging Ask Detail`, `Aging Start`, `Aging SC Receive`, `Aging Ins Approve`, `Aging Finish`, dan `Aging Expired`; `Submission` dikecualikan.
- `Submission.TAT` sekarang dihitung decimal-day dari `claim_submitted_datetime` sampai runtime.
- `CLAIM_EXPIRE` dan `CLAIM_EXPIRE_WALKIN` dipindahkan ke sheet `Expired Claim`.
- `EV-Bike` sekarang menerima klaim token `VVMAR` tanpa pengecualian status; sheet baru `Doss` menerima klaim token `DOSS`.
- `Special Case` kembali fokus MAIN-only; SUB/FORM tidak lagi memproses atau strict-sync sheet tersebut.
- SUB cross-sheet relocation kini smart-reset `Stage Aging`: bucket status lama/baru yang sama memakai aging source sheet tujuan dari `Raw Data`, sedangkan bucket berubah atau referensi kosong tetap reset ke `0`.
- Filter range aktif diexpand ke full used range sebelum MAIN/SUB write/sort agar hidden/out-of-filter rows tetap ikut update.
- Flag `Migration Policy` dari `Claimed Active Policies` menjadi prioritas highlight tertinggi dan digabung dengan note flag lain.
- `Store Name` operational sekarang diisi dari `Raw Data.outlet_name`.
- `B2B` MAIN hanya mengambil row dengan `id_business_partner_category_name = B2B Partnership`; SUB hanya memperbarui `Last Status` dan `Service Center` pada claim yang sudah ada.
- `Special Case` MAIN sekarang memasukkan semua claim yang memiliki flag, termasuk status yang sebelumnya masuk pruning done/closed.
- `optional-project/Service Center Extractor` menambahkan routing `Samsung Exclusive` untuk Samsung Authorized by Unicom Pontianak/Samarinda/Banjarmasin dan override Deltasindo untuk Sorong/Office.

### Fixed
- `Expired Claim` ikut scope operational movement SUB, termasuk saat status berubah keluar dari expired menuju sheet mapping lain.
- SUB refresh ikut meng-upsert `EV-Bike` dan `Doss` dari `Raw OLD` / `Raw NEW`.
- `Service Type` untuk `Start`, `Finish`, dan `Expired Claim` memakai `device_checkin_option_name` dengan fallback status ke `WALKIN` / `PICKUP`.
- `Expired Claim.Service Type` khusus status `CLAIM_EXPIRE` diisi `Ask Detail`.
- `Expired Claim` ikut autofill `Branch` dan `Service Center PIC` berdasarkan service center mapping yang sama dengan operational sheet lain.
- `IMEI/SN` ditulis sebagai plain text dan koma ribuan dibuang agar nomor tidak berubah format.
- Strict sync `Submission Date` diperluas ke operational/optional sheets yang aktif dan mengabaikan nilai boolean existing supaya tidak muncul `TRUE`.
- SUB refresh re-apply highlight/note setelah relocation agar warna flagging tidak hilang.
- SUB finish-status relocation memprioritaskan sheet `Finish`, sehingga claim finish tidak terseret kosong/pindah ke SC universe.
- `SC - Unmapped` tidak lagi menerima/menahan klaim token `VVMAR` atau `DOSS`.
- Kolom detail `Start Date`, `End Date`, dan `Details` dipulihkan untuk writer `Special Case`.

---

## 2026-05-11

### Fixed
- SUB relocation tidak lagi mengecualikan sheet `Exclusion`. Ini memastikan perpindahan status lintas sheet tetap jalan saat claim berubah ke status domain `Exclusion` (contoh `DONE_REPAIR` -> `DONE`) dan tidak tertahan di sheet SC.

---

## 2026-05-06

### Changed
- `refreshReportBaseFromOperational06_` sekarang menulis 6 helper column tambahan di `Report Base`: `Position Detail`, `Position Detail Order`, `Status Aging Days`, `Status Aging Bucket`, `Submission Aging Days`, dan `Submission Aging Bucket`.
- Derivasi `Position Detail` dipertegas: case-insensitive + trim-safe, `Middle - Unassigned` saat PIC kosong, serta canonical casing PIC Middle (`Farhan`/`Meilani`/`Meindar`) agar mapping pivot order stabil.
- `Position Detail Order` memakai mapping konfiguratif + fallback (`Middle - PIC lain` = `3.9`, `Middle - Unassigned` = `3.99`, lainnya = `99`) untuk menjaga urutan pivot tetap deterministic.
- Tambah proses historis `fillWeeklyReportBase(snapshotDateOverride, sourceFileName)` berbasis agregasi dari `Daily Report Base` ke `Weekly Report Base` (replace-by-snapshot-date, zero-row hilang, recalculate helper previous/change/last7days, sorting final).
- Hook weekly snapshot di akhir MAIN sekarang membawa konteks file utama (`sourceFileName`) dan snapshot date hasil extract filename (`yyyy-MM-dd` sebelum `T`) saat tersedia.
- Hardening filter range: setelah write flow MAIN, filter aktif pada operational + optional (`B2B`, `EV-Bike`, `Special Case`) + `Daily/Weekly Report Base` disinkronkan ke full used range tanpa membuang criteria filter yang ada.

### Fixed
- Refresh `Weekly Report Base`: jalur `SUB` tetap dibatasi 1x per hari di jam 09:00 (script timezone), sedangkan jalur `FORM - SUB` boleh refresh saat flow selesai (tanpa gate jam 09:00).
- Tambah manual trigger `runWeeklyReportBaseManual(...)` untuk force refresh `Weekly Report Base` dari `Daily Report Base` di luar jadwal otomatis.
- `fillWeeklyReportBase` + `runWeeklyReportBaseManual` kini fallback buka `CONFIG.masterSpreadsheetId` saat `getActiveSpreadsheet()`/context aktif tidak tersedia (hindari error `Spreadsheet tidak ditemukan`).
- Enrichment helper di `Report Base` tidak lagi menghitung `Position Detail` dua kali per row (mengurangi duplikasi perhitungan saat build output rows).
- Mapping PIC `Report Base` untuk position `Middle` kini tahan variasi casing/spacing pada `Position` dan `Service Center` (termasuk newline/karakter non-alfanumerik), sehingga keyword `MDP`/`deltasindo`/`ezcare`/`b-store` tidak lagi mudah jatuh ke `Unknown`.
- Sinkronisasi snapshot kini kompatibel dengan rename sheet `Report Base` -> `Daily Report Base` (tetap fallback ke nama lama untuk backward compatibility).
- Recalculate helper `Weekly Report Base` dioptimasi dari scan nested ke map index berbasis `(snapshotDate + dimensi kombinasi)` untuk menurunkan kompleksitas saat histori membesar.
- `fillWeeklyReportBase` kini menerima `ssOverride` dari MAIN pipeline agar tidak gagal pada konteks non-active spreadsheet (`Spreadsheet tidak ditemukan`).
- Penulisan `Submission Date` di routing operasional diperketat: hanya menulis nilai yang valid sebagai tanggal (hapus fallback raw string) untuk mencegah nilai non-date (mis. teks bebas) masuk ke kolom tanggal.
- Source `Submission Date` dipertegas strict ke `Raw Data.claim_submission_date` (tanpa fallback sumber tanggal lain) untuk menghilangkan drift antar-sheet.
- Setelah routing+enrichment, kolom `Submission Date` di sheet operasional utama (`Submission`,`Ask Detail`,`Start`,`Finish`,`PO`,`B2B`,`Special Case`) di-overwrite ulang dari mapping `Claim Number -> Raw Data.claim_submission_date` untuk mencegah kebocoran nilai dari kolom lain (mis. `OR`/`Remarks`).
- Sinkronisasi strict `Submission Date`/`Submission by Month` dijalankan ulang setelah optional processors agar row `B2B` hasil rebuild (termasuk flow `FORM - MAIN`) tidak tertinggal kosong.
- Refresh `Daily Report Base` setelah SUB diubah ke full rewrite, sekaligus melepas filter sheet sebelum write agar row tersembunyi filter tidak tertinggal stale.
- Fallback PIC `Daily Report Base` ditambah: bila position kosong/unmapped tapi keyword service center dikenal (mis. `B-Store`), PIC tetap terisi (`Meindar`).
- `Special Case` schema guard diperingan: `Start Date`/`End Date`/`Details` tidak lagi dianggap mandatory (hilangkan noise error `SPECIAL_CASE_SCHEMA_MISSING` untuk kolom legacy yang tidak dipakai).
- Detail penjelasan rule (`First-Month`, `Policy Remaining`, `Second-Year`) kini ditulis sebagai **note** di kolom `Reason`, sehingga tetap informatif tanpa ketergantungan kolom tambahan.

---

## 2026-05-04

### Fixed
- `processB2B_` tidak lagi membersihkan isi sheet sebelum ada row pengganti. Ini mencegah kondisi sheet `B2B` jadi header-only saat run tertentu tidak menghasilkan kandidat row (mis. jendela Raw sementara kosong / semua row terfilter).
- Highlight claim number untuk flag `Second-Year (Market Value)`, `First-Month Policy`, dan `Policy Remaining <= 1 Month` kini tetap mewarnai sel walau note detail memakai varian label tanpa tanda titik (normalisasi matcher note dibuat toleran).
- Default subject filter flow `MAIN` diperbarui ke `3. Daily Claim Pending Monitoring` agar ingest tetap menangkap email queue dengan title terbaru.
- Overview METABASE (`Pulling Time`, `Processing Time`, `Flow`) sekarang konsisten terisi untuk run `MAIN`, `SUB`, `FORM - MAIN`, dan `FORM - SUB` dengan format timestamp/duration yang sama.

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

- Routing operasional `Submission Date` sekarang prioritas ke `claim_submitted_datetime` lalu fallback ke `claim_submission_date`/`submission_date` agar SUB tidak blank dan tidak drift ke kolom non-tanggal.
- Enrichment operational sekarang juga mencakup sheet `B2B` untuk pengisian `Submission by Month` dari Raw agar kolom bulanan tidak kosong.
- Refresh `Weekly Report Base` dipindah menjadi hanya saat flow `SUB` selesai (termasuk trigger jam 9), tidak lagi dieksekusi di flow `MAIN`.
