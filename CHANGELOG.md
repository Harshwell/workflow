# Changelog

Perubahan material repository dicatat di sini berdasarkan outcome, flow/project terdampak, dan dampak operasional. Current state tetap berada di `README.md`.

## Unreleased

- Repair mirror Service Center Extractor mengurutkan output berdasarkan `Last Status Aging` descending, opsional `Branch` A-Z, `Last Status` A-Z, lalu `Status Type` A-Z; kolom `Branch` hanya dibuat/diisi untuk mirror Unicom dari nama bucket sumber. Registrasi mirror SC tambahan seperti Mitracare/iBox tetap tersedia melalui menu dan Script Property, dengan link kosong/tidak dapat diakses diisolasi agar hanya mirror terkait yang dilewati.
- Installer Service Center Extractor tidak lagi memanggil `SpreadsheetApp.getUi()` dari execution context manual; menu ditambahkan oleh open trigger setelah spreadsheet di-reload sehingga instalasi tidak gagal dengan context error.
- Installer menu Service Center Extractor sekarang membuat installable spreadsheet-open trigger unik `onOpenServiceCenterTransfer`, menghapus trigger duplikat, dan menulis status ke Overview agar menu tetap pulih ketika global `onOpen()` bertabrakan dengan script optional lain dalam project yang sama.
- Service Center Extractor mereset `Log - SC Transfer` pada awal setiap run dan menambahkan `onInstall()`/`installServiceCenterTransferMenu()` untuk memulihkan menu global spreadsheet ketika `onOpen()` belum terpanggil.
- Repair mirror setup sekarang memvalidasi akses workbook dan keberadaan sheet `Repair` sebelum menyimpan property; summary MIRROR menjadi warning eksplisit ketika `refreshed=0` agar tidak lagi terlihat seolah mirror berhasil.
- Service Center Extractor menambahkan menu `Setup Repair Mirror IDs` untuk menyimpan URL/Spreadsheet ID ke Script Properties; konfigurasi mirror yang kosong sekarang dicatat sebagai warning dan mirror tersebut dilewati tanpa menggagalkan core transfer.
- Outstanding `runWorker()` sekarang memperlakukan script-lock contention sebagai logged skip yang akan dicoba ulang oleh scheduled worker berikutnya, sedangkan Service Center Extractor menampilkan pesan error aktual pada progress cell agar diagnosis tidak berhenti pada pesan generik.
- Service Center Extractor sekarang memirror hasil bucket Unicom/Samsung Exclusive/Xiaomi Authorized dan Sitcomtara ke sheet `Repair` workbook SC terkait melalui required Script Properties, mempertahankan `Update from Service Center` by `Claim Number`, dan menerapkan dropdown ketat `Status Type`; sync feedback SC standalone diarahkan ke sheet universe owner PIC (`SC - Meilani` atau `SC - Farhan`).
- SC branch standalone merevisi Salvage Repair agar memakai cutoff `Approval Date >= 2026-08-01`, menambahkan flow `Start Repair` dari pilihan branch `Overview!G2:G4` ke sheet `Repair`, dan backup `Update from Service Center` ke Remarks di `SC - Farhan`/`SC - Meilani`/`SC - Meindar`.
- SC branch standalone menambahkan menu `Update SC Universe Remarks` untuk menyalin `Repair`.`Update from Service Center` ke `SC - Universe`.`Remarks` berdasarkan `Claim Number` dengan skip blank/duplicate dan structured log.
- SC branch standalone di-split menjadi profile `SC - Unicom`, `SC - GSI`, `SC - Sitcomtara`, `SC - Mitracare`, dan `SC - iBox`; standalone syntax check juga mencakup semua file profile baru.
- SC-Meilani Salvage Repair memakai mapping raw eksplisit (`claim_submitted_datetime`, `claim_number`, `ins_approve_datetime`, `repairer_location_store_name`, `insurance_partner_code`, `sum_insured_amount`, `device_brand`, `device_type`, `imei_number`, `claim_last_status_name`), output dashboard hyperlink berteks `LINK`, target columns opsional selain identifier, dan sort target by Branch A-Z lalu Approval Date terlama.
- SC-Meilani Salvage Repair sekarang membaca source `Raw Data`, membuat `DB Link`/dashboard hyperlink dari `Claim Number`, dan membatasi eligible row ke `Submission Date` mulai 1 Jun 2026.
- SC-Meilani standalone menambahkan menu manual `Backup Repair Remarks` untuk menyalin cell `Remarks` dari sheet `Repair` ke sheet `SC - Meilani` berdasarkan `Claim Number` dengan copy 1:1 termasuk isi dan format cell.

### Added

- Developer validation interface berbasis Node.js 22 dengan command `check:root`, `check:standalone`, `check:mappings`, `check:docs`, `check:diff`, dan aggregate `check`.
- Regression contracts untuk canonical Service Center mapping, status routing, header aliases, IMEI/SN text preservation, dan ownership manual fields, termasuk negative validator fixtures.
- GitHub Actions checks `code-contracts`, `docs-governance`, dan `sensitive-diff` untuk pull request serta push ke `main`.

### Changed

- MAIN menggunakan continuation dua tahap dengan Script Property token dan one-shot trigger; stage 2 mempertahankan RunID, melanjutkan progress/log stage 1, dan memakai snapshot durable tanpa membaca ulang seluruh operational sheet sebelum clear.
- Backup/restore operational mencakup enam field manual `Update Status`, `Timestamp`, `Status`, `Remarks`, `AWB`, dan `Timestamp AWB`, termasuk formula serta rich formatting yang didukung.
- MAIN/SUB memakai log terpisah (`Log - Main`, `Log - Sub`); handoff `_OPS_MAIN_SUB_TEMP` dipertahankan sampai SUB window pukul 09:00 memprosesnya.
- Optional Service Center Extractor menggunakan mapping canonical di kode untuk CV Berkah/Rejeki Seluler, GSI, Deltasindo/Deltafone, Samsung Unicom variants, dan EzCare Apple/non-Apple.

### Fixed

- SUB refresh EV-Bike/Doss hanya menulis managed fields dan tidak menimpa kolom manual/dropdown.
- `Claim Type` untuk Reject Claim dan `Service Center PIC` pada PO mengikuti mapping aktif.

### Removed

- Dua reference docs yang menduplikasi workflow/column contracts serta local Obsidian state/empty artifacts dikeluarkan dari tracked knowledge architecture.

### Security

- Added-lines gate menolak secret, private key, hardcoded Google resource identifier baru, internal email/URL yang tidak diizinkan, dan raw operational metadata pada logging.
- CI menggunakan read-only permissions dan action references yang dipin; production IDs yang sudah ada tetap menjadi backlog migrasi Script Properties.

### Documentation

- Knowledge architecture dinormalisasi menjadi tiga file: `AGENTS.md`, `README.md`, dan `CHANGELOG.md`.
- README dibangun ulang sebagai current-state handbook dengan architecture, flow, mapping, data/config contracts, recovery, security, governance, UAT, roadmap, serta parity checklist migrasi.
- Changelog dinormalisasi ke `Unreleased` dan histori per tanggal dengan kategori konsisten; diary/debug notes digabung menjadi outcome material.

## 2026-07-06

### Added

- MAIN dan SUB menambahkan operational sheet `Reject Claim` untuk status yang mengandung `reject` dan mempunyai last-status aging atau last-update recency maksimal 30 hari.

### Changed

- Root routing, PIC enrichment, Extractor, dan Salvage memindahkan GSI ke Meilani; Rejeki Seluler/Seluller tetap Farhan.
- Second strict Submission Date/Month sync dibatasi ke optional sheets yang baru diproses untuk mengurangi scan/write berulang.
- `Submission.TAT` menggunakan decimal-day satu digit; sheet lain tetap memakai numeric/integer behavior masing-masing.

### Fixed

- SUB dapat merelokasi claim existing ke `Reject Claim` dan branch/PIC mengenali Rejeki Seluler tanpa jatuh ke unmapped.
- B2B mengecualikan `DONE_EXPIRED`, `CLAIM_EXPIRE`, dan `CLAIM_EXPIRE_WALKIN`.
- Salvage hanya menulis timestamp run pada range kontrol yang benar.

### Removed

- WebApp Movement Tracking dikeluarkan dari runtime dan health check; Daily/Weekly Report Base tetap aktif.

### Security

- Tidak ada perubahan security material.

### Documentation

- Workflow dan column contracts diselaraskan dengan Reject Claim, SC mapping, dan runtime optimization.

## 2026-06-29

### Added

- Optional sheet `Doss` untuk claim token `DOSS`; EV-Bike menerima token `VVMAR` tanpa status exclusion.
- Pending SUB marker dan automatic drain setelah MAIN melepaskan lock.

### Changed

- `Aging Position`/`Aging Post.` dinormalisasi menjadi `Stage Aging`, bersumber dari aging field per destination dan tidak digunakan pada Submission.
- SUB relocation mempertahankan Stage Aging hanya ketika old/new status bucket sama; bucket berubah atau source kosong mereset ke `0`.
- `Submission.TAT` dihitung dari `claim_submitted_datetime` hingga runtime; active filters diperluas ke used range sebelum write/sort.
- `Store Name` menggunakan `outlet_name`; B2B MAIN hanya menerima kategori exact `B2B Partnership`, sedangkan SUB hanya meng-update row existing.
- Special Case menjadi MAIN-only dan mempertahankan seluruh flagged claim; SUB meng-upsert EV-Bike/Doss dari Raw OLD/NEW.
- Extractor menambahkan Samsung Authorized by Unicom variants ke `Samsung Exclusive` dan Deltasindo overrides ke `Deltafone`.

### Fixed

- `CLAIM_EXPIRE` dan `CLAIM_EXPIRE_WALKIN` masuk `Expired Claim`, yang ikut scope relocation agar claim dapat bergerak keluar lagi.
- IMEI/SN dipaksa plain text tanpa comma separator; strict Submission Date sync menolak boolean existing.
- Finish relocation memprioritaskan `Finish`; VVMAR/DOSS tidak ditahan di `SC - Unmapped`; highlight/note dipasang ulang setelah SUB relocation.
- Expired Claim mendapatkan Branch/PIC dan fallback type yang sesuai; Start/Finish/Expired membaca service/check-in source dengan fallback status.

### Removed

- Operational writers tidak lagi membuat/menulis `DB`, `Status Type`, `Update Status Asso`, `Timestamp Asso`, `Update Status Admin`, atau `Timestamp Admin`.

### Security

- Tidak ada perubahan security material.

### Documentation

- Kontrak kolom dan flow diperbarui untuk deprecation, Stage Aging, optional sheets, filter safety, dan mapping standalone.

## 2026-05-11

### Added

- Tidak ada capability baru.

### Changed

- Exclusion dimasukkan ke operational movement scope SUB.

### Fixed

- Status yang berpindah ke domain Exclusion tidak lagi tertahan pada sheet SC lama.

### Removed

- Tidak ada removal material.

### Security

- Tidak ada perubahan security material.

### Documentation

- Flow contract SUB diselaraskan dengan relocation Exclusion.

## 2026-05-06

### Added

- Daily Report Base memperoleh helper `Position Detail`, order, status/submission aging days, dan bucket.
- `fillWeeklyReportBase()` membangun snapshot historis dengan replace-by-date, zero-row continuity, previous/change helpers, sorting, serta manual runner.

### Changed

- Weekly refresh pure SUB dibatasi pukul 09:00 sekali per tanggal; FORM SUB dapat refresh segera setelah selesai.
- Report mapping menormalisasi Middle PIC/casing dan filter range diselaraskan setelah write.

### Fixed

- Weekly functions dapat membuka master workbook ketika active spreadsheet tidak tersedia dan memakai indexed recalculation untuk histori besar.
- Submission Date ditulis hanya dari valid `claim_submitted_datetime` dengan legacy `claim_submission_date` fallback; strict sync dijalankan ulang setelah optional writers.
- Daily refresh setelah SUB menjadi full rewrite dan melepas filter agar tidak meninggalkan stale rows.
- Special Case tidak bergantung pada kolom legacy untuk schema guard; reason details tetap tersedia melalui note/output.

### Removed

- Perhitungan duplicate Position Detail per row dihapus.

### Security

- Tidak ada perubahan security material.

### Documentation

- Report Base runbook dan date-source contract didokumentasikan.

## 2026-05-04

### Added

- Overview menampilkan Pulling Time, Processing Time, dan Flow secara konsisten untuk MAIN, SUB, FORM MAIN, dan FORM SUB.

### Changed

- Default MAIN subject menjadi `3. Daily Claim Pending Monitoring` dan dapat dioverride melalui Script Property.

### Fixed

- B2B tidak di-clear sampai replacement rows siap, sehingga source kosong/filter penuh tidak menghasilkan header-only sheet.
- Highlight policy-age tetap bekerja pada variasi punctuation note.

### Removed

- Tidak ada removal material.

### Security

- Tidak ada perubahan security material.

### Documentation

- Ingestion override dan B2B safety behavior didokumentasikan.

## 2026-05-03

### Added

- Tidak ada capability baru.

### Changed

- Daily Report Base mengenali MDP, Deltasindo, EzCare, dan B-Store sebagai Meindar; Branch mapping menormalisasi MDP dan Deltafone.

### Fixed

- Known SC keywords tersebut tidak lagi mudah jatuh ke PIC `Unknown`.

### Removed

- Tidak ada removal material.

### Security

- Tidak ada perubahan security material.

### Documentation

- Canonical SC/PIC/branch outcome dicatat.

## 2026-04-27

### Added

- Exclusion status menambahkan `INSURANCE_CLAIM_WAITING_PAID` dan `CLAIM_CANCELLED`.

### Changed

- `Submission by Month` menjadi date tanggal pertama dengan format `MMM yy`; `DONE_EXPIRED` disejajarkan ke position Exclusion.
- Optional excluded-status matching menjadi case-insensitive dan B2B partner coverage diperluas.
- Daily Report Base memakai full rewrite setelah SUB; Weekly refresh mengikuti SUB/FORM SUB gates dan mempunyai manual override.

### Fixed

- Submission Date dan Month di-resync dari canonical raw source setelah route/optional processors untuk mencegah kebocoran OR/Remarks atau boolean.
- PIC report mempunyai fallback berbasis Service Center saat Position kosong/unmapped.

### Removed

- Tidak ada removal material.

### Security

- Tidak ada perubahan security material.

### Documentation

- Maintenance reference diperbarui untuk routing, optional sheets, date contract, dan report refresh.

## 2026-03-30

### Added

- Tidak ada capability baru.

### Changed

- Modul `06d`, `06e`, dan `06f` dikonsolidasikan ke `06c_PostProcessAndUtils.gs`; root load-order menjadi `00` sampai `06c`.
- `SC - Ivan` direname menjadi `SC - Meindar` pada root policy/flow.
- Log memakai canonical v2 table, SUB progress per-stage, dan operational enrichment diperluas.

### Fixed

- Cleanup SUB memakai queue label variable yang benar; internal `__*` bucket tidak menjadi physical destination.
- SUB mereset workflow fields ketika Last Status berubah sesuai contract saat itu; routing/header date handling dibuat lebih toleran.

### Removed

- Modul root `06d_IntegratedMaintenance.gs`, `06e_SubHelpers.gs`, dan `06f_RuntimeAssertions.gs` sebagai file terpisah.

### Security

- Tidak ada perubahan security material.

### Documentation

- Architecture map diselaraskan dengan konsolidasi modul.

## 2026-03-29

### Added

- Shared DB/date helpers, runtime preflight assertions, DATE_AUTO resolver, per-sheet highlight isolation, dan bounded Past scan.

### Changed

- Header/status/insurance utilities mulai dikonsolidasikan; SUB helper dipisah bertahap dengan delegator kompatibel; EV-Bike dedupe diperketat.

### Fixed

- Excluded-status cache menjadi lazy per-run, Raw backup menghindari parameter reassignment, dan hardcoded status-type fallback dihapus.

### Removed

- Duplicate helper/fallback paths yang sudah digantikan shared source.

### Security

- Tidak ada perubahan security material.

### Documentation

- Refactor boundary dan validation path diperbarui.

## 2026-03-26

### Added

- `appsscript.json` eksplisit untuk V8, Asia/Jakarta, dan Stackdriver exception logging.
- Local static smoke harness yang memuat seluruh root source dan menjalankan `runSelfCheck_()`.

### Changed

- `CONFIG` dibangun setelah dependent constants; aliases source-of-truth diperluas; highlight/status/date helpers diarahkan ke canonical policy.

### Fixed

- Load-time TDZ crash, false-negative self-check global symbols, canonical highlight policy lookup, header matching, datetime parsing, dan sort di bawah filter.

### Removed

- Duplicate bootstrap override yang men-shadow helper utama.

### Security

- Manifest runtime dan exception logging menjadi eksplisit; OAuth scopes belum diaudit pada perubahan ini.

### Documentation

- Initial architecture, validation, dan refactor priorities didokumentasikan.
