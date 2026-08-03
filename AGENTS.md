# Workflow Repository — Agent Guidance and Second Brain

## Scope

Instruksi ini berlaku untuk seluruh repository `Harshwell/workflow`, termasuk:

- seluruh root Google Apps Script `00_*.gs` sampai `06c_*.gs`;
- `optional-project/Service Center Extractor.js`;
- `optional-project/SC-Meilani.js`;
- `optional-project/salvage`;
- `optional-project/Outstanding`;
- script validasi, manifest, konfigurasi, dan seluruh dokumentasi repository.

Instruksi eksplisit pengguna untuk task aktif selalu menjadi prioritas. Jangan memperluas scope tanpa kebutuhan yang jelas.

## Repository Mission

Repository ini mengelola automation Google Apps Script untuk claim workflow:

- ingestion dari Gmail, Form, dan Drive;
- parsing serta normalisasi data;
- MAIN, SUB, dan FORM orchestration;
- routing claim ke operational sheets;
- backup dan restore field manual;
- enrichment, validation, logging, report base, dan maintenance;
- automation standalone untuk Service Center, Salvage, Outstanding, dan SC-Meilani.

Prioritas utama setiap perubahan:

1. correctness dan integritas data;
2. idempotency serta aman untuk rerun;
3. konsistensi MAIN, SUB, FORM, dan standalone consumers;
4. efisiensi Spreadsheet API dan runtime Apps Script;
5. observability dan error traceability;
6. maintainability tanpa overengineering.

## Required Start Workflow

Sebelum mengubah apa pun:

1. Periksa `git status` dan pertahankan perubahan pengguna yang tidak terkait.
2. Identifikasi flow, script, sheet, mapping, atau data contract yang terdampak.
3. Gunakan bagian repository context di file ini sebagai starting point.
4. Baca hanya bagian `README.md` yang relevan dengan domain task.
5. Periksa bagian `Unreleased` di `CHANGELOG.md` untuk perubahan yang belum dirilis.
6. Inspeksi implementasi aktual dan seluruh consumer langsung sebelum membuat asumsi.
7. Tentukan validasi yang bisa dijalankan sebelum mulai mengedit.

Jangan membaca seluruh repository untuk perubahan kecil. Gunakan pencarian terarah berdasarkan nama fungsi, status, header, sheet, atau mapping.

## Source Precedence

Jika informasi berbeda:

1. instruksi task aktif;
2. behavior kode aktual dan test yang terverifikasi;
3. durable rules di `AGENTS.md`;
4. current-state documentation di `README.md`;
5. histori di `CHANGELOG.md`.

Jangan diam-diam memilih salah satu jika konflik memengaruhi behavior. Jelaskan konflik, tentukan source of truth, lalu selaraskan kode dan dokumentasi dalam scope yang sama.

## Architecture Map

| File/domain | Ownership utama |
|---|---|
| `00_Config.gs` | Policy registry, routing, status, PIC, sheet names, flags, runtime knobs, ingestion configuration |
| `01_Utils.gs` | Generic safe I/O, normalization, coercion, retry, header matching, Gmail/Drive helpers |
| `02_LogAndDetails.gs` | Log lifecycle, audit details, progress, error reporting |
| `03_SheetsAndValidation.gs` | Sheet assurance, templates, schema, dropdown, checkbox, layout, formatting |
| `04_ParseAndAging.gs` | Source parsing, date handling, aging derivation |
| `05a_Pipeline_RawMutate_Backup.gs` | Raw mutation, manual-field backup, restore preparation |
| `05b_Pipeline_RoutingOperational.gs` | Operational routing, SC destination selection, highlight, write behavior |
| `05c_Pipeline_OptionalSheets.gs` | B2B, EV-Bike, Doss, Special Case, dan optional-sheet processing |
| `06a_EntryPoints.gs` | Triggers, entrypoints, MAIN/SUB/FORM orchestration, queue behavior |
| `06b_PipelineAndEnrichment.gs` | Main pipeline, enrichment, continuation stage |
| `06c_PostProcessAndUtils.gs` | Post-process, Report Base, maintenance, runtime assertions, self-check |
| `optional-project/*` | Apps Script standalone; tidak ikut root deployment/load order |
| `static_smoke_check.js` | Local static runtime stub dan regression guard |

## Runtime Flow Contracts

### MAIN

- Mengambil queued MAIN email atau input FORM/MANUAL.
- Mengonversi attachment dan mengisi `Raw Data`.
- Menjalankan backup, routing, restore, enrichment, optional processing, reporting, dan finalization.
- MAIN dapat berjalan dalam continuation dua tahap menggunakan token dan one-shot trigger.
- Stage berikutnya harus mempertahankan RunID, progress, retryability, dan snapshot yang diperlukan SUB.
- Email/temp cleanup hanya dilakukan setelah success boundary tercapai.

### SUB

- Membutuhkan pasangan OLD dan NEW input yang valid.
- Mengisi `Raw OLD` dan `Raw NEW`.
- Meng-update claim existing, merelokasi row berdasarkan status, memperbarui optional consumers yang relevan, dan menyortir output.
- Jika MAIN memegang lock, SUB harus menggunakan pending/handoff behavior yang sudah disediakan; jangan membuat jalur paralel baru.
- Internal bucket seperti `__SC_SHARED__` tidak boleh dibuat sebagai physical sheet.

### FORM / MANUAL

- Mendeteksi jenis flow lalu menggunakan shared MAIN atau SUB core.
- Jangan membuat implementasi business logic terpisah jika behavior dapat menggunakan shared flow.
- Logging dan timing harus konsisten dengan flow asalnya.

### Standalone Projects

- `optional-project/*` mempunyai deployment dan trigger sendiri.
- Jangan mengasumsikan root constant/function tersedia di standalone project.
- Perubahan mapping bersama harus diperiksa pada seluruh consumer standalone.
- Pertahankan runtime guard, lock, logging, dan cleanup masing-masing project.

## Durable Business Invariants

### Service Center

- `GSI` dimiliki Meilani.
- `Rejeki Seluler`, variasi `Rejeki Seluller`, dan `CV Berkah Athallah` dimiliki Farhan.
- `PT DELTASINDO...` dimiliki Meindar dan menggunakan output/branch `Deltafone` sesuai contract saat ini.
- Samsung Authorized by Unicom Pontianak/Samarinda/Banjarmasin mengikuti mapping `Samsung Exclusive` yang berlaku.
- EzCare/EZ Care mengikuti device/service split yang didefinisikan pada canonical mapping.
- Perubahan SC mapping harus diperiksa minimal pada root routing, Service Center PIC, Extractor, Salvage, Outstanding, dan SC-Meilani jika relevan.

### Operational Data

- `Submission Date` harus mengikuti strict source contract yang didokumentasikan di README dan implementasi aktif.
- Manual operational fields yang dikelola workflow harus dipertahankan saat MAIN/SUB reroute sesuai policy backup/restore.
- Field manual aktif mencakup `Update Status`, `Timestamp`, `Status`, `Remarks`, `AWB`, dan `Timestamp AWB` selama contract belum berubah.
- IMEI/SN harus diperlakukan sebagai text agar leading zero dan digit tidak berubah.
- Rerun tidak boleh menghasilkan duplicate claim atau menghapus data manual yang masih valid.

### Status and Routing

- `Reject Claim` menerima claim reject yang memenuhi active aging/update window yang berlaku.
- Status expired, finish, exclusion, and SC routing harus tetap sinkron dengan status type dan position mapping.
- Menambah status baru wajib mengevaluasi routing, status type, position, exclusions, optional sheets, relocation, report propagation, dan test.

### Logging and Reliability

- MAIN dan SUB menggunakan log lifecycle yang terpisah sesuai implementasi aktif.
- Setiap flow harus memperlihatkan start, current step, progress, success/failure, dan alasan error yang actionable.
- Jangan menelan exception tanpa log atau context.
- Cleanup destructive hanya dilakukan setelah target write/processing terverifikasi berhasil.

## Change Impact Rules

### Jika Status atau Routing Berubah

Periksa:

- `00_Config.gs` policy terkait;
- routing index dan relocation;
- status type dan position mapping;
- MAIN, SUB, dan FORM behavior;
- optional-sheet exclusions;
- Report Base propagation;
- smoke-check assertions;
- README mapping/contract;
- CHANGELOG Unreleased.

### Jika Service Center Mapping Berubah

Periksa:

- root SC routing;
- Service Center PIC derivation;
- branch/output naming;
- `Service Center Extractor.js`;
- `SC-Meilani.js`;
- `salvage`;
- `Outstanding`;
- README canonical mapping;
- static assertions untuk mencegah drift.

### Jika Kolom atau Header Berubah

Periksa:

- config/header aliases;
- parser dan header resolver;
- template/schema enforcement;
- writer;
- backup/restore;
- formatting dan validation;
- optional sheets;
- MAIN/SUB relocation;
- Report Base;
- README data contract.

### Jika Trigger atau Orchestration Berubah

Periksa:

- lock dan pending behavior;
- continuation token/properties;
- duplicate trigger prevention;
- retry and cleanup boundaries;
- RunID dan progress continuity;
- email label/read/trash behavior;
- runtime limit dan cooperative stop.

## Coding Rules

- Pertahankan Google Apps Script V8 compatibility.
- Hindari dependency baru jika native JavaScript dan helper existing cukup.
- Reuse generic helper hanya jika contract-nya benar-benar sama.
- Business rule harus berada di policy/owner yang jelas, bukan tersebar sebagai fallback hardcoded.
- Prefer batch reads/writes daripada Spreadsheet API call per row.
- Hindari full-sheet scan bila dapat menggunakan bounded range, index, atau map.
- Pertahankan idempotency dan deterministic output.
- Jangan membuat silent fallback untuk required configuration atau required headers.
- Legacy compatibility harus diberi alasan dan removal condition jika sifatnya sementara.
- Jangan melakukan refactor lintas modul yang tidak diperlukan oleh task.
- Jangan mengubah nama public trigger/entrypoint tanpa audit caller dan migration plan.

## Configuration and Security

- Jangan menambahkan token, password, credential, atau secret ke repository.
- Untuk konfigurasi environment-specific, prioritaskan Script Properties.
- Spreadsheet/file IDs bukan password, tetapi hindari menambah hardcoded production identifiers baru.
- Jangan menampilkan nilai sensitif dalam log.
- Perubahan scope OAuth atau `appsscript.json` harus dijelaskan dan divalidasi secara eksplisit.

## Validation Matrix

### Root Workflow Changes

Minimal jalankan:

```bash
node static_smoke_check.js
```

### Standalone Script Changes

Minimal jalankan syntax check pada file yang berubah:

```bash
node --check "optional-project/Service Center Extractor.js"
node --check "optional-project/SC-Meilani.js"
node --check "optional-project/salvage"
node --check "optional-project/Outstanding"
```

Jalankan hanya command yang relevan dengan file yang berubah.

### Runtime-Dependent Changes

Static test tidak membuktikan behavior Gmail, Drive, Spreadsheet, trigger, quota, atau Apps Script runtime sebenarnya. Jika runtime test tidak dapat dilakukan:

- nyatakan bahwa validasi hanya static;
- tuliskan UAT Apps Script yang masih diperlukan;
- jangan mengklaim end-to-end success.

## Documentation Ownership

Repository hanya memakai tiga dokumen inti:

### `AGENTS.md`

Berisi durable context dan aturan kerja agent:

- architecture ownership;
- flow contracts;
- business invariants;
- change-impact rules;
- validation dan completion workflow.

Update hanya jika konteks durable atau aturan kerja berubah. Jangan isi dengan diary atau daftar fix harian.

### `README.md`

Berisi current state:

- tujuan dan setup;
- repository map;
- current architecture;
- workflow dan Mermaid mapping chart;
- current SC/status/data contract;
- configuration properties;
- operational runbook;
- validation/UAT;
- change-impact reference.

README menjelaskan kondisi sekarang, bukan histori. Jangan menambah section `Update terbaru` berdasarkan tanggal.

### `CHANGELOG.md`

Berisi histori perubahan:

- `Unreleased` untuk perubahan aktif;
- kategori Added, Changed, Fixed, Removed, Security, atau Documentation sesuai kebutuhan;
- behavior yang berubah, flow terdampak, dan dampak operasional singkat.

Jangan menduplikasi arsitektur atau mapping table lengkap di changelog.

## Mandatory Knowledge Sync

Setelah implementasi dan validasi:

1. Review `AGENTS.md`, `README.md`, dan `CHANGELOG.md`.
2. Update `CHANGELOG.md` untuk setiap perubahan material pada kode atau behavior.
3. Update `README.md` jika current behavior, mapping, flow, configuration, schema, runbook, atau validation berubah.
4. Update `AGENTS.md` jika invariant, ownership, agent workflow, impact rule, atau validation requirement berubah.
5. Jika dokumen tidak perlu berubah, jangan touch file hanya untuk mengubah tanggal. Jelaskan singkat pada final report bahwa file telah diperiksa dan tidak memerlukan update.
6. Dokumentasi dan kode harus mendeskripsikan behavior akhir yang sama.

Jangan menyimpan commit SHA sebagai freshness marker di dokumen yang sama dengan commit tersebut karena menghasilkan self-reference yang tidak stabil. Gunakan Git history sebagai audit trail.

## Obsidian Policy

- Markdown repository tetap menjadi source of truth.
- Obsidian hanya viewer, search, backlinks, dan graph layer.
- Gunakan Markdown link biasa: `[README](README.md)` dan `[CHANGELOG](CHANGELOG.md)`.
- Jangan bergantung pada plugin Obsidian untuk memahami atau menjalankan workflow.
- Jangan commit workspace state, cache, plugin binaries, atau UI-specific files yang tidak diperlukan bersama.
- Canvas/Base boleh dipakai sebagai derived visualization, tetapi bukan source of truth.
- External notes/Graphy boleh mirror repository secara read-only; jangan jadikan external copy sebagai canonical version.

## Completion Checklist

Sebelum menyelesaikan task:

- scope request sudah terpenuhi;
- unrelated user changes tetap aman;
- seluruh consumer langsung sudah diperiksa;
- source-of-truth tidak diduplikasi tanpa alasan;
- error path dan retry behavior dipertimbangkan;
- validasi relevan dijalankan dan hasilnya dicatat;
- runtime-dependent gaps dijelaskan;
- tiga dokumen inti sudah direview dan diperbarui sesuai ownership;
- `git diff` diperiksa untuk perubahan tidak disengaja;
- tidak ada commit/push/PR kecuali pengguna meminta.

## Final Response Format

Laporkan secara ringkas:

1. hasil utama;
2. file dan behavior yang berubah;
3. validasi yang dijalankan beserta hasil;
4. dokumentasi yang diperbarui atau alasan tidak perlu;
5. risiko/UAT tersisa;
6. commit/PR hanya jika memang dibuat.
