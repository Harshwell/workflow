# workflow

Google Apps Script workflow untuk claim pipeline, ingestion, routing, validation, enrichment, dan operational sheet maintenance.

Repo ini sudah punya fondasi yang cukup serius: single master workbook, policy-driven config, safe writer utilities, template-driven sheet assurance, serta flow orchestration untuk **MAIN**, **SUB**, dan **FORM**. Masalah utamanya bukan kurang fitur, tapi **kurang lapisan dokumentasi yang membuat perubahan jadi mudah dipahami dan dirawat**.

Dokumen ini sengaja dibuat tetap ramping. Untuk self-project, dokumentasi yang efektif lebih penting daripada dokumentasi yang banyak tapi jadi dekorasi.

---

## Update implementasi terbaru (2026-05-06)

Tambahan update untuk helper analitik `Report Base`:

- `Report Base` sekarang mengisi 6 helper column otomatis saat refresh:
  - `Position Detail`
  - `Position Detail Order`
  - `Status Aging Days`
  - `Status Aging Bucket`
  - `Submission Aging Days`
  - `Submission Aging Bucket`
- `Position Detail` dibuat case-insensitive + trim-safe:
  - `Position` kosong -> `Unknown`
  - `Position = Middle` + `PIC` kosong -> `Middle - Unassigned`
  - `Position = Middle` + `PIC` terisi -> `Middle - <PIC>`
  - selain `Middle` -> nama posisi rapi (title case)
- urutan pivot `Position Detail Order` memakai mapping tetap:
  - `Front` (1), `Expedition` (2), `Middle - Farhan` (3.1), `Middle - Meilani` (3.2), `Middle - Meindar` (3.3), `Back` (4), `Closed` (5)
  - fallback: `Middle - PIC lain` (3.9), `Middle - Unassigned` (3.99), lainnya (99)
- `Status Aging Days` diambil dari `Last Status Aging` / `LSA`, dan `Submission Aging Days` dari `TAT` (invalid -> blank), lalu dibucket untuk sorting pivot yang konsisten.

---

## Update implementasi terbaru (2026-05-04)

Tambahan fix untuk stabilitas optional sheet:

- `B2B` sekarang **tidak di-clear duluan**. Data hanya di-rebuild saat ada row hasil filter yang siap ditulis. Ini mencegah kondisi sheet jadi header-only ketika source run sementara kosong atau semua row ter-filter.
- matcher highlight berdasarkan note untuk `Second-Year (Market Value)`, `First-Month Policy`, dan `Policy Remaining <= 1 Month` dibuat lebih toleran (dengan/tanpa tanda titik), supaya kasus note tampil tapi warna kosong tidak terjadi lagi.
- default subject filter untuk email ingest `MAIN` digeser ke title terbaru: `3. Daily Claim Pending Monitoring`.
- jika prefix subject berubah lagi (contoh `4.`), cukup set Script Property `EMAIL_INGEST_SUBJECT` tanpa edit code.
- panel METABASE di `Overview` sekarang terisi konsisten: `Pulling Time` (start datetime), `Processing Time` (durasi eksekusi), dan `Flow` (`MAIN` / `SUB` / `FORM - MAIN` / `FORM - SUB`).

---

## Update implementasi terbaru (2026-05-03)

Tambahan update setelah perapihan mapping SC `Report Base`:

- mapping PIC `Middle` diperluas agar SC berikut tidak jatuh ke `Unknown`: `MDP`, `PT DELTASINDO...` (`deltasindo`), `EzCare` / `EZ Care`, dan `B-Store` (diarahkan ke `Meindar`).
- autofill `Branch` di sheet SC ditambah aturan: `MDP` -> `MDP` dan `PT DELTASINDO...` (`deltasindo`) -> `Deltafone`.

---


## Update implementasi terbaru (2026-04-27)

Tambahan update setelah batch Part 9:

- `POSITION_BY_LAST_STATUS['DONE_EXPIRED']` disejajarkan ke `Exclusion` agar konsisten dengan routing sheet Exclusion.
- `Submission by Month` sekarang disimpan sebagai **date** (tanggal 1 tiap bulan) dengan display format `MMM yy` di operational + Report Base.
- `Submission Date` operasional sekarang **strict** hanya dari `Raw Data.claim_submission_date` (tanpa fallback source lain), tetap ditulis sebagai tanggal valid.
- Setelah routing+enrichment, kolom `Submission Date` di sheet operasional utama (`Submission`,`Ask Detail`,`Start`,`Finish`,`PO`,`B2B`,`Special Case`) di-overwrite ulang dari mapping `Claim Number -> Raw Data.claim_submission_date` untuk mencegah kebocoran nilai dari kolom lain (mis. `OR`/`Remarks`).
- Sinkronisasi strict `Submission Date`/`Submission by Month` dijalankan ulang setelah optional processors (`B2B`/`EV-Bike`/`Special Case`) supaya row hasil rebuild `B2B` (termasuk FORM - MAIN) tetap terisi.
- Refresh `Daily Report Base` setelah SUB sekarang full rewrite (bukan incremental upsert) agar tidak menyisakan row stale saat user sedang pakai filter.
- Enrichment juga menulis `Submission by Month` ke sheet `B2B` agar konsisten dengan sheet operasional lain.
- Mapping PIC `Daily Report Base` ditambah fallback berbasis keyword `Service Center` jika position kosong/unmapped (contoh `B-Store` -> `Meindar`).
- Refresh `Weekly Report Base`: untuk **pure SUB** dijalankan maksimal 1x/hari di jam 09:00 (script timezone); untuk **FORM - SUB** dijalankan saat flow selesai (tidak terikat jam 09:00).
- Tambah manual trigger `runWeeklyReportBaseManual(snapshotDateOverride, sourceFileNameOverride)` untuk force update `Weekly Report Base` dari `Daily Report Base` saat diperlukan.

- normalisasi kolom bulan existing ditambahkan saat enforcement layout supaya nilai lama (text/date campur) dikonversi konsisten jadi date-format bulan.
- hardening optional sheets: excluded-status check di B2B & EV-Bike sekarang case-insensitive (normalisasi uppercase sebelum compare).
- partner mapping B2B diperluas sesuai kebutuhan operasional terbaru: `Bhinneka`, `PSMS`, `DIGIMAP EnE`, `Parastar`, `GSE`, `KPD`, `Tukar Ind`, `Bumilindo`.
- fallback `Submission Date` diperketat untuk skenario MAIN/SUB: prioritas `claim_submitted_datetime`, fallback aman ke `claim_submission_date`/`submission_date`.
- routing `Exclusion` ditambah status: `INSURANCE_CLAIM_WAITING_PAID` dan `CLAIM_CANCELLED` (termasuk position/excluded set).

---

## Update implementasi terbaru (2026-03-30)

Perubahan terbaru yang sudah masuk ke script:

- rename sheet routing `SC - Ivan` -> `SC - Meindar` pada policy + flow SUB
- hardening SUB relocate agar bucket internal seperti `__SC_SHARED__` tidak dianggap sebagai target sheet nyata
- update progress SUB di `Log` (`Progress`, `%`, `Current Step`, `Updated At`) dibuat konsisten sepanjang lifecycle run
- compaction log default ke mode v2-only (mengurangi duplikasi log legacy B:H)
- perbaikan mapping `Submission Date` (MAIN + SUB) agar lebih toleran pada variasi header source
- SUB reset workflow columns (`Update Status`, `Timestamp`, `Status`, `Remarks`) saat `Last Status` berubah
- optional sheet schema self-heal untuk `B2B` dan `EV-Bike` pada kolom `Claim Number`, `Start Date`, `End Date`, `Details`
- tambah final pass hardening check di `runSelfCheck_` untuk area UAT Part 9 (manual reset, B2B fallback, EV-Bike overlay, Report Base sync)
- tambah UAT checklist MAIN/SUB/FORM Part 9 di `docs/WORKFLOW_MAP.md`

---

## Tujuan dokumentasi

Layer dokumentasi repo ini dibuat dengan prinsip:

1. **Sedikit file, tinggi sinyal**
2. **Satu pintu masuk utama**
3. **Perubahan mudah ditelusuri**
4. **Setiap area punya source of truth yang jelas**
5. **Cocok untuk self-project, tapi tetap terasa enterprise**

Struktur dokumentasi yang dipakai:

- `README.md` → entrypoint utama, repo map, aturan maintenance, review temuan penting
- `docs/WORKFLOW_MAP.md` → flow map, diagram, dan change impact map

---

## Repository map

| File | Peran utama | Ubah di sini ketika... |
|---|---|---|
| `00_Config.gs` | source of truth untuk policy, mapping, konstanta, flags, routing, workbook/sheet config | menambah status, ubah routing, ubah policy bisnis, ubah IDs/sheet names, ubah feature flags |
| `01_Utils.gs` | utility layer: safe I/O, coercion, date parsing, header matching, retry, idempotency, gmail/drive helpers | butuh helper generik reusable, bukan business rule |
| `02_LogAndDetails.gs` | logging dan details reporting | ubah perilaku log/detail, struktur audit output |
| `03_SheetsAndValidation.gs` | sheet assurance, template header, dropdown propagation, layout enforcement | ubah template sheet, schema heal, dropdown/checkbox/layout |
| `04_ParseAndAging.gs` | parsing input dan aging derivation | ubah cara membaca file sumber / parsing dataset |
| `05a_Pipeline_RawMutate_Backup.gs` | raw mutation / backup stage | ubah tahap transform raw sebelum routing lanjutan |
| `05c_Pipeline_OptionalSheets.gs` | optional sheet processors | ubah logika B2B / EV-Bike / Special Case |
| `06a_EntryPoints.gs` | trigger entrypoints dan flow orchestration MAIN / SUB / FORM | ubah trigger, orchestration flow, queue consumer, attachment process orchestration |
| `06b_PipelineAndEnrichment.gs` | enrichment / main pipeline logic | ubah enrichment atau tahap pipeline utama |
| `06c_PostProcessAndUtils.gs` | post-process, status type, movement/webapp helpers, final utilities | ubah finalization, movement tracking, atau util pasca pipeline |

---

## Arsitektur singkat

Secara konseptual repo ini terdiri dari 4 layer:

### 1. Policy layer
Berada terutama di `00_Config.gs`.

Isi layer ini:
- status routing
- status type / position mapping
- workbook profile
- sheet template expectations
- feature flags
- optional sheet policy
- ingestion policy
- runtime knobs

### 2. Utility layer
Berada terutama di `01_Utils.gs`.

Isi layer ini:
- safe write helpers
- normalization helpers
- header matching
- typed coercion
- retry / idempotency
- generic Gmail / Drive helpers

### 3. Schema & presentation layer
Berada terutama di `03_SheetsAndValidation.gs`.

Isi layer ini:
- ensure sheet
- template header
- dropdown propagation
- checkbox enforcement
- number format / alignment
- profile-based sheet provisioning

### 4. Flow orchestration + processing layer
Berada terutama di `04_*`, `05*`, `06*`.

Isi layer ini:
- file ingestion
- parsing
- raw update
- routing
- enrichment
- optional sheets
- post-process
- trigger execution

---

## Flow yang ada

### MAIN
Flow utama untuk daily claim monitoring.

Ringkasnya:
1. cari email queued MAIN
2. ambil attachment dashboard
3. convert XLSX ke temp spreadsheet
4. jalankan pipeline ke master workbook
5. cleanup email jika sukses

### SUB
Flow operational dashboard incremental.

Ringkasnya:
1. cari email queued SUB
2. ambil 2 attachment: OLD + NEW
3. copy ke `Raw OLD` dan `Raw NEW`
4. update operational sheets by Claim Number
5. relocate row berdasarkan status routing
6. sort sheet operasional
7. movement snapshot / tracking
8. cleanup email jika sukses penuh

### FORM / MANUAL
Flow alternatif untuk submit file manual dari Form/Drive.

Ringkasnya:
1. baca flow selector dan file upload
2. autodetect MAIN atau SUB jika perlu
3. jalankan flow yang sama dengan orchestration utama
4. simpan timing dan log seperti flow lain

Detail diagram dan impact map ada di `docs/WORKFLOW_MAP.md`.

---

## Aturan maintenance

### Ubah status atau routing?
Mulai dari `00_Config.gs`.

Cek minimal bagian berikut:
- `OPS_ROUTING_POLICY`
- `STATUS_TYPE_BY_LAST_STATUS`
- `POSITION_BY_LAST_STATUS`
- `FINISH_STATUSES`
- policy sheet khusus seperti SC / PO / Exclusion / Special Case

### Ubah template kolom sheet?
Mulai dari `03_SheetsAndValidation.gs`.

Cek minimal bagian berikut:
- `SV03_TEMPLATES`
- `ensurePicSheets_`
- `sv03_ensureSheetWithHeader_`
- `sv03_enforceStandardLayoutForSheet_`

### Ubah parsing / datetime / header matching?
Mulai dari `01_Utils.gs` dan `04_ParseAndAging.gs`.

### Ubah trigger atau orchestration flow?
Mulai dari `06a_EntryPoints.gs`.

### Ubah optional sheet behavior?
Mulai dari `05c_Pipeline_OptionalSheets.gs` dan policy pendukung di `00_Config.gs`.

---

## Dokumentasi maintenance model

Supaya dokumentasi tetap rapi dan tidak jadi museum file markdown, gunakan aturan ini:

### Selalu update `README.md` jika:
- ada module baru
- ada flow baru
- ada perubahan source of truth
- ada sheet/route penting yang berpindah ownership
- ada perubahan cara maintainer harus melakukan modifikasi

### Selalu update `docs/WORKFLOW_MAP.md` jika:
- sequence flow berubah
- ada node proses baru
- ada branch logic baru
- ada perubahan dependency antar layer

### Jangan tambah file dokumentasi baru kecuali benar-benar perlu
Default-nya cukup 2 file ini.

Kalau suatu hari butuh file tambahan, prioritaskan urutan ini:
1. update file existing dulu
2. tambah section baru di file existing
3. baru buat file baru kalau memang tidak masuk akal digabung

---

## Checklist sebelum merge perubahan besar

- source of truth perubahan sudah jelas
- perubahan policy tidak duplikatif
- template sheet tidak bertentangan dengan routing
- flow MAIN / SUB / FORM tetap konsisten
- backward compatibility memang sengaja dipertahankan, bukan kebetulan
- perubahan status baru sudah ikut:
  - routing
  - status type
  - position
  - optional sheet logic jika relevan
  - dokumentasi jika dampaknya lintas modul

---

## Code review summary

Berikut temuan paling penting dari review awal repo ini.

### Update hardening terbaru (2026-03-26)

Putaran hardening terbaru sudah menutup beberapa sumber masalah yang paling berbahaya:

- konstruksi `CONFIG` tidak lagi memicu load-time crash karena referensi konstanta yang belum siap
- bootstrap patch maintenance tidak lagi men-shadow helper utama
- self-check sekarang membaca simbol global `const` / `let` dengan benar
- parsing datetime dan sort-under-filter diperketat di area operasional yang paling rawan drift
- repo sekarang punya `appsscript.json` eksplisit dan `tools/static_smoke_check.js` untuk smoke-check lokal

Catatan: daftar temuan di bawah tetap berguna sebagai konteks desain, tetapi beberapa poin prioritas tingginya sudah ditangani oleh hardening terbaru.


### Update stabilisasi berikutnya (2026-03-29, fase lanjutan)

Perubahan lanjutan yang sudah diterapkan setelah patch critical sebelumnya:

- `sv03_getDateAutoNumberFormatForColumn_` ditambahkan agar format `DATE_AUTO` tidak lagi melempar ReferenceError saat schema-format enforcement berjalan.
- `applyOperationalClaimHighlightsByRaw_` sekarang mengisolasi error per-sheet (try/catch per sheet), jadi kegagalan satu sheet tidak menghentikan highlight di sheet lain.
- EV-Bike sekarang dedup claim lebih ketat saat overlay dari `Submission` (menghindari proses ulang claim yang sudah diproses dari Raw pada run yang sama).
- Scan Event ID untuk sheet `Past` dibatasi (windowed read) agar post-process movement tracking lebih scalable pada sheet historis besar.
- Resolusi excluded-last-status di 05a diubah jadi lazy per-call + cache runtime (menghindari cache stale module-level).
- `backupOpsToRawFull_` tidak lagi me-reassign parameter `rawValues`; sekarang pakai buffer lokal (`workingRawValues`) supaya alur data lebih eksplisit.


### Update konsolidasi helper (2026-03-29, phase 3)

- Helper DB classifier dan parser datetime claim dipusatkan ke `01_Utils.gs` untuk memangkas duplikasi lintas modul.
- Helper map insurance di modul operasional diarahkan ke `mapInsuranceShort_()` yang sudah jadi source bersama.
- `getStatusTypeMap06c_` sekarang strict ke source-of-truth config (tanpa hardcoded fallback map lokal).
- `enrichOperationalSheetsFromRaw06_` mulai dipisah bertahap dengan resolver indeks raw terpusat agar maintenance lebih aman.
- Header matching mulai dikonsolidasikan ke util shared untuk mengurangi mismatch lintas modul.


### Update struktur SUB helper (2026-03-29, phase 4A)

- Implementasi helper SUB untuk append ke `Submission` dan sort operational dipisah dari `06a_EntryPoints.gs` dengan delegator kompatibilitas.
- Ditambahkan runtime preflight non-fatal di flow MAIN/SUB untuk deteksi dini missing symbol tanpa memutus run.

### Update konsolidasi modul 06 (2026-03-30)

- Fungsi maintenance/self-check, helper SUB, dan runtime preflight dikonsolidasikan ke `06c_PostProcessAndUtils.gs`.
- Modul script kini disederhanakan menjadi `00` sampai `06c` untuk mengurangi fragmentasi load-order dan biaya maintenance.
- Implementasi helper SUB untuk append ke `Submission` dan sort operational kini menjadi bagian dari blok konsolidasi di `06c_PostProcessAndUtils.gs`.
- `06a_EntryPoints.gs` tetap mempertahankan nama fungsi existing sebagai delegator supaya caller lama tidak pecah.
- Runtime preflight non-fatal di flow MAIN/SUB kini dijalankan dari blok konsolidasi `06c_PostProcessAndUtils.gs` untuk deteksi dini missing symbol tanpa memutus run.

### Yang sudah bagus

- Struktur modul numerik sudah memberi urutan mental yang cukup jelas
- Utility layer cukup kaya dan niatnya benar: safe write, normalization, retry, idempotency
- Config cukup kuat untuk jadi policy registry
- Sheet validation/template layer sudah lumayan matang
- Ada usaha observability, introspection, dan movement tracking
- Ada banyak guard untuk menjaga backward compatibility

### Yang perlu direvisi paling cepat

#### 1. Header validation masih berpotensi false negative
Ada indikasi util validasi schema masih mencampur **exact header map** dengan **normalized key lookup**. Secara praktis, ini bisa bikin header dianggap hilang padahal variasinya cuma beda casing/spacing.

**Prioritas:** tinggi

#### 2. Lookup policy details log tampak tidak konsisten
Ada indikasi sebagian kode membaca policy lewat `CONFIG.*`, sementara source of truth aktualnya berdiri sebagai constant global terpisah. Efeknya: fallback bisa terus kepakai tanpa sadar.

**Prioritas:** tinggi

#### 3. `00_Config.gs` terlalu besar untuk discovery cepat
Sebagai source of truth, file ini kuat. Sebagai file yang harus dipahami manusia, file ini terlalu padat. Mencari satu aturan terasa seperti audit forensik kecil-kecilan.

**Prioritas:** menengah

**Saran minimal:** jangan langsung pecah jadi banyak file. Mulai dengan section index yang stabil, naming convention yang lebih tegas, dan blok “change here when...” di setiap domain config.

#### 4. `06a_EntryPoints.gs` masih terlalu gemuk
File entrypoint ini tidak lagi murni entrypoint. Ia juga menampung cukup banyak orchestration detail, helper flow, sorting, relocation, dan beberapa concern operasional lain.

**Prioritas:** menengah

**Saran minimal:** pisahkan secara bertahap menjadi:
- entrypoints / installers
- flow runners
- sub-flow orchestration helpers

Tanpa perlu membuat terlalu banyak file sekaligus.

#### 5. Kontrak antar layer belum cukup eksplisit
Banyak fungsi sebenarnya sudah reusable, tapi kontraknya masih tersirat.

**Saran:** tambahkan docblock yang lebih tegas untuk fungsi yang jadi titik integrasi, misalnya:
- input assumptions
- output shape
- side effects
- source of truth dependency
- allowed caller layer

#### 6. Backward compatibility branch terlalu banyak di beberapa area
Ini wajar untuk repo Apps Script yang berkembang organik, tapi lama-lama jadi mahal dibaca dan diuji.

**Saran:** tandai dengan jelas mana yang:
- legacy but required
- temporary fallback
- safe to remove later

---

## Refactor strategy yang saya rekomendasikan

Untuk repo ini, pendekatan terbaik **bukan** “pecah semua sekarang”. Itu overkill untuk self-project dan malah bikin maintenance makin nyebelin.

Pakai strategi 3 tahap berikut:

### Tahap 1 — stabilisasi discovery
- rapikan dokumentasi
- tandai source of truth
- perjelas module ownership
- perjelas change impact

### Tahap 2 — kecilkan cognitive load
- kurangi helper ganda
- rapikan lookup policy yang campur antara constant vs config object
- rapikan contract function yang jadi boundary

### Tahap 3 — split hanya yang paling padat
Prioritas split nanti:
1. `06a_EntryPoints.gs`
2. `00_Config.gs`
3. baru area lain jika memang masih sakit dibaca

---

## Prinsip commit ke depan

Agar repo ini tetap enak dipelihara, usahakan commit mengikuti pola ini:

- `docs: ...` untuk perubahan dokumentasi
- `fix: ...` untuk bug yang mengubah behavior
- `refactor: ...` untuk rapih-rapih tanpa ubah behavior
- `feat: ...` untuk capability baru

Dan idealnya satu commit punya satu niat utama. Ya, konsep kuno tapi masih bekerja karena ternyata codebase tidak otomatis jadi rapi hanya karena niatnya baik.

---

## Next recommended actions

Prioritas paling masuk akal setelah dokumentasi ini:

1. perbaiki bug header validation di utility layer
2. rapikan lookup policy details log supaya source of truth konsisten
3. tambahkan section index di `00_Config.gs`
4. kurangi kepadatan `06a_EntryPoints.gs` tanpa meledakkan jumlah file

---

## Related docs

- [Workflow Map](docs/WORKFLOW_MAP.md)

---

## Maintenance note

Kalau repo ini terus tumbuh, jangan buru-buru menambah file dokumentasi. Biasanya masalahnya bukan kurang file, tapi kurang disiplin menjaga dua file utama tetap hidup.
