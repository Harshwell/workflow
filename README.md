# Workflow Second Brain

Google Apps Script untuk ingestion, normalisasi, routing, enrichment, dan monitoring claim pada satu master workbook, ditambah empat automation standalone yang mempunyai deployment sendiri.

Dokumen ini adalah handbook current-state. Konfigurasi dan business rule tetap dieksekusi dari kode; README membantu manusia menemukan contract, owner, failure boundary, dan cara memvalidasinya tanpa perlu melakukan arkeologi kecil setiap kali ada perubahan.

## Quick Validation

Requirement lokal: Node.js `22.x` dan npm.

```bash
npm ci
npm run check
```

Command yang lebih sempit:

| Command | Scope |
| --- | --- |
| `npm run check:root` | Load-order dan static smoke check root Apps Script. |
| `npm run check:standalone` | Syntax check empat script di `optional-project/`. |
| `npm run check:mappings` | Contract tests mapping, routing, header, dan manual-field ownership. |
| `npm run check:docs` | Markdown, local link/anchor, heading policy, Mermaid, dan knowledge-file policy. |
| `npm run check:diff` | Documentation-drift dan sensitive-added-lines gate. |

Static validation tidak membuktikan behavior Gmail, Drive, Spreadsheet, trigger, quota, permission, atau Apps Script runtime. Gunakan [Apps Script UAT](#apps-script-uat) sebelum release yang menyentuh runtime.

## Document Ownership and Source Precedence

Repository hanya mempunyai tiga knowledge files yang dilacak Git:

| File | Isi | Bukan tempat untuk |
| --- | --- | --- |
| `AGENTS.md` | Durable invariants, ownership, agent workflow, impact rules, definition of done. | Mapping rinci, schema volatil, atau diary fix. |
| `README.md` | Current state, architecture, flow, mapping, data/config contract, runbook, dan UAT. | Histori berbasis tanggal. |
| `CHANGELOG.md` | `Unreleased` dan histori outcome per tanggal. | Duplikasi handbook atau debugging diary. |

Jika informasi bertentangan, gunakan urutan:

1. instruksi task aktif;
2. regression test dan behavior kode yang terverifikasi;
3. durable contract di `AGENTS.md`;
4. current-state handbook ini;
5. histori di `CHANGELOG.md`.

Expected contract dan observed implementation bukan hal yang sama. Implementasi yang bertentangan dengan contract tidak otomatis menjadi benar hanya karena kebetulan sedang berjalan.

## Glossary

| Istilah | Arti |
| --- | --- |
| MAIN | Full ingest dari daily monitoring export ke `Raw Data`, lalu rebuild/routing dan post-process. |
| SUB | Incremental ingest pasangan OLD/NEW untuk update dan relocation claim existing. |
| FORM | Form/Drive adapter yang memakai shared MAIN atau SUB core. |
| Operational sheet | Destination utama seperti `Submission`, `Start`, SC owner sheets, `Finish`, `PO`, dan exclusion queues. |
| Optional sheet | Output dengan contract khusus: `B2B`, `EV-Bike`, `Doss`, dan `Special Case`. |
| Canonical header | Nama field utama yang diharapkan runtime; alias hanya compatibility layer. |
| Manual field | Nilai/formula user-managed yang harus bertahan melewati clear, route, dan relocation. |
| Success boundary | Titik ketika output penting sudah berhasil sehingga cleanup input boleh dilakukan. |
| Pending SUB | Marker durable ketika SUB tidak memperoleh script lock karena MAIN sedang berjalan. |
| Continuation | MAIN dua eksekusi yang dihubungkan token Script Property dan one-shot trigger. |
| UAT | Verifikasi di Apps Script/workbook nyata; berbeda dari static test lokal. |

## System Context and Architecture

```mermaid
architecture-beta group inputs(cloud)[Input Sources] service gmail_main(database)[Gmail MAIN Queue] in inputs service gmail_sub(database)[Gmail SUB Queue] in inputs service manual_upload(disk)[Form / Manual Drive Upload] in inputs group pipeline(cloud)[Root Apps Script Pipeline] service entry(server)[06a Entry Points] in pipeline service parser(server)[04 Parser & Aging] in pipeline service raw(database)[Raw Data / Raw OLD / Raw NEW] in pipeline service mutation(server)[05a Backup & Raw Mutation] in pipeline service routing(server)[05b Operational Routing] in pipeline service optional(server)[05c Optional Sheets] in pipeline service enrichment(server)[06b Enrichment & Continuation] in pipeline service reporting(server)[06c Reporting / Recovery / Self Check] in pipeline service output(database)[Operational & Report Sheets] in pipeline group shared(cloud)[Shared Modules] service policy(server)[00 Policy & Config] in shared service utils(server)[01 Shared Utilities] in shared service logs(database)[02 Structured Logs] in shared service schema(server)[03 Schema & Validation] in shared service standalone(server)[Standalone Apps Script Projects] gmail_main:R --> L:entry gmail_sub:R --> L:entry manual_upload:R --> L:entry entry:R --> L:parser parser:R --> L:raw raw:R --> L:mutation mutation:R --> L:routing routing:R --> L:optional optional:R --> L:enrichment enrichment:R --> L:reporting reporting:R --> L:output output:R --> L:standalone policy:R --> L:entry policy:R --> L:routing policy:R --> L:optional utils:R --> L:parser utils:R --> L:routing schema:R --> L:routing logs:R --> L:entry logs:R --> L:enrichment
```

Policy dan config berada terutama di `00_Config.gs`; generic utility di `01_Utils.gs`; log lifecycle di `02_LogAndDetails.gs`; schema/layout di `03_SheetsAndValidation.gs`; parsing di `04_ParseAndAging.gs`; dan processing di `05*` serta `06*`. Standalone scripts tidak berbagi load order maupun global root.

## Repository Map

| Path | Owner utama | Ubah ketika |
| --- | --- | --- |
| `00_Config.gs` | Policy registry, status/routing maps, SC keywords, flags, Script Property defaults. | Business policy atau environment knob berubah. |
| `01_Utils.gs` | Safe I/O, normalization, coercion, retry, header matching, Gmail/Drive helpers. | Helper generik dengan contract lintas flow berubah. |
| `02_LogAndDetails.gs` | Structured log, progress, details, RunID, error context. | Observability atau layout log berubah. |
| `03_SheetsAndValidation.gs` | `SV03_TEMPLATES`, schema/layout, dropdown, checkbox, formatting. | Destination column/template berubah. |
| `04_ParseAndAging.gs` | Source parsing, dates, aging. | Input interpretation berubah. |
| `05a_Pipeline_RawMutate_Backup.gs` | Raw mutation serta backup ke Raw. | Pre-route mutation atau durable manual backup berubah. |
| `05b_Pipeline_RoutingOperational.gs` | Operational destination selection, row writer, SC split, reject window, highlight. | Status/SC routing atau operational output berubah. |
| `05c_Pipeline_OptionalSheets.gs` | B2B, EV-Bike, Doss, Special Case. | Optional-sheet eligibility atau output berubah. |
| `06a_EntryPoints.gs` | Trigger/entrypoint, queue, lock, pending SUB, FORM/SUB orchestration. | Trigger, ingestion, cleanup, atau relocation berubah. |
| `06b_PipelineAndEnrichment.gs` | MAIN pipeline, enrichment, strict date sync, continuation stage. | Pipeline sequence atau enrichment berubah. |
| `06c_PostProcessAndUtils.gs` | Manual restore, Daily/Weekly Report Base, maintenance, self-check. | Recovery, reporting, atau runtime assertions berubah. |
| `optional-project/Service Center Extractor.js` | SC transfer dari root outputs ke workbook Service Center. | Destination SC tab/PIC/branch mapping berubah. |
| `optional-project/SC-Meilani.js`, `optional-project/SC-GSI.js`, `optional-project/SC-Sitcomtara.js`, `optional-project/SC-Mitracare.js`, `optional-project/SC-iBox.js` | Manual mirror Salvage, upsert Salvage Repair, dan backup Remarks per branch/service-center standalone. `SC-Meilani.js` sekarang menjadi profile `SC - Unicom` untuk Unicom/Samsung Exclusive/Xiaomi Authorized; file lain mewakili GSI, Sitcomtara, Mitracare, dan iBox. | Branch/status/header contract project ini berubah. |
| `optional-project/salvage` | Gmail salvage ingest dan upsert `Salvage 25-26`. | Salvage source, target, mapping, atau scheduling berubah. |
| `optional-project/Outstanding` | Hourly queue mirroring ke Claim Outstanding. | Region/PIC routing, queue, retry, atau preserved fields berubah. |
| `optional-project/Project Apple/*` | Apple claim append-only sync ke `General` dan OnEdit formatting khusus `REQ FU` dalam satu Apps Script project. | Filter, target, trigger, atau formatting Apple Claim berubah. |
| `optional-project/Project Samsung/*` | Samsung-only claim append-only sync ke `General` dan OnEdit formatting khusus `REQ FU` dalam satu Apps Script project. | Filter, target, trigger, atau formatting Samsung Claim berubah. |
| `static_smoke_check.js` | Legacy root static harness yang dipanggil tooling Node. | Root symbol/load-order assertion berubah. |
| `scripts/`, `tests/` | Local validators dan regression contracts. | Developer interface atau guard berubah. |

## Panduan Operasional: MAIN, SUB, dan Outstanding

Bagian ini adalah pintu masuk untuk operator baru. Anggap tiga flow ini sebagai tiga pekerjaan berbeda:

| Flow | Analogi sederhana | Kapan dipakai |
| --- | --- | --- |
| **MAIN** | Membangun ulang papan kerja dari snapshot utama terbaru. | Saat file claim utama masuk, normalnya melalui antrean email MAIN. |
| **SUB** | Memperbarui claim yang sudah ada dan memindahkannya ke tahap terbaru. | Saat file monitoring OLD/NEW masuk, normalnya setiap jam. |
| **Outstanding** | Membuat dashboard ringkas lintas region dari hasil operasional. | Otomatis setiap jam pada project standalone Outstanding. |

### Peta besar data

```text
File MAIN (Gmail / Form / manual)
        |
        v
     Raw Data ------------------------------+
        |                                    |
        v                                    v
Operational sheets                    Optional sheets
(Submission, Ask Detail, OR,          (B2B, EV-Bike, Doss,
 Start, SC, Finish, PO, dll.)          Special Case)
        |
        +--------------------+---------------+
                             |
File SUB OLD/NEW ------------+  memperbarui dan merelokasi claim
                             |
                             v
                    Daily/Weekly Report Base
                             |
                             v
              Standalone Outstanding membaca
                  sebagian operational sheets
                             |
                             v
             Digi / Region / Unmapped dashboard
```

MAIN dan SUB bekerja pada **master workflow workbook** yang sama. Outstanding adalah project terpisah: ia membaca output operational dari workbook tersebut, lalu menulis snapshot ke workbook Claim Outstanding.

### Flow MAIN — membangun snapshot operasional

#### Tujuan

MAIN adalah refresh utama. Ia mengambil file export claim terbaru, menjadikannya `Raw Data`, lalu membangun ulang seluruh sheet kerja berdasarkan status, Service Center, dan aturan routing saat ini. Karena sheet operasional dibangun ulang, MAIN lebih berat daripada SUB.

#### Source dan target

| Jenis | Source/target | Penjelasan operator |
| --- | --- | --- |
| Input utama | Email berlabel `QUEUED_MAIN`, Form, atau `runManual()` | Email normalnya diproses satu thread valid per run; Form/manual memakai pipeline inti yang sama. |
| Input file | File utama, dengan aging/standardization bila tersedia | File utama wajib ditemukan; file tambahan memperkaya aging. |
| Landing source | `Raw Data` | Snapshot mentah/canonical dan tempat backup durable field manual. |
| Target operasional | `Submission`, `Ask Detail`, `OR - OLD`, `Start`, `Finish`, `Expired Claim`, `Reject Claim`, `SC - Farhan`, `SC - Meilani`, `SC - Meindar`, `SC - Unmapped`, `PO`, `Exclusion` | Claim masuk ke sheet sesuai `Last Status`; claim SC juga dipisahkan menurut Service Center. |
| Target opsional | `B2B`, `EV-Bike`, `Doss`, `Special Case` | Dibentuk dari aturan eligibility masing-masing, bukan diedit sebagai source utama. |
| Target laporan | `Daily Report Base`, `Weekly Report Base` | Snapshot untuk kebutuhan reporting setelah routing/enrichment. |
| Observability | `Overview`, `Log - Main`, details/progress terkait | Tempat operator melihat flow, progress, durasi, hasil, dan error. |

#### Flowchart MAIN

```text
Ambil satu input MAIN
        |
        v
Validasi konfigurasi + file
        |
        v
Parse file dan samakan header ke schema Raw
        |
        v
Backup field manual lama berdasarkan Claim Number
        |
        v
Tulis snapshot terbaru ke Raw Data
        |
        v
Ada minimal satu row yang bisa diroute?
   | tidak                         | ya
   v                               v
STOP AMAN: jangan hapus ops     Simpan snapshot/continuation
                                   |
                                   v
                          Clear operational sheets
                                   |
                                   v
                       Route Raw Data berdasarkan status
                                   |
                                   v
                     Restore field manual + formatting
                                   |
                                   v
                Enrichment + optional sheets + reports + sort
                                   |
                                   v
                       Success boundary terverifikasi
                                   |
                                   v
                         Cleanup email/temp input
```

#### Kolom yang perlu dipahami

Raw input menggunakan header teknis seperti `claim_number`, `claim_submitted_datetime`, `claim_last_status_name`, `repairer_location_store_name`, `days_aging_from_last_activity`, dan `device_imei`. Output operasional memakai nama yang ramah pengguna. Kelompok kolom utamanya:

| Kelompok | Kolom penting | Dipakai untuk |
| --- | --- | --- |
| Identitas | `Claim Number`, `DB Link` | Kunci deduplikasi, update, restore, dan navigasi dashboard. |
| Customer/partner | `Partner Name`, `Buss. Category`, `PM Name`, `APM Name`, `Insurance` | Konteks bisnis dan routing tertentu. |
| Device | `Device Type`, `Product`, `Device Brand`, `IMEI/SN` | Identitas perangkat; IMEI/SN wajib tetap text. |
| Workflow | `Last Status`, `Last Status Date`, `Service Center` | Menentukan sheet tujuan, PIC/SC, dan posisi claim. |
| Aging | `Last Status Aging`, `Activity Log Aging`, `TAT`, `Stage Aging` | Prioritas dan pemantauan SLA. |
| Finance/policy | `Sum Insured Amount`, `Claim Amount`, `Claim Own Risk Amount`, `Nett Claim Amount`, `% Approval`, `Start Date`, `End Date` | Analisis nilai claim dan policy. |
| Manual | `Update Status`, `Timestamp`, `Status`, `Remarks`, `AWB`, `Timestamp AWB` | Input manusia yang dibackup sebelum rebuild dan direstore berdasarkan `Claim Number`. |

Tidak semua sheet memiliki seluruh kolom. Sheet SC menambahkan `Type` dan `Branch`; `PO` menambahkan `OR` serta `Service Center PIC`; workflow sheet yang lebih sederhana memakai subset kolom sesuai template canonical.

#### Periode dan cara menjalankan

- Jalur normal: `runEmailIngest()` membaca antrean `QUEUED_MAIN`.
- Jalur alternatif: Form dan `runManual()` tetap masuk ke core MAIN yang sama.
- MAIN dapat dibagi menjadi dua execution. Stage 1 menulis `Raw Data` dan backup durable; one-shot continuation menjalankan clear, route, restore, enrichment, dan finalization dengan RunID yang sama.
- Jangan menganggap email selesai hanya karena sudah terbaca. Cleanup email/temp baru boleh terjadi setelah success boundary; pada failure, input dipertahankan agar bisa dicoba ulang.

#### Checklist operator MAIN

1. Pastikan input yang benar berada di queue MAIN.
2. Pantau `Overview` dan `Log - Main` sampai stage 2/finalization selesai.
3. Periksa jumlah row `Raw Data` dan beberapa sampel `Claim Number` di destination.
4. Periksa field manual lama tidak hilang setelah refresh.
5. Bila log menyebut `0 routable rows`, jangan melakukan clear manual; safety gate sengaja mempertahankan output lama.

### Flow SUB — update cepat dari snapshot OLD/NEW

#### Tujuan

SUB tidak membangun ulang semuanya. Ia membaca snapshot monitoring OLD dan/atau NEW, mencari claim berdasarkan `Claim Number`, memperbarui field status penting, kemudian memindahkan row ke sheet yang cocok dengan status terbaru. Ini membuat dashboard operasional tetap segar di antara refresh MAIN.

#### Source dan target

| Jenis | Source/target | Penjelasan operator |
| --- | --- | --- |
| Input normal | Email label `QUEUED_SUB` dengan subject monitoring yang dikonfigurasi | Consumer mengambil satu thread deterministik per run. |
| Attachment OLD | Nama mengandung `List of Claims with Aging` | Disalin penuh ke `Raw OLD`. |
| Attachment NEW | Nama mengandung `(Standardization)` | Disalin penuh ke `Raw NEW`. |
| Landing source | `Raw OLD`, `Raw NEW` | Snapshot sementara/canonical untuk update SUB. |
| Target update/relocation | `Submission`, `Ask Detail`, `OR - OLD`, `SC - Farhan`, `SC - Meilani`, `SC - Meindar`, `SC - Unmapped`, `Start`, `Finish`, `PO`, `Exclusion`, `Expired Claim`, `Reject Claim` | Row existing di-update dan dapat berpindah sheet. |
| Target opsional | `EV-Bike`, `Doss` | Managed fields direfresh tanpa menimpa field manual. |
| Target laporan | Report Base | Direfresh setelah update dan relocation. |
| Observability | `Overview`, `Log - Sub` | Menampilkan start, progress, relocation, result, dan error. |

#### Flowchart SUB

```text
Trigger SUB tiap jam (kecuali jam 08:00)
        |
        v
Cari satu email QUEUED_SUB
        |
        v
Kenali attachment OLD dan/atau NEW
        |
        v
Copy full data ke Raw OLD / Raw NEW
        |
        v
Deduplikasi snapshot berdasarkan Claim Number
        |
        v
Update kolom status pada row operasional existing
        |
        +--> OLD + status SUBMITTED: append ke Submission bila belum ada
        +--> NEW + status CLAIM_INITIATE: append ke Submission bila belum ada
        |
        v
Relokasi full row sesuai Last Status terbaru
        |
        v
Refresh EV-Bike/Doss + sort + Report Base
        |
        v
Jam 09:00? Restore handoff manual dari MAIN
        |
        v
Jika semua sukses: bersihkan email; jika gagal: biarkan queued
```

#### Kolom source dan kolom yang di-update

| Raw OLD/NEW | Kolom operasional | Fungsi |
| --- | --- | --- |
| `claim_number` | `Claim Number` | Kunci pencarian wajib. |
| `claim_submitted_datetime` | `Submission Date` | Tanggal submit dan data row baru Submission. |
| `dashboard_link` | `DB Link` | Link claim. |
| `partner_name` | `Partner Name` | Konteks partner. |
| `insurance_partner_code` / `insurance_code` | `Insurance` | OLD dan NEW dapat memakai nama raw berbeda. |
| `device_type` | `Device Type` | Informasi perangkat. |
| `device_imei` | `IMEI/SN` | Identifier perangkat, dipertahankan sebagai text. |
| `claim_last_status_name` | `Last Status` | Menentukan update dan perpindahan sheet. |
| `days_aging_from_last_activity` | `Last Status Aging` | Aging status terbaru. |
| `activity_log_aging` | `Activity Log Aging` | Aging aktivitas. |
| `repairer_location_store_name` | `Service Center` | Update SC dan routing SC. |
| `days_aging_from_submission` | `TAT` | Umur claim sejak submission. |

SUB juga mengisi `Activity Log` dari `activity_log` dengan fallback sesuai kontrak. Field manual pada row yang dipindahkan harus tetap dipertahankan; SUB tidak boleh mengganti business logic MAIN atau melakukan full clear.

#### Periode, lock, dan retry

- Installer memasang `runSubEmailIngest()` setiap jam di sekitar menit 20.
- Eksekusi jam **08:00** dilewati agar tidak bertabrakan dengan MAIN.
- Jika MAIN sedang memegang lock, SUB menulis pending handoff; MAIN mencoba menjalankannya setelah selesai.
- Handoff backup manual MAIN hanya direstore pada window **09:00**; SUB di jam lain tidak memakai backup tersebut.
- Kontrak operasional mewajibkan pasangan OLD dan NEW yang valid. Implementasi saat ini masih menoleransi OLD-only atau NEW-only dengan warning; anggap ini compatibility behavior yang perlu direkonsiliasi, bukan prosedur normal yang boleh diandalkan.

#### Checklist operator SUB

1. Pastikan marker nama attachment OLD/NEW benar.
2. Cek `Raw OLD` dan `Raw NEW` setelah proses.
3. Cek beberapa claim yang berubah status sudah pindah ke sheet baru dan tidak tersisa ganda.
4. Cek field manual pada row yang dipindahkan.
5. Jika gagal, jangan menghapus label/email queue; failure path sengaja membuat input tetap retryable.

### Flow Outstanding — dashboard snapshot per region

#### Tujuan

Outstanding membaca hasil operasional, menyaring claim yang masih relevan, lalu membuat workbook dashboard yang dibagi menurut region. Flow ini **tidak mengubah workbook utama**; ia hanya membaca Overview Claim dan menulis workbook Claim Outstanding.

#### Source, referensi, dan target

| Jenis | Sheet | Peran |
| --- | --- | --- |
| Source operational | `Submission`, `Ask Detail`, `OR - OLD`, `Start`, `SC - Farhan`, `SC - Meilani`, `SC - Ivan`, `Finish`, `PO` | Sembilan sheet yang dibaca oleh standalone Outstanding saat ini. |
| Referensi routing | `Store Region` pada workbook referensi | Row 1 berisi region, row 2 berisi subheader `Partner Name`, row 3+ berisi daftar partner. Mapping di-cache 10 menit. |
| Target khusus | `Digi` | Untuk partner `Digimap` dan `Digiplus`. |
| Target region | Nama region dari `Store Region` | Contoh canonical: `Sumatra`, `Jabalnusra`, `Special Project`, `B2B`, `National Retailer`, `Sulawesi`, `Kalimantan`. |
| Target fallback | `Unmapped` | Partner kosong atau tidak ditemukan tidak dibuang. |
| Internal | `_Runs`, `_Queue`, `_Staging`, `_Audit` | Sheet tersembunyi untuk status run, task, hasil sementara, dan audit. |

> Catatan: source Outstanding masih menyebut legacy `SC - Ivan`, sedangkan master workflow memakai canonical `SC - Meindar`. Ini adalah technical debt yang harus direkonsiliasi sebagai perubahan behavior terpisah, bukan diganti diam-diam saat menjalankan flow.

#### Flowchart Outstanding

```text
Trigger enqueue setiap 1 jam
        |
        v
Buat run unik untuk jam tersebut
        |
        v
Pecah 9 source sheet menjadi task @ maksimum 600 row
        |
        v
Worker setiap 5 menit mengambil maksimum 6 task
        |
        v
Untuk setiap row:
  wajib ada Claim Number
  -> terjemahkan Last Status
  -> keluarkan Claim Expired
  -> hitung PIC
  -> route Partner ke Digi / Region / Unmapped
        |
        v
Simpan ke _Staging
        |
        v
Queue habis? Deduplikasi per destination + Claim Number
        |
        v
Backup Update Status + Timestamp Status
        |
        v
Clear data lama pada managed destination sheets
        |
        v
Tulis snapshot baru + restore dua field manual
        |
        v
Run = COMPLETED atau PARTIAL; event masuk _Audit
```

#### Kolom output Outstanding

| Urutan | Kolom | Catatan |
| ---: | --- | --- |
| 1 | `Submission Date` | Tanggal submission. |
| 2 | `Claim Number` | Kunci deduplikasi dan restore manual. |
| 3 | `Partner Name` | Dasar routing region. |
| 4 | `Insurance` | Insurer. |
| 5 | `Last Status` | Kode raw diterjemahkan menjadi label ramah pengguna; kode tak dikenal tetap ditulis apa adanya. |
| 6 | `Last Status Date` | Menentukan record terbaru ketika claim duplikat. |
| 7 | `Service Center` | Membantu penentuan PIC pada status middle/SC. |
| 8 | `Last Status Aging` | Aging status. |
| 9 | `Activity Log` | Aktivitas terakhir. |
| 10 | `Timestamp` | Timestamp source occurrence pertama. |
| 11 | `TAT` | Dinormalisasi menjadi angka jika memungkinkan. |
| 12 | `Update Status` | Field manual yang dipertahankan lintas refresh/region. |
| 13 | `Timestamp Status` | Field manual yang dipertahankan; fallback source dapat memakai occurrence kedua `Timestamp`. |
| 14 | `PIC` | `Adi & Adit`, `Suci & Yudha`, PIC SC, atau `Unknown`, berdasarkan raw status dan Service Center. |

#### Filter, deduplikasi, dan refresh

- Row tanpa `Claim Number` dilewati.
- Status yang diterjemahkan menjadi `Claim Expired` dilewati; tidak ada filter tanggal atau minimum aging tambahan.
- Deduplikasi memakai kombinasi **destination sheet + Claim Number**. Row dengan `Update Status` diprioritaskan; bila setara, `Last Status Date` terbaru menang.
- Finalization adalah **full snapshot refresh**, bukan append-only. Data lama pada managed sheet dibersihkan setelah nilai manual ditangkap.
- Hanya `Update Status` dan `Timestamp Status` yang secara eksplisit dipertahankan oleh standalone ini.

#### Periode, kapasitas, dan recovery

- `install()` memasang enqueue setiap **1 jam** dan worker setiap **5 menit**.
- Satu task maksimum **600 row**; satu worker maksimum **6 task**.
- Task gagal dicoba sampai **5 attempt** dengan exponential backoff, lalu menjadi `DEAD`.
- Task `IN_PROGRESS` lebih dari **20 menit** direcovery; run aktif lebih dari **90 menit** dibatalkan sebagai stale.
- Lima failure beruntun membuka circuit breaker selama **30 menit**.
- `runNow()` cocok untuk test/manual refresh; `runNowFresh()` mereset state internal lebih agresif dan tidak boleh menjadi pilihan rutin tanpa diagnosis.

#### Checklist operator Outstanding

1. Jalankan `install()` sekali pada deployment dan `runNow()` untuk smoke test.
2. Pantau `_Runs`: hasil normal `COMPLETED`; `PARTIAL` berarti ada task `DEAD`/`CANCELED`.
3. Pantau `_Audit` untuk source hilang, header hilang, partner unmapped, retry, atau finalization failure.
4. Periksa `Unmapped`; isinya biasanya berarti mapping `Store Region` perlu diperbaiki, bukan claim harus dihapus.
5. Setelah mengubah `Store Region`, beri waktu cache maksimal sekitar 10 menit atau lakukan force reload/debug terkontrol.
6. Uji bahwa `Update Status` dan `Timestamp Status` tetap ada setelah claim berpindah region.

### Pilih flow yang benar

| Situasi | Jalankan/cek |
| --- | --- |
| Ada export claim utama terbaru dan seluruh dashboard perlu dibangun ulang | **MAIN** |
| Hanya ada update monitoring status OLD/NEW dan dashboard harus disegarkan | **SUB** |
| Workbook utama sudah benar, tetapi dashboard per-region belum terbaru | **Outstanding** |
| Field manual hilang setelah refresh utama | Cek backup/restore dan `Log - Main`; jangan mencoba memperbaiki lewat Outstanding. |
| Claim tidak muncul di Outstanding tetapi ada di operational | Cek `Claim Number`, status `Claim Expired`, routing `Store Region`, lalu `_Audit`. |
| Claim ada di sheet operational yang salah setelah update status | Cek SUB relocation dan canonical status routing, bukan mapping region Outstanding. |

## Flow Registry

### MAIN

| Contract | Current behavior |
| --- | --- |
| Trigger/input | `runEmailIngest()` membaca maksimal satu thread dari label `QUEUED_MAIN`, unread, ber-attachment; `runManual()` dan FORM MAIN memakai core yang sama. |
| Lock/idempotency | Script lock dengan timeout 30 detik; queue query deterministic; transaction/idempotency guard aktif secara default. |
| Processing | XLSX dikonversi, ditulis ke `Raw Data`, manual state dibackup, lalu route, restore, enrich, optional processors, Daily Report Base, sort, dan finalization. |
| Continuation | MAIN dapat berhenti setelah stage 1, menyimpan `MAIN_PIPELINE_STAGE2`, lalu one-shot trigger menjalankan stage 2 dengan RunID yang sama. Progress bersifat kumulatif. |
| Success boundary | Cleanup Gmail/temp hanya sesudah route/finalization sukses. Jika gagal, queued email dipertahankan untuk retry. |
| Cleanup | Success: mark read, remove queue label, trash thread/temp sesuai policy. Failure: input tidak dikonsumsi. |
| Recovery | Stage 2 membaca snapshot durable stage 1. `_OPS_MAIN_SUB_TEMP` dipertahankan untuk handoff SUB pukul 09:00. |
| UAT minimum | Satu email valid, satu invalid attachment, rerun yang sama, manual-field/formula restore, stage-2 continuation, report refresh, dan cleanup success/failure. |

### SUB

| Contract | Current behavior |
| --- | --- |
| Trigger/input | `runSubEmailIngest()` membaca queue `QUEUED_SUB`; membutuhkan attachment NEW yang mengandung `(Standardization)` dan OLD yang mengandung `List of Claims with Aging`. FORM SUB memakai file Drive. |
| Lock/idempotency | Script lock; bila MAIN sibuk, `WORKFLOW_SUB_PENDING_AFTER_MAIN` dibuat dan didrain sekali setelah MAIN melepas lock. |
| Processing | Isi `Raw OLD`/`Raw NEW`, append kandidat baru ke `Submission`, update claim existing, relocate lintas operational sheets, refresh EV-Bike/Doss/B2B existing sesuai contract, sort, dan refresh reports. |
| Success boundary | Kedua input valid dan seluruh core update/relocation selesai. |
| Cleanup | Email/temp hanya dibersihkan setelah success; state retryable dipertahankan ketika gagal. |
| Recovery | Handoff `_OPS_MAIN_SUB_TEMP` hanya direstore pada window jam 09:00; fallback manual backup tetap tersedia. Pure SUB menjalankan Weekly Report Base hanya jam 09:00 dan maksimal sekali per tanggal. |
| UAT minimum | Missing OLD/NEW, same-bucket vs changed-bucket `Stage Aging`, pending lock, cross-sheet relocation, reject/expired movement, manual fields, optional refresh, dan email retry. |

### FORM and MANUAL

| Contract | Current behavior |
| --- | --- |
| Trigger/input | `onFormSubmit(e)` membaca field `Flow` dan upload `Metabase - Upload Claim Data`; field OLD/NEW terpisah bersifat optional. `runManual()` menerima file IDs. |
| Lock/idempotency | Memakai lock dan shared core flow asal; tidak mempunyai business-rule fork sendiri. |
| Processing | Deteksi MAIN/SUB, bentuk request, lalu delegasi ke pipeline yang sama. |
| Success boundary | Sama dengan flow yang dipilih. |
| Cleanup | File upload mengikuti cleanup policy shared flow setelah sukses. |
| Recovery | Error dicatat dengan flow/source context; FORM SUB boleh refresh Weekly Report Base segera setelah selesai. |
| UAT minimum | MAIN upload, SUB two-file upload, auto-detection, field yang hilang, file duplikat, log/timing, dan cleanup. |

### Standalone Projects

| Project | Entry/trigger | Lock/idempotency | Success and recovery contract |
| --- | --- | --- | --- |
| Apple Claim Sync | Dua file dalam `optional-project/Project Apple` dipasang bersama dalam satu Apps Script project: `Apple-Claim-Sync.js` menangani setup, manual recheck, dan time-driven sync pukul 09:00; `Apple-Claim-OnEdit.js` hanya menangani edit target `REQ FU`.`Status` kolom N untuk strikethrough/fill abu-abu row `CLOSED`. Isi Script Properties wajib `APPLE_CLAIM_SOURCE_SPREADSHEET_ID` dan `APPLE_CLAIM_TARGET_SPREADSHEET_ID` dengan ID atau URL Google Sheets, lalu `setupAppleClaimSync()` memasang ulang satu OnEdit khusus workbook target dan satu daily trigger sesuai timezone project. | Script lock men-serialize initial setup, scheduled, dan manual sync; OnEdit tidak lagi memantau `Raw Data` atau menjalankan sync. Installer menghapus trigger dengan handler sama sebelum membuat pengganti agar tidak terduplikasi; property invalid gagal eksplisit. Setiap run menulis waktu mulai ke `General!G2` dengan format `yyyy-MM-dd HH:mm:ss` dan lifecycle `ON PROGRESS`, `UPDATED`, atau `FAILED` ke `H2`. | Source `Raw Data` tetap memakai header raw exact seperti `claim_number` pada row 1. Di target `General`, script mencari header display `Claim Number` pada 20 row pertama dan seluruh kolom, sehingga header boleh berada di row 2, row 6, atau posisi kolom lain; data mulai satu row setelah header tersebut. Claim eligible dideduplikasi dan identifier baru ditambahkan tepat setelah Claim Number terakhir yang terisi, bukan setelah last row akibat formula/control kolom lain. Row lama tetap utuh walaupun claim hilang dari source; duplicate target gagal eksplisit dan kolom detail/formula di luar `Claim Number` tidak dimutasi. |
| Samsung Claim Sync | Dua file dalam `optional-project/Project Samsung` dipasang bersama dalam satu Apps Script project. `Samsung-Claim-Sync.js` membaca `Raw Data`, menerima hanya exact-normalized `device_brand = Samsung` dengan `claim_submitted_datetime >= 2026-06-01`, lalu menulis ke workbook target Samsung. Source dan target memakai Spreadsheet ID tetap yang ditetapkan di konfigurasi project, sehingga setup tidak memerlukan Script Property. Setup memasang OnEdit target dan daily sync pukul 09:00. | Script lock men-serialize initial setup, scheduled, dan manual sync. Installer menghapus trigger handler lama sebelum membuat pengganti. Setiap run menulis waktu mulai ke `General!G1` dan lifecycle `ON PROGRESS`, `UPDATED`, atau `FAILED` ke `G2`. | Kontrak source/header, status exclusion, deduplikasi, fail-fast duplicate target, pencarian header `General`.`Claim Number` pada 20 row pertama, dan append-only preservation sama dengan Apple. `Samsung-Claim-OnEdit.js` hanya memformat row `REQ FU` berdasarkan `Status` kolom N dan tidak menjalankan sync. |
| Service Center Extractor | `runServiceCenterTransfer()`; menu global spreadsheet via `onOpenServiceCenterTransfer()`; `installServiceCenterTransferMenu()` memasang installable open trigger unik tanpa memanggil UI dari execution context installer, sehingga tetap bekerja bila project yang sama memiliki `onOpen()` dari standalone lain. Jalankan installer sekali dari editor lalu reload spreadsheet agar open trigger menambahkan menu. | Script lock, minimum run interval, RunID properties. | Batch read/write, reset `Log - SC Transfer` pada awal setiap run, unmapped quarantine, dan verify destination. Bucket Unicom/Samsung Exclusive/Xiaomi Authorized dimirror ke workbook property `SC_REPAIR_MIRROR_UNICOM_SPREADSHEET_ID`, sedangkan Sitcomtara memakai `SC_REPAIR_MIRROR_SITCOMTARA_SPREADSHEET_ID`; hanya mirror Unicom yang memastikan dan mengisi kolom opsional `Branch` dari nama bucket sumber. Menu `Setup Repair Mirror IDs` mengelola mirror bawaan, sedangkan `Add/Update Optional SC Mirror` menerima nama bucket SC seperti Mitracare/iBox dan URL atau Spreadsheet ID yang disimpan pada `SC_REPAIR_MIRROR_CUSTOM_CONFIG`. Link boleh kosong dan setiap mirror yang kosong, tidak dapat diakses, atau kehilangan sheet `Repair` dilewati secara independen tanpa menggagalkan mirror lain maupun core transfer. Output `Repair` diurutkan berdasarkan `Last Status Aging` terbesar-ke-terkecil, lalu `Branch` A-Z pada mirror Unicom, `Last Status` A-Z, dan `Status Type` A-Z. Refresh mempertahankan `Update from Service Center` by `Claim Number`, membuat hyperlink dashboard, serta memasang dropdown ketat `Status Type`. |
| SC branch standalone | Menu `Salvage`, `Salvage Repair`, `Start Repair`, `Backup Repair Remarks`, `Update SC Universe Remarks`, `Run All` per profile (`SC - Unicom`, `SC - GSI`, `SC - Sitcomtara`, `SC - Mitracare`, `SC - iBox`); profile Sitcomtara mempunyai source dan destination workbook wajib yang terisi di konfigurasi deployment. | Script lock, active-run/stop properties per deployment. | Upsert by claim; `SC-Meilani` Salvage mempertahankan target-only claim, meng-update atau append claim source yang masih ber-Remarks `Unit belum ada`, dan hanya menghapus target claim ketika claim tersebut ada di source tetapi Remarks-nya sudah berubah. `SC-Meilani` Salvage Repair hanya menerima Approval Date mulai 1 Agustus 2026, meng-update atau append managed columns tanpa menghapus row target, serta menandai cell Claim Number pink dan memberi note jika claim target tidak ada di `Raw Data`; marker dibersihkan otomatis saat claim kembali. Salvage Repair reads explicit `Raw Data` raw headers, derives hyperlink text `LINK` menuju partner portal dari Claim Number tanpa formula `LET`/`REGEXMATCH`, supports optional target columns/sheet override, dan sort Branch A-Z lalu Approval Date tertua. `Start Repair` membaca branch `Overview!G2:G4`, mempertahankan feedback, dan menerapkan dropdown ketat `Status Type`. Trigger `Update SC Universe Remarks` menyalin feedback by `Claim Number` ke owner: Unicom/GSI/Sitcomtara → `SC - Meilani`, Mitracare/iBox → `SC - Farhan`. |
| Salvage | `setupSalvageAutomation()` memasang daily trigger; `runQueuedSalvage()` consumer. | Script lock dan dedupe thread/message window. | Duplicate target claim fail-fast; success baru men-trash temp/thread; failure tetap tercatat di `Log Salvage`. |
| Outstanding | `install()` memasang hourly enqueue dan worker 5 menit; manual `runNow()`/`runNowFresh()`. | Per-hour idempotency, durable queue, retry/backoff, stale-run rescue, circuit breaker; worker yang bertabrakan dengan eksekusi aktif melakukan logged skip dan dicoba ulang oleh trigger berikutnya, bukan runtime error. | Staging lalu finalizer; preserve manual columns; system sheets menyimpan run, queue, staging, audit. |

## Canonical Mapping Registry

### Status to Destination

Daftar status lengkap dieksekusi oleh `OPS_ROUTING_POLICY.LAST_STATUS_BY_SHEET`; `STATUS_TYPE_BY_LAST_STATUS` dan `POSITION_BY_LAST_STATUS` harus berubah bersama. Ringkasan domain:

| Destination/domain | Status contract | Additional consumers |
| --- | --- | --- |
| `Submission` | `SUBMITTED`, `CLAIM_INITIATE`. | SUB append rules dan report base. |
| `Ask Detail` | Ask-detail, resubmit-document, dan reopen depan. | Position/Status Type maps. |
| `OR - OLD` | `WAITING_PAYMENT`. | SUB relocation. |
| `Start` | Walk-in/pickup/courier start statuses. `COURIER_PICKUP_START_DONE` juga terlihat di SC universe. Pada SUB berikutnya, row dengan dropdown manual `Status = Delivered` dimirror ke SC owner: `Service Center PIC` diprioritaskan, lalu `Branch`, dan fallback ke `Service Center`. Row sumber tidak dihapus; `AWB`, `Timestamp AWB`, `Branch`, `Claim Type`, dan `Service Center PIC` tidak ikut dimirror, sedangkan SC `Type` diisi `Start`. | Service/Claim Type dan SC mirror rule. |
| SC universe | Receive, estimate, repair/on-progress, insurance review/approval, OR-repair, dan finish tracking statuses. | Split oleh `SC_NAME_KEYWORDS`, PIC, branch, Daily Report Base, standalone projects. |
| `Finish` | Repair/checkout/finish statuses; lima replacement-delivery status hanya masuk `Finish` dan tidak masuk SC universe. | SUB clone/relocate, managed `Repair Type`, dan reporting. |
| `Expired Claim` | `CLAIM_EXPIRE`, `CLAIM_EXPIRE_WALKIN`. | SUB relocation dapat memindahkan claim keluar lagi. |
| `Reject Claim` | Status yang mengandung `reject` dan `Last Status Aging <= 30`; jika aging tidak tersedia, last-update datetime harus berada dalam 30 hari. | MAIN route, SUB relocation, `REJECT_CLAIM_TYPE_BY_LAST_STATUS`. |
| `PO` | Replacement/back-stage statuses. | `OR` serta `Service Center PIC`. |
| `Exclusion` | Done/closed/paid/cancelled/rejected domain yang tidak memenuhi active Reject Claim window. | Position, optional exclusions, reports. |
| `SC - Unmapped` | Status tidak terpetakan atau SC-universe tanpa keyword match. | Structured mapping error; token `VVMAR`/`DOSS` tidak ditahan di sini. |

Unknown status/SC harus terlihat sebagai unmapped/error evidence, bukan disamarkan oleh fallback owner. Ketika menambah status, audit routing, type, position, exclusion set, optional sheets, SUB relocation, reporting, dan regression test.

### Service Center to PIC, Branch, and Output

| Canonical match | Root operational owner | Branch/output | Standalone consumers |
| --- | --- | --- | --- |
| Mitracare, iBox | Farhan | Nama canonical masing-masing | Extractor, Salvage, Outstanding bila relevan. |
| Sitcomtara | Meilani | `Sitcomtara` | Root routing dan consumer SC terkait. |
| Rejeki Seluler / Seluller | Farhan | `Rejeki Seluler` | Extractor, Salvage, Outstanding. |
| CV Berkah Athallah / CV Berkah | Farhan | `CV Berkah` atau tab `CV Berkah Athallah` sesuai project | Extractor, Salvage, Outstanding. |
| GSI | Meilani | `GSI` | Extractor, Salvage. |
| Andalas, Unicom, Xiaomi Authorized, Samsung Exclusive, Carlcare | Meilani | Nama canonical | Extractor, SC-Meilani, Salvage. |
| Samsung Authorized by Unicom variants | Meilani | `Samsung Exclusive` untuk variant/override yang dikontrak | Extractor, SC-Meilani, Salvage. |
| Klikcare, J-Bros, Makmur Era Abadi, Manado Mitra Bersama, Kayu Awet Sejahtera, MDP, B-Store, Multikom, GH Store | Meindar | Nama canonical | Extractor, Salvage, Outstanding. |
| PT Deltasindo / Deltasindo | Meindar | `Deltafone` | Extractor, Salvage, Outstanding. |
| EzCare / Ez Care | Default root: Meindar | `Device Brand` yang mengandung Apple selalu diarahkan Farhan tanpa date gate; non-Apple tetap Meindar. | Root routing, Extractor, Salvage. |
| Tidak match | Tidak ada owner | `SC - Unmapped` / `Unmapped` | Masing-masing project wajib fail-closed. |

Source dan consumer utama: `OPS_ROUTING_POLICY.SC_NAME_KEYWORDS`, `BRANCH_KEYWORDS`, `05b` SC filter/override, PIC enrichment di `06b/06c`, serta mapping lokal di empat standalone scripts. Karena standalone tidak mengimpor root constants, mapping bersama harus dilindungi contract tests.

### Optional-Sheet Routing

| Sheet | Eligibility | Writer/consumer contract |
| --- | --- | --- |
| `B2B` | MAIN: `id_business_partner_category_name = B2B Partnership`; closed/expired exclusions berlaku. | `processB2B_`; tidak memakai partner-pattern/claim-token fallback. SUB tidak rebuild/append, hanya update claim existing. |
| `EV-Bike` | Claim token `VVMAR`, plus Submission overlay; configured policy-number exclusions berlaku. | `processEVBike_`; upsert by Claim Number, manual `Status` protected, deprecated Start/End/Details removed. |
| `Doss` | Claim token `DOSS`. | Memakai EV-Bike writer shape; manual `Status` protected. |
| `Special Case` | MAIN-only flags: Flex, `month_policy_aging > 12`, first-month policy, atau policy remaining under 30 days. | Fixed schema, upsert, all flagged claims retained; `Reason`, Start/End/Details remain active. |

## Data-Contract Registry

### Dataset and Identity

| Dataset/sheet | Identity key | Writer | Primary consumers | Null/failure behavior | Privacy class |
| --- | --- | --- | --- | --- | --- |
| `Raw Data` | `claim_number` | MAIN parser/pipeline | Routing, optional writers, enrichment, reports. | Required identity blank is skipped/logged; aliases normalized before lookup. | Restricted operational. |
| `Raw OLD`, `Raw NEW` | `claim_number` | SUB ingest | Submission append, updates, relocation. | Missing required attachment/header aborts flow; input remains retryable. | Restricted operational. |
| Operational sheets | `Claim Number` | `05b`, SUB relocation, `06b/06c`. | Users, Report Base, Extractor/Outstanding. | Writers only set columns present; unknown route goes quarantine. | Restricted operational. |
| `B2B` | `Claim Number` | `processB2B_` | Operations/reporting. | No replacement candidates must not erase existing dataset. | Restricted operational. |
| `EV-Bike`, `Doss` | `Claim Number` | Token writer + SUB refresh. | Operations. | Upsert/dedupe; manual Status not overwritten. | Restricted operational. |
| `Special Case` | `Claim Number` | `processSpecialCase_`. | Operations/highlight context. | Fixed schema; missing optional date does not silently invent a flag. | Restricted, policy/financial. |
| `Daily Report Base` | One current row per claim | `refreshReportBaseFromOperational06_`. | Weekly aggregation/pivots. | Full rewrite; filter removed/synchronized to avoid stale hidden rows. | Internal analytical. |
| `Weekly Report Base` | Snapshot date + aggregation dimensions | `fillWeeklyReportBase`. | Historical reporting. | Same-date run replaces snapshot; other history preserved; required column missing fails. | Internal analytical. |
| Structured logs/details | RunID + sequence/event | `02_LogAndDetails.gs` and standalone loggers. | Operators/debugging. | Start/progress/failure must remain visible; avoid raw sensitive payload. | Restricted operational metadata. |

### Canonical Header and Alias Contract

| Logical value | Canonical source | Accepted compatibility aliases | Destination/type | Null behavior |
| --- | --- | --- | --- | --- |
| Claim identity | `claim_number` | `Claim Number` | `Claim Number`, text | Blank row is not routable. |
| Submission date | `claim_submitted_datetime` | Legacy `claim_submission_date`; SUB/display variants are compatibility-only. | `Submission Date`, valid Date | Invalid/boolean values are rejected; no unrelated-field fallback. |
| Submission month | `claim_submitted_month` | `claim_submission_months`, display `Submission by Month` | First day of month, format `MMM yy` | Derived from valid Submission Date when source month absent. |
| Last status | `claim_last_status_name` | `last_status`, `Last Status` | Text | Blank/unmapped creates visible mapping evidence. |
| Last-status date | `claim_last_updated_datetime` | `last_update_datetime`, legacy last-activity variants | Date/datetime | Invalid value stays blank; reject fallback fails closed. |
| Service Center | `repairer_location_store_name` | `sc_name`, `service_center`, display variants | Text | SC-universe blank/unmatched goes `SC - Unmapped`. |
| Last Status Aging | `days_aging_from_last_activity` | `last_status_aging`, `LSA` | Number | Blank allowed; Reject Claim may fall back to last-update date. |
| Activity Log Aging | `activity_log_aging` | `ALA` | Number | Blank allowed. |
| IMEI/SN | `imei_number` | `device_imei`, IMEI/SN/serial aliases | Plain text | Comma separators removed; preserve leading zero/digits. |
| DB Link | `dashboard_link` | `db_link` variants | Link/text | Derived from Claim Number only when source blank. |
| Partner | `business_partner_name` | SUB `partner_name` | `Partner Name`, text | Blank allowed but mapping/details may log it. |
| Insurance | `insurance_partner_name` | `insurance_partner_code`, `insurance_code` | Normalized short text | Code fallback when name absent. |
| Stage Aging | Sheet-specific raw aging column | Legacy display `Aging Position`, `Aging Post.` | Number; not on Submission | Same status bucket reuses target-source value; bucket change/missing reference resets to `0`. |
| TAT | `days_aging_from_submission` | Derived for Submission/EV-Bike when required | Submission decimal-day; others numeric | Invalid source blank; Exclusion computes last-status minus submission and clamps at zero. |

Reconciled contract: current runtime treats `claim_submitted_datetime` as primary source and keeps `claim_submission_date` only as legacy fallback. Older documentation that described `claim_submission_date` as the sole strict source is historical, not current behavior.

### Managed, Manual, and Deprecated Fields

| Field group | Ownership | Preservation/write contract |
| --- | --- | --- |
| `Update Status`, `Timestamp`, `Status`, `Remarks` | Manual/restored where columns exist. | Snapshot before clear; restore fills blank destination by claim, preserving rich text, formula, wrap, format, and validation where supported. |
| `Repair Type` | Managed/derived khusus `Finish`. | Exact normalized Last Status pada repair-status policy menghasilkan `Repair`; Last Status lain yang terisi menghasilkan `Replace`; blank tetap blank. |
| `AWB`, `Timestamp AWB` | Manual/restored, primarily Start contract. | Backed up to Raw and restored after routing; formula retention is required. |
| `_OPS_MAIN_SUB_TEMP` | Hidden handoff state. | Match by Claim Number + Service Center; consumed by SUB only in the 09:00 handoff window. |
| `_OPS_MANUAL_BACKUP` | Hidden fallback state. | Match by PIC + Claim Number when normal snapshot restore misses. |
| `OR` | Template/manual field on relevant layouts. | Do not treat as universal raw field. |
| `Start Date`, `End Date`, `Details` | Layout-dependent. | Active for operational/Special Case context; deprecated and removed from B2B/EV-Bike/Doss. |
| `DB`, operational `Status Type` | Deprecated output columns. | May still exist in internal classification/maps, but MAIN/SUB/FORM operational writers must not create/write them. |
| `Update Status Asso`, `Timestamp Asso`, `Update Status Admin`, `Timestamp Admin` | Deprecated. | Removed/ignored by layout enforcement and writers. |

Dropdown `Status` pada Raw Data dan seluruh target memakai urutan canonical: `Pending Logistic`, `Pending Front`, `DONE`, `Pending Insurance`, `Pending TO`, `Pending Finance`, `Pending PIC`, `Pending Cust`, `Pending Buss. Team`, `Pending SC (TA)`, `Pending SC (Repair)`, `Pending SC (Estimation)`, `Delivering`, `Delivered`, `Waiting Courier`, dan `Re-pickup`. Pembaruan validation tidak memutasi nilai manual lama yang telah dibackup.

Dropdown `Status` bersifat general dan dibangun dari satu owner list `STATUS_DROPDOWN_OPTIONS`; tidak ada sheet atau row seperti `Finish.Status` maupun `Raw Data.Status` yang menjadi template. MAIN menerapkan rule yang sama langsung ke Raw Data dan setiap operational consumer. Restore `Status` membersihkan validation lama sebelum menulis backup, lalu memasang kembali dropdown general dengan allow-invalid agar nilai legacy tetap dipertahankan 1:1; kegagalan restore menyertakan source restore, sheet, range/cell, Claim Number, dan nilai pada structured log.

Template enforcement membersihkan data validation dari kolom operational yang bukan owner validation, lalu hanya mengembalikan validation untuk `Status`, `OR`, dan `Type`. Ini mencegah dropdown row-template bocor ke kolom seperti `Submission Date` atau `AWB`. Raw column reorder tetap opt-in; kegagalan move mencatat sheet, header, source column, destination column, dan row bound.

Restore only fills blank destination manual cells. Existing non-empty destination value wins. Rerun tidak boleh membuat duplicate claim atau menghapus manual state yang valid.

## Configuration Registry

Script Properties are environment overrides. Current code still contains several production identifier fallbacks; treat those as migration debt, not permission to add more IDs.

| Property/group | Required | Default/current behavior | Consumer and validation | Sensitivity |
| --- | --- | --- | --- | --- |
| `MASTER_SPREADSHEET_ID`, `MASTER_RAW_SHEET_NAME` | Operationally required | ID fallback exists; raw sheet `Raw Data`. | Root open/preflight and all flows. Invalid ID fails when workbook opens. | Internal identifier. |
| `RESPONSES_SPREADSHEET_ID`, `RESPONSES_SHEET_NAME` | FORM only | ID fallback; `Form Responses 1`. | FORM response lookup. | Internal identifier. |
| `LOG_SPREADSHEET_ID`, `LOG_SHEET_NAME_MAIN`, `LOG_SHEET_NAME_SUB` | Logging | ID fallback; `Log - Main`, `Log - Sub`. `LOG_SHEET_NAME` legacy. | Log sheet assurance/lazy migration. | Internal identifier. |
| `GMAIL_QUEUE_LABEL_MAIN`, `GMAIL_QUEUE_LABEL_SUB` | Email flows | `QUEUED_MAIN`, `QUEUED_SUB`. | Queue query and cleanup. | Internal metadata. |
| `EMAIL_INGEST_ENABLE`, `SUB_EMAIL_INGEST_ENABLE`, `FORM_INGEST_ENABLE` | No | `true`. | Entrypoint gate. | Public config. |
| `EMAIL_INGEST_FROM`, `EMAIL_INGEST_SUBJECT`, `EMAIL_INGEST_ATTACHMENT_NAME_PREFIX`, `EMAIL_INGEST_SEARCH_QUERY` | MAIN email | Code defaults. | Eligible-email selector; UAT exact matching. | Internal email metadata. |
| `SUB_EMAIL_INGEST_FROM`, `SUB_EMAIL_INGEST_SUBJECT`, `SUB_ATTACH_NEW_CONTAINS`, `SUB_ATTACH_OLD_CONTAINS`, `SUB_EMAIL_INGEST_SEARCH_QUERY` | SUB email | Code defaults. | Pair detection; missing pair aborts. | Internal email metadata. |
| `SUB_RAW_OLD_SHEET_NAME`, `SUB_RAW_NEW_SHEET_NAME` | SUB | `Raw OLD`, `Raw NEW`. | Sheet assurance and SUB readers. | Public config. |
| `FORM_FLOW_FIELD_NAME`, `FORM_FILE_UPLOAD_FIELD_NAME`, `FORM_SUB_OLD_FILE_UPLOAD_FIELD_NAME`, `FORM_SUB_NEW_FILE_UPLOAD_FIELD_NAME` | FORM | `Flow`, primary upload field; split fields blank. | Case-insensitive form request builder. | Internal schema. |
| `EVBIKE_ONLY_FOR_PIC`, `B2B_ONLY_FOR_PIC` | No | `*`. | Optional policy compatibility; current single-master flow processes all. | Public config. |
| `SC_ENABLE_*`, `SC_MIN_SUBMISSION_YEAR`, `SC_SKIP_EXCLUDED_LAST_STATUSES` | No | Policy defaults. | Special Case flag gates. Validate with mapping tests/UAT. | Business-sensitive. |
| `EXCLUDED_LAST_STATUSES_CSV` | No | Empty = code set. | Overrides whole exclusion set; risky and requires full contract test. | Business-sensitive. |
| `CLAIM_HIGHLIGHT_MODE` | No | `FOLLOW_NOTE`. | Highlight behavior. | Public config. |
| `DETAILS_UNMAPPED_PARTNER_MIN_SUBMISSION_DATE` | No | `2025-06-01`. | Suppresses legacy unmapped-partner noise before cutoff. | Business-sensitive. |
| `STRICT_SCHEMA_VALIDATION`, `ENABLE_RUN_METRICS`, `USE_TASK_QUEUE`, `ENSURE_ACTIVITY_LOG_COLUMN`, `ENABLE_TXN_IDEMPOTENCY` | No | `false`, `true`, `false`, `false`, `true`. | Feature gates; experimental queue should not be enabled casually. | Public config. |
| `RAW_DATA_REORDER_ENABLE` | No | `false`. | Raw layout finalizer. | Public config. |
| `MAPPING_SPREADSHEET_ID`, `MAPPING_SHEET_NAME` | Legacy | Mapping feature disabled. | Compatibility only; do not revive without design review. | Internal identifier. |
| `MAIN_PIPELINE_STAGE2`, `WORKFLOW_SUB_PENDING_AFTER_MAIN`, `WEEKLY_REPORT_BASE_LAST_RUN_DATE` | Runtime-managed | No static default. | Continuation, pending SUB, weekly once/day state. Never configure manually during normal operation. | Runtime state. |

Do not commit token, password, API key, private key, raw customer data, email subject/attachment metadata from production, or new hardcoded Google resource IDs. Prefer required Script Properties plus fail-fast preflight in the security-hardening follow-up.

## Logging, Continuation, Recovery, and Cleanup

Root logging uses separate `Log - Main` and `Log - Sub`. Canonical rows include Flow, Run ID, stage, duration, counts, status, error, metrics, notes, dan severity. Progress area exposes percentage, current step, updated time, dan source metadata. Runtime evidence belongs in structured logs—not in README.

Recovery checklist:

1. Cari RunID dan stage terakhir di log flow yang benar.
2. Jika MAIN berhenti setelah stage 1, periksa `MAIN_PIPELINE_STAGE2` dan duplicate one-shot trigger sebelum menjalankan ulang.
3. Jika SUB melewati run karena lock, periksa `WORKFLOW_SUB_PENDING_AFTER_MAIN`; MAIN seharusnya mendrain marker sekali.
4. Jika manual data hilang, periksa `RESTORE_AUDIT`, `_OPS_MAIN_SUB_TEMP`, `_OPS_MANUAL_BACKUP`, perubahan Claim Number/Service Center, dan apakah destination sudah berisi nilai non-empty.
5. Jika row salah destination, periksa last status, reject window, SC keyword, token exclusion, lalu routing/type/position maps.
6. Jika report stale, refresh Daily Report Base terlebih dulu; gunakan `runWeeklyReportBaseManual(snapshotDateOverride, sourceFileNameOverride)` hanya setelah daily source benar.

Cleanup invariants:

- destructive cleanup hanya setelah target processing terverifikasi sukses;
- failure mempertahankan unread/queue state agar retryable;
- rerun harus idempotent atau replace-by-key/date;
- active filter range diselaraskan ke used range sebelum write/sort;
- temp, continuation token, dan hidden backup hanya dihapus oleh lifecycle owner-nya.

## Security and Privacy Classification

| Class | Examples | Repository/log policy |
| --- | --- | --- |
| Secret | Password, OAuth token, API key, private key. | Dilarang di Git/log. Gunakan secure property/store. |
| Internal identifier | Spreadsheet/file IDs, internal URL, label, workbook/sheet topology. | Jangan tambahkan identifier produksi baru; Script Properties preferred. Jangan cetak jika tidak diperlukan. |
| Restricted operational data | Claim/customer identity, IMEI/SN, policy number, raw email subject, attachment metadata. | Tidak boleh masuk fixture atau log mentah. Gunakan synthetic/redacted sample dan counts. |
| Business-sensitive | Routing/PIC, financial values, policy-age flags, exclusions. | Boleh didokumentasikan sebagai contract; perubahan wajib review/test. |
| Public technical metadata | Function names, generic schema, validation commands. | Aman di docs. |

Current OAuth/runtime scopes belum diaudit ulang di perubahan dokumentasi ini. Manifest, deployment, dan Apps Script permissions harus diverifikasi terpisah sebelum hardening dinyatakan selesai.

## Change-Impact Guide

| Perubahan | Minimum owner/consumer audit | Required docs/tests |
| --- | --- | --- |
| Status/routing | `00`, `05b`, SUB relocation `06a`, type/position/exclusion, optional, reports. | README mapping + changelog + mapping tests. |
| SC/PIC/branch | Root keywords/override/PIC enrichment dan seluruh standalone consumers relevan. | README SC table + changelog + cross-consumer tests. |
| Header/column | `00` aliases/types, `03` templates, writer, backup/restore, formatting, SUB, reports. | README data contract + changelog + header/manual tests. |
| Trigger/orchestration | Lock, pending marker, continuation token, retry, RunID, cleanup. | README flow/runbook + changelog + root test dan runtime UAT. |
| Optional sheet | `00` policy, `05c` writer, SUB refresh, fixed-schema policy, reports. | README optional/data contract + changelog + mapping tests. |
| Reporting | Operational source list, Daily full rewrite, Weekly gate/key/idempotency/filter. | README runbook + changelog + report UAT. |
| Standalone | Local config, entrypoint, lock, log, cleanup, deployment. | README standalone registry + changelog + syntax/mapping tests. |

Safe editing sequence: temukan source of truth, cari seluruh direct consumer, tentukan expected contract, tambah/ubah regression test, edit minimal owner, jalankan check, lalu sinkronkan current state dan changelog.

## Validation and Governance

CI berjalan pada pull request dan push ke `main` dengan permissions read-only dan concurrency cancellation. Required checks yang ditargetkan:

- `code-contracts`: root/standalone syntax dan mapping regression;
- `docs-governance`: knowledge files, Markdown/link/anchor/Mermaid, dan documentation drift;
- `sensitive-diff`: scan added lines untuk secret, identifier, dan raw operational metadata.

Documentation drift rules:

- perubahan root `.gs`, `optional-project/**`, manifest, tests, tooling, atau CI wajib mengubah `CHANGELOG.md`;
- config, routing, schema, flow, atau standalone contract wajib mengevaluasi `README.md`;
- ownership, invariant, validation workflow, atau agent contract wajib mengevaluasi `AGENTS.md`;
- README/AGENTS escape hatch hanya melalui PR body dengan alasan spesifik:

```text
Docs-Impact-README: none — <alasan spesifik>
Docs-Impact-AGENTS: none — <alasan spesifik>
```

Alasan kosong atau generik gagal. Business-code change tidak dapat melewati kewajiban changelog.

## Apps Script UAT

Jalankan dengan data synthetic/terkontrol pada staging copy bila memungkinkan.

1. MAIN: valid email, invalid attachment, same-input rerun, continuation stage 1/2, output counts, cleanup boundary.
2. Manual restore: isi formula/value pada enam manual fields yang tersedia, jalankan MAIN dan SUB relocation, lalu verifikasi formula, rich text, dropdown, wrap, dan timestamp format.
3. SUB: OLD+NEW valid, satu file hilang, same/change status bucket, `Stage Aging`, pending lock, exit dari Expired/Exclusion, dan optional existing-row update.
4. Reject: reject dengan aging `<= 30`, reject dengan aging `> 30`, dan blank aging dengan recent/stale last-update date.
5. SC: GSI, Rejeki Seluler/Seluller, CV Berkah, Deltasindo, Samsung Unicom variants, EzCare Apple/non-Apple, dan unknown fallback.
6. Optional: B2B category gate, VVMAR overlay/dedupe, DOSS token, Special Case four flags termasuk exact `month_policy_aging = 12` yang tidak boleh ter-flag.
7. Data: IMEI leading zero/plain text, valid/invalid Submission Date, month-date format, filters yang hanya mencakup sebagian used range.
8. Reports: Daily uniqueness/Position/PIC; Weekly same-date replace, previous/change helpers, pure SUB 09:00 once/day, FORM SUB immediate, dan manual refresh.
9. Standalone: lock contention, no-data, duplicate key, unmapped route, retry/queue, post-write verification, serta cleanup success/failure per project.
10. Observability: start/progress/failure/success terlihat dengan RunID yang sama dan tanpa raw restricted data.

Record runtime result, sample IDs yang sudah direduksi/redacted, rollback action, dan gap yang tersisa di Issue/PR—not in this handbook.

## Known Limitations and Roadmap

1. Security/config hardening: pindahkan production IDs ke required Script Properties, tambah fail-fast preflight, audit OAuth scopes, sanitasi error, dan tetapkan retention log 90 hari atau 5.000 rows sesuai kebutuhan project.
2. Deployment reproducibility: inventaris root + empat standalone deployments, trigger, property, manifest, environment, staging/production, dan rollback; evaluasi `clasp` tanpa mengekspos script ID.
3. Observability: stable error codes, retryable flag, last-success, queue/continuation freshness; kumpulkan baseline 14 hari sebelum menetapkan SLO.
4. Refactor setelah coverage: pecah orchestration `06a`, reporting/maintenance/restore `06c`, kelompokkan `00_Config.gs`, konsolidasikan standalone mapping, dan audit nama script tanpa ekstensi.
5. Current technical debt: beberapa hardcoded Google resource IDs masih ada; standalone Outstanding masih memuat legacy `SC - Ivan`/PIC naming dan mapping yang perlu direkonsiliasi pada revisi behavior terpisah.
6. Gmail, Drive, Spreadsheet, trigger, quota, permission, dan deployment belum tervalidasi end-to-end oleh local checks.

## Appendix

<details>
<summary>Expanded raw and destination field registry</summary>

| Domain | Important fields | Contract |
| --- | --- | --- |
| Raw identity/policy | `claim_number`, `qoala_policy_number`, `source_system_name`, `business_partner_name`, `id_business_partner_category_name`. | Identity, duplicate/source classification, B2B, partner/flag routing. |
| Raw dates | `claim_submitted_datetime`, `claim_submitted_month`, `claim_last_updated_datetime`, `last_update_datetime`, `policy_start_datetime`, `policy_end_datetime`. | Valid Date parsing; ambiguous/invalid input fails closed. |
| Raw activity/aging | `days_aging_from_submission`, `days_aging_from_last_activity`, `activity_log_aging`, `last_activity_log_name`, `last_activity_log_datetime`, sheet-specific `Aging *`. | Numeric aging, operational monitoring, Stage Aging. |
| Device/customer | `device_type`, `device_brand`, `imei_number`, `device_imei`, `holder_name`, `customer_name`, `outlet_name`, `pa_name`, `spa_name`. | Destination fields; customer/device values classified restricted. |
| Finance | `sum_insured_amount`, `claim_amount`, `claim_own_risk_amount`, `nett_claim_amount`. | Optional/layout-dependent; `Selisih` and `% Approval` derived when numeric. |
| Classification | Claim token `SFP/SFX/SMR` = OLD; `VVMAR/GADLD` = NEW; duplicate comparison uses policy/source/claim/submission/status with configured 62-day window. | Internal classification can remain even though operational `DB` column is deprecated. |
| Highlight flags | Migration Policy, expired, Flex, B2B, duplicate, second-year, first-month, remaining-one-month. | MAIN exclusively calculates and applies note/fill after final operational formatting and sorting; priority comes from configured policy, and `month_policy_aging > 12` is strict. SUB only preserves the existing Claim Number note/fill during updates and relocation. |
| SC-specific output | `Type`, `Branch`, `Service Center PIC`. | Derived only where destination header exists. |
| Workflow output | `Claim Type` on Reject Claim; service/claim type on Start/Finish/Expired depending template. | Status/source fallback must be tested per sheet. |

</details>

<details>
<summary>Operational debugging index</summary>

| Symptom | Check first |
| --- | --- |
| Destination field blank | Destination header exists, canonical raw header/alias resolves, value type parses. |
| Claim missing | Claim key, flow eligibility, status route, optional gate, exclusion, and fallback log. |
| Wrong SC owner | Normalized service center, EzCare date/device split, keyword ordering, standalone local mapping. |
| Manual field missing | `RESTORE_AUDIT`, hidden backups, formula/rich snapshot, claim/SC key change, non-empty destination precedence. |
| `TRUE` in Submission Date | Strict sync source/type and old checkbox/data validation. |
| IMEI changed | Plain-text number format and comma normalization before write. |
| B2B empty | Category exact match, excluded status, MAIN vs SUB behavior, replacement row guard. |
| EV-Bike/Doss missing | Claim token, configured exclusion, Raw/Submission overlay, SC-Unmapped exclusion. |
| Special Case missing | Flag inputs and dates; exact 12 months is not second-year; sheet fixed schema. |
| Weekly stale/duplicate | Daily source, snapshot key/date parsing, once/day property, same-date replacement. |

</details>

<details>
<summary>Migration parity checklist</summary>

Every heading from the retired workflow/column references was classified before deletion:

| Legacy section | Classification | Destination |
| --- | --- | --- |
| High-level map, layer dependency map | Retained/merged | System Context and Repository Map. |
| MAIN, SUB, FORM flow and touchpoints | Retained/merged | Flow Registry and Change-Impact Guide. |
| Change impact map, safe editing, governance | Merged | Change-Impact Guide and Validation/Governance. |
| Dated flow updates and hardening notes | Historical | Normalized Changelog outcomes. |
| Refactor priority | Retained | Known Limitations and Roadmap. |
| UAT Part 9, SC-unmapped FAQ, Weekly quick reference | Retained/merged | Apps Script UAT; Mapping; Recovery. |
| Contract ownership and flow map | Retained/merged | Document Ownership, Repository Map, Flow Registry. |
| Current contract updates | Merged after code reconciliation | Mapping and Data-Contract registries. |
| Column source categories, manual restore, aliases | Retained | Data-Contract Registry and debugging appendix. |
| Raw/destination/optional sheet tables | Retained/condensed | Dataset, Header, Managed Fields, and expanded appendix. |
| Derived classifications/highlighting | Retained | Mapping registry and expanded appendix. |
| Old sole-source `claim_submission_date` statements | Stale | Replaced by verified primary `claim_submitted_datetime` + legacy fallback contract. |
| Deprecated Asso/Admin fields listed as active tails | Stale | Classified deprecated; not preserved as active contract. |
| Operational `Status Type` listed as active output | Stale | Classified deprecated output; internal map remains for compatibility/analytics. |
| Special Case done-status pruning statements | Stale | Current writer retains all flagged claims. |
| Debugging questions and change checklist | Retained/merged | Debugging appendix and Change-Impact Guide. |

</details>

## Second-Brain Lifecycle

Capture ide/incident melalui GitHub Issue. Perubahan yang diterima dicatat di `CHANGELOG.md#Unreleased`. Runtime evidence masuk structured logs. Review dilakukan melalui PR dan CI. Current state ditemukan di README. Histori outcome diarsipkan per tanggal dalam changelog. Git history tetap menjadi audit trail teknis; jangan menaruh commit SHA sebagai freshness marker yang self-referential.
