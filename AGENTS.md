# Workflow Repository — Agent Control Plane

## Scope and mission

Instruksi ini berlaku untuk seluruh repository `Harshwell/workflow`: root Google Apps Script `00_*.gs`–`06c_*.gs`, `optional-project/*`, manifest, validation tooling, CI, dan tiga knowledge files repository.

Repository mengelola automation claim workflow untuk ingestion Gmail/Form/Drive, parsing, MAIN/SUB/FORM orchestration, routing ke operational sheets, backup/restore field manual, enrichment, reporting, maintenance, dan automation standalone.

Urutan prioritas perubahan:

1. integritas dan correctness data;
2. idempotency dan rerun safety;
3. konsistensi seluruh consumer root dan standalone;
4. efisiensi Apps Script/Spreadsheet API;
5. observability, recovery, dan error traceability;
6. maintainability tanpa overengineering.

Instruksi task aktif mengalahkan file ini, tetapi tidak memperluas authority untuk deployment, credential, branch, commit, push, atau PR tanpa permintaan eksplisit.

## Start workflow

Sebelum mengedit:

1. Periksa `git status`, branch, dan perubahan pengguna yang harus dipertahankan.
2. Identifikasi flow, policy owner, data contract, dan consumer yang terdampak.
3. Baca bagian [README](README.md) yang relevan dan `CHANGELOG.md#Unreleased`.
4. Inspeksi implementasi aktual dan direct consumers; jangan bergantung pada dokumentasi saja.
5. Tentukan validation command dan runtime UAT gap sebelum membuat perubahan.

Gunakan pencarian terarah berdasarkan function, constant, status, header, sheet, atau mapping. Jangan membaca seluruh repository untuk perubahan kecil.

## Source and conflict model

Pisahkan dua hal berikut:

- **Expected contract**: instruksi task, approved business rule, invariant di file ini, current-state contract di README, dan regression tests.
- **Observed implementation**: behavior kode yang sedang ada.

Observed implementation bukan otomatis benar. Jika kode, tests, dan dokumentasi berbeda:

1. hentikan asumsi pada area konflik;
2. tentukan expected contract dari task dan evidence yang dapat diverifikasi;
3. jelaskan konflik bila berdampak pada behavior;
4. selaraskan code, tests, README, dan changelog dalam scope yang sama.

Executable policy/config dan regression tests adalah source of truth teknis. README menjelaskan current state dan provenance-nya; jangan membuat salinan business mapping baru tanpa owner constant/function yang jelas.

## Architecture ownership

| Domain | Owner utama |
| --- | --- |
| Policy, routing, schema constants, flags, runtime configuration | `00_Config.gs` |
| Generic safe I/O, normalization, coercion, retry, Gmail/Drive helpers | `01_Utils.gs` |
| Log lifecycle, progress, details, error reporting | `02_LogAndDetails.gs` |
| Sheet assurance, schema, validation, formatting | `03_SheetsAndValidation.gs` |
| Parsing, dates, aging | `04_ParseAndAging.gs` |
| Raw mutation dan manual backup/restore preparation | `05a_Pipeline_RawMutate_Backup.gs` |
| Operational routing, SC selection, write behavior | `05b_Pipeline_RoutingOperational.gs` |
| B2B, EV-Bike, Doss, Special Case, optional sheets | `05c_Pipeline_OptionalSheets.gs` |
| Triggers, queues, MAIN/SUB/FORM orchestration | `06a_EntryPoints.gs` |
| Pipeline dan enrichment | `06b_PipelineAndEnrichment.gs` |
| Reporting, post-process, maintenance, assertions | `06c_PostProcessAndUtils.gs` |
| Standalone deployments | `optional-project/*` |
| Current-state handbook dan detailed registries | `README.md` |

Standalone projects tidak menerima root globals secara otomatis. Perubahan policy bersama wajib mengaudit seluruh standalone consumer yang relevan.

## Durable runtime invariants

- MAIN dapat memakai continuation dua tahap; RunID, progress, retryability, snapshot, dan cleanup boundary harus tetap konsisten.
- SUB membutuhkan OLD/NEW input valid, meng-update claim existing, dan merelokasi row sesuai routing. Jika MAIN memegang lock, gunakan pending/handoff yang sudah ada.
- FORM/MANUAL harus memakai shared MAIN/SUB core, bukan fork business logic baru.
- Rerun tidak boleh membuat duplicate claim atau menghapus field manual valid.
- IMEI/SN harus dipertahankan sebagai text.
- Internal routing bucket tidak boleh dibuat sebagai physical sheet.
- Cleanup destructive hanya boleh berjalan setelah success boundary terverifikasi.
- Required configuration/header harus fail visibly; jangan membuat silent fallback baru.
- Setiap flow harus mencatat start, step/progress, result, duration, dan error context yang actionable tanpa membocorkan data sensitif.

Detail current mapping, field ownership, data contract, configuration, dan recovery procedure berada di README dan source code yang ditautkan di sana.

## Change-impact matrix

### Status, routing, atau Service Center

Audit minimal:

- policy owner di `00_Config.gs`;
- routing index, fan-out, relocation, exclusions, status type, dan position;
- MAIN, SUB, FORM, optional sheets, Report Base;
- seluruh root/standalone consumer mapping terkait;
- mapping contract tests, README registry, dan changelog.

### Header, column, atau schema

Audit minimal:

- canonical header dan aliases;
- parser/resolver, template enforcement, writer, formatting/validation;
- backup/restore dan manual-field ownership;
- optional sheets, relocation, reporting;
- data-contract tests dan README registry.

### Trigger, orchestration, atau cleanup

Audit minimal:

- lock, pending/handoff, continuation property/token;
- duplicate trigger prevention dan cooperative stop;
- RunID/progress continuity, retry, success boundary;
- Gmail label/read/trash dan temporary-file lifecycle;
- runtime quota risk serta UAT.

### Logging, configuration, OAuth, atau privacy

Audit minimal:

- required/optional property semantics dan fail-fast behavior;
- log allowlist, redaction, retention, and error sanitization;
- OAuth scope impact dan reauthorization;
- public/internal/restricted/secret classification;
- README registry, security diff gate, dan runtime UAT.

## Engineering rules

- Pertahankan Google Apps Script V8 compatibility dan public trigger names kecuali migration plan disetujui.
- Reuse helper hanya bila contract benar-benar sama; business rule harus punya owner jelas.
- Prefer batch read/write, bounded range, index, dan map daripada API call per row atau full-sheet scan.
- Output harus deterministic, idempotent, dan aman untuk retry.
- Legacy compatibility harus mempunyai alasan dan removal condition bila sementara.
- Jangan melakukan refactor lintas modul yang tidak dibutuhkan task.
- Jangan menambahkan secret, production identifier, internal URL/email, atau raw sensitive metadata ke repository/log.
- Environment-specific configuration memakai Script Properties dan explicit validation.
- Jangan mengubah deployment, trigger production, credential, branch policy, atau remote state tanpa scope eksplisit.

## Second-brain lifecycle

- **Capture**: request, defect, atau gap durable dicatat sebagai GitHub Issue; runtime evidence tetap di structured runtime log.
- **Process**: impact matrix menentukan owner, consumer, tests, documentation, dan UAT.
- **Record**: accepted material change masuk `CHANGELOG.md#Unreleased`.
- **Review**: PR dan CI memeriksa code contract, documentation governance, sensitive diff, dan final diff.
- **Retrieve**: README adalah front door current state; file ini adalah agent control plane; changelog adalah temporal memory.
- **Archive**: saat release, pindahkan Unreleased ke section tanggal; jangan menyimpan diary atau audit selesai di README/AGENTS.

Repository hanya memiliki tiga tracked Markdown knowledge files:

- `AGENTS.md`: durable agent contract;
- `README.md`: current-state handbook;
- `CHANGELOG.md`: temporal history.

Jangan membuat `REPO_CONTEXT.md`, ADR Markdown, atau doc tambahan. Rationale besar disimpan di GitHub Issue/PR; runtime evidence tidak disalin ke repository docs.

## Validation

Gunakan Node 22.

```bash
npm ci
npm run check
```

Gunakan command lebih sempit hanya saat iterasi lokal; sebelum handoff, jalankan suite penuh. `npm run check` mencakup root smoke test, standalone syntax termasuk profile SC branch di `optional-project/SC-*.js`, mapping contracts, documentation validation, dan governance self-tests.

Static validation tidak membuktikan Gmail, Drive, Spreadsheet, trigger, quota, deployment, atau Apps Script runtime sebenarnya. Jika runtime UAT tidak dilakukan, nyatakan batasannya dan jangan mengklaim end-to-end success.

Documentation-impact enforcement:

- material code/tooling/CI change memerlukan changelog;
- config/routing/schema/flow contract change memerlukan README atau alasan PR yang spesifik;
- ownership/invariant/validation workflow change memerlukan AGENTS atau alasan PR yang spesifik.

## Completion checklist

Sebelum menyelesaikan task:

- request dan root cause tertangani tanpa scope creep;
- unrelated user changes aman;
- direct consumers, failure path, retry, idempotency, dan cleanup dipertimbangkan;
- source-of-truth tidak diduplikasi;
- tests relevan dan `npm run check` dijalankan;
- runtime UAT gap dan risk dicatat;
- AGENTS, README, dan CHANGELOG direview sesuai ownership;
- sensitive diff dan final `git diff` diperiksa;
- tidak ada deployment, credential, commit, push, PR, branch, atau remote mutation tanpa authority.

Final response harus menyebut hasil, file/behavior berubah, validation beserta hasil, documentation sync, dan runtime risks/UAT yang tersisa.
