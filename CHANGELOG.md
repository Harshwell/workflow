# Changelog

Repo ini memakai changelog ringan yang fokus ke perubahan yang benar-benar relevan buat maintainer.
Formatnya sengaja sederhana: **Added / Changed / Fixed**. Tidak perlu sok formal kalau ujungnya tidak pernah dibaca.

---

## 2026-03-27

### Added
- `README.md` sebagai entrypoint dokumentasi repository.
- `docs/WORKFLOW_MAP.md` untuk flow map, dependency map, dan change-impact guide.
- `98_PipelineCoreHardening` untuk hardening targeted di pipeline core.
- `99_RuntimeFixes` untuk runtime guard dan compatibility override yang sifatnya surgical.
- `97_SelfCheck` untuk smoke check non-destructive terhadap health, config, symbol, dan coverage mapping.

### Changed
- `00_Config` dirapikan dengan section index dan alias config resmi agar akses source-of-truth lebih konsisten.
- Layer dokumentasi dibuat tetap ramping agar cocok untuk self-project tapi tetap enak dirawat.
- Runtime hardening diarahkan ke area yang paling rawan silent failure: header resolution, datetime parsing, status-type resolution, dan filtered sort.

### Fixed
- Header validation dibuat lebih toleran terhadap variasi casing, whitespace, dan alias header.
- Guard details-log dirapikan agar tidak salah baca policy source-of-truth.
- Resolver compatibility lama ditutup supaya caller lawas tidak gampang meledak di runtime.
- Filter-preserving sort di flow operasional diperbaiki agar memakai index relatif terhadap filter range, bukan asumsi kolom absolut.
- Parsing datetime diperketat untuk menekan drift timezone dan native parse yang ambigu.
- Status Type mapping diarahkan ke source-of-truth utama agar tidak fallback diam-diam ke mapping lama.

### Notes
- File hardening tambahan (`98_` dan `99_`) sengaja dibuat sebagai patch layer yang sempit agar risk perubahan tetap rendah.
- Kalau fix yang sama nanti sudah aman dipindah ke source module asli, patch layer ini bisa diperkecil atau dihapus bertahap.
