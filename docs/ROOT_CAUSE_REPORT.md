# PDGM investigation — 2026-09-19

Investigation completed before editing application code. References below refer to
the original files preserved in `.investigation/baseline/`. No production workbook
was supplied. The directory is not a Git repository; baseline copies enable diff
review and code rollback. Synthetic experiments used a TemporaryDirectory only.

## Data flow and evidence

`app.yukle` (1054–1104) saves each upload under a unique filename, then calls
`excelden_aktar` under `_import_kilidi`. COM opens that specific file read-only,
without external-link updates, copies Value2 cells into a temporary workbook,
and skips hidden columns. `_excelden_aktar` reads the snapshot with openpyxl,
parses all sheets, then commits once via `depo.excel_import_uygula` under `_kilit`.
The store writes kartlar/log/uploads through temp files, backups and os.replace;
exceptions restore memory, and ordinary replace failures roll back disk files.
Startup/reload reads kartlar.xlsx and validates it. `kartlari_getir` filters and
builds independent view-model dictionaries, `_pano_verisi` counts distinct stock
numbers, and Flask renders panel/operator/monitor. Non-static responses already
have Cache-Control: no-store. No evidence of an upload filename/cache bug.

| Symptom | Root cause and original reference | Fix strategy |
|---|---|---|
| Duplicate on date update | excel_araclari.py:426–464 derives later keys from mutable dates; first row gets bare Talep/Stok | Immutable row ID scoped to source sheet |
| Wrong row receives data | excel_araclari.py:619 shares key groups across sheets; depo.py:2098–2103 matches only that key | Persist source_sheet/source_row_id/source_key; reject ambiguous migration |
| Deleted rows stay operational | depo.py:2251 onward sets source_active=0, but :356–361 ignores it | Common visibility and operation guard |
| Delivered remains planned | excel_araclari.py:558–560 discards valid status without actual delivery date; depo.py:2110 onward preserves old workflow on None | Valid status wins; missing date remains unknown, with warning |
| Actual delivery date stays stale | depo.py:2131 onward preserves it unless delivered status is accepted; delivered branch uses `new or old or today` | Explicit source overwrite including clearing, without fabricated date |
| EUM start missing | Header map :99–124 lacks T.planlanan tarih; parser only reads explicit start :526–538 | Parse planned date, derive Monday once |
| Monitor contains other types | app.py:743–750 supplies all types | Filter backend dataset by MAKINE |
| Monitor too fast/incomplete cycle | monitor.html:240 uses 5000ms, :259 reloads at 50000ms and resets to page zero | 12000ms constant; cycle-aligned reload and saved position |
| Misleading counts/feedback | _pano_verisi uses a set of stok_no; workflow_korundu increments for every match | Clarify count labels and count only preserved workflows |

Experiments on original code:
- Two duplicate rows, change second delivery date: IDs 1,2 became 1,2,3;
  ID 2 had source_active=0 but still appeared in kartlari_getir().
- Reorder two rows: ID 1 changed from NO=1 to NO=2; third card appeared.
- PLANA ALINDI -> TESLİM EDİLDİ with blank actual date: stayed PLANA ALINDI.
- Remove MAKİNE row from 5 MAKİNE + 2 ELDE (delivered): ELDE took ID 1
  formerly belonging to MAKİNE, while old ELDE ID 2 stayed visible.
- Reload persisted the same incorrect state.
- EÜM T.planlanan tarih=17.09.2026 returned plan_baslama=None.

## Original merge ownership

| Fields | Existing-card behavior |
|---|---|
| sira, talep_sahibi, toplam_adet, adet_metin, plan_hafta, plan_baslama, plan_teslim, excel_durum, pcb, dizgi_tipi, dizgi_sorumlusu, malzeme_bekliyor | Always overwritten by plan, if key matches |
| id, anahtar, talep_no, stok_no | Preserved on match |
| durum | Accepted ilk_durum overrides; None preserves manual workflow |
| gerceklesen_teslim | Mostly preserved; accepted delivered uses new/old/today fallback; transition back clears |
| completed/start quantities and workflow timestamps | Conditional status transition; counts preserved on unknown status, delivered completed=total |
| operator, aciklama | Preserved; blank operator sometimes becomes Excel |
| admin_gizli | Preserved |
| source_active, aktif, kaynak, guncelleme | Reset to 1,1,EXCEL,current time on match |
| absent Excel rows | Only source_active becomes 0; physical history kept |

## Design / implementation plan

User confirmed NO is stable and unique within each sheet. This is a data-model
change within the existing Flask/XLSX architecture. Identity will be a canonical
JSON tuple [sheet code, immutable NO], stored as source_key and used as anahtar.
Dates, quantity, order, Talep/Stok and upload filename are not identity inputs.
An optional explicit PDGM_ROW_ID supports sheets without NO; otherwise a missing
or duplicate ID rejects the entire upload. Sıra alone is not proof of identity.
Missing IDs cannot be invented on every read: that would not survive reuploads.

Legacy migration preserves IDs/notes/operator history by sheet type + exact NO
(and Talep/Stok consistency); a unique legacy group without NO can migrate only
when both old and incoming group have one member. Ambiguous duplicate groups
abort before mutation; dates/amounts are never tie-breakers. Persist old anahtar
as legacy_anahtar. Existing source keys never migrate to a different key.

- [x] Add unittest regressions with real temporary XLSX writes and COM boundary fake;
      observe failures before changing product code.
- [x] Implement parser identity, EÜM Monday normalization, status semantics.
- [x] Add persisted identity columns, conservative migration, Excel field ownership,
      inactive filtering and per-card sync audit under existing locks/rollback.
- [x] Backend monitor filtering, local badge text, 12-second rotation, safe refresh,
      distinct-stock labels; validate Flask rendered output and JS timer behavior.
- [x] Test reload, legacy files, ambiguous identity rejection, failure rollback,
      sheet separation, hidden identity column and date edges; review complete diff.

Review focus: duplicate/missing source IDs; legacy ambiguity; existing manual cards;
Excel quantities conflicting with manual partial completion; failures after first
file replacement; date1904 workbooks; empty but structurally valid source sheets.

Decisions: valid Excel status is authoritative; unknown/blank status preserves
manual workflow. Unknown actual delivery dates remain None rather than today's
date. Source fields overwrite even when blank. Manual notes/operator and hidden
flag survive; quantities remain protected against invalid partial-completion
reductions. A full, valid empty workbook deactivates all Excel cards. Missing
optional sheets mean those source rows are absent, as in the existing full-snapshot
contract. No automatic repair of already misassigned historical workflow is claimed.


## Reanalysis — 2026-09-20

A second independent review found two additional paths: unchanged admin status
with only a note edit invented unknown event dates (`admin_kart_duzenle`), and
stale operator pages could append notes to inactive source cards (`kart_not_guncelle`).
Both were reproduced by failing regressions and corrected. Date defaults now apply
only to actual admin status transitions. Operators cannot append to inactive cards;
admins retain historical audit-note access. Actual operator start/finish/delivery
still records event timestamps, covered by a separate workflow test.

Final validation: 39 automatic Python tests pass; 1 opt-in real COM test passes
separately; Node slider behavior passes; Python syntax checks pass for 9 files.
No production data directory was created. See docs/EXCEL_SYNC.md for remaining
format/migration and multi-file crash-recovery limitations.
