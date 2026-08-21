# CRF Specializations → P21 Form Template Converter

Converts a CDISC COSMOS **CRF Specializations** export into a **Pinnacle 21 (P21) Form Template** workbook (Forms / Sections / Questions / Codelists / Terms / Units / FormSpec, plus the other template sheets) and a companion **Value Level Metadata (VLM)** workbook.

Two implementations are provided, kept in sync:

| | |
|---|---|
| [`java/`](java/) | Dependency-free Java port. No Apache POI, no Maven - reads/writes `.xlsx` using only the JDK's zip and XML support. |
| [`crfdss_to_form_clean.sas`](crfdss_to_form_clean.sas) | Cleaned, commented, macro-wrapped rewrite of the original SAS proof of concept. |
| [`crfdss_to_form.sas`](crfdss_to_form.sas) | The original draft, kept for reference. Has the bugs described below. |

## Why this exists

The original SAS draft worked, but had two problems that blocked a real P21 import:

1. **Duplicate codelists.** Codelist ID/name were built from `crf_group_id` + `crf_item`, so a codelist got recreated once per *question* instead of once per distinct set of terms. The standard Yes/No ("NY") codelist alone came out as ~310 near-identical copies.
2. **Section IDs.** P21 requires Section IDs to be unique. `crf_group_id` is documented as unique in the COSMOS CRF specialization model itself, so it's used directly as the Section ID (just sanitized for P21-safe characters) - this one turned out not to need a design change, just confirmation that the model already guarantees it.

The codelist problem is fixed here at the design level, not patched around:

- **Codelists are now keyed by content**: (submission value, or variable name as a fallback) + the sorted, case-normalized set of terms. Identical codelists collapse into one row no matter how many questions reference them; genuinely different subsets (e.g. a restricted "Sex Male/Female only" scenario vs. the full Sex codelist) still get their own entry, correctly.
- **Codelist IDs are readable, not hashed, whenever that's safe.** Content-based identity (above) is essential for codelists genuinely *shared* across many CRF groups (the standard Yes/No "NY" list is reused by ~140 of them) - there's no single "owning" group to name those after, so they keep a content-derived ID (hashed only if the term list is too long to use directly, e.g. the 30+ term CMTRT drug list). But a codelist used by only *one* `crf_group_id` doesn't need that: it now gets a readable `<variable>_<crf_group_id>` ID instead (variable/codelist first, then the qualifying suffix, matching P21's usual convention) - e.g. `QSORRES_EQ5D0201`, `FTORRES_ADCDRL` - rather than an opaque hash (`QSORRES_163C2CA3`) or a meaningless direct key (`FTORRES_0_1_10_2_3_4_5_6_7_8_9` for an 11-point scale). Verified against the full source file: 257 distinct codelists (matches a real SAS run), zero ID collisions.
- **Codelist names no longer borrow the wrong CRF group's short name.** This showed up two ways. Codelists built from `prepopulated_term` (e.g. `LBCAT` defaulting to `CHEMISTRY`) are a single default value, not a real picklist, and have nothing to do with whichever CRF group's `short_name` happens to be attached to them - naming one "Subset for Albumin/Creatinine in Urine (Denormalized)" instead of "Subset for Chemistry" affects 69 such shared default-value codelists in the source (lab categories, inclusion/exclusion categories, body locations, substance-use categories, and more), not just one; these are now named after the actual default value instead. Separately, every other codelist shared across multiple CRF groups was named after whichever group's `short_name` happened to survive deduplication - arbitrary, and outright wrong for the 5 of 19 shared codelist "bases" (`UNIT`, `LOC`, `POSITION`, `VSRESU`, `IEORRES`) that each cover more than one genuinely different term set, so naming them all after their base alone would just trade one ambiguous name for another. These are now named after their base plus a short preview of the actual terms - e.g. "Subset for NY: No, Yes", "Subset for LOC: Brachial Artery, Femoral Artery, Peripheral Artery, Radial Artery" - which is both meaningful and guaranteed distinct by construction.

Three follow-up refinements to the codelist ID/naming fix above: a shared codelist whose term list is too long to use directly now tries combining its (few) owning `crf_group_id`s into the id before falling back to a hash - e.g. `ACN_AE_DENORMALIZED_AE_NORMALIZED` for a list shared by exactly 2 groups - which resolves 5 of 7 real long-shared codelists in the source; the remaining 2 (one shared by 15 groups, one by 6) still need the hash, since no reasonably short combination of that many group names would help. The term preview used for shared-codelist names only read `value_display_list`/`value_list`, so a codelist whose term came from `prepopulated_term` with neither of those populated (e.g. Prothrombin Time's bare "s"/"ms" unit defaults) got a blank-looking name; it now falls through to `prepopulated_term` too. And the Terms sheet's `Order` column no longer matches physical row order after deduplication - fixed by re-sorting back to `(codelist, order)` and renumbering contiguously.

Other bugs fixed along the way (see the header comments in each file for details): a SUBSTR call that crashed on short `crf_item` values like `SEX`/`AGE`, two column names that didn't exist in the source (`mandatory_value`, `sdtm_variable`) so the VLM export came out blank in those spots, the Terms sheet landing in a different workbook than everything else, a numeric→text export bug where missing "Digits" values showed up as a literal `.` instead of a blank cell, (SAS only) a variable-name collision that silently truncated the Questions sheet's `Codelist` column to 7 characters - the source has its own short column literally named `codelist` (the raw NCI id, e.g. `C66731`), and SAS variable names are case-insensitive, so the derived `Codelist` value inherited that column's short length unless given an explicit one - and a Questions `SDTM Target` bug where multi-variable targets are delimited with `;` in the source (never `,`), so the domain-prefix logic never actually split them; every variable is now individually prefixed (`<domain>.<variable>`) and re-joined with `,`, which is the delimiter P21 expects.

A `QC_Checks` row also flags (without altering) `AESTDAT`/`AEENDAT` in the AE domain, where the source's own `sdtm_target_variable` is self-mapped to the CRF item instead of the expected `*DTC` variable - a source data quality gap (276 of 280 "date" rows map correctly, e.g. `CMSTDAT`→`CMSTDTC`; only these two don't), not something this tool silently corrects.

## Naming, in plain terms

The source file's vocabulary doesn't match P21's out of the box. This tool maps:

- **Form = `domain`** (e.g. `AE`, `DM`). Each distinct CDASH domain in the source becomes one Form.
- **Section = `crf_group_id`**, used as-is (just sanitized for P21-safe characters). A `crf_group_id` groups related CRF items together (e.g. all the fields for one scenario), which is what P21 calls a Section, and it's unique by design in the COSMOS CRF specialization model.

Everything downstream (Questions' `Form`/`Section` columns, the Sections sheet's `Form`/`ID`) follows from those two mappings.

## Template alignment

Sheet names, columns, and column order match Ben's `Form-Spec-Template` workbook:

`FormSpec → Events → Forms → Sections → Questions → Units → Codelists → Methods → Conditions → Terms → ScheduleOfActivities`

`Events`, `Methods`, `Conditions`, and `ScheduleOfActivities` have no corresponding data in a CRF Specializations export, so they're written as header-only sheets (correct columns, zero rows) rather than skipped - the workbook still imports as a complete template.

The template's `Codelists` sheet includes a **Recommended Codelist** column, populated from the source's CDISC/NCI controlled terminology reference (the `codelist` column, e.g. `C66731`) - traceability the original draft discarded.

Three fields from the source aren't part of the standard P21 template but are kept as **trailing extension columns** on the Questions sheet (after `Developer Notes`), since P21 tolerates extra columns and these may be useful downstream: `Short Name`, `SDTM Annotation`, `Prepopulated Term`.

A `QC_Checks` sheet is also appended (data type / digits mismatches flagged during processing) - a bonus sheet, not part of the template.

## Usage

### Java

Requires a JDK (11+). No other dependencies.

```bash
cd java
javac *.java
java CrfSpecToP21 "/path/to/cdisc_crf_specializations_draft.xlsx" /path/to/output/dir
```

Produces `crf_form.xlsx` and `crf_vlm.xlsx` in the output directory, plus a console summary (form/section/codelist/term counts).

### SAS

```sas
%include "crfdss_to_form_clean.sas";   /* defines the macro */

%crfdss_to_form(
  fpath   = C:\Users\bmant\Downloads,
  dssfile = cdisc_crf_specializations_draft
  /* optional: outfile_form=my_form.xlsx, outfile_vlm=my_vlm.xlsx */
);
```

`outfile_form=` and `outfile_vlm=` default to `crf_form.xlsx` / `crf_vlm.xlsx` and only need to be set if you want different output names.

## What still needs manual completion

The source data doesn't cover everything a P21 template can express. Left blank deliberately (not an oversight) for manual completion:

- Forms: `Description`, `Class`, `Repeating`, `Condition`, `Developer Notes`
- Sections: `Condition`, `Developer Notes`
- Questions: `Description`, `Core`, `Method`, `Condition`, `Implementation Notes`, `Mapping Instructions`, `Reason Not Mapped`, `Developer Notes`
- The `Events`, `Methods`, `Conditions`, and `ScheduleOfActivities` sheets entirely

## Verification note

Neither implementation could be compiled/executed end-to-end in the environment they were built in (no JDK, no SAS). Both were validated by: a line-by-line manual audit, an equivalent Python port of the algorithm run against the real 2073-row source file (confirming codelist counts collapse as expected and all section IDs are unique), and round-tripping the Java writer's exact OOXML templates through openpyxl and LibreOffice. Worth a real compile/run before relying on either in production.
