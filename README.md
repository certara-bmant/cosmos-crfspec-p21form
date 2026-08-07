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

Other bugs fixed along the way (see the header comments in each file for details): a SUBSTR call that crashed on short `crf_item` values like `SEX`/`AGE`, two column names that didn't exist in the source (`mandatory_value`, `sdtm_variable`) so the VLM export came out blank in those spots, the Terms sheet landing in a different workbook than everything else, a numeric→text export bug where missing "Digits" values showed up as a literal `.` instead of a blank cell, and (SAS only) a variable-name collision that silently truncated the Questions sheet's `Codelist` column to 7 characters - the source has its own short column literally named `codelist` (the raw NCI id, e.g. `C66731`), and SAS variable names are case-insensitive, so the derived `Codelist` value inherited that column's short length unless given an explicit one.

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
