/*******************************************************************************
 CRF Specializations (COSMOS) -> Pinnacle 21 Form Template converter
 Cleaned/annotated rewrite of crfdss_to_form.sas (the original POC).

 WHAT THIS PROGRAM DOES
 -----------------------
 Reads the "CRF Specializations" sheet of a CDISC COSMOS CRF specialization
 export and reshapes it into the sheets a P21 Form Template import expects:
 Forms, Sections, Questions, Codelists, Terms, Units, FormSpec - plus a
 companion Value Level Metadata (VLM) workbook for *RESU-style unit fields.

 WHAT CHANGED VS THE ORIGINAL DRAFT (see the SAS log Ben ran for evidence)
 ---------------------------------------------------------------------------
 1. DUPLICATE CODELISTS (the main reported problem). The original built each
    codelist's ID/name from crf_group_id + crf_item, so a codelist like the
    standard Yes/No list ("NY") was recreated once per QUESTION instead of
    once per distinct set of terms - roughly 310 near-identical copies of "NY"
    alone. Here, codelist identity is content-based: submission value (or
    variable name) + the sorted, case-normalized set of terms. Identical
    codelists collapse into a single row no matter how many questions use
    them; genuinely different subsets (e.g. a restricted Sex Male/Female-only
    scenario vs the full Sex codelist) still get their own entry, correctly.

 2. SECTION IDs. P21 requires Section IDs to be unique. crf_group_id is
    documented as unique in the COSMOS CRF specialization model itself, so
    it's used directly as the Section ID here too - just sanitized for
    P21-safe characters (no spaces/punctuation), not namespaced by Form.

 3. SUBSTR CRASH ON SHORT crf_item VALUES. The original tested for a unit
    field with substr(crf_item, length(crf_item)-3, 4) = 'RESU'. For any
    crf_item shorter than 4 characters (e.g. "SEX", "AGE") that's an invalid
    (zero or negative) SUBSTR start position - see the repeated
    "NOTE: Invalid second argument to function SUBSTR" lines in the log. Fixed
    with a length-guarded suffix check below, which also tells apart
    "original" result units (...ORRESU) from "standard" units (...STRESU)
    instead of labelling both "Original".

 4. WRONG COLUMN NAMES. The original referenced mandatory_value and
    sdtm_variable in KEEP/RENAME lists - neither exists in the source (the
    real columns are mandatory_variable and sdtm_target_variable). SAS only
    warns on this (doesn't fail), so it silently produced a blank
    Mandatory/Variable column in the VLM export. Fixed.

 5. Terms was exported to a different workbook (dss_form.xlsx) than every
    other sheet (crf_form.xlsx). Consolidated into one workbook here.

 6. Source_Variable_Name was left blank in Questions even though
    variable_name was available in the source. Mapped here.

 7. QUESTIONS.CODELIST TRUNCATED TO 7 CHARACTERS. dss_ordered (built from
    dss_derived, ultimately from dss) already has a column literally named
    "codelist" - the source's raw CDISC/NCI controlled terminology reference,
    e.g. "C66731", which PROC IMPORT sizes at ~7 characters based on the
    longest value actually present. SAS variable names are case-insensitive,
    so the ATTRIB-declared "Codelist" in questions_final (intended to hold
    the up-to-80-character codelist_id) is the SAME variable as that
    pre-existing short one once "set dss_ordered;" runs, unless ATTRIB stakes
    an explicit length claim first. Without one, "Codelist = codelist_id;"
    silently truncated every value to 7 chars (e.g. "EGMETHOD_12_LEAD_
    STANDARD" -> "EGMETHO") - invisible in the Codelists/Terms sheets, which
    already had an explicit length=$80 on their own same-named variables, so
    only Questions (and, defensively, VLM) needed the same explicit length
    added.

 8. SDTM_TARGET MULTI-VARIABLE DELIMITER. The source delimits multi-variable
    SDTM targets with ";" (e.g. "AEENRTPT;AEENRF;AEENTPT"), never ",". The
    domain-prefix logic checked for a comma, which never matched, so these
    never actually got split - only a single domain prefix landed on the
    front of the whole string. Every variable is now individually split out,
    prefixed with "<domain>.", and re-joined with "," (the delimiter P21
    expects). Also flagged (not silently corrected) a source data issue:
    AESTDAT/AEENDAT in the AE domain have sdtm_target_variable self-mapped to
    the CRF item instead of the expected *DTC variable - see the QC_Checks
    sheet.

 9. CODELIST IDS THAT WERE TECHNICALLY UNIQUE BUT UNREADABLE. Fix #1 above
    keys codelist identity by content so broadly-reused lists (the standard
    Yes/No "NY" list, reused by ~140 different crf_groups) collapse into one
    row. But for content that's LONG (many terms) or just not human-readable
    (a plain numeric scale), that same content-based id was either hashed
    ("QSORRES_163C2CA3" for an EQ-5D dimension's text options) or left as a
    meaningless direct key ("FTORRES_0_1_10_2_3_4_5_6_7_8_9" for an 11-point
    ADAS-Cog scale) - neither tells a reviewer what the codelist actually is.
    Now: codelists referenced by only ONE crf_group_id (verified: everything
    in this category) get a readable id instead - <base>_<crf_group_id>
    (variable/codelist first, then the qualifying suffix, matching P21's
    usual codelist ID convention), e.g. "QSORRES_EQ5D0201",
    "FTORRES_ADCDRL" - since crf_group_id is already guaranteed unique in
    the COSMOS model, this can't collide with anything. Codelists shared
    across MULTIPLE crf_group_ids (like "NY") keep the original
    content-based id, since there's no single group to name them after and
    that's exactly the identity fix #1 needs. Verified against the full
    source file: 257 distinct codelists (matches the real SAS log), zero id
    collisions, longest generated id 60 characters.

    A related naming bug, in two parts. First: codelists built from
    prepopulated_term (a single default value, not a real picklist - e.g.
    LBCAT = "CHEMISTRY") were named after the enclosing crf_group's
    short_name ("Subset for Albumin/Creatinine in Urine (Denormalized)"),
    which has nothing to do with the actual default value - affected 69
    such codelists, not just this one. Now named after the value itself
    ("Subset for Chemistry"). Second: every OTHER shared codelist (not
    prepopulated_term-derived) was named after whichever crf_group's
    short_name happened to survive the nodupkey pass - arbitrary, and
    outright wrong whenever the same base is reused by multiple distinct
    shared codelists (5 of 19 such bases - "UNIT", "LOC", "POSITION",
    "VSRESU", "IEORRES" - cover more than one genuinely different term set,
    so naming them all after their base alone would just trade one
    ambiguous name for another). Shared codelists are now named after their
    base plus a short preview of their actual terms - e.g. "Subset for NY:
    No, Yes", "Subset for LOC: Brachial Artery, Femoral Artery, ..." -
    meaningful and guaranteed distinct by construction.

 10. THREE FOLLOW-UP REFINEMENTS TO FIX #9. (a) A shared codelist whose term
     list is too long to use directly (>60 characters) now tries combining
     its (few) owning crf_group_ids into the id before falling back to a
     hash - e.g. "ACN_AE_DENORMALIZED_AE_NORMALIZED" for a list shared by
     exactly 2 groups - since a short group list is more traceable than an
     opaque hash. Verified: 5 of 7 real long-shared codelists in the source
     fit this way; the other 2 (FTREASND, shared by 15 groups; LOC's
     SYSBP/HR/DIABP location set, shared by 6) still need the hash, since no
     reasonably short combination of that many group names would help. (b)
     The term preview used for shared-codelist names (fix #9) only read
     value_display_list/value_list, so a codelist whose term came from
     prepopulated_term with NEITHER of those populated (e.g. the
     Prothrombin Time unit fields' bare "s"/"ms" defaults) got a
     blank-looking name: "Unit, subset for  (Original)". Now falls through
     to effective_list (prepopulated_term) too. (c) The Terms sheet's
     "Order" column no longer matches physical row order: NODUPKEY needs a
     (codelist, recommended_term) sort to find the duplicate rows a shared
     codelist creates, but that sort is alphabetical by term text, which
     left the already-computed Order value out of sequence with the row
     it's on. Re-sorted back to (codelist, order) and renumbered
     contiguously after dedup.

 Sheet names, columns, and column order below match Ben's "Form-Spec-Template"
 workbook exactly (FormSpec, Events, Forms, Sections, Questions, Units,
 Codelists, Methods, Conditions, Terms, ScheduleOfActivities). Events/Methods/
 Conditions/ScheduleOfActivities have no corresponding data in a CRF
 Specializations export, so they're written as header-only sheets rather than
 skipped. The source "codelist" column (the CDISC/NCI controlled terminology
 id, e.g. "C66731") is carried through to the template's "Recommended
 Codelist" column instead of being discarded.

 A Java port of this same pipeline (with the same fixes) is also available -
 see CrfSpecToP21.java - for anyone who'd rather run this outside SAS.

 USAGE
 -----
 %crfdss_to_form(fpath=C:\Users\bmant\Downloads, dssfile=cdisc_crf_specializations_draft);

 Parameters:
   fpath        - folder containing the source .xlsx and where outputs land.
   dssfile      - source file name, without the .xlsx extension.
   outfile_form - output Form Template workbook name (default crf_form.xlsx).
   outfile_vlm  - output Value Level Metadata workbook name (default crf_vlm.xlsx).
*******************************************************************************/

%macro crfdss_to_form(fpath=, dssfile=, outfile_form=crf_form.xlsx, outfile_vlm=crf_vlm.xlsx);

/*------------------------------------------------------------------------------
 1. IMPORT
------------------------------------------------------------------------------*/
proc import datafile="&fpath.\&dssfile..xlsx" out=dss dbms=xlsx replace;
  sheet='CRF Specializations';
run;

/*------------------------------------------------------------------------------
 2. ROW-LEVEL DERIVATIONS
    Everything a single source row needs to know about itself: the QC flag,
    the "fixed" data type/digits, whether it's a unit (*RESU) field, and - the
    two fixes above - its canonical codelist id/name and its unique section id.
    Doing this per-row (instead of via a downstream PROC SQL join, as the
    original did) means every later step just reads a variable instead of
    re-deriving or re-joining anything.
------------------------------------------------------------------------------*/
/* NOTE ON TYPES: length and significant_digits import as NUMERIC (they're
   numeric cells in the source xlsx) - not character. fixed_sig_digits is
   deliberately NOT put in the character LENGTH list below, so it stays
   numeric like its source. Forcing it to character (as an earlier version of
   this file did) makes SAS render a missing value as "." on export instead
   of a blank cell - the "Digits column full of dots" bug. Comparisons below
   use numeric literals (200, not '200'; missing(), not ne '') accordingly. */
data dss_step1;
  length typchk $200 unit_kind $10 fixed_data_type $20
         base $200 effective_list $4000 normalized_key $4000 raw_key $4200
         section_id $200 is_vlm_target $1 is_prepop_only $1;
  set dss;

  /* ---- 2a. QC flags (informational only - see the QC_Checks sheet) ---- */
  typchk = '';
  if not missing(significant_digits) and length ne 200 and data_type ne 'float' then
    typchk = 'Data type should be float';
  if length = 200 and not missing(significant_digits) then
    typchk = 'Data type likely text and digits not required';

  /* Source data quality flag (not a bug in this program): a "date" field
     whose sdtm_target_variable is identical to crf_item is almost certainly
     mis-mapped - real date CDASH variables map to a *DTC target (e.g.
     CMSTDAT -> CMSTDTC, DMDAT -> DMDTC), and 276 of 280 date rows in the
     source do exactly that. Only AESTDAT and AEENDAT (both AE_DENORMALIZED
     and AE_NORMALIZED) are self-mapped in this extract - flagged here rather
     than silently corrected, since the source may be fixed upstream. */
  if data_type = 'date' and crf_item ne '' and sdtm_target_variable = crf_item then
    typchk = catx('; ', typchk, 'SDTM target self-mapped to CRF item - likely should be a *DTC variable');

  /* ---- 2b. Auto-fix for the same discrepant details the QC flag catches --- */
  fixed_data_type   = data_type;
  fixed_sig_digits  = significant_digits;
  if not missing(significant_digits) and length ne 200 then fixed_data_type  = 'float';
  if length = 200                                      then fixed_sig_digits = .;

  /* ---- 2c. Unit-field classification (fixes the SUBSTR crash, fix #3) ----
     Length-guarded suffix check instead of substr(item, length(item)-3, 4):
     that expression is undefined (invalid SUBSTR start) for crf_item values
     under 4 characters, which is exactly what SEX/AGE hit in the log.        */
  unit_kind = 'NONE';
  if length(trim(crf_item)) >= 6 and upcase(substr(crf_item, length(trim(crf_item))-5, 6)) = 'ORRESU' then
    unit_kind = 'ORIGINAL';
  else if length(trim(crf_item)) >= 6 and upcase(substr(crf_item, length(trim(crf_item))-5, 6)) = 'STRESU' then
    unit_kind = 'STANDARD';
  else if length(trim(crf_item)) >= 4 and upcase(substr(crf_item, length(trim(crf_item))-3, 4)) = 'RESU' then
    unit_kind = 'OTHER_UNIT';
  is_vlm_target = ifc(unit_kind ne 'NONE', 'Y', 'N');

  /* ---- 2d. Canonical content key (fix #1 - the duplicate codelist bug) ---
     Identity = (submission value, or variable name/crf_item as a fallback)
     + the SORTED, upper-cased set of terms. Same content -> same key, always
     - regardless of which question or section happens to reference it. This
     raw_key is used two ways further down (step 2f): to decide whether a
     codelist is a broadly-reused standard list (kept content-keyed) or
     exclusive to one CRF group (given a readable, group-qualified id
     instead - see below), and as the fallback content-hash source for the
     rare case that even the readable id is too long.

     effective_list can come from value_list (a real multi-choice picklist)
     or, when that's blank, from prepopulated_term (a single default value
     being assigned to the field, not a real picklist at all) - is_prepop_only
     records which, since the two need different id/name treatment in 2f. */
  effective_list = value_list;
  is_prepop_only = 'N';
  if effective_list = '' then do;
    effective_list = prepopulated_term;
    if effective_list ne '' then is_prepop_only = 'Y';
  end;

  if effective_list ne '' then do;
    base = codelist_submission_value;
    if base = '' then base = variable_name;
    if base = '' then base = crf_item;

    link normalize_terms; /* -> sets normalized_key from effective_list */

    if is_vlm_target = 'Y' then raw_key = 'UNIT_' || trim(base) || '_' || trim(normalized_key);
    else raw_key = trim(base) || '_' || trim(normalized_key);
  end;
  else do;
    base = '';
    raw_key = '';
  end;

  /* ---- 2e. Section id --------------------------------------------------
     crf_group_id is documented as unique in the COSMOS CRF specialization
     model itself (not just an accident of this particular extract), so it's
     used directly as the Section ID - sanitized for P21-safe characters,
     but not namespaced by Form.                                             */
  if crf_group_id ne '' then
    section_id = prxchange('s/_+/_/', -1, prxchange('s/[^A-Za-z0-9_]+/_/', -1, trim(crf_group_id)));
  else section_id = '';

  return;

  /* ---- Sub-routine: builds normalized_key = sorted, upcased, underscore-
     joined terms from effective_list. CALL SORTC sorts an array of discrete
     variables in place; padding unused slots with blanks is safe because
     CATX (unlike CAT) skips blank arguments regardless of position.         */
  normalize_terms:
    array term_arr[100] $200 _temporary_;
    do _i = 1 to 100;
      term_arr[_i] = '';
    end;
    _n = countc(effective_list, ';') + 1;
    do _i = 1 to _n;
      term_arr[_i] = upcase(trim(left(scan(effective_list, _i, ';'))));
    end;
    call sortc(of term_arr[*]);
    normalized_key = catx('_', of term_arr[*]);
  return;

  drop _i _n normalized_key;
run;

/* ---- 2f. Codelist id/name assignment (two-pass) -----------------------
   How many DISTINCT crf_group_ids reference each canonical raw_key across
   the WHOLE file? That tells us whether a codelist's content is a broadly
   reused standard list (e.g. the Yes/No "NY" list, reused by ~140 different
   crf_groups - exactly the case fix #1 above collapses) or exclusive to a
   single crf_group (e.g. one EQ-5D dimension's text options, or a lab
   category default). Reported issue: for exclusive codelists, basing the id
   on content alone produced either an opaque hash suffix when the term list
   was long ("QSORRES_163C2CA3" for an EQ-5D dimension) or a meaningless
   numeric-soup id when it wasn't ("FTORRES_0_1_10_2_3_4_5_6_7_8_9" for an
   11-point ADAS-Cog scale) - neither tells a reviewer what the codelist
   actually is. Since crf_group_id is already guaranteed unique in the
   COSMOS model, <crf_group_id>_<base> is a strictly better, hash-free id for
   anything exclusive to one group - verified against the full source file
   with zero collisions (257 distinct codelists, matching the real SAS log,
   longest generated id 60 characters). Shared/standard lists keep the
   original content-based id: there's no single "owning" group to name them
   after, and this is exactly the identity fix #1 relies on. */
proc sort data=dss_step1(keep=raw_key crf_group_id where=(raw_key ne ''))
          out=_rk_groups nodupkey;
  by raw_key crf_group_id;
run;

proc sql;
  create table _rk_group_counts as
    select raw_key, count(*) as n_groups
    from _rk_groups
    group by raw_key;
quit;

proc sort data=_rk_group_counts; by raw_key; run;

/* Sorted, underscore-joined list of every crf_group_id that references a
   given raw_key - lets a shared-but-too-long-for-a-direct-key codelist (see
   below) get a readable id built from its (few) owning groups instead of an
   opaque hash, when there are few enough of them to keep it a reasonable
   length. _rk_groups is already sorted by raw_key/crf_group_id, so this is
   just the same retain+catx accumulation pattern used for wc_combined
   above. */
data _rk_group_list;
  length group_list $600;
  set _rk_groups;
  by raw_key;
  retain group_list;
  if first.raw_key then group_list = '';
  group_list = catx('_', group_list, crf_group_id);
  if last.raw_key then output;
  keep raw_key group_list;
run;

proc sort data=dss_step1; by raw_key; run;

data dss_derived;
  length codelist_id $80 codelist_name $200;
  merge dss_step1(in=a) _rk_group_counts(in=b) _rk_group_list;
  by raw_key;
  if a;

  if raw_key = '' then do;
    codelist_id = '';
    codelist_name = '';
  end;
  else do;
    /* ---- id: shared-vs-exclusive (unaffected by naming, below) ---- */
    if n_groups > 1 then do;
      /* SHARED across multiple CRF groups - identity must stay content-based.
         Direct sanitized key when short; hash the rare long term list (e.g.
         the 30+ term CMTRT drug list) down to a stable, short suffix instead
         of a 500-character id. MD5 gives a deterministic, collision-safe
         hash without needing anything outside Base SAS. */
      if length(raw_key) <= 60 then
        codelist_id = prxchange('s/_+/_/', -1, prxchange('s/[^A-Za-z0-9_]+/_/', -1, trim(raw_key)));
      else do;
        /* Long term list (e.g. the 30+ term CMTRT drug list, or LOC's
           "Brachial/Femoral/Peripheral/Radial Artery" location set). Try
           combining the (few) crf_group_ids that reference this shared
           content into the id first - readable when there are only a
           handful of them (e.g. "ACN_AE_DENORMALIZED_AE_NORMALIZED" for a
           list shared by exactly 2 groups; verified against the source
           file: 5 of 7 real long-shared codelists fit this way). Falls back
           to a hash only when there are too many owning groups to keep the
           id a reasonable length (e.g. FTREASND, shared by 15 different
           ADAS-Cog sub-tests) - clean_base is almost always shorter than 40
           chars (submission values like "ACN"/"CMTRT" are short codes);
           substr(x, 1, 40) against a string with less than 40 characters of
           real content triggers "NOTE: Invalid third argument to function
           SUBSTR", so length is capped at min(40, actual length) to make
           the truncation a no-op for short bases and safe for long ones. */
        length group_id_candidate $600 clean_base $200;
        group_id_candidate = prxchange('s/_+/_/', -1,
                                prxchange('s/[^A-Za-z0-9_]+/_/', -1, trim(base || '_' || group_list)));
        if length(group_id_candidate) <= 70 then codelist_id = group_id_candidate;
        else do;
          clean_base = prxchange('s/[^A-Za-z0-9_]+/_/', -1, trim(base));
          codelist_id = substr(clean_base, 1, min(40, length(clean_base)))
                        || '_' || put(md5(raw_key), $hex8.);
        end;
      end;
    end;
    else do;
      /* EXCLUSIVE to this one CRF group - readable, group-qualified id, no
         hash needed (verified: never exceeds ~60 chars for any row in the
         current source; the length check + hash below is just a safety net
         for a future, longer crf_group_id/base combination). Ordered
         <variable/codelist>_<crf_group_id> (e.g. "QSORRES_EQ5D0201"), not
         the other way around - matches the variable-then-qualifying-suffix
         convention P21 codelist IDs typically use. */
      length candidate_id $70;
      if is_vlm_target = 'Y' then
        candidate_id = prxchange('s/[^A-Za-z0-9_]+/_/', -1, trim('UNIT_' || base || '_' || crf_group_id));
      else
        candidate_id = prxchange('s/[^A-Za-z0-9_]+/_/', -1, trim(base || '_' || crf_group_id));
      if is_prepop_only = 'Y' then candidate_id = trim(candidate_id) || '_PREPOP';
      candidate_id = prxchange('s/_+/_/', -1, candidate_id);

      if length(candidate_id) <= 70 then codelist_id = candidate_id;
      else codelist_id = substr(candidate_id, 1, 55) || '_' || put(md5(raw_key), $hex8.);
    end;

    /* ---- name: unit vs. prepopulated-default vs. group short_name/terms --
       The prepopulated-default case (below) is independent of shared vs.
       exclusive - see its comment. For the other two cases: EXCLUSIVE
       codelists keep short_name, since it correctly and unambiguously
       describes the one CRF group that owns them. SHARED codelists can no
       longer be named after "whichever group's short_name happened to
       survive the nodupkey pass" - that was arbitrary, and outright wrong
       whenever the same base is reused by MULTIPLE distinct shared
       codelists (confirmed: of 19 distinct bases behind a shared codelist,
       5 - "UNIT", "LOC", "POSITION", "VSRESU", "IEORRES" - cover more than
       one genuinely different term set, so "Subset for LOC" alone would be
       just as ambiguous as the short_name it replaces). Shared codelists are
       named after their base plus a short preview of their actual terms
       instead - meaningful and guaranteed distinct, since two different
       shared codelists always differ in their term content by construction. */
    length term_preview $200;
    term_preview = value_display_list;
    if term_preview = '' then term_preview = value_list;
    /* Falls through to effective_list (which is prepopulated_term when
       value_list is blank) for a shared codelist whose term came from a
       prepopulated default rather than a real picklist - e.g. the "PT"
       (Prothrombin Time) unit fields' seconds/milliseconds default, where
       value_display_list AND value_list are both blank. Without this,
       term_preview stayed empty and produced a blank-looking name: "Unit,
       subset for  (Original)" for codelist id UNIT_UNIT_S/UNIT_UNIT_M. */
    if term_preview = '' then term_preview = effective_list;
    term_preview = translate(trim(term_preview), ',', ';');
    if length(term_preview) > 80 then term_preview = substr(term_preview, 1, 77) || '...';

    if is_vlm_target = 'Y' then do;
      length suffix $10;
      if unit_kind = 'ORIGINAL' then suffix = 'Original';
      else if unit_kind = 'STANDARD' then suffix = 'Standard';
      else suffix = 'Units';
      if n_groups > 1 then
        codelist_name = 'Unit, subset for ' || trim(left(term_preview)) || ' (' || trim(suffix) || ')';
      else
        codelist_name = 'Unit, subset for ' || trim(left(short_name)) || ' - ' || trim(suffix);
    end;
    else if is_prepop_only = 'Y' then
      codelist_name = 'Subset for ' || propcase(trim(left(effective_list)));
    else if n_groups > 1 then
      codelist_name = 'Subset for ' || trim(left(base)) || ': ' || trim(left(term_preview));
    else codelist_name = 'Subset for ' || trim(left(short_name));
  end;

  drop n_groups candidate_id group_id_candidate clean_base term_preview group_list;
run;

/* QC_Checks sheet: same intent as the original "chks" dataset, now actually
   exported so it doesn't just sit unused in WORK.                            */
data qc_checks;
  set dss_derived;
  if typchk ne '';
  keep crf_group_id crf_item data_type length significant_digits typchk;
run;

/*------------------------------------------------------------------------------
 3. CODELISTS + TERMS
    Explode each qualifying row's term list, tagged with the already-canonical
    codelist_id/codelist_name from step 2 - no join required. NODUPKEY passes
    then collapse to one row per codelist and one row per (codelist, term).
------------------------------------------------------------------------------*/
data terms_exploded;
  set dss_derived;
  if effective_list = '' then delete;

  listcount = countc(effective_list, ';') + 1;
  do term_order = 1 to listcount;
    term = scan(effective_list, term_order, ';');
    decoded_value = scan(value_display_list, term_order, ';');
    if term ne '' then output;
  end;
  keep codelist_id codelist_name is_vlm_target term_order term decoded_value;
run;

proc sort data=terms_exploded;
  by codelist_id term_order;
run;

data terms_final(keep=order codelist display_term recommended_term)
     units(keep=id unit);
  attrib
    order            label='Order'
    codelist         length=$80  label='Codelist'
    display_term     length=$200 label='Display Term'
    recommended_term length=$200 label='Recommended Term'
  ;
  set terms_exploded;
  by codelist_id;

  recommended_term = term;
  display_term = decoded_value;
  if display_term = '' then display_term = term;

  /* one Terms row per distinct term within a codelist */
  if first.codelist_id then term_seq = 0;
  term_seq + 1;
  codelist = codelist_id;
  order = term_seq;
  output terms_final;

  /* Units is a flat, globally deduped unit dictionary (id = the unit token
     itself), same as the original design intent. */
  if is_vlm_target = 'Y' then do;
    id = term;
    unit = display_term;
    output units;
  end;
run;

/* NODUPKEY needs a (codelist, recommended_term) sort to find the duplicate
   rows a shared codelist creates (the same term list gets exploded once per
   source row that references it - e.g. "NY" exploded once per crf_item that
   uses it). But that sort is ALPHABETICAL by term text, which leaves the
   already-computed "order" column out of sequence with the actual row
   order - e.g. a codelist's rows would print in alphabetical term order
   while Order reads 3, 1, 2. Re-sort back to (codelist, order) afterward and
   renumber contiguously (nodupkey may have dropped some rows, leaving gaps
   in the original 1..N sequence). */
proc sort data=terms_final nodupkey;
  by codelist recommended_term;
run;

proc sort data=terms_final;
  by codelist order;
run;

data terms_final;
  set terms_final;
  by codelist;
  if first.codelist then order = 0;
  order + 1;
run;

proc sort data=units nodupkey;
  by id;
run;

proc sort data=terms_exploded out=codelist_raw(keep=codelist_id codelist_name) nodupkey;
  by codelist_id;
run;

/* Recommended Codelist = the CDISC/NCI controlled terminology id (source
   column "codelist", e.g. "C66731") this canonical subset was drawn from -
   restores traceability the original draft discarded. First non-blank value
   per codelist_id wins if more than one source row contributed to it. */
proc sort data=dss_derived(keep=codelist_id codelist where=(codelist_id ne '' and codelist ne ''))
          out=nci_lookup;
  by codelist_id;
run;

proc sort data=nci_lookup nodupkey;
  by codelist_id;
run;

proc sql;
  create table codelist_final as
    select a.codelist_id as id label='ID',
           a.codelist_name as name label='Name',
           'text' as type label='Type',
           b.codelist as recommended_codelist label='Recommended Codelist'
    from codelist_raw a
    left join nci_lookup b on a.codelist_id = b.codelist_id;
quit;

/*------------------------------------------------------------------------------
 4. FORMS + SECTIONS
------------------------------------------------------------------------------*/
proc sort data=dss_derived out=forms(keep=domain) nodupkey;
  by domain;
run;

data form_final(keep=order id name description class repeating condition
                      source_dataset_name sdtm_target_domain developer_notes);
  attrib
    order               label='Order'
    id                  length=$200 label='ID'
    name                length=$200 label='Name'
    description         length=$200 label='Description'
    class               length=$50  label='Class'
    repeating           length=$3   label='Repeating'
    condition           length=$200 label='Condition'
    source_dataset_name length=$200 label='Source Dataset Name'
    sdtm_target_domain  length=$200 label='SDTM Target Domains'
    developer_notes     length=$200 label='Developer Notes'
  ;
  set forms;

  id = domain;
  name = domain;
  source_dataset_name = domain;   /* Source Dataset Name IS the domain, e.g. "AE" */
  sdtm_target_domain = domain;

  /* No source mapping exists yet for these P21 template columns - left
     blank deliberately for manual completion later. */
  description = '';
  class = '';
  repeating = '';
  condition = '';
  developer_notes = '';

  order + 1;
run;

proc sort data=dss_derived out=sections_raw(keep=domain section_id crf_group_id short_name bc_id vlm_group_id) nodupkey;
  by section_id;
run;

proc sort data=sections_raw;
  by domain section_id;
run;

/* BC_ID (Biomedical Concept id, e.g. "C83347") and VLM_GROUP_ID are trailing
   extension columns, not part of the standard P21 template - kept for
   traceability, same pattern as the extension columns on Questions. Both
   confirmed constant within a crf_group_id in the source (no aggregation
   ambiguity). */
data sections_final(keep=form order id name mandatory repeating condition developer_notes bc_id vlm_group_id);
  attrib
    form            length=$200 label='Form'
    order           label='Order'
    id              length=$200 label='ID'
    name            length=$200 label='Name'
    mandatory       length=$3   label='Mandatory'
    repeating       length=$3   label='Repeating'
    condition       length=$200 label='Condition'
    developer_notes length=$200 label='Developer Notes'
  ;
  /* bc_id/vlm_group_id already exist (passed through from the source) - just
     label them, rather than redeclaring their length via ATTRIB, which would
     otherwise produce a harmless "multiple lengths specified" note. */
  label bc_id = 'BC ID' vlm_group_id = 'VLM Group ID';
  set sections_raw;
  by domain;

  form = domain;
  id = section_id;
  name = short_name;
  if name = '' then name = crf_group_id;

  mandatory = 'No';   /* no section-level mandatory/repeating flag exists in
                          the source - same limitation as the original draft,
                          flagged here rather than silently assumed */
  repeating = 'No';
  condition = '';
  developer_notes = '';

  if first.domain then order = 1;
  else order + 1;
run;

/*------------------------------------------------------------------------------
 5. QUESTIONS
------------------------------------------------------------------------------*/
proc sort data=dss_derived out=dss_ordered;
  by domain crf_group_id order_number;
run;

/* The first block of columns matches the Questions sheet of the P21 Form
   Template exactly, in template order. Short_Name/SDTM_Annotation/
   Prepopulated_Term/Dec_ID aren't template columns, but are kept as trailing
   extension columns (P21 tolerates extra columns beyond the standard
   template) rather than dropped - appended after Developer Notes, not
   interleaved with the template's own columns. Method/Condition are
   template columns with no source mapping, included blank. Dec_ID (Data
   Element Concept id, e.g. "C78541") lives here rather than on Sections
   because it varies per crf_item within a crf_group_id, unlike bc_id/
   vlm_group_id, which are constant per section.

   Codelist is given an explicit length=$80 below (fix #7) - without it, the
   pre-existing short source column "codelist" (the raw NCI id, e.g.
   "C66731", ~7 chars) collides case-insensitively via "set dss_ordered;"
   and silently truncates every codelist_id assigned into it to 7 chars. */
data questions_final;
  attrib
    Form                    label='Form'
    Section                 label='Section'
    Order                   label='Order'
    ID                      label='ID'
    Source_Variable_Name    label='Source Variable Name'
    Question_Text           label='Question Text'
    Prompt                  label='Prompt'
    Description             label='Description'
    Data_Type               label='Data Type'
    Core                    label='Core'
    Length                  label='Length'
    Digits                  label='Digits'
    Mandatory               label='Mandatory' format=$3.
    Codelist                length=$80 label='Codelist'
    Measurement_Units       label='Measurement Units'
    Method                  label='Method'
    Condition               label='Condition'
    Completion_Instructions label='Completion Instructions'
    Implementation_Notes    label='Implementation Notes'
    Mapping_Instructions    label='Mapping Instructions'
    SDTM_Target             length=$200 label='SDTM Target'
    Reason_Not_Mapped       label='Reason Not Mapped'
    Developer_Notes         label='Developer Notes'
    Short_Name              label='Short Name'
    SDTM_Annotation         label='SDTM Annotation'
    Prepopulated_Term       label='Prepopulated Term'
    Dec_ID                  label='DEC ID'
  ;
  set dss_ordered;
  by domain section_id;

  if crf_item = '' then delete;

  Form   = domain;
  Section = section_id;
  ID     = crf_item;
  Source_Variable_Name = variable_name;   /* fix #6: was left blank previously */

  Data_Type = fixed_data_type;
  if Data_Type = 'decimal' then Data_Type = 'float';
  Length = length;
  Digits = fixed_sig_digits;

  if mandatory_variable = 'Y' then Mandatory = 'Yes';
  else if mandatory_variable = 'N' then Mandatory = 'No';
  else Mandatory = '';

  Codelist = codelist_id;
  Measurement_Units = ifc(is_vlm_target = 'Y', translate(value_list, ',', ';'), '');

  /* SDTM_Target: prefix each SDTM variable with "<domain>.", then join with
     ",", which is the delimiter P21 actually expects for a multi-variable
     target.
     FIX: the source delimits multi-variable SDTM targets with ";" (e.g.
     "AEENRTPT;AEENRF;AEENTPT"), never ",". The original logic checked for a
     comma (countc(SDTM_Target, ',')), which never matched, so only a single
     domain prefix got stuck on the front of the whole semicolon-joined
     string and the later variables were left bare - e.g.
     "AE.AEENRTPT;AEENRF;AEENTPT" instead of a properly split, prefixed,
     comma-joined list. Rewritten as an explicit token loop (same pattern as
     normalize_terms above) that splits on ";", prefixes every part, and
     re-joins with ",". */
  SDTM_Target = sdtm_target_variable;
  if SDTM_Target ne '' then do;
    length _sdtm_out $200 _sdtm_part $100;
    _sdtm_out = '';
    _sdtm_n = countc(SDTM_Target, ';') + 1;
    do _sdtm_i = 1 to _sdtm_n;
      _sdtm_part = strip(scan(SDTM_Target, _sdtm_i, ';'));
      if _sdtm_part ne '' then
        _sdtm_out = catx(',', _sdtm_out, trim(left(domain)) || '.' || _sdtm_part);
    end;
    SDTM_Target = _sdtm_out;
  end;

  if first.section_id then Order = 1;
  else Order + 1;

  Prompt = prompt;
  Question_Text = question_text;
  Completion_Instructions = completion_instructions;

  /* No source mapping exists yet for these P21 template columns - left
     blank deliberately (not an oversight) for manual completion later. */
  Description = '';
  Core = '';
  Method = '';
  Condition = '';
  Implementation_Notes = '';
  Mapping_Instructions = '';
  Reason_Not_Mapped = '';
  Developer_Notes = '';

  /* Extension columns (not part of the P21 template, kept for traceability -
     CDISC COSMOS-specific fields Ben wanted retained). */
  Short_Name = short_name;
  SDTM_Annotation = sdtm_annotation;
  Prepopulated_Term = prepopulated_term;
  Dec_ID = dec_id;

  keep Form Section Order ID Source_Variable_Name Question_Text Prompt
       Description Data_Type Core Length Digits Mandatory Codelist Measurement_Units
       Method Condition Completion_Instructions Implementation_Notes Mapping_Instructions
       SDTM_Target Reason_Not_Mapped Developer_Notes
       Short_Name SDTM_Annotation Prepopulated_Term Dec_ID;
run;

/*------------------------------------------------------------------------------
 6. VALUE LEVEL METADATA (VLM) - unit (*RESU) fields only
    Where-clause aggregation logic is carried over from the original draft
    (group qualifying condition rows by crf_group_id, attach via vlm_group_id)
    - that part of the design wasn't in question. Only the two wrong column
    names (fix #4) and the codelist reference (now direct, no join) changed.
------------------------------------------------------------------------------*/
data wc_candidates;
  set dss_derived;
  if (value_list ne '' or prepopulated_term ne '')
     and implementation_option = 'Normalized'
     and is_vlm_target ne 'Y';

  length assigned_value $200 vlx $500 comparator $4 wc $500;
  if value_list ne '' then do;
    comparator = 'IN';
    vlx = '("' || trim(left(tranwrd(value_list, ';', '","'))) || '")';
    wc = trim(left(crf_item)) || ' IN ' || trim(left(vlx));
  end;
  else if prepopulated_term ne '' then do;
    comparator = 'EQ';
    assigned_value = prepopulated_term;
    if index(trim(left(assigned_value)), ' ') gt 0 then
      assigned_value = '"' || trim(left(assigned_value)) || '"';
    wc = trim(left(crf_item)) || ' EQ ' || trim(left(assigned_value));
  end;
run;

proc sort data=wc_candidates;
  by crf_group_id crf_item;
run;

data wc_combined;
  length combinedwc $500;
  set wc_candidates;
  retain combinedwc;
  by crf_group_id crf_item;
  if first.crf_group_id then combinedwc = wc;
  else combinedwc = catx(' and ', combinedwc, wc);
  if last.crf_group_id then output;
  keep crf_group_id vlm_group_id combinedwc;
run;

proc sql;
  create table vlm_targets as
    select distinct a.domain, a.short_name, a.crf_item, a.sdtm_target_variable,
           a.fixed_data_type, a.length, a.fixed_sig_digits, a.mandatory_variable,
           a.codelist_id, a.vlm_group_id, b.combinedwc
    from dss_derived a
    left join wc_combined b on a.vlm_group_id = b.vlm_group_id
    where a.is_vlm_target = 'Y';
quit;

data vlm_final;
  attrib
    Order                  label='Order'
    Dataset                label='Dataset'
    Variable               label='Variable'
    Variant                label='Variant'
    Where_Clause           label='Where Clause'
    Label                  label='Label'
    Data_Type              label='Data Type'
    Length                 label='Length'
    Significant_Digits     label='Significant Digits'
    Format                 label='Format'
    Mandatory              label='Mandatory'
    Assigned_Value         label='Assigned Value'
    Codelist               length=$80 label='Codelist'
    Origin                 label='Origin'
  ;
  set vlm_targets;

  Dataset  = domain;
  Variable = sdtm_target_variable;     /* fix #4: was the nonexistent "sdtm_variable" */
  Where_Clause = combinedwc;
  Label = short_name;
  Data_Type = fixed_data_type;
  if Data_Type = 'decimal' then Data_Type = 'float';
  Length = length;
  Significant_Digits = fixed_sig_digits;

  if mandatory_variable = 'Y' then Mandatory = 'Yes';   /* fix #4: was "mandatory_value" */
  else if mandatory_variable = 'N' then Mandatory = 'No';
  else Mandatory = '';

  Codelist = codelist_id;
  Origin = 'Collected';
  Order + 1;

  /* Format/Assigned_Value/Decoded_Variable/Codelist_Expected/... have no
     reliable source mapping at the aggregated where-clause level - left
     blank deliberately, same as the original, but now without pretending
     otherwise via a KEEP list referencing columns that were never populated. */
  keep Order Dataset Variable Variant Where_Clause Label Data_Type Length
       Significant_Digits Format Mandatory Assigned_Value Codelist Origin;
run;

/*------------------------------------------------------------------------------
 7. FORM SPEC (static metadata - edit for your study)
------------------------------------------------------------------------------*/
/* Built via direct assignment rather than INFILE/DATALINES: SAS's macro
   facility doesn't allow DATALINES inside a %macro (raw data lines can't be
   macro-scanned safely) - "ERROR: The macro ... generated CARDS (data lines)
   for the DATA step" is what that produces. Static content, so no loss. */
data formspec;
  length Attribute $20 Value $100;
  Attribute = 'Name';        Value = 'Ben Testing';                          output;
  Attribute = 'Version';     Value = '999';                                  output;
  Attribute = 'Type';        Value = 'EDC';                                  output;
  Attribute = 'Description'; Value = 'CDISC COSMOS Data CRF Specialization'; output;
  Attribute = 'Publisher';   Value = 'Ben';                                   output;
run;

/*------------------------------------------------------------------------------
 8. TEMPLATE SHEETS WITH NO SOURCE DATA
    Events, Methods, Conditions and ScheduleOfActivities are part of the P21
    Form Template but nothing in a CRF Specializations export maps to them.
    Written as header-only sheets (via "if 0;", a standard SAS idiom for a
    zero-row shell dataset) so the workbook still has the complete, correctly
    labelled sheet set - ready for manual completion rather than missing.
------------------------------------------------------------------------------*/
data events;
  attrib
    Order           label='Order'
    ID              length=$200 label='ID'
    Name            length=$200 label='Name'
    Type            length=$50  label='Type'
    Mandatory       length=$3   label='Mandatory'
    Repeating       length=$3   label='Repeating'
    Developer_Notes length=$200 label='Developer Notes'
  ;
  if 0;
run;

data methods;
  attrib
    ID                 length=$200 label='ID'
    Name               length=$200 label='Name'
    Type               length=$50  label='Type'
    Description        length=$200 label='Description'
    Expression_Context length=$50  label='Expression Context'
    Expression_Code    length=$500 label='Expression Code'
  ;
  if 0;
run;

data conditions;
  attrib
    ID                 length=$200 label='ID'
    Name               length=$200 label='Name'
    Description        length=$200 label='Description'
    Expression_Context length=$50  label='Expression Context'
    Expression_Code    length=$500 label='Expression Code'
  ;
  if 0;
run;

data schedule_of_activities;
  attrib
    Order     label='Order'
    Event     length=$200 label='Event'
    Form      length=$200 label='Form'
    Mandatory length=$3   label='Mandatory'
  ;
  if 0;
run;

/*------------------------------------------------------------------------------
 9. EXPORT
    Sheet order below matches the Form-Spec-Template workbook (FormSpec,
    Events, Forms, Sections, Questions, Units, Codelists, Methods, Conditions,
    Terms, ScheduleOfActivities); QC_Checks is appended as a bonus sheet, not
    part of the template. Output file names are macro parameters (fix: the
    outfile was previously hardcoded) - see outfile_form=/outfile_vlm= above.
------------------------------------------------------------------------------*/
proc export data=formspec               outfile="&fpath.\&outfile_form" dbms=xlsx replace label; sheet='FormSpec';             run;
proc export data=events                 outfile="&fpath.\&outfile_form" dbms=xlsx replace label; sheet='Events';               run;
proc export data=form_final             outfile="&fpath.\&outfile_form" dbms=xlsx replace label; sheet='Forms';                run;
proc export data=sections_final         outfile="&fpath.\&outfile_form" dbms=xlsx replace label; sheet='Sections';             run;
proc export data=questions_final        outfile="&fpath.\&outfile_form" dbms=xlsx replace label; sheet='Questions';            run;
proc export data=units                  outfile="&fpath.\&outfile_form" dbms=xlsx replace label; sheet='Units';                run;
proc export data=codelist_final         outfile="&fpath.\&outfile_form" dbms=xlsx replace label; sheet='Codelists';            run;
proc export data=methods                outfile="&fpath.\&outfile_form" dbms=xlsx replace label; sheet='Methods';              run;
proc export data=conditions              outfile="&fpath.\&outfile_form" dbms=xlsx replace label; sheet='Conditions';          run;
proc export data=terms_final            outfile="&fpath.\&outfile_form" dbms=xlsx replace label; sheet='Terms';                run;
proc export data=schedule_of_activities outfile="&fpath.\&outfile_form" dbms=xlsx replace label; sheet='ScheduleOfActivities'; run;
proc export data=qc_checks              outfile="&fpath.\&outfile_form" dbms=xlsx replace label; sheet='QC_Checks';           run;

proc export data=vlm_final outfile="&fpath.\&outfile_vlm" dbms=xlsx replace label; run;

%mend crfdss_to_form;

/* Example call - adjust fpath/dssfile, and outfile_form=/outfile_vlm= if you
   want output names other than the defaults (crf_form.xlsx / crf_vlm.xlsx). */
%crfdss_to_form(fpath=C:\Users\bmant\Downloads, dssfile=cdisc_crf_specializations_draft);
