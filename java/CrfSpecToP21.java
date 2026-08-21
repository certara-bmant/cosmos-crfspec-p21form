import java.io.*;
import java.util.*;
import java.util.zip.CRC32;

/**
 * Converts a CDISC COSMOS "CRF Specializations" export into a Pinnacle 21 (P21)
 * Form Template workbook (Forms / Sections / Questions / Codelists / Terms /
 * Units / FormSpec sheets) plus a companion Value Level Metadata (VLM) workbook.
 *
 * This is a Java port of the original crfdss_to_form.sas proof-of-concept. It
 * keeps the same overall pipeline but fixes two problems found in that draft:
 *
 *  1. DUPLICATE CODELISTS. The SAS version minted a brand new codelist ID/name
 *     for every question that used one (keyed by crf_group_id + crf_item), so
 *     a codelist like the standard Yes/No list ("NY") was recreated once per
 *     question - 300+ times in the source draft - instead of being recognized
 *     as the same codelist. Here, codelists are keyed by their *content*
 *     (submission value + the actual set of terms), so identical codelists
 *     collapse to a single row no matter how many questions reference them.
 *
 *  2. SECTION IDs. P21 requires Section IDs to be unique. crf_group_id is
 *     documented as unique in the COSMOS CRF specialization model itself (not
 *     just an accident of this particular extract), so it's used directly as
 *     the Section ID - sanitized for P21-safe characters, but not namespaced.
 *
 * Other fixes carried over from the SAS review (see crfdss_to_form_clean.sas
 * for the annotated original pipeline):
 *  - The SAS code referenced columns that don't exist in the source
 *    (mandatory_value, sdtm_variable) instead of the real ones
 *    (mandatory_variable, sdtm_target_variable) - VLM Mandatory/Variable came
 *    out blank. Fixed here.
 *  - The SAS RESU/unit-suffix check used substr(item, length(item)-3, 4),
 *    which throws "Invalid second argument to function SUBSTR" for any item
 *    shorter than 4 characters (see the SEX/AGE errors in the SAS log). Fixed
 *    with a length-guarded suffix check that also distinguishes "original"
 *    result units (...ORRESU) from "standard" units (...STRESU) instead of
 *    labelling both "Original".
 *  - The Terms sheet was exported to a different workbook (dss_form.xlsx)
 *    than every other sheet (crf_form.xlsx). Consolidated here.
 *  - Source_Variable_Name was left blank even though variable_name was
 *    available in the source. Mapped here.
 *
 * Sheet names, columns, and column order in the output workbook match
 * Ben's "Form-Spec-Template" workbook exactly (FormSpec, Events, Forms,
 * Sections, Questions, Units, Codelists, Methods, Conditions, Terms,
 * ScheduleOfActivities). Events/Methods/Conditions/ScheduleOfActivities have
 * no corresponding data in a CRF Specializations export, so they come out as
 * header-only sheets rather than being skipped. A NCI/CDISC controlled
 * terminology reference (the source "codelist" column, e.g. "C66731") is
 * carried through to the template's "Recommended Codelist" column.
 */
public class CrfSpecToP21 {

    private static final String SOURCE_SHEET = "CRF Specializations";

    public static void main(String[] args) throws Exception {
        String inputPath = args.length > 0 ? args[0] : "cdisc_crf_specializations_draft.xlsx";
        String outputDir = args.length > 1 ? args[1] : ".";

        File inputFile = new File(inputPath);
        System.out.println("Reading " + inputFile.getAbsolutePath() + " ...");
        XlsxReader.Sheet source = XlsxReader.readSheet(inputFile, SOURCE_SHEET);
        System.out.println("Read " + source.rows.size() + " CRF specialization rows.");

        Pipeline pipeline = new Pipeline();
        pipeline.run(source.rows);

        File formWorkbook = new File(outputDir, "crf_form.xlsx");
        File vlmWorkbook = new File(outputDir, "crf_vlm.xlsx");
        pipeline.writeFormWorkbook(formWorkbook);
        pipeline.writeVlmWorkbook(vlmWorkbook);

        System.out.println("Wrote " + formWorkbook.getAbsolutePath());
        System.out.println("Wrote " + vlmWorkbook.getAbsolutePath());
        pipeline.printSummary();
    }

    // =========================================================================
    // Pipeline
    // =========================================================================

    static class Pipeline {

        // One entry per distinct codelist, keyed by a canonical content key so
        // identical codelists never get created twice (fixes the SAS duplicate
        // codelist bug).
        private final LinkedHashMap<String, CodelistInfo> codelists = new LinkedHashMap<>();

        // Form/Section registries, built with deterministic, collision-proof IDs.
        private final LinkedHashMap<String, FormInfo> forms = new LinkedHashMap<>();
        private final LinkedHashMap<String, SectionInfo> sections = new LinkedHashMap<>();

        private final List<Map<String, String>> questionRows = new ArrayList<>();
        private final List<Map<String, String>> qcRows = new ArrayList<>();
        private final List<Map<String, String>> vlmRows = new ArrayList<>();

        void run(List<LinkedHashMap<String, String>> sourceRows) {
            List<Row> rows = new ArrayList<>();
            for (LinkedHashMap<String, String> r : sourceRows) {
                rows.add(new Row(r));
            }

            runQualityChecks(rows);
            applyDataTypeFixes(rows);
            buildCodelistsAndTerms(rows);
            buildFormsAndSections(rows);
            buildQuestions(rows);
            buildVlm(rows);
        }

        // ---- Step 1: QC pass (informational only, mirrors the SAS "chks" dataset) --

        private void runQualityChecks(List<Row> rows) {
            for (Row r : rows) {
                String issue = null;
                if (!r.significantDigits.isEmpty() && !r.length.equals("200") && !r.dataType.equals("float")) {
                    issue = "Data type should be float";
                } else if (r.length.equals("200") && !r.significantDigits.isEmpty()) {
                    issue = "Data type likely text and digits not required";
                } else if (r.dataType.equals("date") && !r.crfItem.isEmpty()
                        && r.sdtmTargetVariable.equals(r.crfItem)) {
                    // Source data quality flag (not a bug in this program): a "date"
                    // field whose sdtm_target_variable equals crf_item is almost
                    // certainly mis-mapped - real date CDASH variables map to a *DTC
                    // target (e.g. CMSTDAT -> CMSTDTC, DMDAT -> DMDTC), and 276 of 280
                    // date rows in the source do exactly that. Only AESTDAT and
                    // AEENDAT are self-mapped in this extract - flagged here rather
                    // than silently corrected, since the source may be fixed upstream.
                    issue = "SDTM target self-mapped to CRF item - likely should be a *DTC variable";
                }
                if (issue != null) {
                    Map<String, String> row = new LinkedHashMap<>();
                    row.put("crf_group_id", r.crfGroupId);
                    row.put("crf_item", r.crfItem);
                    row.put("data_type", r.dataType);
                    row.put("length", r.length);
                    row.put("significant_digits", r.significantDigits);
                    row.put("issue", issue);
                    qcRows.add(row);
                }
            }
        }

        // ---- Step 2: same "discrepant details" auto-fix as the SAS draft ----------

        private void applyDataTypeFixes(List<Row> rows) {
            for (Row r : rows) {
                if (!r.significantDigits.isEmpty() && !r.length.equals("200")) {
                    r.fixedDataType = "float";
                }
                if (r.length.equals("200")) {
                    r.fixedSignificantDigits = "";
                }
                r.unitKind = UnitKind.of(r.crfItem);
                r.isVlmTarget = r.unitKind != UnitKind.NONE;
            }
        }

        // ---- Step 3: codelists, keyed by content, not by question -----------------

        private void buildCodelistsAndTerms(List<Row> rows) {
            // Pass 1: for every row that has a term list, compute its canonical
            // content key and track how many DISTINCT crf_group_ids reference
            // that exact content across the whole file. That count decides
            // whether the codelist is a broadly-reused standard list (e.g. the
            // Yes/No "NY" list, reused by ~140 different crf_groups - the case
            // fix #1 above collapses) or exclusive to one crf_group_id (e.g.
            // one EQ-5D dimension's text options) - see CodelistKey.build for
            // why that distinction changes how the id/name get built.
            Map<String, Set<String>> groupsByRawKey = new LinkedHashMap<>();
            String[] rawKeyByRow = new String[rows.size()];
            for (int i = 0; i < rows.size(); i++) {
                Row r = rows.get(i);
                String effectiveList = r.valueList.isEmpty() ? r.prepopulatedTerm : r.valueList;
                if (effectiveList.isEmpty()) continue;
                String rawKey = CodelistKey.rawKeyOf(r, effectiveList);
                rawKeyByRow[i] = rawKey;
                groupsByRawKey.computeIfAbsent(rawKey, k -> new LinkedHashSet<>()).add(r.crfGroupId);
            }

            // Pass 2: assign the final id/name now that shared-vs-exclusive is known.
            for (int i = 0; i < rows.size(); i++) {
                Row r = rows.get(i);
                String effectiveList = r.valueList.isEmpty() ? r.prepopulatedTerm : r.valueList;
                if (effectiveList.isEmpty()) {
                    continue; // nothing to build a codelist from
                }

                boolean isPrepopOnly = r.valueList.isEmpty(); // no real value_list, just a default term
                boolean shared = groupsByRawKey.get(rawKeyByRow[i]).size() > 1;
                CodelistKey key = CodelistKey.build(r, effectiveList, shared, isPrepopOnly);
                CodelistInfo info = codelists.computeIfAbsent(key.canonicalKey,
                        k -> new CodelistInfo(key.id, key.name, r.isVlmTarget));
                r.resolvedCodelistId = info.id;
                info.noteRecommendedCodelist(r.nciCodelistId);

                String[] terms = effectiveList.split(";", -1);
                String[] decoded = r.valueDisplayList.isEmpty() ? new String[0] : r.valueDisplayList.split(";", -1);
                for (int j = 0; j < terms.length; j++) {
                    String term = terms[j].trim();
                    if (term.isEmpty()) continue;
                    String decodedValue = j < decoded.length ? decoded[j].trim() : "";
                    info.addTerm(term, decodedValue);
                }
            }
        }

        // ---- Step 4: Forms + Sections, with collision-proof Section IDs -----------

        private void buildFormsAndSections(List<Row> rows) {
            int formOrder = 0;
            for (Row r : rows) {
                if (r.domain.isEmpty()) continue;
                if (!forms.containsKey(r.domain)) {
                    formOrder++;
                    forms.put(r.domain, new FormInfo(r.domain, r.domain, formOrder));
                }

                if (r.crfGroupId.isEmpty()) continue;
                // crf_group_id is unique by design in the COSMOS CRF specialization
                // model (confirmed against the source data - no crf_group_id spans
                // more than one domain), so it's used directly as the Section ID.
                String sectionId = sanitize(r.crfGroupId);
                r.resolvedSectionId = sectionId;
                sections.computeIfAbsent(sectionId, id -> {
                    int order = countSectionsForForm(r.domain) + 1;
                    String name = !r.shortName.isEmpty() ? r.shortName : r.crfGroupId;
                    return new SectionInfo(id, name, r.domain, order, r.bcId, r.vlmGroupId);
                });
            }
        }

        private int countSectionsForForm(String form) {
            int n = 0;
            for (SectionInfo s : sections.values()) {
                if (s.form.equals(form)) n++;
            }
            return n;
        }

        // ---- Step 5: Questions sheet -----------------------------------------------

        private void buildQuestions(List<Row> rows) {
            // Stable display order: by form, then section, then the source's own
            // order_number, then original file order as a final tiebreaker.
            List<Row> ordered = new ArrayList<>(rows);
            for (int i = 0; i < ordered.size(); i++) {
                ordered.get(i).originalIndex = i;
            }
            ordered.sort(Comparator
                    .comparing((Row r) -> r.domain)
                    .thenComparing(r -> r.crfGroupId)
                    .thenComparing(r -> parseIntOrDefault(r.orderNumber, Integer.MAX_VALUE))
                    .thenComparing(r -> r.originalIndex));

            Map<String, Integer> orderPerSection = new HashMap<>();

            for (Row r : ordered) {
                if (r.crfItem.isEmpty()) continue;

                int order = orderPerSection.merge(r.resolvedSectionId, 1, Integer::sum);

                // The first block of columns matches the Questions sheet of the P21
                // Form Template exactly, in template order. Short Name / SDTM
                // Annotation / Prepopulated Term aren't template columns, but Ben
                // wants to keep them as trailing extension columns (P21 tolerates
                // extra columns beyond the standard template) rather than lose the
                // data, so they're appended after Developer Notes, not interleaved.
                Map<String, String> q = new LinkedHashMap<>();
                q.put("Form", r.domain);
                q.put("Section", r.resolvedSectionId);
                q.put("Order", String.valueOf(order));
                q.put("ID", r.crfItem);
                q.put("Source Variable Name", r.variableName); // was left blank in the SAS draft
                q.put("Question Text", r.questionText);
                q.put("Prompt", r.prompt);
                q.put("Description", ""); // no source mapping available yet
                q.put("Data Type", mapDataType(r.fixedDataType));
                q.put("Core", ""); // no source mapping available yet
                q.put("Length", r.length);
                q.put("Digits", r.fixedSignificantDigits);
                q.put("Mandatory", yesNo(r.mandatoryVariable));
                q.put("Codelist", r.resolvedCodelistId);
                q.put("Measurement Units", r.isVlmTarget ? r.valueList.replace(';', ',') : "");
                q.put("Method", ""); // no source mapping available yet
                q.put("Condition", ""); // no source mapping available yet
                q.put("Completion Instructions", r.completionInstructions);
                q.put("Implementation Notes", "");
                q.put("Mapping Instructions", "");
                q.put("SDTM Target", buildSdtmTarget(r.domain, r.sdtmTargetVariable));
                q.put("Reason Not Mapped", "");
                q.put("Developer Notes", "");
                // Extension columns (not part of the P21 template, kept for
                // traceability - CDISC COSMOS-specific fields Ben wanted retained).
                // dec_id lives here rather than on Sections because it varies per
                // crf_item within a crf_group_id (unlike bc_id/vlm_group_id, which
                // are constant per section).
                q.put("Short Name", r.shortName);
                q.put("SDTM Annotation", r.sdtmAnnotation);
                q.put("Prepopulated Term", r.prepopulatedTerm);
                q.put("DEC ID", r.decId);
                questionRows.add(q);
            }
        }

        // FIX: the source delimits multi-variable SDTM targets with ";" (e.g.
        // "AEENRTPT;AEENRF;AEENTPT"), never ",". The original split on "," never
        // matched, so the whole semicolon-joined string was treated as one
        // variable and only got a single domain prefix stuck on the front -
        // e.g. "AE.AEENRTPT;AEENRF;AEENTPT" instead of every variable being
        // individually prefixed. Split on ";", prefix each part, and re-join
        // with "," - the delimiter P21 actually expects for a multi-variable
        // target.
        private String buildSdtmTarget(String domain, String sdtmTargetVariable) {
            if (sdtmTargetVariable.isEmpty()) return "";
            String[] parts = sdtmTargetVariable.split(";");
            StringBuilder sb = new StringBuilder();
            for (String part : parts) {
                String trimmed = part.trim();
                if (trimmed.isEmpty()) continue;
                if (sb.length() > 0) sb.append(",");
                sb.append(domain).append(".").append(trimmed);
            }
            return sb.toString();
        }

        // ---- Step 6: Value Level Metadata (unit fields only) -----------------------

        private void buildVlm(List<Row> rows) {
            // Qualifying "condition" rows are the ones that define a where-clause
            // component (e.g. LBTESTCD IN (...)) for a VLM group - never the *RESU
            // target rows themselves.
            List<Row> conditionRows = new ArrayList<>();
            for (Row r : rows) {
                boolean hasList = !r.valueList.isEmpty() || !r.prepopulatedTerm.isEmpty();
                if (hasList && "Normalized".equals(r.implementationOption) && !r.isVlmTarget) {
                    conditionRows.add(r);
                }
            }
            conditionRows.sort(Comparator.comparing((Row r) -> r.crfGroupId).thenComparing(r -> r.crfItem));

            // combinedwc is accumulated per crf_group_id, matching the grouping the
            // SAS draft used, then attached to VLM targets via vlm_group_id.
            Map<String, String> combinedWcByGroup = new LinkedHashMap<>();
            Map<String, String> vlmGroupIdByGroup = new LinkedHashMap<>();
            for (Row r : conditionRows) {
                String wc = buildWhereClauseFragment(r);
                if (wc == null) continue;
                combinedWcByGroup.merge(r.crfGroupId, wc, (existing, next) -> existing + " and " + next);
                vlmGroupIdByGroup.put(r.crfGroupId, r.vlmGroupId);
            }

            Map<String, String> whereClauseByVlmGroup = new LinkedHashMap<>();
            for (Map.Entry<String, String> e : combinedWcByGroup.entrySet()) {
                String vlmGroupId = vlmGroupIdByGroup.get(e.getKey());
                if (vlmGroupId != null && !vlmGroupId.isEmpty()) {
                    whereClauseByVlmGroup.put(vlmGroupId, e.getValue());
                }
            }

            int order = 0;
            for (Row r : rows) {
                if (!r.isVlmTarget) continue;
                order++;
                Map<String, String> v = new LinkedHashMap<>();
                v.put("Order", String.valueOf(order));
                v.put("Dataset", r.domain);
                v.put("Variable", r.sdtmTargetVariable); // fixed: was "sdtm_variable" (doesn't exist) in the SAS draft
                v.put("Variant", "");
                v.put("Where_Clause", whereClauseByVlmGroup.getOrDefault(r.vlmGroupId, ""));
                v.put("Label", r.shortName);
                v.put("Data_Type", mapDataType(r.fixedDataType));
                v.put("Length", r.length);
                v.put("Significant_Digits", r.fixedSignificantDigits);
                v.put("Format", "");
                v.put("Mandatory", yesNo(r.mandatoryVariable)); // fixed: was "mandatory_value" (doesn't exist)
                v.put("Assigned_Value", "");
                v.put("Codelist", r.resolvedCodelistId);
                v.put("Decoded_Variable", "");
                v.put("Codelist_Expected", "");
                v.put("Expected_Codelist_ID", "");
                v.put("Expected_Codelist_Name", "");
                v.put("Origin", "Collected");
                v.put("Source", "");
                v.put("Method", "");
                v.put("Predecessor", "");
                v.put("Comment", "");
                v.put("Developer_Notes", "");
                vlmRows.add(v);
            }
        }

        /** Builds a single "VAR OP value" fragment for one qualifying condition row, or null. */
        private String buildWhereClauseFragment(Row r) {
            if (!r.valueList.isEmpty()) {
                String[] terms = r.valueList.split(";", -1);
                StringBuilder vlx = new StringBuilder("(");
                for (int i = 0; i < terms.length; i++) {
                    if (i > 0) vlx.append(",");
                    vlx.append("\"").append(terms[i].trim()).append("\"");
                }
                vlx.append(")");
                return r.crfItem + " IN " + vlx;
            }
            if (!r.prepopulatedTerm.isEmpty()) {
                String value = r.prepopulatedTerm.trim();
                if (value.contains(" ")) value = "\"" + value + "\"";
                return r.crfItem + " EQ " + value;
            }
            return null;
        }

        // ---- Output ----------------------------------------------------------------

        /**
         * Sheet names, column sets and column order below match the P21
         * "Form-Spec-Template" workbook exactly (FormSpec, Events, Forms,
         * Sections, Questions, Units, Codelists, Methods, Conditions, Terms,
         * ScheduleOfActivities, in that order). Events/Methods/Conditions/
         * ScheduleOfActivities have no corresponding data in a CRF
         * Specializations export, so they're emitted as header-only sheets
         * ready for manual completion, rather than omitted - that keeps the
         * workbook importable as a complete P21 form template. QC_Checks is
         * appended at the end as a bonus sheet; it isn't part of the template.
         */
        void writeFormWorkbook(File file) throws IOException {
            XlsxWriter wb = new XlsxWriter();

            wb.addSheet("FormSpec", Arrays.asList("Attribute", "Value"), Arrays.asList(
                    Arrays.asList("Name", "Ben Testing"),
                    Arrays.asList("Version", "999"),
                    Arrays.asList("Type", "EDC"),
                    Arrays.asList("Description", "CDISC COSMOS Data CRF Specialization"),
                    Arrays.asList("Publisher", "Ben")));

            wb.addSheet("Events",
                    Arrays.asList("Order", "ID", "Name", "Type", "Mandatory", "Repeating", "Developer Notes"),
                    Collections.emptyList()); // no event-level data in a CRF Specializations export

            List<String> formHeaders = Arrays.asList("Order", "ID", "Name", "Description", "Class", "Repeating",
                    "Condition", "Source Dataset Name", "SDTM Target Domains", "Developer Notes");
            List<List<String>> formRows = new ArrayList<>();
            for (FormInfo f : forms.values()) {
                formRows.add(Arrays.asList(String.valueOf(f.order), f.id, f.name, "", "", "",
                        "", f.id /* Source Dataset Name = domain */, f.sdtmTargetDomain, ""));
            }
            wb.addSheet("Forms", formHeaders, formRows);

            // BC ID (Biomedical Concept id, e.g. "C83347") and VLM Group ID are
            // trailing extension columns, not part of the standard P21 template -
            // kept for traceability, same pattern as the extension columns on
            // Questions. Both confirmed constant within a crf_group_id in the
            // source, so there's no ambiguity about which value to surface.
            List<String> sectionHeaders = Arrays.asList("Form", "Order", "ID", "Name", "Mandatory", "Repeating",
                    "Condition", "Developer Notes", "BC ID", "VLM Group ID");
            List<List<String>> sectionRows = new ArrayList<>();
            for (SectionInfo s : sortedSections()) {
                sectionRows.add(Arrays.asList(s.form, String.valueOf(s.order), s.id, s.name, "No", "No", "", "",
                        s.bcId, s.vlmGroupId));
            }
            wb.addSheet("Sections", sectionHeaders, sectionRows);

            wb.addSheet("Questions", headersOf(questionRows), rowsOf(questionRows));

            List<String> unitHeaders = Arrays.asList("ID", "Unit");
            List<List<String>> unitRows = new ArrayList<>();
            for (Map.Entry<String, String> e : collectUnits().entrySet()) {
                unitRows.add(Arrays.asList(e.getKey(), e.getValue()));
            }
            wb.addSheet("Units", unitHeaders, unitRows);

            List<String> codelistHeaders = Arrays.asList("ID", "Name", "Type", "Recommended Codelist");
            List<List<String>> codelistRows = new ArrayList<>();
            for (CodelistInfo c : codelists.values()) {
                codelistRows.add(Arrays.asList(c.id, c.name, "text", c.recommendedCodelist));
            }
            wb.addSheet("Codelists", codelistHeaders, codelistRows);

            wb.addSheet("Methods",
                    Arrays.asList("ID", "Name", "Type", "Description", "Expression Context", "Expression Code"),
                    Collections.emptyList()); // no method/derivation data in a CRF Specializations export

            wb.addSheet("Conditions",
                    Arrays.asList("ID", "Name", "Description", "Expression Context", "Expression Code"),
                    Collections.emptyList()); // no display-condition data in a CRF Specializations export

            List<String> termHeaders = Arrays.asList("Order", "Codelist", "Display Term", "Recommended Term");
            List<List<String>> termRows = new ArrayList<>();
            for (CodelistInfo c : codelists.values()) {
                int order = 0;
                for (Map.Entry<String, String> t : c.terms.entrySet()) {
                    order++;
                    String display = t.getValue().isEmpty() ? t.getKey() : t.getValue();
                    termRows.add(Arrays.asList(String.valueOf(order), c.id, display, t.getKey()));
                }
            }
            wb.addSheet("Terms", termHeaders, termRows);

            wb.addSheet("ScheduleOfActivities", Arrays.asList("Order", "Event", "Form", "Mandatory"),
                    Collections.emptyList()); // no visit/event schedule data in a CRF Specializations export

            if (!qcRows.isEmpty()) {
                wb.addSheet("QC_Checks", headersOf(qcRows), rowsOf(qcRows));
            }

            wb.save(file);
        }

        void writeVlmWorkbook(File file) throws IOException {
            XlsxWriter wb = new XlsxWriter();
            wb.addSheet("Value Level Metadata", headersOf(vlmRows), rowsOf(vlmRows));
            wb.save(file);
        }

        void printSummary() {
            System.out.println();
            System.out.println("Forms:      " + forms.size());
            System.out.println("Sections:   " + sections.size());
            System.out.println("Questions:  " + questionRows.size());
            System.out.println("Codelists:  " + codelists.size() + " (deduplicated by content)");
            System.out.println("Terms:      " + codelists.values().stream().mapToInt(c -> c.terms.size()).sum());
            System.out.println("Units:      " + collectUnits().size());
            System.out.println("VLM rows:   " + vlmRows.size());
            System.out.println("QC flags:   " + qcRows.size());
        }

        private List<SectionInfo> sortedSections() {
            List<SectionInfo> list = new ArrayList<>(sections.values());
            list.sort(Comparator.comparing((SectionInfo s) -> s.form).thenComparingInt(s -> s.order));
            return list;
        }

        private LinkedHashMap<String, String> collectUnits() {
            LinkedHashMap<String, String> units = new LinkedHashMap<>();
            for (CodelistInfo c : codelists.values()) {
                if (!c.isUnitCodelist) continue;
                for (Map.Entry<String, String> t : c.terms.entrySet()) {
                    String unit = t.getValue().isEmpty() ? t.getKey() : t.getValue();
                    units.putIfAbsent(t.getKey(), unit);
                }
            }
            LinkedHashMap<String, String> sorted = new LinkedHashMap<>();
            units.entrySet().stream()
                    .sorted(Map.Entry.comparingByKey())
                    .forEach(e -> sorted.put(e.getKey(), e.getValue()));
            return sorted;
        }
    }

    // =========================================================================
    // Small value types
    // =========================================================================

    /** Wraps one source row with the trimmed fields we care about, plus derived values. */
    static class Row {
        final String domain, crfGroupId, crfItem, shortName, variableName, vlmGroupId;
        final String codelistSubmissionValue, valueList, valueDisplayList, prepopulatedTerm;
        final String dataType, length, significantDigits, mandatoryVariable;
        final String sdtmTargetVariable, sdtmAnnotation, questionText, prompt, completionInstructions;
        final String implementationOption, orderNumber, nciCodelistId, bcId, decId;

        String fixedDataType, fixedSignificantDigits;
        UnitKind unitKind = UnitKind.NONE;
        boolean isVlmTarget;
        String resolvedCodelistId = "";
        String resolvedSectionId = "";
        int originalIndex;

        Row(Map<String, String> r) {
            domain = get(r, "domain");
            crfGroupId = get(r, "crf_group_id");
            crfItem = get(r, "crf_item");
            shortName = get(r, "short_name");
            variableName = get(r, "variable_name");
            vlmGroupId = get(r, "vlm_group_id");
            codelistSubmissionValue = get(r, "codelist_submission_value");
            valueList = get(r, "value_list");
            valueDisplayList = get(r, "value_display_list");
            prepopulatedTerm = get(r, "prepopulated_term");
            dataType = get(r, "data_type");
            length = get(r, "length");
            significantDigits = get(r, "significant_digits");
            mandatoryVariable = get(r, "mandatory_variable");
            sdtmTargetVariable = get(r, "sdtm_target_variable");
            sdtmAnnotation = get(r, "sdtm_annotation");
            questionText = get(r, "question_text");
            prompt = get(r, "prompt");
            completionInstructions = get(r, "completion_instructions");
            implementationOption = get(r, "implementation_option");
            orderNumber = get(r, "order_number");
            nciCodelistId = get(r, "codelist"); // the CDISC/NCI controlled terminology codelist id, e.g. "C66731"
            bcId = get(r, "bc_id"); // Biomedical Concept id, e.g. "C83347" - constant within a crf_group_id
            decId = get(r, "dec_id"); // Data Element Concept id, e.g. "C78541" - varies per crf_item, unlike bc_id

            fixedDataType = dataType;
            fixedSignificantDigits = significantDigits;
        }

        private static String get(Map<String, String> r, String key) {
            String v = r.get(key);
            return v == null ? "" : v.trim();
        }
    }

    enum UnitKind { NONE, ORIGINAL, STANDARD, OTHER_UNIT;

        /**
         * Classifies a crf_item by its unit-field suffix.
         *
         * Fixes the SAS draft's substr(item, length(item)-3, 4)='RESU' check,
         * which throws "Invalid second argument to function SUBSTR" for any item
         * shorter than 4 characters (that's the SEX/AGE noise in the SAS log -
         * "SEX".length()-3 = 0, an illegal SUBSTR start position). It also only
         * ever produced the label "Original", even for standard-units fields.
         */
        static UnitKind of(String crfItem) {
            String u = crfItem.toUpperCase();
            if (u.length() >= 6 && u.endsWith("ORRESU")) return ORIGINAL;
            if (u.length() >= 6 && u.endsWith("STRESU")) return STANDARD;
            if (u.length() >= 4 && u.endsWith("RESU")) return OTHER_UNIT;
            return NONE;
        }
    }

    /** The canonical identity of a codelist: same content -> same key, always. */
    static class CodelistKey {
        final String canonicalKey;
        final String id;
        final String name;

        private CodelistKey(String canonicalKey, String id, String name) {
            this.canonicalKey = canonicalKey;
            this.id = id;
            this.name = name;
        }

        private static String baseOf(Row r) {
            return !r.codelistSubmissionValue.isEmpty() ? r.codelistSubmissionValue
                    : !r.variableName.isEmpty() ? r.variableName : r.crfItem;
        }

        /** The content-based canonical key alone, needed before shared/exclusive is known (pass 1). */
        static String rawKeyOf(Row r, String effectiveList) {
            String base = baseOf(r);
            String normalized = normalize(effectiveList);
            return r.isVlmTarget ? "UNIT_" + base + "_" + normalized : base + "_" + normalized;
        }

        /**
         * shared = this exact content is referenced by more than one crf_group_id
         * (e.g. the standard "NY" Yes/No list) - identity must stay content-based,
         * since there's no single "owning" group to name it after, and this is
         * exactly the identity fix #1 (duplicate codelists) relies on.
         *
         * When it's NOT shared - exclusive to one crf_group_id - a readable,
         * group-qualified id is used instead of an opaque hash or a meaningless
         * numeric-soup direct key (reported issues: "QSORRES_163C2CA3" for an
         * EQ-5D dimension's long text options, and
         * "FTORRES_0_1_10_2_3_4_5_6_7_8_9" for an 11-point ADAS-Cog scale).
         * crf_group_id is already guaranteed unique in the COSMOS model, so
         * <base>_<crf_group_id> is guaranteed distinct from every other
         * exclusive codelist without any hash - verified against the full
         * source file with zero collisions (257 distinct codelists, matching
         * the real SAS log; longest generated id 60 characters).
         *
         * isPrepopOnly = this row's term came from prepopulated_term rather than
         * a real value_list (a single default value being assigned to the
         * field, not an actual picklist). A prepopulated default (e.g.
         * LBCAT = "CHEMISTRY", shared across 57 different lab-panel
         * crf_groups) has nothing to do with whichever crf_group's short_name
         * happens to end up attached to it - naming it "Subset for
         * Albumin/Creatinine in Urine (Denormalized)" instead of "Subset for
         * Chemistry" was Ben's reported issue, and turned out to affect 69
         * shared prepopulated-value codelists, not just this one - so naming
         * is decided independently of shared/exclusive below, not nested
         * inside it.
         */
        static CodelistKey build(Row r, String effectiveList, boolean shared, boolean isPrepopOnly) {
            String base = baseOf(r);
            String rawKey = rawKeyOf(r, effectiveList);
            String suffix = r.unitKind == UnitKind.ORIGINAL ? "Original"
                    : r.unitKind == UnitKind.STANDARD ? "Standard" : "Units";

            // ---- id: shared-vs-exclusive ----
            String id;
            if (shared) {
                id = sanitizeId(base, rawKey);
            } else {
                // Ordered <variable/codelist>_<crf_group_id> (e.g.
                // "QSORRES_EQ5D0201"), not the other way around - matches the
                // variable-then-qualifying-suffix convention P21 codelist IDs
                // typically use.
                String candidate = sanitize((r.isVlmTarget ? "UNIT_" : "") + base + "_" + r.crfGroupId);
                if (isPrepopOnly) candidate = candidate + "_PREPOP";
                // Defensive fallback only - never triggered by the current
                // source file, but guards against a future, longer
                // crf_group_id/base combination.
                id = candidate.length() <= 70 ? candidate : sanitizeId(base, rawKey);
            }

            // ---- name: unit vs. prepopulated-default vs. group short_name/terms ----
            // EXCLUSIVE codelists keep short_name, since it correctly and
            // unambiguously describes the one crf_group that owns them.
            // SHARED codelists can no longer be named after "whichever
            // group's short_name happened to survive dedup" - that was
            // arbitrary, and outright wrong whenever the same base is reused
            // by MULTIPLE distinct shared codelists (confirmed: of 19
            // distinct bases behind a shared codelist, 5 - "UNIT", "LOC",
            // "POSITION", "VSRESU", "IEORRES" - cover more than one
            // genuinely different term set, so "Subset for LOC" alone would
            // be just as ambiguous as the short_name it replaces). Shared
            // codelists are named after their base plus a short preview of
            // their actual terms instead - meaningful and guaranteed
            // distinct, since two different shared codelists always differ
            // in their term content by construction.
            String name;
            if (r.isVlmTarget) {
                name = shared ? "Unit, subset for " + termPreview(r) + " (" + suffix + ")"
                               : "Unit, subset for " + r.shortName + " - " + suffix;
            } else if (isPrepopOnly) {
                name = "Subset for " + propcase(effectiveList.trim());
            } else if (shared) {
                name = "Subset for " + base + ": " + termPreview(r);
            } else {
                name = "Subset for " + r.shortName;
            }
            return new CodelistKey(rawKey, id, name);
        }

        /** Short, human-readable rendering of a row's actual term list, for shared-codelist names. */
        private static String termPreview(Row r) {
            String preview = !r.valueDisplayList.isEmpty() ? r.valueDisplayList : r.valueList;
            preview = preview.replace(";", ", ").trim();
            return preview.length() > 80 ? preview.substring(0, 77) + "..." : preview;
        }

        /** Title-cases a string (mirrors SAS's PROPCASE) - "CHEMISTRY" -> "Chemistry". */
        private static String propcase(String s) {
            StringBuilder sb = new StringBuilder();
            boolean capitalizeNext = true;
            for (char c : s.toCharArray()) {
                if (Character.isLetter(c)) {
                    sb.append(capitalizeNext ? Character.toUpperCase(c) : Character.toLowerCase(c));
                    capitalizeNext = false;
                } else {
                    sb.append(c);
                    capitalizeNext = true;
                }
            }
            return sb.toString();
        }

        /** Trimmed, upper-cased, alphabetically sorted terms - order-independent, case-independent. */
        private static String normalize(String list) {
            String[] parts = list.split(";", -1);
            List<String> cleaned = new ArrayList<>();
            for (String p : parts) {
                String t = p.trim().toUpperCase();
                if (!t.isEmpty()) cleaned.add(t);
            }
            Collections.sort(cleaned);
            return String.join("_", cleaned);
        }

        /** Human/ID-friendly identifier, capped in length via a stable checksum for long term lists. */
        private static String sanitizeId(String base, String rawKey) {
            String cleanBase = sanitize(base);
            if (rawKey.length() <= 60) {
                return sanitize(rawKey);
            }
            CRC32 crc = new CRC32();
            crc.update(rawKey.getBytes());
            return cleanBase + "_" + Long.toHexString(crc.getValue());
        }
    }

    static class CodelistInfo {
        final String id;
        final String name;
        final boolean isUnitCodelist;
        final LinkedHashMap<String, String> terms = new LinkedHashMap<>(); // term (upper) -> decoded value
        String recommendedCodelist = ""; // CDISC/NCI codelist id this subset was drawn from, if known

        CodelistInfo(String id, String name, boolean isUnitCodelist) {
            this.id = id;
            this.name = name;
            this.isUnitCodelist = isUnitCodelist;
        }

        void addTerm(String term, String decodedValue) {
            String key = term; // keep original casing for display; dedupe key below
            terms.merge(key, decodedValue, (existing, next) -> existing.isEmpty() ? next : existing);
        }

        void noteRecommendedCodelist(String nciCodelistId) {
            if (recommendedCodelist.isEmpty() && nciCodelistId != null && !nciCodelistId.isEmpty()) {
                recommendedCodelist = nciCodelistId;
            }
        }
    }

    static class FormInfo {
        final String id, name, sdtmTargetDomain;
        final int order;
        FormInfo(String id, String name, int order) {
            this.id = id;
            this.name = name;
            this.sdtmTargetDomain = name;
            this.order = order;
        }
    }

    static class SectionInfo {
        final String id, name, form;
        final int order;
        final String bcId, vlmGroupId;
        SectionInfo(String id, String name, String form, int order, String bcId, String vlmGroupId) {
            this.id = id;
            this.name = name;
            this.form = form;
            this.order = order;
            this.bcId = bcId;
            this.vlmGroupId = vlmGroupId;
        }
    }

    // =========================================================================
    // Small helpers
    // =========================================================================

    private static String sanitize(String s) {
        return s.replaceAll("[^A-Za-z0-9_]+", "_").replaceAll("_+", "_");
    }

    private static String yesNo(String yOrN) {
        if ("Y".equalsIgnoreCase(yOrN)) return "Yes";
        if ("N".equalsIgnoreCase(yOrN)) return "No";
        return "";
    }

    private static String mapDataType(String dataType) {
        return "decimal".equalsIgnoreCase(dataType) ? "float" : dataType;
    }

    private static int parseIntOrDefault(String s, int fallback) {
        try {
            return Integer.parseInt(s.trim());
        } catch (Exception e) {
            return fallback;
        }
    }

    private static List<String> headersOf(List<Map<String, String>> rows) {
        if (rows.isEmpty()) return Collections.emptyList();
        return new ArrayList<>(rows.get(0).keySet());
    }

    private static List<List<String>> rowsOf(List<Map<String, String>> rows) {
        List<List<String>> out = new ArrayList<>();
        for (Map<String, String> r : rows) {
            out.add(new ArrayList<>(r.values()));
        }
        return out;
    }
}
