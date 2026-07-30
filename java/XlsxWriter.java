import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.*;
import java.util.zip.*;

/**
 * Minimal, dependency-free writer for .xlsx workbooks.
 *
 * Produces a single valid OOXML workbook containing one or more sheets, using
 * inline strings for every cell (no shared-string table needed). Every cell
 * is written as text; this keeps the writer simple and is perfectly fine for
 * feeding Excel or Pinnacle 21's import, at the cost of numbers not being
 * "real" numeric cells (cosmetic only - sorting/filtering on numeric columns
 * in Excel will treat them as text unless the user reformats the column).
 *
 * No external libraries (e.g. Apache POI) are required.
 */
public class XlsxWriter {

    private final LinkedHashMap<String, SheetData> sheets = new LinkedHashMap<>();

    private static class SheetData {
        final List<String> headers;
        final List<List<String>> rows;
        SheetData(List<String> headers, List<List<String>> rows) {
            this.headers = headers;
            this.rows = rows;
        }
    }

    /** Adds a sheet. Row lists are matched to headers by position (missing cells -> blank). */
    public XlsxWriter addSheet(String sheetName, List<String> headers, List<List<String>> rows) {
        if (sheetName.length() > 31) {
            sheetName = sheetName.substring(0, 31); // Excel's hard sheet-name limit
        }
        sheets.put(sheetName, new SheetData(headers, rows));
        return this;
    }

    public void save(File file) throws IOException {
        try (ZipOutputStream zip = new ZipOutputStream(new FileOutputStream(file))) {
            writeEntry(zip, "[Content_Types].xml", contentTypesXml());
            writeEntry(zip, "_rels/.rels", rootRelsXml());
            writeEntry(zip, "xl/workbook.xml", workbookXml());
            writeEntry(zip, "xl/_rels/workbook.xml.rels", workbookRelsXml());
            writeEntry(zip, "xl/styles.xml", stylesXml());

            int i = 1;
            for (SheetData sheet : sheets.values()) {
                writeEntry(zip, "xl/worksheets/sheet" + i + ".xml", worksheetXml(sheet));
                i++;
            }
        }
    }

    // ---- OOXML parts -------------------------------------------------------

    private String contentTypesXml() {
        StringBuilder sb = new StringBuilder();
        sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
        sb.append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
        sb.append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
        sb.append("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
        sb.append("<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
        int i = 1;
        for (int n = 0; n < sheets.size(); n++) {
            sb.append("<Override PartName=\"/xl/worksheets/sheet").append(i).append(".xml\" ")
              .append("ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
            i++;
        }
        sb.append("</Types>");
        return sb.toString();
    }

    private String rootRelsXml() {
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
            + "</Relationships>";
    }

    private String workbookXml() {
        StringBuilder sb = new StringBuilder();
        sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.append("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ")
          .append("xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
        sb.append("<sheets>");
        int i = 1;
        for (String name : sheets.keySet()) {
            sb.append("<sheet name=\"").append(escapeXml(name)).append("\" sheetId=\"").append(i)
              .append("\" r:id=\"rId").append(i + 1).append("\"/>");
            i++;
        }
        sb.append("</sheets></workbook>");
        return sb.toString();
    }

    private String workbookRelsXml() {
        StringBuilder sb = new StringBuilder();
        sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
        sb.append("<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");
        int i = 1;
        for (int n = 0; n < sheets.size(); n++) {
            sb.append("<Relationship Id=\"rId").append(i + 1).append("\" ")
              .append("Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" ")
              .append("Target=\"worksheets/sheet").append(i).append(".xml\"/>");
            i++;
        }
        sb.append("</Relationships>");
        return sb.toString();
    }

    /** The bare minimum styles.xml Excel requires to open the file without a repair prompt. */
    private String stylesXml() {
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
            + "<fonts count=\"1\"><font><sz val=\"10\"/><name val=\"Calibri\"/></font></fonts>"
            + "<fills count=\"1\"><fill><patternFill patternType=\"none\"/></fill></fills>"
            + "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>"
            + "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>"
            + "<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs>"
            + "</styleSheet>";
    }

    private String worksheetXml(SheetData sheet) {
        StringBuilder sb = new StringBuilder();
        sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>");

        sb.append(rowXml(1, sheet.headers));
        int r = 2;
        for (List<String> row : sheet.rows) {
            sb.append(rowXml(r, row));
            r++;
        }
        sb.append("</sheetData></worksheet>");
        return sb.toString();
    }

    private String rowXml(int rowNum, List<String> values) {
        StringBuilder sb = new StringBuilder();
        sb.append("<row r=\"").append(rowNum).append("\">");
        for (int c = 0; c < values.size(); c++) {
            String v = values.get(c);
            if (v == null) v = "";
            String ref = columnLetter(c) + rowNum;
            if (v.isEmpty()) {
                sb.append("<c r=\"").append(ref).append("\"/>");
            } else {
                sb.append("<c r=\"").append(ref).append("\" t=\"inlineStr\"><is><t xml:space=\"preserve\">")
                  .append(escapeXml(v)).append("</t></is></c>");
            }
        }
        sb.append("</row>");
        return sb.toString();
    }

    private static String columnLetter(int zeroBasedCol) {
        int col = zeroBasedCol + 1;
        StringBuilder sb = new StringBuilder();
        while (col > 0) {
            int rem = (col - 1) % 26;
            sb.insert(0, (char) ('A' + rem));
            col = (col - 1) / 26;
        }
        return sb.toString();
    }

    private static String escapeXml(String s) {
        StringBuilder sb = new StringBuilder(s.length());
        for (int i = 0; i < s.length(); i++) {
            char ch = s.charAt(i);
            switch (ch) {
                case '&': sb.append("&amp;"); break;
                case '<': sb.append("&lt;"); break;
                case '>': sb.append("&gt;"); break;
                case '"': sb.append("&quot;"); break;
                case '\'': sb.append("&apos;"); break;
                default:
                    // Strip control characters XML can't legally contain.
                    if (ch >= 0x20 || ch == '\t' || ch == '\n' || ch == '\r') {
                        sb.append(ch);
                    }
            }
        }
        return sb.toString();
    }

    private void writeEntry(ZipOutputStream zip, String name, String content) throws IOException {
        ZipEntry entry = new ZipEntry(name);
        zip.putNextEntry(entry);
        zip.write(content.getBytes(StandardCharsets.UTF_8));
        zip.closeEntry();
    }
}
