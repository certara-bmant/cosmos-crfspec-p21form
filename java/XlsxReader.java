import java.io.*;
import java.util.*;
import java.util.zip.*;
import javax.xml.parsers.*;
import org.w3c.dom.*;

/**
 * Minimal, dependency-free reader for .xlsx workbooks.
 *
 * Only handles what we actually need for this project: a worksheet made up of
 * <row>/<c> elements where cell values are stored either as inline strings
 * (t="inlineStr", the format Excel/SAS commonly writes) or as bare numbers
 * (t="n" or no "t" attribute). Shared-string tables and rich formatting are
 * intentionally not supported - if you feed it a workbook that uses them for
 * the sheet you care about, extend readCellValue() below.
 *
 * No external libraries (e.g. Apache POI) are required - everything here is
 * built on java.util.zip and the JDK's built-in DOM parser, since a .xlsx
 * file is just a zip archive of XML parts.
 */
public class XlsxReader {

    /** One worksheet, already split into a header row and data rows keyed by header name. */
    public static class Sheet {
        public final List<String> headers;
        public final List<LinkedHashMap<String, String>> rows;
        Sheet(List<String> headers, List<LinkedHashMap<String, String>> rows) {
            this.headers = headers;
            this.rows = rows;
        }
    }

    /** Reads a single named sheet from an .xlsx file, using row 1 as the header row. */
    public static Sheet readSheet(File xlsxFile, String sheetName) throws Exception {
        try (ZipFile zip = new ZipFile(xlsxFile)) {
            String worksheetPart = resolveWorksheetPart(zip, sheetName);
            Document doc = parseXml(zip, worksheetPart);

            List<List<String>> rawRows = new ArrayList<>();
            NodeList rowNodes = doc.getElementsByTagName("row");
            for (int r = 0; r < rowNodes.getLength(); r++) {
                Element rowEl = (Element) rowNodes.item(r);
                rawRows.add(readRow(rowEl));
            }
            if (rawRows.isEmpty()) {
                return new Sheet(Collections.emptyList(), Collections.emptyList());
            }

            List<String> headers = new ArrayList<>();
            for (String h : rawRows.get(0)) {
                headers.add(h == null ? "" : h.trim());
            }

            List<LinkedHashMap<String, String>> rows = new ArrayList<>();
            for (int r = 1; r < rawRows.size(); r++) {
                List<String> raw = rawRows.get(r);
                LinkedHashMap<String, String> row = new LinkedHashMap<>();
                for (int c = 0; c < headers.size(); c++) {
                    String v = c < raw.size() ? raw.get(c) : "";
                    row.put(headers.get(c), v == null ? "" : v);
                }
                rows.add(row);
            }
            return new Sheet(headers, rows);
        }
    }

    // ---- internals -------------------------------------------------------

    /** Maps a sheet name to its worksheetN.xml zip entry via workbook.xml + workbook.xml.rels. */
    private static String resolveWorksheetPart(ZipFile zip, String sheetName) throws Exception {
        Document wb = parseXml(zip, "xl/workbook.xml");
        String targetRid = null;
        NodeList sheets = wb.getElementsByTagName("sheet");
        for (int i = 0; i < sheets.getLength(); i++) {
            Element sh = (Element) sheets.item(i);
            if (sheetName.equals(sh.getAttribute("name"))) {
                targetRid = sh.getAttribute("r:id");
                break;
            }
        }
        if (targetRid == null) {
            throw new IOException("Sheet not found: " + sheetName);
        }
        Document rels = parseXml(zip, "xl/_rels/workbook.xml.rels");
        NodeList relNodes = rels.getElementsByTagName("Relationship");
        for (int i = 0; i < relNodes.getLength(); i++) {
            Element rel = (Element) relNodes.item(i);
            if (targetRid.equals(rel.getAttribute("Id"))) {
                String target = rel.getAttribute("Target");
                target = target.replaceFirst("^/?xl/", "").replaceFirst("^/", "");
                return "xl/" + target;
            }
        }
        throw new IOException("Relationship not found for sheet: " + sheetName);
    }

    private static Document parseXml(ZipFile zip, String entryName) throws Exception {
        ZipEntry entry = zip.getEntry(entryName);
        if (entry == null) {
            throw new IOException("Zip entry not found: " + entryName);
        }
        DocumentBuilderFactory f = DocumentBuilderFactory.newInstance();
        f.setNamespaceAware(false);
        DocumentBuilder builder = f.newDocumentBuilder();
        try (InputStream in = zip.getInputStream(entry)) {
            return builder.parse(in);
        }
    }

    private static List<String> readRow(Element rowEl) {
        NodeList cellNodes = rowEl.getElementsByTagName("c");
        // Cells with no value are simply omitted by Excel, so track by column index.
        Map<Integer, String> byCol = new TreeMap<>();
        int maxCol = -1;
        for (int i = 0; i < cellNodes.getLength(); i++) {
            Element cell = (Element) cellNodes.item(i);
            String ref = cell.getAttribute("r"); // e.g. "AB12"
            int col = columnIndexFromRef(ref);
            String value = readCellValue(cell);
            byCol.put(col, value);
            maxCol = Math.max(maxCol, col);
        }
        List<String> result = new ArrayList<>();
        for (int c = 0; c <= maxCol; c++) {
            result.add(byCol.getOrDefault(c, ""));
        }
        return result;
    }

    private static String readCellValue(Element cell) {
        String type = cell.getAttribute("t");
        if ("inlineStr".equals(type)) {
            NodeList tNodes = cell.getElementsByTagName("t");
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < tNodes.getLength(); i++) {
                sb.append(tNodes.item(i).getTextContent());
            }
            return sb.toString();
        }
        // Numeric (t="n") or untyped cells: value lives in <v>.
        NodeList vNodes = cell.getElementsByTagName("v");
        if (vNodes.getLength() == 0) {
            return "";
        }
        String raw = vNodes.item(0).getTextContent();
        if (raw == null || raw.isEmpty()) {
            return "";
        }
        return normalizeNumber(raw);
    }

    /** Renders "3.0" as "3" so numeric CRF fields (length, order, etc.) come out clean. */
    private static String normalizeNumber(String raw) {
        try {
            double d = Double.parseDouble(raw);
            if (d == Math.rint(d) && !Double.isInfinite(d)) {
                return String.valueOf((long) d);
            }
            return raw;
        } catch (NumberFormatException e) {
            return raw;
        }
    }

    private static int columnIndexFromRef(String ref) {
        int i = 0;
        int col = 0;
        while (i < ref.length() && Character.isLetter(ref.charAt(i))) {
            col = col * 26 + (Character.toUpperCase(ref.charAt(i)) - 'A' + 1);
            i++;
        }
        return col - 1; // zero-based
    }
}
