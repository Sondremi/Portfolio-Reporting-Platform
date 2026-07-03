package report;

import java.io.FileWriter;
import java.io.IOException;

final class ReportTemplateHelper {

    private ReportTemplateHelper() {}

    static void writeHtmlRow(FileWriter writer, String... cells) throws IOException {
        writer.write("<tr>\n");
        for (String cell : cells) {
            writer.write("    <th>" + (cell != null ? cell : "") + "</th>\n");
        }
        writer.write("</tr>\n");
    }

    static void writeHtmlRowWithClassAndAttributes(
            FileWriter writer,
            String rowClass,
            String rowAttributes,
            String... cells) throws IOException {
        String classAttribute = (rowClass == null || rowClass.isBlank()) ? "" : " class=\"" + escapeHtml(rowClass) + "\"";
        String attributeSuffix = (rowAttributes == null || rowAttributes.isBlank()) ? "" : " " + rowAttributes.trim();
        writer.write("<tr" + classAttribute + attributeSuffix + ">\n");
        for (String cell : cells) {
            writer.write("    <td>" + (cell != null ? cell : "") + "</td>\n");
        }
        writer.write("</tr>\n");
    }

    public static String escapeHtml(String text) {
        if (text == null) {
            return "";
        }
        return text
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace("\"", "&quot;");
    }
}
