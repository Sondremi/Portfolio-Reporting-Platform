package report;

import java.io.FileWriter;
import java.io.IOException;

final class ReportStyleHelper {

    private ReportStyleHelper() {}

    static void writeBaseThemeStyles(FileWriter writer) throws IOException {
        writer.write("        :root { --bg:#eef3f7; --line:#d8e0e9; --card:#ffffff; --ink:#16202a; --muted:#5a6877; --good:#1f8b4d; --bad:#b23a31; --spark-text:#5c7187; --spark-axis:#9eb1c3; --spark-axis-soft:#b8c7d6; --spark-grid:#cfdbe6; --spark-line:#223c55; --spark-point:#223c55; }\n");
        writer.write("        body.theme-dark { --bg:#0f1722; --line:#253245; --card:#162231; --ink:#e5edf7; --muted:#aebdce; --good:#59c887; --bad:#f07f7f; --spark-text:#d5e1ef; --spark-axis:#7f95ab; --spark-axis-soft:#9ab0c6; --spark-grid:#8ea4ba; --spark-line:#edf4fc; --spark-point:#edf4fc; }\n");
        writer.write("        * { box-sizing: border-box; }\n");
        writer.write("        body { font-family: 'Segoe UI','Avenir Next','Helvetica Neue',Arial,sans-serif; margin:0; background: radial-gradient(circle at top,#f8fbfe 0%,var(--bg) 58%); color:var(--ink); overflow-x:hidden; }\n");
        writer.write("        body.theme-dark { background: radial-gradient(circle at top,#1c2b3f 0%, var(--bg) 62%); }\n");
    }
}
