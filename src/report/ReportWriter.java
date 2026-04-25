package report;

import csv.TransactionStore;
import model.Events;
import model.Security;

import java.io.FileWriter;
import java.io.IOException;
import java.time.Instant;
import java.time.LocalDate;
import java.time.ZoneOffset;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.NavigableMap;
import java.util.Set;

public class ReportWriter {

    private static final DateTimeFormatter DETAIL_DATE_FORMATTER = DateTimeFormatter.ofPattern("dd.MM.yyyy");
    private static final String DEFAULT_TOTAL_CURRENCY = "NOK";
    private static final String REPORT_TYPE_STANDARD = "standard";
    private static final String REPORT_TYPE_ANNUAL = "annual";
    private static final double EPSILON = 0.0000001;

    private static final class ReportConfig {
        private final String reportType;
        private final int reportYear;
        private final String benchmarkTicker;

        private ReportConfig(String reportType, int reportYear, String benchmarkTicker) {
            this.reportType = reportType;
            this.reportYear = reportYear;
            this.benchmarkTicker = benchmarkTicker;
        }
    }

    public static void writeHtmlReport(TransactionStore store, String outputFile) throws IOException {
        List<OverviewRow> overviewRows = PortfolioCalculator.buildOverviewRows(store);
        Map<String, Double> ratesToNok = CurrencyConversionService.loadRatesToNok(collectCurrencies(store, overviewRows));
        HeaderSummary headerSummary = PortfolioCalculator.buildHeaderSummary(store, overviewRows, ratesToNok);
        ReportConfig reportConfig = resolveReportConfig();
        AnnualPerformanceSummary annualSummary = REPORT_TYPE_ANNUAL.equals(reportConfig.reportType)
            ? PortfolioCalculator.buildAnnualPerformanceSummary(store, ratesToNok, reportConfig.reportYear, reportConfig.benchmarkTicker)
            : null;
        List<AnnualSnapshotRow> annualSnapshotRows = new ArrayList<>();
        if (REPORT_TYPE_ANNUAL.equals(reportConfig.reportType)) {
            int snapshotYear = Math.max(2000, Math.min(2100, reportConfig.reportYear));
            annualSnapshotRows = buildAnnualSnapshotRows(store, LocalDate.of(snapshotYear, 12, 31));
        }

        try (FileWriter writer = new FileWriter(outputFile)) {
            writer.write("<!DOCTYPE html>\n");
            writer.write("<html lang=\"no\">\n");
            writer.write("<head>\n");
            writer.write("    <meta charset=\"UTF-8\">\n");
            writer.write("    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">\n");
            writer.write("    <title>Portfolio Report</title>\n");
            writer.write("    <style>\n");
            ReportStyleHelper.writeBaseThemeStyles(writer);
            writer.write("        .page { width:100%; max-width:100%; margin:0; padding:24px 8px 32px; }\n");
            writer.write("        h2 { margin:26px 2px 12px; font-size:1.14rem; color:var(--ink); }\n");
            writer.write("        table { width:100%; border-collapse:collapse; min-width:0; table-layout:fixed; background:var(--card); }\n");
            writer.write("        th, td { padding:5px 5px; border-bottom:1px solid #edf2f7; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }\n");
            writer.write("        th { background:#f5f8fb; text-align:left; font-size:.72rem; text-transform:uppercase; letter-spacing:.2px; color:#374556; border-bottom:1px solid var(--line); }\n");
            writer.write("        th.sortable-header { cursor:pointer; user-select:none; position:relative; padding-right:16px; }\n");
            writer.write("        th.sortable-header::after { content:'↕'; position:absolute; right:5px; top:50%; transform:translateY(-50%); font-size:.62rem; opacity:.45; }\n");
            writer.write("        th.sortable-header.sort-asc::after { content:'▲'; opacity:.85; }\n");
            writer.write("        th.sortable-header.sort-desc::after { content:'▼'; opacity:.85; }\n");
            writer.write("        td { font-size:.72rem; }\n");
            writer.write("        td.num, th.num { text-align:right; }\n");
            writer.write("        .table-wrap { background:var(--card); border:1px solid var(--line); border-radius:14px; overflow-x:auto; overflow-y:hidden; -webkit-overflow-scrolling:touch; scrollbar-gutter:auto; box-shadow:0 5px 14px rgba(15,23,33,.06); }\n");
            writer.write("        .table-wrap::-webkit-scrollbar { height:12px; }\n");
            writer.write("        .table-wrap::-webkit-scrollbar-track { background:transparent; border-radius:999px; }\n");
            writer.write("        .table-wrap::-webkit-scrollbar-thumb { background:#9db0c3; border-radius:999px; border:2px solid transparent; background-clip:padding-box; }\n");
            writer.write("        .overview-mode-shell { display:flex; align-items:flex-end; gap:0; margin:10px 0 -1px; }\n");
            writer.write("        .overview-mode-btn { border:1px solid var(--line); border-bottom:none; background:#e9f2fb; color:#24435a; font-size:.74rem; font-weight:700; padding:6px 10px; cursor:pointer; }\n");
            writer.write("        .overview-mode-btn + .overview-mode-btn { border-left:none; }\n");
            writer.write("        .overview-mode-btn:first-child { border-top-left-radius:10px; }\n");
            writer.write("        .overview-mode-btn:last-child { border-top-right-radius:10px; }\n");
            writer.write("        .overview-mode-btn.is-active { background:var(--card); color:var(--ink); }\n");
            writer.write("        .overview-details-toggle-btn { margin-left:8px; border-left:1px solid var(--line) !important; background:#eef4fb; }\n");
            writer.write("        .overview-details-toggle-btn:disabled { opacity:.55; cursor:not-allowed; }\n");
            writer.write("        body.theme-dark .overview-mode-btn { background:#1f3347; color:#d8e7f5; border-color:#2f445a; }\n");
            writer.write("        body.theme-dark .overview-mode-btn.is-active { background:#162231; color:#edf5ff; }\n");
            writer.write("        body.theme-dark .overview-details-toggle-btn { border-left-color:#2f445a !important; background:#22384e; }\n");
            writer.write("        .report-standard .overview-table-wrap { width:100%; max-width:100%; overflow-x:auto; }\n");
            writer.write("        .report-standard .overview-table { table-layout:fixed; width:100%; }\n");
            writer.write("        .report-standard .overview-table th, .report-standard .overview-table td { white-space:nowrap; overflow:hidden; text-overflow:ellipsis; font-size:.67rem; padding:4px 4px; }\n");
            writer.write("        .report-standard .overview-table tr > * { min-width:0; }\n");
            writer.write("        .report-standard .overview-summary-table tr > *:nth-child(1)  { width:5%; min-width:80px; max-width:100px; overflow:hidden !important; text-overflow:ellipsis !important; }\n");
            writer.write("        .report-standard .overview-summary-table tr > *:nth-child(2)  { width:11%; min-width:188px; max-width:280px; overflow:hidden !important; text-overflow:ellipsis !important; }\n");
            writer.write("        .report-standard .overview-summary-table tr > *:nth-child(3)  { width:7%; }\n");
            writer.write("        .report-standard .overview-summary-table tr > *:nth-child(4)  { width:7%; }\n");
            writer.write("        .report-standard .overview-summary-table tr > *:nth-child(5)  { width:9%; }\n");
            writer.write("        .report-standard .overview-summary-table tr > *:nth-child(6)  { width:13%; }\n");
            writer.write("        .report-standard .overview-summary-table tr > *:nth-child(7)  { width:7%; }\n");
            writer.write("        .report-standard .overview-summary-table tr > *:nth-child(8)  { width:8%; }\n");
            writer.write("        .report-standard .overview-summary-table tr > *:nth-child(9)  { width:8%; }\n");
            writer.write("        .report-standard .overview-summary-table tr > *:nth-child(10) { width:9%; }\n");
            writer.write("        .report-standard .overview-summary-table tr > *:nth-child(11) { width:10%; }\n");
            writer.write("        .report-standard .overview-holdings-table tr > *:nth-child(1)  { width:1%; min-width:40px; max-width:70px; overflow:hidden !important; text-overflow:ellipsis !important; }\n");
            writer.write("        .report-standard .overview-holdings-table tr > *:nth-child(2)  { width:1.5%; min-width:50px; max-width:130px; overflow:hidden !important; text-overflow:ellipsis !important; }\n");
            writer.write("        .report-standard .overview-holdings-table tr > *:nth-child(3), .report-standard .overview-holdings-table tr > *:nth-child(4), .report-standard .overview-holdings-table tr > *:nth-child(5), .report-standard .overview-holdings-table tr > *:nth-child(6), .report-standard .overview-holdings-table tr > *:nth-child(7), .report-standard .overview-holdings-table tr > *:nth-child(8), .report-standard .overview-holdings-table tr > *:nth-child(9), .report-standard .overview-holdings-table tr > *:nth-child(10), .report-standard .overview-holdings-table tr > *:nth-child(11), .report-standard .overview-holdings-table tr > *:nth-child(12), .report-standard .overview-holdings-table tr > *:nth-child(13) { width:auto; min-width:max-content; max-width:none; }\n");
            writer.write("        .report-standard .overview-table tr > *:nth-child(n+3) { overflow:hidden !important; text-overflow:ellipsis !important; }\n");
            writer.write("        .report-standard .overview-holdings-table th, .report-standard .overview-holdings-table td { white-space:nowrap; overflow:visible !important; text-overflow:clip !important; }\n");
            writer.write("        .report-standard .overview-holdings-table tr > *:nth-child(1), .report-standard .overview-holdings-table tr > *:nth-child(2) { overflow:hidden !important; text-overflow:ellipsis !important; }\n");
            writer.write("        .report-standard .overview-holdings-table tr > *:nth-child(n+3) { overflow:visible !important; text-overflow:clip !important; }\n");
            writer.write("        .report-standard .overview-holdings-table { table-layout:auto; width:max-content; min-width:100%; }\n");
            writer.write("        .report-standard .overview-holdings-table tr.total-row td { overflow:visible !important; text-overflow:clip !important; white-space:nowrap !important; font-variant-numeric:tabular-nums; padding-right:10px; }\n");
            writer.write("        .report-standard .overview-fundamentals-table { table-layout:auto; width:max-content; min-width:100%; }\n");
            writer.write("        .report-standard .overview-fundamentals-table tr > *:nth-child(1)  { width:1%; min-width:40px; max-width:70px; overflow:hidden !important; text-overflow:ellipsis !important; }\n");
            writer.write("        .report-standard .overview-fundamentals-table tr > *:nth-child(2)  { width:1.5%; min-width:50px; max-width:130px; overflow:hidden !important; text-overflow:ellipsis !important; }\n");
            writer.write("        .report-standard .overview-fundamentals-table tr > *:nth-child(n+3) { width:auto; min-width:max-content; max-width:none; }\n");
            writer.write("        .report-standard .overview-fundamentals-table th, .report-standard .overview-fundamentals-table td { white-space:nowrap; overflow:visible !important; text-overflow:clip !important; }\n");
            writer.write("        .report-standard .overview-fundamentals-table tr > *:nth-child(1), .report-standard .overview-fundamentals-table tr > *:nth-child(2) { overflow:hidden !important; text-overflow:ellipsis !important; }\n");
            writer.write("        .report-standard .overview-fundamentals-table tr > *:nth-child(n+3) { overflow:visible !important; text-overflow:clip !important; }\n");
            writer.write("        .wk-range-cell { min-width:118px; }\n");
            writer.write("        .wk-range-track { position:relative; height:4px; border-radius:999px; background:#c7d3df; margin:0 2px 6px; }\n");
            writer.write("        .wk-range-marker { position:absolute; top:50%; width:10px; height:10px; border-radius:50%; background:#2b67bc; transform:translate(-50%, -50%); box-shadow:0 0 0 1px rgba(255,255,255,.85); }\n");
            writer.write("        .wk-range-labels { display:flex; justify-content:space-between; gap:6px; font-size:.68rem; color:#2f3f4f; }\n");
            writer.write("        body.theme-dark .wk-range-track { background:#4a5d72; }\n");
            writer.write("        body.theme-dark .wk-range-labels { color:#d5e3f1; }\n");
            writer.write("        .mini-day-chart { display:block; width:96px; height:30px; overflow:visible; }\n");
            writer.write("        .mini-day-chart-area { stroke:none; opacity:.28; }\n");
            writer.write("        .mini-day-chart-area.positive { fill:#2f9e62; }\n");
            writer.write("        .mini-day-chart-area.negative { fill:#c4514a; }\n");
            writer.write("        .mini-day-chart-line { fill:none; stroke:#2e5f88; stroke-width:1.8; stroke-linecap:round; stroke-linejoin:round; }\n");
            writer.write("        .mini-day-chart-line.positive { stroke:#1f8b4d; }\n");
            writer.write("        .mini-day-chart-line.negative { stroke:#b23a31; }\n");
            writer.write("        .mini-day-chart-open { stroke:#8aa0b5; stroke-width:1; stroke-dasharray:2.5 2.5; opacity:.8; }\n");
            writer.write("        .mini-day-chart-end { stroke:#ffffff; stroke-width:1.1; }\n");
            writer.write("        .mini-day-chart-end.positive { fill:#1f8b4d; }\n");
            writer.write("        .mini-day-chart-end.negative { fill:#b23a31; }\n");
            writer.write("        body.theme-dark .mini-day-chart-open { stroke:#6f879f; opacity:.9; }\n");
            writer.write("        body.theme-dark .mini-day-chart-area.positive { fill:#2a8f57; }\n");
            writer.write("        body.theme-dark .mini-day-chart-area.negative { fill:#a54640; }\n");
            writer.write("        .report-standard .ticker-scroll, .report-standard .security-scroll { display:block; position:relative; width:100%; max-width:100%; overflow-x:auto; overflow-y:hidden; white-space:nowrap; text-overflow:clip; scrollbar-width:none; -ms-overflow-style:none; padding-bottom:6px; cursor:grab; }\n");
            writer.write("        .report-standard .ticker-scroll { max-width:126px; }\n");
            writer.write("        .report-standard .security-scroll { max-width:236px; }\n");
            writer.write("        .report-standard .ticker-scroll::-webkit-scrollbar, .report-standard .security-scroll::-webkit-scrollbar { display:none; width:0; height:0; }\n");
            writer.write("        .report-standard .ticker-scroll::after, .report-standard .security-scroll::after { content:''; position:absolute; left:5px; right:5px; bottom:1px; height:4px; border-radius:999px; background:rgba(140,160,178,.18); opacity:.28; transition:opacity .12s ease, background .12s ease; }\n");
            writer.write("        .report-standard .ticker-scroll:hover::after, .report-standard .security-scroll:hover::after { opacity:.5; background:rgba(140,160,178,.28); }\n");
            writer.write("        @media (max-width:1060px) { .report-standard .overview-table tr > *:nth-child(n+3), .report-standard .realized-table tr > *:nth-child(n+3) { overflow:hidden !important; text-overflow:ellipsis !important; white-space:nowrap !important; } }\n");
            writer.write("        @media (max-width:1060px) { .report-standard .overview-holdings-table tr > *:nth-child(n+3) { overflow:visible !important; text-overflow:clip !important; white-space:nowrap !important; } }\n");
            writer.write("        @media (max-width:1060px) { .report-standard .overview-fundamentals-table tr > *:nth-child(n+3) { overflow:visible !important; text-overflow:clip !important; white-space:nowrap !important; } }\n");
            writer.write("        .report-annual .realized-table { table-layout:auto; }\n");
            writer.write("        .report-annual .realized-table tr > *:nth-child(1) { width:106px; max-width:106px; min-width:106px; overflow:visible; text-overflow:clip; }\n");
            writer.write("        .report-annual .realized-table tr > *:nth-child(2) { width:auto; min-width:9ch; max-width:none; overflow:visible; text-overflow:clip; }\n");
            writer.write("        .report-annual .realized-table tr > *:nth-child(3) { width:auto; min-width:14ch; max-width:none; overflow:visible; text-overflow:clip; }\n");
            writer.write("        .report-annual .realized-table tr > *:nth-child(7) { width:160px; max-width:160px; }\n");
            writer.write("        .report-standard .realized-table { table-layout:auto; width:100%; }\n");
            writer.write("        .report-standard .realized-table th, .report-standard .realized-table td { white-space:nowrap; overflow:visible; text-overflow:clip; }\n");
            writer.write("        .report-standard .realized-table tr > *:nth-child(1)  { width:auto; min-width:108px; max-width:108px; overflow:hidden !important; text-overflow:ellipsis !important; }\n");
            writer.write("        .report-standard .realized-table tr > *:nth-child(2)  { width:auto; max-width:250px; overflow:hidden !important; text-overflow:ellipsis !important; }\n");
            writer.write("        .ticker-scroll { display:block; width:100%; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; padding-bottom:0; }\n");
            writer.write("        .security-scroll { display:block; position:relative; width:100%; max-width:100%; overflow-x:auto; overflow-y:hidden; white-space:nowrap; text-overflow:clip; scrollbar-width:none; -ms-overflow-style:none; padding-bottom:6px; cursor:grab; }\n");
            writer.write("        .security-scroll::-webkit-scrollbar { display:none; width:0; height:0; }\n");
            writer.write("        .security-scroll::after { content:''; position:absolute; left:5px; right:5px; bottom:1px; height:4px; border-radius:999px; background:rgba(140,160,178,.18); opacity:.28; transition:opacity .12s ease, background .12s ease; }\n");
            writer.write("        .security-scroll:hover::after { opacity:.5; background:rgba(140,160,178,.28); }\n");
            writer.write("        .security-scroll.is-dragging::after { opacity:.85; background:rgba(120,145,168,.45); }\n");
            writer.write("        .security-scroll.is-dragging { cursor:grabbing; }\n");
            writer.write("        body.inline-cell-dragging { user-select:none; cursor:grabbing; }\n");
            writer.write("        .total-row { font-weight:700; background:#f3f7fb; color:#1a2b3a; }\n");
            writer.write("        .asset-split td { border-top:3px solid #8a9eb3 !important; }\n");
            writer.write("        .positive { color:var(--good); } .negative { color:var(--bad); }\n");
            writer.write("        .report-hero { display:grid; grid-template-columns:1.25fr 1fr; gap:16px; background:linear-gradient(120deg,#0f2238 0%,#18344f 60%,#164663 100%); border-radius:18px; padding:22px; color:#f4f8fc; box-shadow:0 14px 26px rgba(10,24,38,.2); margin-bottom:18px; }\n");
            writer.write("        .annual-hero { grid-template-columns:1fr; gap:14px; margin-bottom:12px; }\n");
            writer.write("        .annual-hero-header { display:flex; flex-direction:column; gap:2px; }\n");
            writer.write("        .hero-title h1 { margin:0; font-size:1.75rem; letter-spacing:.4px; }\n");
            writer.write("        .hero-meta { margin-top:10px; display:flex; flex-wrap:wrap; gap:8px; }\n");
            writer.write("        .meta-chip { display:inline-flex; align-items:center; gap:6px; padding:6px 11px; border-radius:999px; border:1px solid rgba(235,245,255,.28); background:rgba(255,255,255,.1); color:#d7e6f4; font-size:.84rem; font-weight:600; }\n");
            writer.write("        .meta-chip strong { color:#ffffff; font-weight:700; }\n");
            writer.write("        .currency-input { width:56px; border:1px solid rgba(235,245,255,.45); border-radius:6px; background:rgba(255,255,255,.18); color:#fff; font-weight:700; text-transform:uppercase; padding:2px 6px; outline:none; }\n");
            writer.write("        .currency-input:focus { border-color:#fff; background:rgba(255,255,255,.26); }\n");
            writer.write("        .hero-kpis { margin-top:14px; display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:12px; }\n");
            writer.write("        .annual-headline-grid { margin-top:0; }\n");
            writer.write("        .kpi-card { background:rgba(255,255,255,.08); border:1px solid rgba(235,245,255,.2); border-radius:10px; padding:10px 11px; }\n");
            writer.write("        .report-standard .annual-summary-grid { grid-template-columns:repeat(8,minmax(0,1fr)); }\n");
            writer.write("        .report-standard .annual-summary-grid > .kpi-card { grid-column:span 1; }\n");
            writer.write("        .report-standard .annual-summary-grid > .kpi-card.kpi-card-wide { grid-column:span 2; }\n");
            writer.write("        .report-standard .annual-summary-grid > .kpi-card.kpi-card-bestworst { grid-column:span 2; }\n");
            writer.write("        .report-standard .kpi-card { background:linear-gradient(180deg,#f9fcff 0%,#f2f8fd 100%); border-color:#d4dfeb; color:#1f3549; }\n");
            writer.write("        .report-standard .kpi-label { color:#5b7288; }\n");
            writer.write("        .report-standard .kpi-value { color:#1f3549; }\n");
            writer.write("        .report-standard .performer { color:#314c64; }\n");
            writer.write("        .report-standard .performer strong { color:#1f3549; }\n");
            writer.write("        .report-standard .kpi-label.positive, .report-standard .kpi-value.positive { color:var(--good); }\n");
            writer.write("        .report-standard .kpi-label.negative, .report-standard .kpi-value.negative { color:var(--bad); }\n");
            writer.write("        .report-standard .performer.positive { color:var(--good); }\n");
            writer.write("        .report-standard .performer.negative { color:var(--bad); }\n");
            writer.write("        .report-standard .cash-holdings-add-btn { border-color:#8da9c4; background:#eef5fb; color:#1f3a52; }\n");
            writer.write("        .report-standard .cash-holdings-add-btn:hover { background:#e4eff9; }\n");
            writer.write("        .report-standard .manual-cash-holding-line { color:#4f6780; }\n");
            writer.write("        .report-standard .manual-cash-holding-line.is-portfolio { color:#1f3a52; }\n");
            writer.write("        body.theme-dark.report-standard .kpi-card { background:#1a2d42; border-color:#2e4258; color:#dbe8f4; }\n");
            writer.write("        body.theme-dark.report-standard .kpi-label { color:#b8cde1; }\n");
            writer.write("        body.theme-dark.report-standard .kpi-value { color:#edf5ff; }\n");
            writer.write("        body.theme-dark.report-standard .performer { color:#d4e3f2; }\n");
            writer.write("        body.theme-dark.report-standard .performer strong { color:#edf5ff; }\n");
            writer.write("        body.theme-dark.report-standard .kpi-label.positive, body.theme-dark.report-standard .kpi-value.positive { color:var(--good); }\n");
            writer.write("        body.theme-dark.report-standard .kpi-label.negative, body.theme-dark.report-standard .kpi-value.negative { color:var(--bad); }\n");
            writer.write("        body.theme-dark.report-standard .performer.positive { color:var(--good); }\n");
            writer.write("        body.theme-dark.report-standard .performer.negative { color:var(--bad); }\n");
            writer.write("        body.theme-dark.report-standard .cash-holdings-add-btn { border-color:#56799a; background:#243c55; color:#e2edf8; }\n");
            writer.write("        body.theme-dark.report-standard .cash-holdings-add-btn:hover { background:#2d4a67; }\n");
            writer.write("        body.theme-dark.report-standard .manual-cash-holding-line { color:#c2d6ea; }\n");
            writer.write("        body.theme-dark.report-standard .manual-cash-holding-line.is-portfolio { color:#edf5ff; }\n");
            writer.write("        .cash-holdings-header { display:flex; align-items:center; justify-content:space-between; gap:8px; }\n");
            writer.write("        .cash-holdings-add-btn { border:1px solid rgba(235,245,255,.45); background:rgba(255,255,255,.12); color:#f3f7fc; border-radius:999px; padding:2px 8px; font-size:.72rem; font-weight:700; cursor:pointer; display:none; }\n");
            writer.write("        .cash-holdings-add-btn:hover { background:rgba(255,255,255,.2); }\n");
            writer.write("        .manual-cash-holdings-list { margin-top:7px; display:grid; gap:3px; }\n");
            writer.write("        .manual-cash-holding-line { font-size:.75rem; color:#d3e3f3; line-height:1.3; }\n");
            writer.write("        .manual-cash-holding-line.is-portfolio { font-weight:700; color:#f4f9ff; }\n");
            writer.write("        .cash-manager-overlay[hidden] { display:none !important; }\n");
            writer.write("        .cash-manager-overlay { position:fixed; inset:0; background:rgba(6,14,24,.58); z-index:12500; display:flex; align-items:center; justify-content:center; padding:16px; }\n");
            writer.write("        .cash-manager-dialog { width:min(700px,94vw); max-height:88vh; overflow:auto; background:#f7fbff; color:#1a3348; border:1px solid #a8bfd4; border-radius:12px; box-shadow:0 18px 36px rgba(8,20,33,.34); padding:12px; }\n");
            writer.write("        .cash-manager-header { display:flex; justify-content:space-between; align-items:center; gap:8px; margin-bottom:10px; }\n");
            writer.write("        .cash-manager-header h4 { margin:0; font-size:.95rem; }\n");
            writer.write("        .cash-manager-close { border:1px solid #9cb5ca; background:#eef5fb; color:#20405a; border-radius:8px; width:28px; height:28px; cursor:pointer; font-size:1rem; line-height:1; }\n");
            writer.write("        .cash-manager-form-row { display:flex; flex-wrap:wrap; gap:6px; margin-bottom:10px; }\n");
            writer.write("        .cash-manager-form-row input { border:1px solid #a8bfd4; border-radius:8px; height:32px; padding:0 8px; font-size:.82rem; }\n");
            writer.write("        .cash-manager-form-row select { border:1px solid #a8bfd4; border-radius:8px; height:32px; padding:0 8px; font-size:.82rem; background:#fff; min-width:120px; }\n");
            writer.write("        .cash-manager-currency-input { width:76px; text-transform:uppercase; }\n");
            writer.write("        .cash-manager-btn { border:1px solid #86a4bf; background:#f3f8fd; color:#1e3951; border-radius:8px; height:32px; padding:0 10px; font-size:.78rem; font-weight:700; cursor:pointer; }\n");
            writer.write("        .cash-manager-btn:hover { background:#e6f1fb; }\n");
            writer.write("        .cash-manager-btn.danger { border-color:#cf8080; background:#fff1f1; color:#8f2d2d; }\n");
            writer.write("        .cash-manager-btn.danger:hover { background:#ffe4e4; }\n");
            writer.write("        .cash-manager-message { margin:-2px 0 8px; font-size:.76rem; color:#355572; min-height:1.1em; }\n");
            writer.write("        .cash-manager-message.is-error { color:#9a2f2f; }\n");
            writer.write("        .cash-account-block { border:1px solid #d4dfeb; border-radius:10px; padding:9px; background:#ffffff; margin-bottom:8px; }\n");
            writer.write("        .cash-account-head { display:flex; justify-content:space-between; align-items:center; gap:8px; margin-bottom:8px; }\n");
            writer.write("        .cash-account-title { margin:0; font-size:.84rem; font-weight:700; color:#1e3951; }\n");
            writer.write("        .cash-account-summary { margin:0 0 8px; font-size:.78rem; color:#355572; }\n");
            writer.write("        .cash-manager-empty { margin:0 0 8px; font-size:.78rem; color:#56708a; }\n");
            writer.write("        .cash-transaction-list { margin:0 0 8px; padding-left:16px; display:grid; gap:4px; }\n");
            writer.write("        .cash-transaction-item { display:flex; align-items:center; justify-content:space-between; gap:6px; font-size:.78rem; }\n");
            writer.write("        .annual-headline-grid .kpi-card { min-height:116px; }\n");
            writer.write("        .kpi-label { color:#c8d9eb; font-size:.8rem; text-transform:uppercase; }\n");
            writer.write("        .kpi-value { margin-top:2px; font-size:1.02rem; font-weight:700; color:#fff; }\n");
            writer.write("        .performer { margin-top:6px; font-size:.84rem; color:#dce8f3; }\n");
            writer.write("        .performer strong { display:block; font-size:.9rem; margin-bottom:2px; }\n");
            writer.write("        .report-standard .kpi-card-bestworst .performer strong { white-space:nowrap; overflow-x:auto; overflow-y:hidden; text-overflow:clip; scrollbar-width:none; -ms-overflow-style:none; }\n");
            writer.write("        .report-standard .kpi-card-bestworst .performer strong::-webkit-scrollbar { display:none; width:0; height:0; }\n");
            writer.write("        .performer-metrics { display:block; }\n");
            writer.write("        .hero-side { position:relative; background:rgba(255,255,255,.06); border:1px solid rgba(235,245,255,.22); border-radius:12px; padding:10px; min-height:172px; }\n");
            writer.write("        .timeline-title-row { display:flex; align-items:center; gap:6px; margin-bottom:8px; }\n");
            writer.write("        .hero-side-title { color:#d4e3f0; font-size:.86rem; text-transform:uppercase; margin:0; }\n");
            writer.write("        .timeline-info-btn { width:18px; height:18px; border-radius:999px; border:1px solid rgba(235,245,255,.55); background:rgba(255,255,255,.14); color:#e8f2fb; font-size:.72rem; font-weight:800; line-height:1; cursor:pointer; display:inline-flex; align-items:center; justify-content:center; padding:0; }\n");
            writer.write("        .timeline-info-btn:hover { background:rgba(255,255,255,.24); }\n");
            writer.write("        .annual-graphs-section .timeline-info-btn { border-color:#8fa9c2; background:#eaf2fb; color:#24415a; }\n");
            writer.write("        .annual-graphs-section .timeline-info-btn:hover { background:#ddeaf7; }\n");
            writer.write("        .timeline-info-overlay[hidden] { display:none !important; }\n");
            writer.write("        .timeline-info-overlay { position:fixed; inset:0; background:rgba(6,14,24,.58); z-index:12000; display:flex; align-items:center; justify-content:center; padding:18px; }\n");
            writer.write("        .timeline-info-dialog { width:min(560px,92vw); background:#f7fbff; color:#1a3348; border:1px solid #a8bfd4; border-radius:12px; box-shadow:0 18px 36px rgba(8,20,33,.34); }\n");
            writer.write("        .timeline-info-header { display:flex; align-items:center; justify-content:space-between; gap:10px; padding:12px 14px; border-bottom:1px solid #d5e3ef; }\n");
            writer.write("        .timeline-info-header h4 { margin:0; font-size:.95rem; letter-spacing:.2px; }\n");
            writer.write("        .timeline-info-close { border:1px solid #9cb5ca; background:#eef5fb; color:#20405a; border-radius:8px; width:26px; height:26px; cursor:pointer; font-size:1rem; line-height:1; padding:0; }\n");
            writer.write("        .timeline-info-body { padding:12px 14px 14px; font-size:.86rem; line-height:1.45; }\n");
            writer.write("        .timeline-info-body p { margin:0 0 8px; }\n");
            writer.write("        .timeline-info-body ul { margin:0; padding-left:18px; }\n");
            writer.write("        .timeline-info-body li { margin:0 0 6px; }\n");
            writer.write("        .hero-side-note { color:#d4e3f0; font-size:.92rem; }\n");
            writer.write("        .app-shell-note { color:#3b5570; font-size:.86rem; font-weight:600; line-height:1.35; }\n");
            writer.write("        .sparkline-widget { display:block; }\n");
            writer.write("        .sparkline-metric-controls { display:flex; flex-wrap:wrap; gap:7px; margin:0 0 8px; }\n");
            writer.write("        .sparkline-metric-btn { border:1px solid #b7c7d7; background:#f2f7fc; color:#27415a; border-radius:999px; padding:3px 9px; font-size:.72rem; font-weight:700; letter-spacing:.2px; cursor:pointer; }\n");
            writer.write("        .sparkline-metric-btn:hover { background:#e8f0f8; }\n");
            writer.write("        .sparkline-metric-btn.is-active { background:#24425b; color:#f4f9ff; border-color:#24425b; }\n");
            writer.write("        .sparkline-controls { display:flex; flex-wrap:wrap; gap:6px; margin:0 0 8px; }\n");
            writer.write("        .sparkline-controls.sparkline-controls-bottom { margin:8px 0 0; }\n");
            writer.write("        .sparkline-range-btn { border:1px solid #b7c7d7; background:#f2f7fc; color:#27415a; border-radius:999px; padding:3px 9px; font-size:.72rem; font-weight:700; letter-spacing:.2px; cursor:pointer; }\n");
            writer.write("        .sparkline-range-btn:hover { background:#e8f0f8; }\n");
            writer.write("        .sparkline-range-btn.is-active { background:#dbe9f8; color:#1f3f5b; border-color:#9eb9d5; }\n");
            writer.write("        .sparkline-return-summary { margin:0 0 8px; font-size:.8rem; font-weight:700; color:#2f4a62; }\n");
            writer.write("        .sparkline-return-summary.positive { color:var(--good); }\n");
            writer.write("        .sparkline-return-summary.negative { color:var(--bad); }\n");
            writer.write("        .hero-side { --spark-text:#d5e1ef; --spark-axis:#7f95ab; --spark-axis-soft:#9ab0c6; --spark-grid:#8ea4ba; --spark-line:#edf4fc; --spark-point:#edf4fc; }\n");
            writer.write("        .hero-side .sparkline-metric-btn, .hero-side .sparkline-range-btn { border-color:rgba(235,245,255,.35); background:rgba(255,255,255,.12); color:#e4eef8; }\n");
            writer.write("        .hero-side .sparkline-metric-btn:hover, .hero-side .sparkline-range-btn:hover { background:rgba(255,255,255,.2); }\n");
            writer.write("        .hero-side .sparkline-metric-btn.is-active, .hero-side .sparkline-range-btn.is-active { background:#eaf4ff; color:#16344d; border-color:#ffffff; }\n");
            writer.write("        .sparkline-panel { display:none; }\n");
            writer.write("        .sparkline-panel.is-active { display:block; }\n");
            writer.write("        .overview-charts { display:grid; grid-template-columns:1fr 1fr; gap:14px; margin:12px 0 14px; }\n");
            writer.write("        .overview-chart { padding:14px; border:1px solid var(--line); border-radius:14px; background:var(--card); box-shadow:0 5px 14px rgba(15,23,33,.06); overflow:hidden; }\n");
            writer.write("        .overview-chart h3 { margin:0 0 10px; font-size:1rem; }\n");
            writer.write("        .overview-chart .chart-svg { display:block; width:100%; margin:0 auto 12px; }\n");
            writer.write("        .allocation-card { margin:16px 0 18px; padding:14px; border:1px solid var(--line); border-radius:14px; background:var(--card); box-shadow:0 5px 14px rgba(15,23,33,.06); }\n");
            writer.write("        .allocation-card h3 { margin:0 0 10px; font-size:1rem; }\n");
            writer.write("        .annual-summary { margin:14px 0 18px; padding:14px; border:1px solid var(--line); border-radius:14px; background:var(--card); box-shadow:0 5px 14px rgba(15,23,33,.06); }\n");
            writer.write("        .annual-summary h3 { margin:0 0 10px; font-size:1rem; }\n");
            writer.write("        .annual-kpi-deck { margin:0 0 14px; padding:12px; border:1px solid var(--line); border-radius:14px; background:var(--card); box-shadow:0 5px 14px rgba(15,23,33,.06); }\n");
            writer.write("        .annual-kpi-deck-title { margin:0 0 10px; font-size:1.02rem; font-weight:700; letter-spacing:0; color:var(--ink); text-transform:none; }\n");
            writer.write("        .annual-summary-grid { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:10px; }\n");
            writer.write("        .annual-summary-card { border:1px solid #d4dfeb; border-radius:11px; padding:11px; background:linear-gradient(180deg,#f9fcff 0%,#f2f8fd 100%); }\n");
            writer.write("        .annual-summary-card h4 { margin:0 0 4px; font-size:.82rem; color:#40576c; text-transform:uppercase; }\n");
            writer.write("        .annual-summary-value { font-size:1.05rem; font-weight:700; }\n");
            writer.write("        .annual-summary-sub { margin-top:4px; font-size:.78rem; color:#5f7488; }\n");
            writer.write("        .annual-summary-value.positive, .annual-summary-sub.positive { color:var(--good); }\n");
            writer.write("        .annual-summary-value.negative, .annual-summary-sub.negative { color:var(--bad); }\n");
            writer.write("        .annual-value-warning { margin-top:6px; padding:6px 7px; font-size:.74rem; line-height:1.35; border:1px solid #f0d8a8; border-radius:8px; background:#fff5df; color:#7b4a00; }\n");
            writer.write("        .annual-summary-card .performer { margin-top:5px; color:#253d53; font-size:.82rem; }\n");
            writer.write("        .annual-summary-card .performer strong { margin-bottom:1px; font-size:.88rem; color:#1f3345; }\n");
            writer.write("        .annual-summary-card .performer.positive { color:var(--good); }\n");
            writer.write("        .annual-summary-card .performer.negative { color:var(--bad); }\n");
            writer.write("        .annual-graphs-section { margin:0 0 18px; padding:12px; border:1px solid var(--line); border-radius:14px; background:var(--card); box-shadow:0 5px 14px rgba(15,23,33,.06); }\n");
            writer.write("        .annual-graphs-heading { display:flex; flex-wrap:wrap; align-items:baseline; justify-content:space-between; gap:8px; margin:0 0 10px; }\n");
            writer.write("        .annual-graphs-heading h2 { margin:0; font-size:1.02rem; color:var(--ink); }\n");
            writer.write("        .annual-graphs-heading p { margin:0; font-size:.8rem; color:var(--muted); }\n");
            writer.write("        .annual-graphs-row { display:grid; grid-template-columns:1fr 1fr; gap:14px; margin:0; }\n");
            writer.write("        .annual-graph-card { display:flex; flex-direction:column; min-height:388px; padding:14px; border:1px solid #d4dfeb; border-radius:13px; background:linear-gradient(180deg,#f9fcff 0%,#f2f8fd 100%); box-shadow:0 2px 8px rgba(19,35,51,.06); overflow:hidden; }\n");
            writer.write("        .report-standard .annual-graph-card { min-height:410px; }\n");
            writer.write("        .annual-graph-card.full-span { grid-column:1 / -1; }\n");
            writer.write("        .annual-graph-card h3 { margin:0 0 6px; font-size:.84rem; font-weight:600; text-transform:uppercase; letter-spacing:.3px; color:#41576d; }\n");
            writer.write("        .total-return-graphs-section { background:linear-gradient(180deg,#f7fbff 0%,#eef5fc 100%); border-color:#c7d6e6; }\n");
            writer.write("        .total-return-chart { min-height:426px; }\n");
            writer.write("        .total-return-chart .chart-svg { background:linear-gradient(180deg,#fbfdff 0%,#f1f7ff 100%); border-color:#c5d5e5; border-radius:10px; }\n");
            writer.write("        .total-return-bar-chart .tr-plot-bg { fill:#f6fbff; stroke:#d6e1ed; }\n");
            writer.write("        .total-return-bar-chart .tr-grid-line { stroke:#d8e3ee; }\n");
            writer.write("        .total-return-bar-chart .tr-axis-label { fill:#496077; }\n");
            writer.write("        .total-return-bar-chart .tr-plot-border { stroke:#b9c8d7; }\n");
            writer.write("        .total-return-bar-chart .tr-axis-line { stroke:#4d6073; }\n");
            writer.write("        .annual-graph-note { margin:0 0 10px; font-size:.78rem; color:#5f7488; }\n");
            writer.write("        .annual-graph-content { flex:1; display:flex; flex-direction:column; justify-content:flex-start; min-height:0; }\n");
            writer.write("        .annual-graph-content > svg { display:block; width:100%; margin-top:auto; }\n");
            writer.write("        .annual-graph-content .sparkline-widget { display:flex; flex-direction:column; gap:6px; min-height:100%; }\n");
            writer.write("        .annual-graph-content .sparkline-panel { flex:1; min-height:0; }\n");
            writer.write("        .annual-graph-content .sparkline-panel.is-active { display:flex; align-items:stretch; }\n");
            writer.write("        .annual-graph-content .sparkline-panel > svg { display:block; width:100%; height:100%; min-height:220px; }\n");
            writer.write("        .allocation-visuals { display:grid; gap:10px; }\n");
            writer.write("        .allocation-row { display:grid; gap:10px; }\n");
            writer.write("        .allocation-row-top { grid-template-columns:repeat(3,minmax(0,1fr)); }\n");
            writer.write("        .allocation-row-bottom { grid-template-columns:repeat(2,minmax(0,1fr)); }\n");
            writer.write("        .allocation-panel { border:1px solid var(--line); border-radius:10px; padding:16px; background:#fafcfe; overflow:hidden; }\n");
            writer.write("        .allocation-panel-title { margin:0 0 6px; font-size:.84rem; font-weight:600; text-transform:uppercase; color:#41576d; letter-spacing:.3px; white-space:nowrap; }\n");
            writer.write("        .chart-svg { width:100%; height:auto; background:var(--card); border:1px solid var(--line); border-radius:8px; }\n");
            writer.write("        .allocation-panel .chart-svg { width:96%; margin:6px auto 10px; display:block; }\n");
            writer.write("        .security-pie-panel .chart-svg, .security-bar-panel .chart-svg { height:340px; width:100%; }\n");
            writer.write("        .chart-hover-target { cursor:pointer; transform-box:fill-box; transform-origin:center; transition:transform .14s ease, filter .14s ease, opacity .14s ease; }\n");
            writer.write("        .chart-hover-target.is-hovered { filter:brightness(1.08); opacity:.96; }\n");
            writer.write("        .chart-hover-bar.is-hovered { transform:translateY(-2px); }\n");
            writer.write("        .chart-hover-slice.is-hovered { transform:scale(1.03); }\n");
            writer.write("        .chart-hover-point.is-hovered { transform:scale(1.75); stroke:#ffffff; stroke-width:1; }\n");
            writer.write("        .chart-hover-avg-hit { pointer-events:stroke; }\n");
            writer.write("        .chart-total-return-label, .chart-security-bar-label { paint-order:stroke; stroke:#ffffff; stroke-width:1.5; stroke-linejoin:round; letter-spacing:.04px; }\n");
            writer.write("        .chart-security-label { font-weight:700; letter-spacing:.05px; paint-order:stroke; stroke:#ffffff; stroke-width:1.6; stroke-linejoin:round; }\n");
            writer.write("        .chart-tooltip { position:fixed; pointer-events:none; z-index:10000; max-width:340px; padding:7px 10px; border-radius:8px; background:rgba(16,28,40,.94); color:#f6fbff; font-size:.8rem; font-weight:600; line-height:1.3; box-shadow:0 8px 18px rgba(7,16,26,.28); border:1px solid rgba(255,255,255,.14); opacity:0; transform:translateY(4px); transition:opacity .1s ease, transform .1s ease; }\n");
            writer.write("        .chart-tooltip.visible { opacity:1; transform:translateY(0); }\n");
            writer.write("        .chart-title-row { display:flex; align-items:center; justify-content:space-between; gap:8px; margin:0 0 8px; }\n");
            writer.write("        .chart-title-row > h3, .chart-title-row > h4, .chart-title-row > .hero-side-title { margin:0; }\n");
            writer.write("        .chart-title-row > h3, .chart-title-row > h4 { white-space:nowrap; }\n");
            writer.write("        .chart-download-btn { border:1px solid #86a4bf; background:#f3f8fd; color:#1e3951; border-radius:7px; width:28px; height:28px; display:inline-flex; align-items:center; justify-content:center; cursor:pointer; padding:0; }\n");
            writer.write("        .chart-download-btn:hover { background:#e6f1fb; }\n");
            writer.write("        .chart-download-btn svg { width:15px; height:15px; stroke:currentColor; fill:none; stroke-width:2; stroke-linecap:round; stroke-linejoin:round; }\n");
            writer.write("        .chart-toolbar { display:flex; gap:6px; align-items:center; flex-wrap:wrap; margin:0 0 8px; position:relative; z-index:3; }\n");
            writer.write("        .chart-tool-btn { border:1px solid #86a4bf; background:#f3f8fd; color:#1e3951; border-radius:7px; min-width:28px; height:28px; padding:0 7px; font-size:.74rem; font-weight:700; cursor:pointer; }\n");
            writer.write("        .chart-tool-btn:hover { background:#e6f1fb; }\n");
            writer.write("        .chart-filter-input { border:1px solid #a5bbcf; border-radius:7px; height:28px; min-width:130px; padding:0 8px; font-size:.74rem; background:#ffffff; color:#20384f; }\n");
            writer.write("        .chart-filter-input::placeholder { color:#6f8498; }\n");
            writer.write("        .chart-viewport { position:relative; overflow:hidden; border:1px solid var(--line); border-radius:8px; background:var(--card); z-index:1; }\n");
            writer.write("        .chart-svg { transform-origin:0 0; transition:transform .12s ease-out; }\n");
            writer.write("        .chart-viewport .chart-svg { border:none; border-radius:0; margin:0 !important; }\n");
            writer.write("        .chart-svg.is-panning { cursor:grabbing; }\n");
            writer.write("        .hero-theme-btn { border:1px solid rgba(235,245,255,.45); background:rgba(255,255,255,.12); color:#f3f7fc; border-radius:999px; padding:4px 10px; font-size:.78rem; font-weight:700; cursor:pointer; }\n");
            writer.write("        .hero-theme-btn:hover { background:rgba(255,255,255,.2); }\n");
            writer.write("        .hero-refresh-btn { border:1px solid rgba(235,245,255,.45); background:rgba(255,255,255,.12); color:#f3f7fc; border-radius:999px; padding:4px 10px; font-size:.78rem; font-weight:700; cursor:pointer; }\n");
            writer.write("        .hero-refresh-btn:hover:not(:disabled) { background:rgba(255,255,255,.2); }\n");
            writer.write("        .hero-refresh-btn:disabled { opacity:.6; cursor:not-allowed; }\n");
            writer.write("        .price-refresh-status { font-size:.76rem; color:#d7e6f4; margin-top:8px; min-height:1.1em; }\n");
            writer.write("        body.theme-dark .table-wrap, body.theme-dark .overview-chart, body.theme-dark .allocation-card, body.theme-dark .allocation-panel, body.theme-dark .details-table, body.theme-dark .annual-kpi-deck, body.theme-dark .annual-graphs-section { border-color:#2a3a4f; box-shadow:none; }\n");
            writer.write("        body.theme-dark .total-row { background:#1a2a3b; color:#ecf3fb; }\n");
            writer.write("        body.theme-dark td, body.theme-dark th { border-bottom-color:#2a3a4d; }\n");
            writer.write("        body.theme-dark th { background:#1d2a3a; color:#d8e4f2; }\n");
            writer.write("        body.theme-dark .details-cell { background:#111d2b; }\n");
            writer.write("        body.theme-dark .details-wrap { background:#111d2b; }\n");
            writer.write("        body.theme-dark .details-wrap h4 { color:#c6d8ea; }\n");
            writer.write("        body.theme-dark .details-table { background:#162231; border-color:#2a3a4d; }\n");
            writer.write("        body.theme-dark .details-table th { background:#1b2b3d; color:#d7e4f2; }\n");
            writer.write("        body.theme-dark .details-table td { color:#dbe7f4; border-bottom-color:#2a3a4d; }\n");
            writer.write("        body.theme-dark .allocation-panel { background:#132235; }\n");
            writer.write("        body.theme-dark .allocation-panel-title { color:#c8d8e8; }\n");
            writer.write("        body.theme-dark .timeline-info-dialog { background:#122437; color:#d8e7f5; border-color:#2b4360; }\n");
            writer.write("        body.theme-dark .timeline-info-header { border-bottom-color:#2b4360; }\n");
            writer.write("        body.theme-dark .timeline-info-close { background:#1a3149; border-color:#3a5879; color:#d8e7f5; }\n");
            writer.write("        body.theme-dark .annual-summary { border-color:#2a3a4f; box-shadow:none; }\n");
            writer.write("        body.theme-dark .annual-kpi-deck-title { color:#e5edf7; }\n");
            writer.write("        body.theme-dark .annual-summary-card { border-color:#2e4258; background:#1a2d42; }\n");
            writer.write("        body.theme-dark .annual-summary-card h4 { color:#c8d9eb; }\n");
            writer.write("        body.theme-dark .annual-summary-sub { color:#d6e4f1; }\n");
            writer.write("        body.theme-dark .annual-value-warning { background:#3d2e19; border-color:#8e6a33; color:#ffdca8; }\n");
            writer.write("        body.theme-dark .annual-summary-card .performer { color:#d4e3f2; }\n");
            writer.write("        body.theme-dark .annual-summary-card .performer strong { color:#e9f2fc; }\n");
            writer.write("        body.theme-dark .annual-summary-value.positive, body.theme-dark .annual-summary-sub.positive { color:var(--good); }\n");
            writer.write("        body.theme-dark .annual-summary-value.negative, body.theme-dark .annual-summary-sub.negative { color:var(--bad); }\n");
            writer.write("        body.theme-dark .annual-summary-card .performer.positive { color:var(--good); }\n");
            writer.write("        body.theme-dark .annual-summary-card .performer.negative { color:var(--bad); }\n");
            writer.write("        body.theme-dark .annual-graphs-heading h2 { color:#e5edf7; }\n");
            writer.write("        body.theme-dark .annual-graphs-heading p { color:#bad0e5; }\n");
            writer.write("        body.theme-dark .annual-graphs-section .timeline-info-btn { border-color:#4d6a87; background:#21374e; color:#d7e8f8; }\n");
            writer.write("        body.theme-dark .annual-graphs-section .timeline-info-btn:hover { background:#2a4663; }\n");
            writer.write("        body.theme-dark .annual-graph-card { border-color:#2e4258; background:#1a2d42; box-shadow:none; }\n");
            writer.write("        body.theme-dark .total-return-graphs-section { background:linear-gradient(180deg,#15283b 0%,#122131 100%); border-color:#2e435a; }\n");
            writer.write("        body.theme-dark .total-return-chart .chart-svg { background:linear-gradient(180deg,#132436 0%,#122131 100%); border-color:#30495f; }\n");
            writer.write("        body.theme-dark .total-return-bar-chart .tr-plot-bg { fill:#1a2f44 !important; stroke:#36506a !important; }\n");
            writer.write("        body.theme-dark .total-return-bar-chart .tr-grid-line { stroke:#3e5872 !important; }\n");
            writer.write("        body.theme-dark .total-return-bar-chart .tr-axis-label { fill:#c8d9ea !important; }\n");
            writer.write("        body.theme-dark .total-return-bar-chart .tr-plot-border { stroke:#4b6580 !important; }\n");
            writer.write("        body.theme-dark .total-return-bar-chart .tr-axis-line { stroke:#8da7c1 !important; }\n");
            writer.write("        body.theme-dark .annual-graph-card h3 { color:#bdd1e4; }\n");
            writer.write("        body.theme-dark .annual-graph-note { color:#bad0e5; }\n");
            writer.write("        body.theme-dark .sparkline-metric-btn, body.theme-dark .sparkline-range-btn { border-color:#45627f; background:#22374d; color:#cfe0f2; }\n");
            writer.write("        body.theme-dark .sparkline-metric-btn:hover, body.theme-dark .sparkline-range-btn:hover { background:#2b4560; }\n");
            writer.write("        body.theme-dark .sparkline-metric-btn.is-active { background:#dceafb; color:#173047; border-color:#dceafb; }\n");
            writer.write("        body.theme-dark .sparkline-range-btn.is-active { background:#c3d7ed; color:#173047; border-color:#9bb7d3; }\n");
            writer.write("        body.theme-dark .sparkline-return-summary { color:#cfe0f2; }\n");
            writer.write("        body.theme-dark .sparkline-return-summary.positive { color:var(--good); }\n");
            writer.write("        body.theme-dark .sparkline-return-summary.negative { color:var(--bad); }\n");
            writer.write("        body.theme-dark .chart-title-row > h3, body.theme-dark .chart-title-row > h4, body.theme-dark .chart-title-row > .hero-side-title { color:#dce8f5; }\n");
            writer.write("        body.theme-dark .chart-svg { background:#162231; border-color:#2b3a4d; }\n");
            writer.write("        body.theme-dark .chart-svg text { fill:#d4e1ee !important; }\n");
            writer.write("        body.theme-dark .chart-total-return-label, body.theme-dark .chart-security-bar-label { fill:#e7f0fa !important; stroke:#0b1624 !important; stroke-width:1.0; }\n");
            writer.write("        body.theme-dark .chart-security-label { fill:#e7f0fa !important; stroke:#0b1624 !important; stroke-width:1.05; }\n");
            writer.write("        body.theme-dark .market-value-bar-chart line[stroke='#495057'] { stroke:#dce8f4 !important; }\n");
            writer.write("        body.theme-dark .market-value-bar-chart text[fill='#495057'] { fill:#dce8f4 !important; }\n");
            writer.write("        body.theme-dark .app-shell-note, body.theme-dark .hero-side-note { color:#d1e0ef; }\n");
            writer.write("        .details-link-btn { display:block; width:100%; min-width:0; text-align:left; border:none; background:transparent; color:inherit; font:inherit; padding:0; margin:0; cursor:pointer; }\n");
            writer.write("        .details-link-btn:hover { text-decoration:underline; text-decoration-thickness:1px; text-underline-offset:2px; }\n");
            writer.write("        .details-head { display:inline-flex; align-items:center; gap:6px; }\n");
            writer.write("        .detail-group-toggle { border:1px solid #86a4bf; background:#f3f8fd; color:#1e3951; border-radius:50%; width:18px; height:18px; padding:0; line-height:16px; font-size:.72rem; font-weight:700; cursor:pointer; display:inline-flex; align-items:center; justify-content:center; }\n");
            writer.write("        .detail-group-toggle:hover { background:#e6f1fb; }\n");
            writer.write("        .details-row { display:none; }\n");
            writer.write("        .details-cell { padding:0 !important; background:#f9fcff; }\n");
            writer.write("        .details-wrap { padding:10px 12px 12px; overflow-x:auto; overflow-y:hidden; }\n");
            writer.write("        .details-wrap h4 { margin:0 0 8px; font-size:.88rem; color:#2b4358; text-transform:uppercase; letter-spacing:.25px; }\n");
            writer.write("        .details-table { width:max-content; min-width:100%; border-collapse:collapse; background:#fff; border:1px solid #dfe7ef; }\n");
            writer.write("        .details-table th, .details-table td { padding:6px 7px; border-bottom:1px solid #edf2f7; font-size:.72rem; white-space:nowrap; overflow:visible; text-overflow:clip; }\n");
            writer.write("        .details-table th { background:#f4f8fc; color:#405a70; }\n");
            writer.write("        .details-buy { color:#1d5d92; font-weight:600; }\n");
            writer.write("        .details-dividend { color:#1f8b4d; font-weight:600; }\n");
            writer.write("        @media (max-width:1200px) { .allocation-row-top{grid-template-columns:1fr 1fr;} .annual-summary-grid{grid-template-columns:repeat(3,minmax(0,1fr));} }\n");
            writer.write("        @media (max-width:1060px) { .report-hero{grid-template-columns:1fr;} .hero-kpis,.annual-headline-grid{grid-template-columns:1fr;} .annual-summary-grid{grid-template-columns:repeat(2,minmax(0,1fr));} .annual-graphs-row{grid-template-columns:1fr;} .overview-charts{grid-template-columns:1fr;} .allocation-row-top,.allocation-row-bottom{grid-template-columns:1fr;} .page{width:100%; padding:16px 8px 22px;} .table-wrap{overflow-x:auto;} .report-standard .overview-table{min-width:0;} .report-standard .realized-table{min-width:0;} .report-annual .table-wrap table{min-width:980px;} }\n");
            writer.write("        @media (max-width:760px) { .annual-summary-grid{grid-template-columns:1fr;} .annual-graphs-heading{flex-direction:column; align-items:flex-start;} }\n");
            writer.write("    </style>\n");
            writer.write("</head>\n");
            writer.write("<body class=\"report-" + reportConfig.reportType + "\">\n");
            writer.write("<main class=\"page\">\n");

            if (REPORT_TYPE_ANNUAL.equals(reportConfig.reportType)) {
                writeAnnualHeaderSummaryHtml(writer, store, ratesToNok, reportConfig.reportYear);
                writeAnnualSummarySectionHtml(writer, store, ratesToNok, reportConfig.reportYear, annualSummary, annualSnapshotRows);
                writeAnnualTimelineChartsHtml(writer, store, ratesToNok, reportConfig.reportYear);
                writeAnnualPortfolioSnapshotTableHtml(writer, annualSnapshotRows, ratesToNok, reportConfig.reportYear);
                writeAnnualRealizedSummaryTableHtml(writer, store, ratesToNok, reportConfig.reportYear);
            } else {
                // Standard portfolio report
                writeHeaderSummaryHtml(writer, headerSummary, overviewRows, store, ratesToNok);
                writeOverviewTableHtml(writer, overviewRows, store, ratesToNok);
                writeRealizedSummaryTableHtml(writer, store, ratesToNok);
            }

            writer.write("</main>\n");
            writer.write("<script>\n");
            ReportScriptHelper.writeDetailsToggleScript(writer);
            writeCurrencyConversionScript(writer, ratesToNok);
            writer.write("</script>\n");
            writer.write("</body>\n");
            writer.write("</html>\n");
        }
    }

    private static void writeAnnualSummaryCardsHtml(
            FileWriter writer,
            AnnualPerformanceSummary summary,
            AnnualHeroMetrics metrics,
            Map<String, Double> ratesToNok,
            List<AnnualSnapshotRow> snapshotRows) throws IOException {

        if (summary == null || metrics == null) {
            return;
        }

        LinkedHashMap<String, Double> cashBuckets = new LinkedHashMap<>();
        cashBuckets.put(DEFAULT_TOTAL_CURRENCY, metrics.cashHoldingsNok);
        LinkedHashMap<String, Double> valueBuckets = singleCurrencyBuckets(DEFAULT_TOTAL_CURRENCY, summary.endValueNok);
        LinkedHashMap<String, Double> portfolioReturnBuckets = singleCurrencyBuckets(DEFAULT_TOTAL_CURRENCY, summary.portfolioReturnNok);
        LinkedHashMap<String, Double> realizedGainBuckets = singleCurrencyBuckets(DEFAULT_TOTAL_CURRENCY, summary.realizedGainNok);
        LinkedHashMap<String, Double> dividendsBuckets = singleCurrencyBuckets(DEFAULT_TOTAL_CURRENCY, summary.dividendsNok);
        LinkedHashMap<String, Double> realizedTotalBuckets = singleCurrencyBuckets(DEFAULT_TOTAL_CURRENCY, summary.realizedTotalNok);
        LinkedHashMap<String, Double> bestReturnBuckets = singleCurrencyBuckets(DEFAULT_TOTAL_CURRENCY, metrics.best.returnNok);
        LinkedHashMap<String, Double> worstReturnBuckets = singleCurrencyBuckets(DEFAULT_TOTAL_CURRENCY, metrics.worst.returnNok);

        String portfolioClass = signedClass(summary.portfolioReturnNok);
        double benchmarkDelta = summary.hasBenchmarkData ? (summary.portfolioReturnPct - summary.benchmarkReturnPct) : 0.0;
        String deltaClass = signedClass(benchmarkDelta);
        String bestClass = signedClass(metrics.best.returnNok);
        String worstClass = signedClass(metrics.worst.returnNok);
        String valueWarningHtml = buildAnnualValueWarningHtml(snapshotRows);

        writer.write("<article class=\"kpi-card annual-summary-card\"><h4>Value</h4><div class=\"annual-summary-value\">"
            + renderConvertibleMoneyCell(valueBuckets, 2, ratesToNok)
            + "</div><div class=\"annual-summary-sub\">Portfolio value at end of year</div>"
            + valueWarningHtml
            + "</article>\n");

        writer.write("<article class=\"kpi-card annual-summary-card\"><h4>Cash Holdings</h4><div class=\"annual-summary-value\">"
            + renderConvertibleMoneyCell(cashBuckets, 0, ratesToNok)
            + "</div><div class=\"annual-summary-sub\">Available cash at year end</div></article>\n");

        writer.write("<article class=\"kpi-card annual-summary-card\"><h4>Portfolio Return</h4><div class=\"annual-summary-value " + portfolioClass + "\">"
            + renderConvertibleMoneyCell(portfolioReturnBuckets, 2, ratesToNok)
            + "</div><div class=\"annual-summary-sub " + portfolioClass + "\">"
            + HtmlFormatter.formatPercent(summary.portfolioReturnPct)
            + "</div><div class=\"annual-summary-sub\">Time-weighted annual return, adjusted for external cash flows.</div></article>\n");

        String realizedGainClass = signedClass(summary.realizedGainNok);
        writer.write("<article class=\"kpi-card annual-summary-card\"><h4>Realized Gain/Loss</h4><div class=\"annual-summary-value " + realizedGainClass + "\">"
            + renderConvertibleMoneyCell(realizedGainBuckets, 2, ratesToNok)
            + "</div><div class=\"annual-summary-sub\">Closed sales in selected year</div></article>\n");

        writer.write("<article class=\"kpi-card annual-summary-card\"><h4>Dividends</h4><div class=\"annual-summary-value\">"
            + renderConvertibleMoneyCell(dividendsBuckets, 2, ratesToNok)
            + "</div><div class=\"annual-summary-sub\">Dividend cash flows in selected year</div></article>\n");

        String realizedTotalClass = signedClass(summary.realizedTotalNok);
        writer.write("<article class=\"kpi-card annual-summary-card\"><h4>Total Realized</h4><div class=\"annual-summary-value " + realizedTotalClass + "\">"
            + renderConvertibleMoneyCell(realizedTotalBuckets, 2, ratesToNok)
            + "</div><div class=\"annual-summary-sub\">Realized gain/loss plus dividends</div></article>\n");

        if (summary.hasBenchmarkData) {
            String benchmarkClass = signedClass(summary.benchmarkReturnPct);
            writer.write("<article class=\"kpi-card annual-summary-card\"><h4>Benchmark (" + escapeHtml(summary.benchmarkTicker) + ")</h4><div class=\"annual-summary-value " + benchmarkClass + "\">"
                + HtmlFormatter.formatPercent(summary.benchmarkReturnPct)
                + "</div><div class=\"annual-summary-sub\">Selected year performance</div></article>\n");

            writer.write("<article class=\"kpi-card annual-summary-card\"><h4>Relative vs Benchmark</h4><div class=\"annual-summary-value " + deltaClass + "\">"
                + HtmlFormatter.formatPercent(benchmarkDelta)
                + "</div><div class=\"annual-summary-sub\">Portfolio minus benchmark</div></article>\n");
        } else {
            writer.write("<article class=\"kpi-card annual-summary-card\"><h4>Benchmark (" + escapeHtml(summary.benchmarkTicker) + ")</h4><div class=\"annual-summary-value\">0.00%</div><div class=\"annual-summary-sub\">No benchmark data available for this year.</div></article>\n");
            writer.write("<article class=\"kpi-card annual-summary-card\"><h4>Relative vs Benchmark</h4><div class=\"annual-summary-value " + portfolioClass + "\">"
                + HtmlFormatter.formatPercent(summary.portfolioReturnPct)
                + "</div><div class=\"annual-summary-sub\">Portfolio minus 0.00% fallback benchmark</div></article>\n");
        }

        writer.write("<article class=\"kpi-card annual-summary-card\"><h4>Best / Worst</h4><div class=\"performer " + bestClass + "\"><strong>"
            + escapeHtml(metrics.best.label)
            + "</strong><span class=\"performer-metrics\">"
            + renderConvertibleMoneyCell(bestReturnBuckets, 0, ratesToNok)
            + " | " + HtmlFormatter.formatPercent(metrics.best.returnPct)
            + "</span></div><div class=\"performer " + worstClass + "\"><strong>"
            + escapeHtml(metrics.worst.label)
            + "</strong><span class=\"performer-metrics\">"
            + renderConvertibleMoneyCell(worstReturnBuckets, 0, ratesToNok)
            + " | " + HtmlFormatter.formatPercent(metrics.worst.returnPct)
            + "</span></div></article>\n");

        if (summary.hasAnalytics) {
            writer.write("<article class=\"kpi-card annual-summary-card\"><h4>Volatility (Ann.)</h4><div class=\"annual-summary-value\">"
                + HtmlFormatter.formatPercent(summary.annualizedVolatilityPct, 2)
                + "</div><div class=\"annual-summary-sub\">Annualized from monthly return variance</div></article>\n");

            writer.write("<article class=\"kpi-card annual-summary-card\"><h4>Sharpe Ratio</h4><div class=\"annual-summary-value\">"
                + String.format(Locale.US, "%.2f", summary.sharpeRatio)
                + "</div><div class=\"annual-summary-sub\">Risk-adjusted return (monthly, annualized)</div></article>\n");

            if (summary.hasBeta) {
                writer.write("<article class=\"kpi-card annual-summary-card\"><h4>Beta vs " + escapeHtml(summary.benchmarkTicker) + "</h4><div class=\"annual-summary-value\">"
                    + String.format(Locale.US, "%.2f", summary.beta)
                    + "</div><div class=\"annual-summary-sub\">Sensitivity vs benchmark monthly returns</div></article>\n");
            } else {
                writer.write("<article class=\"kpi-card annual-summary-card\"><h4>Beta vs " + escapeHtml(summary.benchmarkTicker) + "</h4><div class=\"annual-summary-value\">N/A</div><div class=\"annual-summary-sub\">Insufficient benchmark overlap</div></article>\n");
            }
        }

        if (summary.hasMonteCarlo) {
            LinkedHashMap<String, Double> medianBuckets = singleCurrencyBuckets(DEFAULT_TOTAL_CURRENCY, summary.monteCarloMedianEndValueNok);
            LinkedHashMap<String, Double> p10Buckets = singleCurrencyBuckets(DEFAULT_TOTAL_CURRENCY, summary.monteCarloP10EndValueNok);
            LinkedHashMap<String, Double> p90Buckets = singleCurrencyBuckets(DEFAULT_TOTAL_CURRENCY, summary.monteCarloP90EndValueNok);
            writer.write("<article class=\"kpi-card annual-summary-card\"><h4>Monte Carlo (" + summary.monteCarloHorizonMonths + "m)</h4><div class=\"annual-summary-value\">"
                + renderConvertibleMoneyCell(medianBuckets, 0, ratesToNok)
                + "</div><div class=\"annual-summary-sub\">Median terminal value (" + summary.monteCarloIterations + " iterations)</div>"
                + "<div class=\"annual-summary-sub\">P10: " + renderConvertibleMoneyCell(p10Buckets, 0, ratesToNok) + " | P90: " + renderConvertibleMoneyCell(p90Buckets, 0, ratesToNok) + "</div></article>\n");
        }

    }

    private static String buildAnnualValueWarningHtml(List<AnnualSnapshotRow> snapshotRows) {
        if (snapshotRows == null || snapshotRows.isEmpty()) {
            return "";
        }

        for (AnnualSnapshotRow row : snapshotRows) {
            if (row != null && row.hasEstimatedPrice) {
                return "<div class=\"annual-value-warning\">Estimated value: prices closest to 31.12 were used where exact year-end closes were unavailable.</div>";
            }
        }
        return "";
    }

    private static final class AnnualSecurityPerformance {
        private final String label;
        private final double returnNok;
        private final double returnPct;

        private AnnualSecurityPerformance(String label, double returnNok, double returnPct) {
            this.label = label;
            this.returnNok = returnNok;
            this.returnPct = returnPct;
        }
    }

    private static final class AnnualHeroMetrics {
        private final int transactionCount;
        private final int holdingsCount;
        private final double cashHoldingsNok;
        private final AnnualSecurityPerformance best;
        private final AnnualSecurityPerformance worst;

        private AnnualHeroMetrics(
                int transactionCount,
                int holdingsCount,
                double cashHoldingsNok,
                AnnualSecurityPerformance best,
                AnnualSecurityPerformance worst) {
            this.transactionCount = transactionCount;
            this.holdingsCount = holdingsCount;
            this.cashHoldingsNok = cashHoldingsNok;
            this.best = best;
            this.worst = worst;
        }
    }

        private static void writeAnnualHeaderSummaryHtml(
            FileWriter writer,
            TransactionStore store,
            Map<String, Double> ratesToNok,
            int reportYear) throws IOException {

        int safeYear = Math.max(2000, Math.min(2100, reportYear));
        AnnualHeroMetrics metrics = buildAnnualHeroMetrics(store, ratesToNok, safeYear);

        writer.write("<section class=\"report-hero annual-hero\">\n");
        writer.write("<div class=\"hero-title annual-hero-header\">\n");
        writer.write("<h1>Annual Report - " + safeYear + "</h1>\n");
        writer.write("<div class=\"hero-meta\">\n");
        writer.write("<span class=\"meta-chip\">Files: <strong>" + store.getLoadedCsvFileCount() + "</strong></span>\n");
        writer.write("<span class=\"meta-chip\">Transactions: <strong>" + metrics.transactionCount + "</strong></span>\n");
        writer.write("<span class=\"meta-chip\">Holdings: <strong>" + metrics.holdingsCount + "</strong></span>\n");
        writer.write("<span class=\"meta-chip\">Currency: <strong><input id=\"portfolio-currency-input\" class=\"currency-input\" type=\"text\" value=\"" + DEFAULT_TOTAL_CURRENCY + "\" maxlength=\"3\" autocomplete=\"off\" spellcheck=\"false\" title=\"Skriv valutakode (f.eks. NOK, USD) og trykk Enter\"></strong></span>\n");
        writer.write("<button id=\"report-theme-toggle\" class=\"hero-theme-btn\" type=\"button\">Dark mode</button>\n");
        writer.write("</div>\n");

        writer.write("</div>\n");
        writer.write("</section>\n");
    }

    private static void writeAnnualSummarySectionHtml(
            FileWriter writer,
            TransactionStore store,
            Map<String, Double> ratesToNok,
            int reportYear,
            AnnualPerformanceSummary annualSummary,
            List<AnnualSnapshotRow> snapshotRows) throws IOException {

        if (annualSummary == null) {
            return;
        }

        int safeYear = Math.max(2000, Math.min(2100, reportYear));
        AnnualHeroMetrics metrics = buildAnnualHeroMetrics(store, ratesToNok, safeYear);

        writer.write("<section class=\"annual-kpi-deck\">\n");
        writer.write("<h2 class=\"annual-kpi-deck-title\">Annual Performance</h2>\n");
        writer.write("<div class=\"annual-summary-grid\">\n");
        writeAnnualSummaryCardsHtml(writer, annualSummary, metrics, ratesToNok, snapshotRows);
        writer.write("</div>\n");
        writer.write("</section>\n");
    }

    private static void writeAnnualTimelineChartsHtml(
            FileWriter writer,
            TransactionStore store,
            Map<String, Double> ratesToNok,
            int reportYear) throws IOException {

        int safeYear = Math.max(2000, Math.min(2100, reportYear));
        String valueChartSvg = PortfolioCalculator.buildAnnualPortfolioValueSparklineSvg(store, ratesToNok, safeYear);
        String returnChartSvg = PortfolioCalculator.buildAnnualPortfolioReturnSparklineSvg(store, ratesToNok, safeYear);

        writer.write("<section class=\"annual-graphs-section\">\n");
        writer.write("<div class=\"annual-graphs-heading\"><h2>Yearly Trend</h2><p>Value and return development month by month for " + safeYear + ".</p></div>\n");
        writer.write("<div class=\"annual-graphs-row\">\n");

        writer.write("<article class=\"annual-graph-card\">\n");
        writer.write("<h3>Portfolio Value</h3>\n");
        writer.write("<p class=\"annual-graph-note\">Month-end portfolio value in selected year.</p>\n");
        writer.write("<div class=\"annual-graph-content\">\n");
        if (valueChartSvg == null || valueChartSvg.isBlank()) {
            writer.write("<div class=\"app-shell-note\">Timeline data is not available for the selected year.</div>\n");
        } else {
            writer.write(valueChartSvg + "\n");
        }
        writer.write("</div>\n");
        writer.write("</article>\n");

        writer.write("<article class=\"annual-graph-card\">\n");
        writer.write("<h3>Portfolio Return</h3>\n");
        writer.write("<p class=\"annual-graph-note\">Month-end portfolio return in selected year.</p>\n");
        writer.write("<div class=\"annual-graph-content\">\n");
        if (returnChartSvg == null || returnChartSvg.isBlank()) {
            writer.write("<div class=\"app-shell-note\">Return timeline is not available for the selected year.</div>\n");
        } else {
            writer.write(returnChartSvg + "\n");
        }
        writer.write("</div>\n");
        writer.write("</article>\n");

        writer.write("</div>\n");
        writer.write("</section>\n");
    }

    private static AnnualHeroMetrics buildAnnualHeroMetrics(
            TransactionStore store,
            Map<String, Double> ratesToNok,
            int reportYear) {

        LocalDate snapshotDate = LocalDate.of(reportYear, 12, 31);
        List<AnnualSnapshotRow> snapshotRows = buildAnnualSnapshotRows(store, snapshotDate);
        int holdingsCount = snapshotRows.size();
        int transactionCount = countAnnualTransactions(store, reportYear);
        double cashHoldingsNok = computeCashHoldingsAtDate(store, snapshotDate);

        AnnualSecurityPerformance best = new AnnualSecurityPerformance("No yearly realized data", 0.0, 0.0);
        AnnualSecurityPerformance worst = best;
        boolean hasBestWorst = false;

        for (Security security : store.getSecurities()) {
            if (security == null) {
                continue;
            }

            String securityCurrency = normalizeCurrencyCode(security.getCurrencyCode());
            double rateToNok = ratesToNok == null ? 0.0 : ratesToNok.getOrDefault(securityCurrency, 0.0);
            if (rateToNok <= 0.0) {
                rateToNok = ratesToNok == null ? 1.0 : ratesToNok.getOrDefault(DEFAULT_TOTAL_CURRENCY, 1.0);
            }
                final double rateToNokFinal = rateToNok;

            double yearGainNok = security.getSaleTradesSortedByDate().stream()
                    .filter(trade -> trade != null && trade.getTradeDate() != null && trade.getTradeDate().getYear() == reportYear)
                    .mapToDouble(trade -> trade.getGainLoss() * rateToNokFinal)
                    .sum();

            double yearCostNok = security.getSaleTradesSortedByDate().stream()
                    .filter(trade -> trade != null && trade.getTradeDate() != null && trade.getTradeDate().getYear() == reportYear)
                    .mapToDouble(trade -> trade.getCostBasis() * rateToNokFinal)
                    .sum();

            double yearDividendsNok = security.getAllDividendEventsSortedByDate().stream()
                    .filter(event -> event != null && event.getTradeDate() != null && event.getTradeDate().getYear() == reportYear)
                    .mapToDouble(event -> event.getAmount() * rateToNokFinal)
                    .sum();

            double yearReturnNok = yearGainNok + yearDividendsNok;
            if (Math.abs(yearReturnNok) < 1e-9 && Math.abs(yearCostNok) < 1e-9) {
                continue;
            }

            double yearReturnPct = yearCostNok > 0.0
                    ? (yearReturnNok / yearCostNok) * 100.0
                    : (yearReturnNok > 0.0 ? 100.0 : 0.0);

            AnnualSecurityPerformance current = new AnnualSecurityPerformance(
                    security.getDisplayName(),
                    yearReturnNok,
                    yearReturnPct
            );

            if (!hasBestWorst || current.returnNok > best.returnNok) {
                best = current;
            }
            if (!hasBestWorst || current.returnNok < worst.returnNok) {
                worst = current;
            }
            hasBestWorst = true;
        }

        return new AnnualHeroMetrics(
                transactionCount,
                holdingsCount,
                cashHoldingsNok,
                best,
            worst
        );
    }

    private static int countAnnualTransactions(TransactionStore store, int reportYear) {
        int unitEventCount = 0;
        for (Events.UnitEvent event : store.getUnitEvents()) {
            if (event != null && event.tradeDate() != null && event.tradeDate().getYear() == reportYear) {
                unitEventCount++;
            }
        }

        int externalCashCount = 0;
        for (Events.CashEvent event : store.getCashEvents()) {
            if (event != null
                    && event.tradeDate() != null
                    && event.tradeDate().getYear() == reportYear
                    && event.externalFlow()) {
                externalCashCount++;
            }
        }

        int dividendCount = 0;
        for (Security security : store.getSecurities()) {
            if (security == null) {
                continue;
            }
            for (Security.DividendEvent event : security.getAllDividendEventsSortedByDate()) {
                if (event != null && event.getTradeDate() != null && event.getTradeDate().getYear() == reportYear) {
                    dividendCount++;
                }
            }
        }

        return unitEventCount + externalCashCount + dividendCount;
    }

    private static double computeCashHoldingsAtDate(TransactionStore store, LocalDate snapshotDate) {
        List<Events.PortfolioCashSnapshot> snapshots = store.getPortfolioCashSnapshots();
        if (!snapshots.isEmpty()) {
            LinkedHashMap<String, Events.PortfolioCashSnapshot> latestByPortfolio = new LinkedHashMap<>();
            for (Events.PortfolioCashSnapshot snapshot : snapshots) {
                if (snapshot == null
                        || snapshot.tradeDate() == null
                        || snapshot.tradeDate().isAfter(snapshotDate)
                        || snapshot.portfolioId() == null
                        || snapshot.portfolioId().isBlank()) {
                    continue;
                }

                Events.PortfolioCashSnapshot existing = latestByPortfolio.get(snapshot.portfolioId());
                if (existing == null
                        || snapshot.tradeDate().isAfter(existing.tradeDate())
                        || (snapshot.tradeDate().equals(existing.tradeDate()) && snapshot.sortId() >= existing.sortId())) {
                    latestByPortfolio.put(snapshot.portfolioId(), snapshot);
                }
            }

            double total = 0.0;
            for (Events.PortfolioCashSnapshot snapshot : latestByPortfolio.values()) {
                total += snapshot.balance();
            }
            return total;
        }

        double total = 0.0;
        for (Events.CashEvent event : store.getCashEvents()) {
            if (event != null && event.tradeDate() != null && !event.tradeDate().isAfter(snapshotDate)) {
                total += event.cashDelta();
            }
        }
        return total;
    }

    private static final class AnnualSnapshotRow {
        private final String ticker;
        private final String securityName;
        private final String assetType;
        private final String currencyCode;
        private final double units;
        private final double averageCost;
        private final double latestPrice;
        private final double costBasis;
        private final double marketValue;
        private final double unrealized;
        private final double unrealizedPct;
        private final boolean hasPrice;
        private final boolean hasEstimatedPrice;

        private AnnualSnapshotRow(
                String ticker,
                String securityName,
                String assetType,
                String currencyCode,
                double units,
                double averageCost,
                double latestPrice,
                double costBasis,
                double marketValue,
                double unrealized,
                double unrealizedPct,
                boolean hasPrice,
                boolean hasEstimatedPrice) {
            this.ticker = ticker;
            this.securityName = securityName;
            this.assetType = assetType;
            this.currencyCode = currencyCode;
            this.units = units;
            this.averageCost = averageCost;
            this.latestPrice = latestPrice;
            this.costBasis = costBasis;
            this.marketValue = marketValue;
            this.unrealized = unrealized;
            this.unrealizedPct = unrealizedPct;
            this.hasPrice = hasPrice;
            this.hasEstimatedPrice = hasEstimatedPrice;
        }
    }

    private static void writeAnnualPortfolioSnapshotTableHtml(
            FileWriter writer,
            List<AnnualSnapshotRow> rows,
            Map<String, Double> ratesToNok,
            int reportYear) throws IOException {

        int safeYear = Math.max(2000, Math.min(2100, reportYear));

        writer.write("<h2>PORTFOLIO OVERVIEW - 31.12." + safeYear + "</h2>\n");
        if (rows.isEmpty()) {
            writer.write("<p class=\"app-shell-note\">No holdings found at 31.12." + safeYear + ".</p>\n");
            return;
        }

        writer.write("<div class=\"table-wrap\">\n<table class=\"overview-table\">\n");
        ReportTemplateHelper.writeHtmlRow(writer, true,
            "Ticker", "Security", "Shares", "Avg Cost", "Price/Share", "Cost Basis", "Market Value", "Unrealized");

        LinkedHashMap<String, Double> totalCostBasisBuckets = new LinkedHashMap<>();
        LinkedHashMap<String, Double> totalMarketValueBuckets = new LinkedHashMap<>();
        LinkedHashMap<String, Double> totalUnrealizedBuckets = new LinkedHashMap<>();
        String previousAssetType = null;

        for (AnnualSnapshotRow row : rows) {
            addToCurrencyBuckets(totalCostBasisBuckets, row.currencyCode, row.costBasis);
            addToCurrencyBuckets(totalMarketValueBuckets, row.currencyCode, row.marketValue);
            addToCurrencyBuckets(totalUnrealizedBuckets, row.currencyCode, row.unrealized);

            String rowClass = isStockFundBoundary(previousAssetType, row.assetType) ? "asset-split" : null;
                String unrealizedText = row.hasPrice
                    ? signedSpan(HtmlFormatter.formatMoney(row.unrealized, row.currencyCode, 2) + " (" + HtmlFormatter.formatPercent(row.unrealizedPct, 2) + ")", row.unrealized)
                    : "-";

                String rowAttributes = "data-asset-group=\"" + escapeHtml(normalizeAssetBoundaryGroup(row.assetType)) + "\"";
                ReportTemplateHelper.writeHtmlRowWithClassAndAttributes(writer, rowClass, rowAttributes,
                    "<span class=\"ticker-scroll\">" + escapeHtml(row.ticker) + "</span>",
                    "<span class=\"security-scroll\">" + escapeHtml(row.securityName) + "</span>",
                    HtmlFormatter.formatUnits(row.units),
                    HtmlFormatter.formatMoney(row.averageCost, row.currencyCode, 2),
                    row.hasPrice ? HtmlFormatter.formatMoney(row.latestPrice, row.currencyCode, 2) : "-",
                    HtmlFormatter.formatMoney(row.costBasis, row.currencyCode, 2),
                    row.hasPrice ? HtmlFormatter.formatMoney(row.marketValue, row.currencyCode, 2) : "-",
                    unrealizedText);

            previousAssetType = row.assetType;
        }

        double totalCostBasisForPct = convertBucketsToTarget(totalCostBasisBuckets, DEFAULT_TOTAL_CURRENCY, ratesToNok);
        double totalUnrealizedForPct = convertBucketsToTarget(totalUnrealizedBuckets, DEFAULT_TOTAL_CURRENCY, ratesToNok);
        double totalUnrealizedPct = totalCostBasisForPct > 0.0 ? (totalUnrealizedForPct / totalCostBasisForPct) * 100.0 : 0.0;

        writer.write("<tr class=\"total-row\">\n");
        writer.write("    <td></td><td><strong>TOTAL</strong></td><td></td><td></td><td></td>\n");
        writer.write("    <td>" + renderConvertibleMoneyCell(totalCostBasisBuckets, 2, ratesToNok) + "</td>\n");
        writer.write("    <td>" + renderConvertibleMoneyCell(totalMarketValueBuckets, 2, ratesToNok) + "</td>\n");
        writer.write("    <td>"
            + signedWrapHtml(renderConvertibleMoneyCell(totalUnrealizedBuckets, 2, ratesToNok), totalUnrealizedForPct)
            + " "
            + signedSpan("(" + HtmlFormatter.formatPercent(totalUnrealizedPct, 2) + ")", totalUnrealizedForPct)
            + "</td>\n");
        writer.write("</tr>\n");
        writer.write("</table>\n</div>\n\n");
    }

    private static List<AnnualSnapshotRow> buildAnnualSnapshotRows(TransactionStore store, LocalDate snapshotDate) {
        Map<String, Security> securityByKey = buildSecurityLookupByKey(store);
        LinkedHashMap<String, Double> unitsBySecurity = new LinkedHashMap<>();
        HashMap<String, NavigableMap<LocalDate, Double>> priceSeriesCache = new HashMap<>();

        for (Events.UnitEvent event : store.getUnitEvents()) {
            if (event == null || event.tradeDate() == null || event.tradeDate().isAfter(snapshotDate)) {
                continue;
            }
            unitsBySecurity.merge(event.securityKey(), event.unitsDelta(), Double::sum);
        }

        ArrayList<AnnualSnapshotRow> rows = new ArrayList<>();
        for (Map.Entry<String, Double> entry : unitsBySecurity.entrySet()) {
            double units = entry.getValue();
            if (units <= EPSILON) {
                continue;
            }

            Security security = securityByKey.get(entry.getKey());
            if (security == null) {
                continue;
            }
            if (isTemporaryRightsSecurity(security)) {
                continue;
            }

            double averageCost = resolveAnnualSnapshotAverageCost(security, snapshotDate, units, priceSeriesCache);
            double costBasis = units * averageCost;
        PortfolioCalculator.PriceResolution priceResolution = PortfolioCalculator.resolvePriceAtDateDetailed(security, snapshotDate, priceSeriesCache);
        double price = priceResolution.getPrice();
            boolean hasPrice = price > 0.0;
            double marketValue = hasPrice ? units * price : 0.0;
            double unrealized = hasPrice ? marketValue - costBasis : 0.0;
            double unrealizedPct = costBasis > 0.0 ? (unrealized / costBasis) * 100.0 : 0.0;
        boolean hasEstimatedPrice = !hasPrice || priceResolution.isEstimated();

            rows.add(new AnnualSnapshotRow(
                    security.getTicker(),
                    security.getDisplayName(),
                    security.getAssetType().name(),
                    security.getCurrencyCode(),
                    units,
                    averageCost,
                    price,
                    costBasis,
                    marketValue,
                    unrealized,
                    unrealizedPct,
                    hasPrice,
                    hasEstimatedPrice
            ));
        }

        rows.sort(Comparator
                .comparingInt((AnnualSnapshotRow row) -> getAssetPriority(row.assetType))
                .thenComparing((AnnualSnapshotRow row) -> row.marketValue, Comparator.reverseOrder())
                .thenComparing(row -> row.securityName, String.CASE_INSENSITIVE_ORDER));

        return rows;
    }

    private static double resolveAnnualSnapshotAverageCost(
            Security security,
            LocalDate snapshotDate,
            double snapshotUnits,
            Map<String, NavigableMap<LocalDate, Double>> priceSeriesCache) {
        if (security == null || snapshotDate == null || snapshotUnits <= EPSILON) {
            return 0.0;
        }

        double reconstructedUnits = 0.0;
        double reconstructedCost = 0.0;

        for (Security.CurrentHoldingLot lot : security.getCurrentHoldingLotsSortedByDate()) {
            if (lot == null || lot.getTradeDate() == null || lot.getUnits() <= EPSILON) {
                continue;
            }
            if (lot.getTradeDate().isAfter(snapshotDate)) {
                continue;
            }

            reconstructedUnits += lot.getUnits();
            reconstructedCost += lot.getCostBasis();
        }

        double missingUnits = Math.max(0.0, snapshotUnits - reconstructedUnits);
        if (missingUnits > EPSILON) {
            for (Security.SaleTrade saleTrade : security.getSaleTradesSortedByDate()) {
                if (saleTrade == null || saleTrade.getTradeDate() == null || !saleTrade.getTradeDate().isAfter(snapshotDate)) {
                    continue;
                }

                double soldUnits = Math.max(0.0, saleTrade.getUnits());
                if (soldUnits <= EPSILON) {
                    continue;
                }

                double soldCostBasis = Math.max(0.0, saleTrade.getCostBasis());
                double restoredUnits = Math.min(missingUnits, soldUnits);
                double unitCost = soldCostBasis / soldUnits;

                reconstructedUnits += restoredUnits;
                reconstructedCost += restoredUnits * unitCost;
                missingUnits -= restoredUnits;

                if (missingUnits <= EPSILON) {
                    break;
                }
            }
        }

        if (reconstructedUnits > EPSILON && reconstructedCost > 0.0) {
            double reconstructedAvg = reconstructedCost / reconstructedUnits;
            if (reconstructedUnits + EPSILON < snapshotUnits) {
                reconstructedCost += (snapshotUnits - reconstructedUnits) * reconstructedAvg;
                reconstructedUnits = snapshotUnits;
            }
            return Math.max(0.0, reconstructedCost / Math.max(snapshotUnits, EPSILON));
        }

        double currentAverageCost = Math.max(0.0, security.getAverageCost());
        if (currentAverageCost > 0.0) {
            return currentAverageCost;
        }

        double fallbackPrice = PortfolioCalculator.resolvePriceAtDate(security, snapshotDate, priceSeriesCache);
        if (fallbackPrice > 0.0) {
            return fallbackPrice;
        }

        return 0.0;
    }

    private static void writeAnnualRealizedSummaryTableHtml(
            FileWriter writer,
            TransactionStore store,
            Map<String, Double> ratesToNok,
            int reportYear) throws IOException {

        int safeYear = Math.max(2000, Math.min(2100, reportYear));
        writer.write("<h2>REALIZED OVERVIEW - SALES IN " + safeYear + "</h2>\n");
        writer.write("<div class=\"table-wrap\">\n<table class=\"realized-table\">\n");
        ReportTemplateHelper.writeHtmlRow(writer, true, ReportTemplateHelper.buildTickerHeaderCell("realized-details-year"), "Security", "Cost Basis", "Sales Value", "Gain/Loss", "Dividends", "Total Return");

        LinkedHashMap<String, Double> totalSalesValueBuckets = new LinkedHashMap<>();
        LinkedHashMap<String, Double> totalCostBasisBuckets = new LinkedHashMap<>();
        LinkedHashMap<String, Double> totalRealizedGainBuckets = new LinkedHashMap<>();
        LinkedHashMap<String, Double> totalRealizedDividendsBuckets = new LinkedHashMap<>();

        ArrayList<Security> sortedSecurities = new ArrayList<>(store.getSecurities());
        sortedSecurities.sort(Comparator
                .comparingInt((Security s) -> getAssetPriority(s.getAssetType().name()))
                .thenComparing(Security::getDisplayName, String.CASE_INSENSITIVE_ORDER));

        String previousAssetType = null;
        int detailsIndex = 0;
        int includedRows = 0;

        for (Security security : sortedSecurities) {
            List<Security.SaleTrade> yearlySales = security.getSaleTradesSortedByDate().stream()
                    .filter(trade -> trade != null && trade.getTradeDate() != null && trade.getTradeDate().getYear() == safeYear)
                    .toList();

            double realizedDividends = security.getAllDividendEventsSortedByDate().stream()
                    .filter(event -> event != null && event.getTradeDate() != null && event.getTradeDate().getYear() == safeYear)
                    .mapToDouble(Security.DividendEvent::getAmount)
                    .sum();

            if (yearlySales.isEmpty() && Math.abs(realizedDividends) < EPSILON) {
                continue;
            }

            double salesValue = yearlySales.stream().mapToDouble(Security.SaleTrade::getSaleValue).sum();
            double costBasis = yearlySales.stream().mapToDouble(Security.SaleTrade::getCostBasis).sum();
            double gain = yearlySales.stream().mapToDouble(Security.SaleTrade::getGainLoss).sum();

            double totalReturnValue = gain + realizedDividends;
            double rowTotalReturnPct = costBasis > 0.0 ? (totalReturnValue / costBasis) * 100.0 : 0.0;
            String currency = security.getCurrencyCode();
            String currentAssetType = security.getAssetType().name();
            String rowClass = isStockFundBoundary(previousAssetType, currentAssetType) ? "asset-split" : null;
                String totalReturnCombined = signedSpan(
                    HtmlFormatter.formatMoney(totalReturnValue, currency, 2)
                        + " (" + HtmlFormatter.formatPercent(rowTotalReturnPct, 2) + ")",
                    totalReturnValue);

            addToCurrencyBuckets(totalSalesValueBuckets, currency, salesValue);
            addToCurrencyBuckets(totalCostBasisBuckets, currency, costBasis);
            addToCurrencyBuckets(totalRealizedGainBuckets, currency, gain);
            addToCurrencyBuckets(totalRealizedDividendsBuckets, currency, realizedDividends);

            String detailsRowId = "realized-year-details-" + detailsIndex;
                String rowAttributes = "data-asset-group=\"" + escapeHtml(normalizeAssetBoundaryGroup(currentAssetType)) + "\"";
                String tickerToggle = "<button class=\"details-link-btn\" data-target=\"" + detailsRowId + "\" onclick=\"toggleOverviewDetails('" + detailsRowId + "', null)\"><span class=\"ticker-scroll\">" + escapeHtml(security.getTicker()) + "</span></button>";
                String securityToggle = "<button class=\"details-link-btn\" data-target=\"" + detailsRowId + "\" onclick=\"toggleOverviewDetails('" + detailsRowId + "', null)\"><span class=\"security-scroll\">" + escapeHtml(security.getDisplayName()) + "</span></button>";
                ReportTemplateHelper.writeHtmlRowWithClassAndAttributes(writer, rowClass, rowAttributes,
                    tickerToggle,
                    securityToggle,
                    HtmlFormatter.formatMoney(costBasis, currency, 2),
                    HtmlFormatter.formatMoney(salesValue, currency, 2),
                    signedSpan(HtmlFormatter.formatMoney(gain, currency, 2), gain),
                    signedSpan(HtmlFormatter.formatMoney(realizedDividends, currency, 2), realizedDividends),
                    totalReturnCombined);

            writer.write("<tr id=\"" + detailsRowId + "\" class=\"details-row\" data-group=\"realized-details-year\">\n");
            writer.write("    <td class=\"details-cell\" colspan=\"7\">\n");
            writer.write(buildRealizedSaleTradesDetailsHtml(security, safeYear));
            writer.write("    </td>\n");
            writer.write("</tr>\n");

            previousAssetType = currentAssetType;
            detailsIndex++;
            includedRows++;
        }

        if (includedRows == 0) {
            writer.write("<tr><td colspan=\"7\" class=\"app-shell-note\">No sales or dividends were recorded for " + safeYear + ".</td></tr>\n");
            writer.write("</table>\n</div>\n\n");
            return;
        }

        double totalCostBasisForPct = convertBucketsToTarget(totalCostBasisBuckets, DEFAULT_TOTAL_CURRENCY, ratesToNok);
        double totalRealizedGainForPct = convertBucketsToTarget(totalRealizedGainBuckets, DEFAULT_TOTAL_CURRENCY, ratesToNok);
        double totalRealizedDividendsForPct = convertBucketsToTarget(totalRealizedDividendsBuckets, DEFAULT_TOTAL_CURRENCY, ratesToNok);
        LinkedHashMap<String, Double> totalRealizedReturnBuckets = sumCurrencyBuckets(totalRealizedGainBuckets, totalRealizedDividendsBuckets);
        double totalRealizedReturnForPct = totalRealizedGainForPct + totalRealizedDividendsForPct;
        double totalReturnPct = totalCostBasisForPct > 0
            ? (totalRealizedReturnForPct / totalCostBasisForPct) * 100.0
            : 0.0;

        writer.write("<tr class=\"total-row\">\n");
        writer.write("    <td></td><td><strong>TOTAL</strong></td>\n");
        writer.write("    <td>" + renderConvertibleMoneyCell(totalCostBasisBuckets, 2, ratesToNok) + "</td>\n");
        writer.write("    <td>" + renderConvertibleMoneyCell(totalSalesValueBuckets, 2, ratesToNok) + "</td>\n");
        writer.write("    <td>" + signedWrapHtml(renderConvertibleMoneyCell(totalRealizedGainBuckets, 2, ratesToNok), totalRealizedGainForPct) + "</td>\n");
        writer.write("    <td>" + signedWrapHtml(renderConvertibleMoneyCell(totalRealizedDividendsBuckets, 2, ratesToNok), totalRealizedDividendsForPct) + "</td>\n");
        writer.write("    <td>"
            + signedWrapHtml(renderConvertibleMoneyCell(totalRealizedReturnBuckets, 2, ratesToNok), totalRealizedReturnForPct)
            + " "
            + signedSpan("(" + HtmlFormatter.formatPercent(totalReturnPct, 2) + ")", totalRealizedReturnForPct)
            + "</td>\n");
        writer.write("</tr>\n");
        writer.write("</table>\n</div>\n\n");
    }

    private static ReportConfig resolveReportConfig() {
        String rawType = System.getProperty("portfolio.report.type", REPORT_TYPE_STANDARD);
        String reportType = rawType == null ? REPORT_TYPE_STANDARD : rawType.trim().toLowerCase(Locale.ROOT);
        if (!REPORT_TYPE_ANNUAL.equals(reportType)) {
            reportType = REPORT_TYPE_STANDARD;
        }

        int defaultYear = LocalDate.now().getYear();
        int reportYear;
        try {
            reportYear = Integer.parseInt(System.getProperty("portfolio.report.year", String.valueOf(defaultYear)).trim());
        } catch (Exception ignored) {
            reportYear = defaultYear;
        }
        reportYear = Math.max(2000, Math.min(2100, reportYear));

        String benchmarkTicker = System.getProperty("portfolio.report.benchmark", "^OSEAX");
        if (benchmarkTicker == null || benchmarkTicker.isBlank()) {
            benchmarkTicker = "^OSEAX";
        }

        return new ReportConfig(reportType, reportYear, benchmarkTicker.trim());
    }

    private static void writeHeaderSummaryHtml(FileWriter writer, HeaderSummary s, List<OverviewRow> overviewRows, TransactionStore store, Map<String, Double> ratesToNok) throws IOException {
        String bestClass = signedClass(s.bestReturn);
        String worstClass = signedClass(s.worstReturn);
        String bestPctLabel = "N/A";
        String worstPctLabel = "N/A";
        double bestPctValue = Double.NEGATIVE_INFINITY;
        double worstPctValue = Double.POSITIVE_INFINITY;
        double bestPctReturnAmount = 0.0;
        double worstPctReturnAmount = 0.0;
        String bestPctCurrency = DEFAULT_TOTAL_CURRENCY;
        String worstPctCurrency = DEFAULT_TOTAL_CURRENCY;
        for (OverviewRow row : overviewRows) {
            if (row == null || !Double.isFinite(row.totalReturnPct)) {
                continue;
            }
            if (row.totalReturnPct > bestPctValue) {
                bestPctValue = row.totalReturnPct;
                bestPctLabel = row.securityDisplayName;
                bestPctReturnAmount = row.totalReturn;
                bestPctCurrency = normalizeCurrencyCode(row.currencyCode);
            }
            if (row.totalReturnPct < worstPctValue) {
                worstPctValue = row.totalReturnPct;
                worstPctLabel = row.securityDisplayName;
                worstPctReturnAmount = row.totalReturn;
                worstPctCurrency = normalizeCurrencyCode(row.currencyCode);
            }
        }
        boolean hasPctExtremes = Double.isFinite(bestPctValue) && Double.isFinite(worstPctValue);
        String bestPctClass = hasPctExtremes ? signedClass(bestPctValue) : "";
        String worstPctClass = hasPctExtremes ? signedClass(worstPctValue) : "";

        writer.write("<section class=\"report-hero annual-hero\">\n");
        writer.write("<div class=\"hero-title annual-hero-header\">\n");
        writer.write("<h1>Portfolio Report</h1>\n");
        writer.write("<div class=\"hero-meta\">\n");
        writer.write("<span class=\"meta-chip\">Date: <strong id=\"report-date-value\">" + escapeHtml(s.generatedDate) + "</strong></span>\n");
        writer.write("<span class=\"meta-chip\">Files: <strong>" + s.fileCount + "</strong></span>\n");
        writer.write("<span class=\"meta-chip\">Transactions: <strong>" + s.transactionCount + "</strong></span>\n");
        writer.write("<span class=\"meta-chip\">Holdings: <strong>" + s.holdingsCount + "</strong></span>\n");
        writer.write("<span class=\"meta-chip\">Currency: <strong><input id=\"portfolio-currency-input\" class=\"currency-input\" type=\"text\" value=\"" + DEFAULT_TOTAL_CURRENCY + "\" maxlength=\"3\" autocomplete=\"off\" spellcheck=\"false\" title=\"Skriv valutakode (f.eks. NOK, USD) og trykk Enter\"></strong></span>\n");
        writer.write("<button id=\"refresh-prices-btn\" class=\"hero-refresh-btn\" type=\"button\">Update</button>\n");
        writer.write("<button id=\"report-theme-toggle\" class=\"hero-theme-btn\" type=\"button\">Dark mode</button>\n");
        writer.write("</div>\n");
        writer.write("<div id=\"refresh-prices-status\" class=\"price-refresh-status\"></div>\n");
        writer.write("</div>\n");
        writer.write("</section>\n");
        writer.write("<section class=\"annual-kpi-deck\">\n");
        writer.write("<h2 class=\"annual-kpi-deck-title\">Portfolio Highlights</h2>\n");
        writer.write("<div class=\"annual-summary-grid\">\n");
        LinkedHashMap<String, Double> totalMarketBuckets = new LinkedHashMap<>();
        LinkedHashMap<String, Double> totalReturnBuckets = new LinkedHashMap<>();
        LinkedHashMap<String, Double> totalUnrealizedBuckets = new LinkedHashMap<>();
        LinkedHashMap<String, Double> totalRealizedBuckets = new LinkedHashMap<>();
        LinkedHashMap<String, Double> totalDividendsBuckets = new LinkedHashMap<>();
        LinkedHashMap<String, Double> totalCostBasisBuckets = new LinkedHashMap<>();
        LinkedHashMap<String, Double> totalHistoricalCostBuckets = new LinkedHashMap<>();
        LinkedHashMap<String, Double> soldOnlyReturnBuckets = new LinkedHashMap<>();
        LinkedHashMap<String, Double> soldOnlyDividendsBuckets = new LinkedHashMap<>();
        LinkedHashMap<String, Double> soldOnlyHistoricalCostBuckets = new LinkedHashMap<>();
        Set<String> activeSecurityKeys = new LinkedHashSet<>();
        for (OverviewRow row : overviewRows) {
            activeSecurityKeys.add(row.securityKey);
            addToCurrencyBuckets(totalMarketBuckets, row.currencyCode, row.marketValue);
            addToCurrencyBuckets(totalReturnBuckets, row.currencyCode, row.totalReturn);
            addToCurrencyBuckets(totalUnrealizedBuckets, row.currencyCode, row.unrealized);
            addToCurrencyBuckets(totalRealizedBuckets, row.currencyCode, row.realized);
            addToCurrencyBuckets(totalDividendsBuckets, row.currencyCode, row.dividends);
            addToCurrencyBuckets(totalCostBasisBuckets, row.currencyCode, row.positionCostBasis);
            addToCurrencyBuckets(totalHistoricalCostBuckets, row.currencyCode, row.historicalCostBasis);
        }

        // Include realized overview securities that are not currently active in portfolio.
        // This mirrors the realized overview formula exactly to keep header total in sync.
        for (Security security : getSortedSoldSecurities(store)) {
            String securityKey = getTrackingSecurityKey(security);
            if (securityKey.isBlank() || activeSecurityKeys.contains(securityKey)) {
                continue;
            }

            double realizedDividends = security.isFullyRealized() ? security.getDividends() : 0.0;
            double realizedOnlyReturn = security.getRealizedGain() + realizedDividends;
            double realizedCostBasis = security.getRealizedCostBasis();
            if (Math.abs(realizedOnlyReturn) < 1e-9 && Math.abs(realizedCostBasis) < 1e-9) {
                continue;
            }

            String currency = security.getCurrencyCode();
            addToCurrencyBuckets(totalReturnBuckets, currency, realizedOnlyReturn);
            addToCurrencyBuckets(totalDividendsBuckets, currency, realizedDividends);
            addToCurrencyBuckets(totalHistoricalCostBuckets, currency, realizedCostBasis);
            addToCurrencyBuckets(soldOnlyReturnBuckets, currency, realizedOnlyReturn);
            addToCurrencyBuckets(soldOnlyDividendsBuckets, currency, realizedDividends);
            addToCurrencyBuckets(soldOnlyHistoricalCostBuckets, currency, realizedCostBasis);
        }

        LinkedHashMap<String, Double> cashBuckets = new LinkedHashMap<>();
        cashBuckets.put(DEFAULT_TOTAL_CURRENCY, store.getCurrentCashHoldings());
        LinkedHashMap<String, Double> portfolioValueBuckets = sumCurrencyBuckets(totalMarketBuckets, cashBuckets);

        double totalReturnInDefaultCurrency = convertBucketsToTarget(totalReturnBuckets, DEFAULT_TOTAL_CURRENCY, ratesToNok);
        double totalHistoricalCostInDefaultCurrency = convertBucketsToTarget(totalHistoricalCostBuckets, DEFAULT_TOTAL_CURRENCY, ratesToNok);
        double totalReturnPct = totalHistoricalCostInDefaultCurrency > 0.0
            ? (totalReturnInDefaultCurrency / totalHistoricalCostInDefaultCurrency) * 100.0
            : 0.0;
        double totalUnrealizedInDefaultCurrency = convertBucketsToTarget(totalUnrealizedBuckets, DEFAULT_TOTAL_CURRENCY, ratesToNok);
        double totalRealizedInDefaultCurrency = convertBucketsToTarget(totalRealizedBuckets, DEFAULT_TOTAL_CURRENCY, ratesToNok);
        double totalCostBasisInDefaultCurrency = convertBucketsToTarget(totalCostBasisBuckets, DEFAULT_TOTAL_CURRENCY, ratesToNok);
        double totalUnrealizedPct = totalCostBasisInDefaultCurrency > 0.0 ? (totalUnrealizedInDefaultCurrency / totalCostBasisInDefaultCurrency) * 100.0 : 0.0;
        double totalRealizedPct = totalCostBasisInDefaultCurrency > 0.0 ? (totalRealizedInDefaultCurrency / totalCostBasisInDefaultCurrency) * 100.0 : 0.0;
        LinkedHashMap<String, Double> dayChangeBuckets = new LinkedHashMap<>();
        LinkedHashMap<String, Double> previousDayValueBuckets = new LinkedHashMap<>();
        for (OverviewRow row : overviewRows) {
            if (row == null || row.units <= 0.0 || row.latestPrice <= 0.0 || row.previousClose <= 0.0) {
                continue;
            }
            double changeAmount = row.units * (row.latestPrice - row.previousClose);
            double previousDayValue = row.units * row.previousClose;
            addToCurrencyBuckets(dayChangeBuckets, row.currencyCode, changeAmount);
            addToCurrencyBuckets(previousDayValueBuckets, row.currencyCode, previousDayValue);
        }
        double dayChangeNok = convertBucketsToTarget(dayChangeBuckets, DEFAULT_TOTAL_CURRENCY, ratesToNok);
        double previousDayValueNok = convertBucketsToTarget(previousDayValueBuckets, DEFAULT_TOTAL_CURRENCY, ratesToNok);
        double dayChangePct = previousDayValueNok > 0.0 ? (dayChangeNok / previousDayValueNok) * 100.0 : 0.0;
        PortfolioCalculator.OneYearChangeSummary oneYearChangeSummary = PortfolioCalculator.buildStandardTrailingOneYearChangeSummary(store, ratesToNok);
        LinkedHashMap<String, Double> oneYearChangeBuckets = oneYearChangeSummary.hasData
            ? singleCurrencyBuckets(DEFAULT_TOTAL_CURRENCY, oneYearChangeSummary.returnNok)
            : new LinkedHashMap<>();
        String totalClass = signedClass(totalReturnInDefaultCurrency);
        String unrealizedClass = signedClass(totalUnrealizedInDefaultCurrency);
        String realizedClass = signedClass(totalRealizedInDefaultCurrency);
        String dayChangeClass = signedClass(dayChangeNok);
        String oneYearChangeClass = oneYearChangeSummary.hasData ? signedClass(oneYearChangeSummary.returnNok) : "";

        writer.write("<article class=\"kpi-card\"><div class=\"kpi-label\">Total Market Value</div><div id=\"hero-total-market-value\" class=\"kpi-value js-convert-money\" data-buckets=\""
            + escapeHtml(toBucketsJson(totalMarketBuckets)) + "\" data-decimals=\"0\">"
            + formatBucketsInTarget(totalMarketBuckets, DEFAULT_TOTAL_CURRENCY, 0, ratesToNok) + "</div></article>\n");

        writer.write("<article class=\"kpi-card\"><div class=\"kpi-label\">Portfolio Value</div><div id=\"hero-portfolio-value\" class=\"kpi-value js-convert-money\" data-buckets=\""
            + escapeHtml(toBucketsJson(portfolioValueBuckets)) + "\" data-decimals=\"0\">"
            + formatBucketsInTarget(portfolioValueBuckets, DEFAULT_TOTAL_CURRENCY, 0, ratesToNok) + "</div></article>\n");

        writer.write("<article class=\"kpi-card\"><div class=\"kpi-label\">Total Return</div><div id=\"hero-total-return-value\" class=\"kpi-value js-convert-money " + totalClass
            + "\" data-buckets=\"" + escapeHtml(toBucketsJson(totalReturnBuckets))
            + "\" data-sold-only-return-buckets=\"" + escapeHtml(toBucketsJson(soldOnlyReturnBuckets))
            + "\" data-sold-only-historical-buckets=\"" + escapeHtml(toBucketsJson(soldOnlyHistoricalCostBuckets))
            + "\" data-decimals=\"0\">"
            + formatBucketsInTarget(totalReturnBuckets, DEFAULT_TOTAL_CURRENCY, 0, ratesToNok)
            + "</div><div id=\"hero-total-return-pct\" class=\"kpi-label " + totalClass + "\">" + HtmlFormatter.formatPercent(totalReturnPct) + "</div></article>\n");

        writer.write("<article class=\"kpi-card\"><div class=\"kpi-label\">Dividends</div><div id=\"hero-dividends-value\" class=\"kpi-value js-convert-money\" data-buckets=\""
            + escapeHtml(toBucketsJson(totalDividendsBuckets))
            + "\" data-sold-only-dividends-buckets=\"" + escapeHtml(toBucketsJson(soldOnlyDividendsBuckets))
            + "\" data-decimals=\"0\">"
            + formatBucketsInTarget(totalDividendsBuckets, DEFAULT_TOTAL_CURRENCY, 0, ratesToNok)
            + "</div></article>\n");

        writer.write("<article class=\"kpi-card\"><div class=\"kpi-label\">Unrealized Return</div><div id=\"hero-unrealized-value\" class=\"kpi-value js-convert-money " + unrealizedClass + "\" data-buckets=\""
            + escapeHtml(toBucketsJson(totalUnrealizedBuckets))
            + "\" data-decimals=\"0\">"
            + formatBucketsInTarget(totalUnrealizedBuckets, DEFAULT_TOTAL_CURRENCY, 0, ratesToNok)
            + "</div><div id=\"hero-unrealized-pct\" class=\"kpi-label " + unrealizedClass + "\">" + HtmlFormatter.formatPercent(totalUnrealizedPct) + "</div></article>\n");

        writer.write("<article class=\"kpi-card\"><div class=\"kpi-label\">Realized Return</div><div id=\"hero-realized-value\" class=\"kpi-value js-convert-money " + realizedClass + "\" data-buckets=\""
            + escapeHtml(toBucketsJson(totalRealizedBuckets))
            + "\" data-decimals=\"0\">"
            + formatBucketsInTarget(totalRealizedBuckets, DEFAULT_TOTAL_CURRENCY, 0, ratesToNok)
            + "</div><div id=\"hero-realized-pct\" class=\"kpi-label " + realizedClass + "\">" + HtmlFormatter.formatPercent(totalRealizedPct) + "</div></article>\n");

        writer.write("<article class=\"kpi-card\"><div class=\"kpi-label\">Day Change</div><div id=\"hero-day-change-value\" class=\"kpi-value js-convert-money " + dayChangeClass + "\" data-buckets=\""
            + escapeHtml(toBucketsJson(dayChangeBuckets))
            + "\" data-decimals=\"0\">"
            + formatBucketsInTarget(dayChangeBuckets, DEFAULT_TOTAL_CURRENCY, 0, ratesToNok)
            + "</div><div id=\"hero-day-change-pct\" class=\"kpi-label " + dayChangeClass + "\">" + HtmlFormatter.formatPercent(dayChangePct) + "</div></article>\n");

        if (oneYearChangeSummary.hasData) {
            writer.write("<article class=\"kpi-card\"><div class=\"kpi-label\">One year change</div><div id=\"hero-one-year-change-value\" class=\"kpi-value js-convert-money " + oneYearChangeClass + "\" data-buckets=\""
                + escapeHtml(toBucketsJson(oneYearChangeBuckets))
                + "\" data-decimals=\"0\">"
                + formatBucketsInTarget(oneYearChangeBuckets, DEFAULT_TOTAL_CURRENCY, 0, ratesToNok)
                + "</div><div id=\"hero-one-year-change-pct\" class=\"kpi-label " + oneYearChangeClass + "\">" + HtmlFormatter.formatPercent(oneYearChangeSummary.returnPct, 2) + "</div></article>\n");
        } else {
            writer.write("<article class=\"kpi-card\"><div class=\"kpi-label\">One year change</div><div class=\"kpi-value\">-</div><div class=\"kpi-label\">-</div></article>\n");
        }

        writer.write("<article id=\"cash-holdings-card\" class=\"kpi-card\"><div class=\"cash-holdings-header\"><div class=\"kpi-label\">Cash Holdings</div><button id=\"cash-holdings-add-btn\" class=\"cash-holdings-add-btn\" type=\"button\">Add</button></div><div id=\"cash-holdings-total\" class=\"kpi-value js-convert-money\" data-buckets=\""
            + escapeHtml(toBucketsJson(cashBuckets)) + "\" data-base-buckets=\"" + escapeHtml(toBucketsJson(cashBuckets)) + "\" data-decimals=\"0\">"
            + formatBucketsInTarget(cashBuckets, DEFAULT_TOTAL_CURRENCY, 0, ratesToNok) + "</div><div id=\"manual-cash-holdings-list\" class=\"manual-cash-holdings-list\"></div></article>\n");

        writer.write("<article class=\"kpi-card kpi-card-bestworst\"><div class=\"kpi-label\">Best / Worst</div><div class=\"performer " + bestClass + "\"><strong>"
            + escapeHtml(s.bestLabel)
            + "</strong><span class=\"performer-metrics\"><span class=\"js-convert-money\" data-buckets=\""
            + escapeHtml(toBucketsJson(singleCurrencyBuckets(s.bestCurrencyCode, s.bestReturn)))
            + "\" data-decimals=\"0\">"
            + formatBucketsInTarget(singleCurrencyBuckets(s.bestCurrencyCode, s.bestReturn), DEFAULT_TOTAL_CURRENCY, 0, ratesToNok)
            + "</span> | " + HtmlFormatter.formatPercent(s.bestReturnPct)
            + "</span></div><div class=\"performer " + worstClass + "\"><strong>" + escapeHtml(s.worstLabel)
            + "</strong><span class=\"performer-metrics\"><span class=\"js-convert-money\" data-buckets=\""
            + escapeHtml(toBucketsJson(singleCurrencyBuckets(s.worstCurrencyCode, s.worstReturn)))
            + "\" data-decimals=\"0\">"
            + formatBucketsInTarget(singleCurrencyBuckets(s.worstCurrencyCode, s.worstReturn), DEFAULT_TOTAL_CURRENCY, 0, ratesToNok)
            + "</span> | " + HtmlFormatter.formatPercent(s.worstReturnPct) + "</span></div></article>\n");

        if (hasPctExtremes) {
            LinkedHashMap<String, Double> bestPctBuckets = singleCurrencyBuckets(bestPctCurrency, bestPctReturnAmount);
            LinkedHashMap<String, Double> worstPctBuckets = singleCurrencyBuckets(worstPctCurrency, worstPctReturnAmount);
            writer.write("<article class=\"kpi-card kpi-card-bestworst\"><div class=\"kpi-label\">Best / Worst %</div><div class=\"performer " + bestPctClass + "\"><strong>"
                + escapeHtml(bestPctLabel)
                + "</strong><span class=\"performer-metrics\">"
                + renderConvertibleMoneyCell(bestPctBuckets, 0, ratesToNok)
                + " | "
                + HtmlFormatter.formatPercent(bestPctValue)
                + "</span></div><div class=\"performer " + worstPctClass + "\"><strong>"
                + escapeHtml(worstPctLabel)
                + "</strong><span class=\"performer-metrics\">"
                + renderConvertibleMoneyCell(worstPctBuckets, 0, ratesToNok)
                + " | "
                + HtmlFormatter.formatPercent(worstPctValue)
                + "</span></div></article>\n");
        } else {
            writer.write("<article class=\"kpi-card kpi-card-bestworst\"><div class=\"kpi-label\">Best / Worst %</div><div class=\"performer\"><strong>N/A</strong><span class=\"performer-metrics\">No percentage return data available.</span></div></article>\n");
        }
        writer.write("</div>\n");
        String valueTimelineSvg = PortfolioCalculator.buildStandardPortfolioValueSparklineSvg(store, ratesToNok);
        String returnTimelineSvg = PortfolioCalculator.buildStandardPortfolioReturnSparklineSvg(store, ratesToNok);

        writer.write("</section>\n");
        writeStandardAnalyticsSectionHtml(writer, store, ratesToNok);
        writer.write("<section class=\"annual-graphs-section\">\n");
        writer.write("<div class=\"annual-graphs-heading\"><div class=\"timeline-title-row\"><h2>Yearly Trend</h2><button type=\"button\" class=\"timeline-info-btn\" aria-label=\"Show calculation info\" title=\"Show calculation info\">i</button></div></div>\n");
        writer.write("<div class=\"annual-graphs-row\">\n");
        writer.write("<article class=\"annual-graph-card\">\n");
        writer.write("<h3>Portfolio Value</h3>\n");
        writer.write("<div class=\"annual-graph-content\">\n");
        if (valueTimelineSvg != null && !valueTimelineSvg.isBlank()) {
            writer.write(valueTimelineSvg);
        } else {
            writer.write("<div class=\"hero-side-note\">Timeline data not available yet for this dataset.</div>");
        }
        writer.write("</div>\n");
        writer.write("</article>\n");

        writer.write("<article class=\"annual-graph-card\">\n");
        writer.write("<h3>Portfolio Return</h3>\n");
        writer.write("<div class=\"annual-graph-content\">\n");
        if (returnTimelineSvg != null && !returnTimelineSvg.isBlank()) {
            writer.write(returnTimelineSvg);
        } else {
            writer.write("<div class=\"hero-side-note\">Return timeline is not available for this dataset.</div>");
        }
        writer.write("</div>\n");
        writer.write("</article>\n");
        writer.write("</div>\n");

        writer.write("<div class=\"timeline-info-overlay\" hidden><div class=\"timeline-info-dialog\" role=\"dialog\" aria-modal=\"true\" aria-label=\"Portfolio timeline info\"><div class=\"timeline-info-header\"><h4>Portfolio Value Timeline - Info</h4><button type=\"button\" class=\"timeline-info-close\" aria-label=\"Close\">×</button></div><div class=\"timeline-info-body\"><p>This chart is an indicative estimate based on imported transactions, cash snapshots, and historical prices.</p><ul><li><strong>Value:</strong> Estimated portfolio value at each month-end in the selected display currency.</li><li><strong>Return (<span class=\"js-report-currency-code\">NOK</span>):</strong> Cumulative cashflow-adjusted return (TWR-based) for the selected range, expressed in <span class=\"js-report-currency-code\">NOK</span> from the range start value.</li><li><strong>Return (%):</strong> Cumulative time-weighted return (TWR) from the selected range start.</li><li><strong>External cash flows:</strong> Deposits, withdrawals, and transfers are neutralized in return calculations so contributions/withdrawals do not count as performance.</li><li><strong>Pricing:</strong> Historical close prices are primarily fetched from Yahoo Finance. If data points are missing, transaction-derived fallback pricing is used.</li><li><strong>Disclaimer:</strong> Values are for analysis and may differ from official broker reporting.</li></ul></div></div></div>\n");
        writer.write("</section>\n");
    }

    private static void writeStandardAnalyticsSectionHtml(
            FileWriter writer,
            TransactionStore store,
            Map<String, Double> ratesToNok) throws IOException {

        PortfolioCalculator.StandardAnalyticsSummary analytics = PortfolioCalculator.buildStandardAnalyticsSummary(
                store,
                ratesToNok,
                "^OSEAX"
        );

        if (!analytics.hasAnalytics && !analytics.hasMonteCarlo) {
            return;
        }

        writer.write("<section class=\"annual-kpi-deck\">\n");
        writer.write("<h2 class=\"annual-kpi-deck-title\">Risk Analytics</h2>\n");
        writer.write("<div class=\"annual-summary-grid\">\n");

        if (analytics.hasAnalytics) {
            writer.write("<article class=\"kpi-card annual-summary-card\"><h4>Volatility (Ann.)</h4><div class=\"annual-summary-value\">"
                    + HtmlFormatter.formatPercent(analytics.annualizedVolatilityPct, 2)
                    + "</div><div class=\"annual-summary-sub\">Annualized from monthly return variance</div></article>\n");

            writer.write("<article class=\"kpi-card annual-summary-card\"><h4>Sharpe Ratio</h4><div class=\"annual-summary-value\">"
                    + String.format(Locale.US, "%.2f", analytics.sharpeRatio)
                    + "</div><div class=\"annual-summary-sub\">Risk-adjusted return (monthly, annualized)</div></article>\n");

            if (analytics.hasBeta) {
                writer.write("<article class=\"kpi-card annual-summary-card\"><h4>Beta vs " + escapeHtml(analytics.benchmarkTicker) + "</h4><div class=\"annual-summary-value\">"
                        + String.format(Locale.US, "%.2f", analytics.beta)
                        + "</div><div class=\"annual-summary-sub\">Sensitivity vs benchmark monthly returns</div></article>\n");
            } else {
                writer.write("<article class=\"kpi-card annual-summary-card\"><h4>Beta vs " + escapeHtml(analytics.benchmarkTicker) + "</h4><div class=\"annual-summary-value\">N/A</div><div class=\"annual-summary-sub\">Insufficient benchmark overlap</div></article>\n");
            }
        }

        if (analytics.hasMonteCarlo) {
            LinkedHashMap<String, Double> medianBuckets = singleCurrencyBuckets(DEFAULT_TOTAL_CURRENCY, analytics.monteCarloMedianEndValueNok);
            LinkedHashMap<String, Double> p10Buckets = singleCurrencyBuckets(DEFAULT_TOTAL_CURRENCY, analytics.monteCarloP10EndValueNok);
            LinkedHashMap<String, Double> p90Buckets = singleCurrencyBuckets(DEFAULT_TOTAL_CURRENCY, analytics.monteCarloP90EndValueNok);
            writer.write("<article class=\"kpi-card annual-summary-card\"><h4>Monte Carlo (" + analytics.monteCarloHorizonMonths + "m)</h4><div class=\"annual-summary-value\">"
                    + renderConvertibleMoneyCell(medianBuckets, 0, ratesToNok)
                    + "</div><div class=\"annual-summary-sub\">Median terminal value (" + analytics.monteCarloIterations + " iterations)</div>"
                    + "<div class=\"annual-summary-sub\">P10: " + renderConvertibleMoneyCell(p10Buckets, 0, ratesToNok) + " | P90: " + renderConvertibleMoneyCell(p90Buckets, 0, ratesToNok) + "</div></article>\n");
        }

        writer.write("</div>\n");
        writer.write("</section>\n");
    }

    private static void writeOverviewTableHtml(FileWriter writer, List<OverviewRow> rows, TransactionStore store, Map<String, Double> ratesToNok) throws IOException {
        writer.write("<h2>PORTFOLIO OVERVIEW - CURRENT HOLDINGS</h2>\n");
        Map<String, Security> securityByKey = buildSecurityLookupByKey(store);
        Map<String, Security.FundamentalsSnapshot> fundamentalsByKey = new HashMap<>();
        for (OverviewRow row : rows) {
            Security security = securityByKey.get(row.securityKey);
            if (security == null) {
                continue;
            }
            fundamentalsByKey.put(row.securityKey, security.getFundamentalsSnapshot());
        }

        writer.write("<section class=\"annual-graphs-section total-return-graphs-section\">\n");
        writer.write("<div class=\"annual-graphs-heading\"><h2>Total Return</h2></div>\n");
        writer.write("<div class=\"annual-graphs-row\">\n");
        writer.write("<article class=\"annual-graph-card overview-chart total-return-chart\"><h3 class=\"js-total-return-money-title\">Total Return (" + DEFAULT_TOTAL_CURRENCY + ")</h3>\n");
        writer.write(ChartBuilder.buildOverviewBarChartSvg(rows, false, ratesToNok));
        writer.write("</article>\n");
        writer.write("<article class=\"annual-graph-card overview-chart total-return-chart\"><h3>Total Return (%)</h3>\n");
        writer.write(ChartBuilder.buildOverviewBarChartSvg(rows, true, ratesToNok));
        writer.write("</article>\n");
        writer.write("</div>\n");
        writer.write("</section>\n");

        writer.write("<div class=\"overview-mode-shell\" role=\"tablist\" aria-label=\"Portfolio table mode\">\n");
        writer.write("<button type=\"button\" class=\"overview-mode-btn is-active\" data-overview-mode=\"summary\">Summary</button>\n");
        writer.write("<button type=\"button\" class=\"overview-mode-btn\" data-overview-mode=\"holdings\">Holdings</button>\n");
        writer.write("<button type=\"button\" class=\"overview-mode-btn\" data-overview-mode=\"fundamentals\">Fundamentals</button>\n");
        writer.write("<button type=\"button\" id=\"overview-details-toggle\" class=\"overview-mode-btn overview-details-toggle-btn\" data-detail-label=\"Open all details\" data-detail-group=\"overview-details\">Open all details ▸</button>\n");
        writer.write("</div>\n");
        writer.write("<div class=\"table-wrap overview-table-wrap js-overview-mode-panel\" data-overview-mode-panel=\"summary\">\n<table class=\"overview-table overview-summary-table\">\n");
        ReportTemplateHelper.writeHtmlRow(writer, true,
            "Ticker", "Security", "Change %", "Change", "Day Chart", "52-Wk Range", "Shares", "Avg Cost", "Last Price",
                "Cost Basis", "Market Value");

        LinkedHashMap<String, Double> totalMarketValueBuckets = new LinkedHashMap<>();
        LinkedHashMap<String, Double> totalCostBasisBuckets = new LinkedHashMap<>();
        LinkedHashMap<String, Double> totalUnrealizedBuckets = new LinkedHashMap<>();
        LinkedHashMap<String, Double> totalRealizedBuckets = new LinkedHashMap<>();
        LinkedHashMap<String, Double> totalDividendsBuckets = new LinkedHashMap<>();
        LinkedHashMap<String, Double> totalHistoricalCostBuckets = new LinkedHashMap<>();
        LinkedHashMap<String, Double> totalDayChangeBuckets = new LinkedHashMap<>();
        LinkedHashMap<String, Double> totalPrevCloseValueBuckets = new LinkedHashMap<>();
        String previousAssetType = null;

        int detailsIndex = 0;
        for (OverviewRow row : rows) {
            addToCurrencyBuckets(totalMarketValueBuckets, row.currencyCode, row.marketValue);
            addToCurrencyBuckets(totalCostBasisBuckets, row.currencyCode, row.positionCostBasis);
            addToCurrencyBuckets(totalUnrealizedBuckets, row.currencyCode, row.unrealized);
            addToCurrencyBuckets(totalRealizedBuckets, row.currencyCode, row.realized);
            addToCurrencyBuckets(totalDividendsBuckets, row.currencyCode, row.dividends);
            addToCurrencyBuckets(totalHistoricalCostBuckets, row.currencyCode, row.historicalCostBasis);
            if (row.latestPrice > 0.0 && row.previousClose > 0.0 && row.units > 0.0) {
                double rowDayChangeValue = (row.latestPrice - row.previousClose) * row.units;
                double rowPrevCloseValue = row.previousClose * row.units;
                addToCurrencyBuckets(totalDayChangeBuckets, row.currencyCode, rowDayChangeValue);
                addToCurrencyBuckets(totalPrevCloseValueBuckets, row.currencyCode, rowPrevCloseValue);
            }
            String rowClass = isStockFundBoundary(previousAssetType, row.assetType) ? "asset-split" : null;
            String detailsRowId = "overview-details-" + detailsIndex;
            Security security = securityByKey.get(row.securityKey);
            String dayChangeCell = formatDayChangeCell(row.dayChangePct, row.hasDayChangePct);
            String dayChangeValueCell = formatDayChangeValueCell(row);
            String dayChartCell = formatDayChartCell(row);
            String fiftyTwoWeekRangeCell = format52WeekRangeCell(row, fundamentalsByKey.get(row.securityKey));

            String rowAttributes = "data-overview-row=\"1\""
                + " data-overview-security-key=\"" + escapeHtml(row.securityKey) + "\""
                + " data-ticker=\"" + escapeHtml(row.tickerText) + "\""
                + " data-asset-group=\"" + escapeHtml(normalizeAssetBoundaryGroup(row.assetType)) + "\""
                + " data-currency=\"" + escapeHtml(normalizeCurrencyCode(row.currencyCode)) + "\""
                + " data-units=\"" + String.format(Locale.US, "%.8f", row.units) + "\""
                + " data-position-cost-basis=\"" + String.format(Locale.US, "%.8f", row.positionCostBasis) + "\""
                + " data-realized=\"" + String.format(Locale.US, "%.8f", row.realized) + "\""
                + " data-dividends=\"" + String.format(Locale.US, "%.8f", row.dividends) + "\""
                + " data-historical-cost-basis=\"" + String.format(Locale.US, "%.8f", row.historicalCostBasis) + "\""
                + " data-latest-price=\"" + String.format(Locale.US, "%.8f", Math.max(0.0, row.latestPrice)) + "\""
                + " data-previous-close=\"" + String.format(Locale.US, "%.8f", Math.max(0.0, row.previousClose)) + "\"";

            String tickerToggle = "<button class=\"details-link-btn\" data-target=\"" + detailsRowId + "\" onclick=\"toggleOverviewDetails('" + detailsRowId + "', null)\"><span class=\"ticker-scroll\">" + escapeHtml(row.tickerText) + "</span></button>";
            String securityToggle = "<button class=\"details-link-btn\" data-target=\"" + detailsRowId + "\" onclick=\"toggleOverviewDetails('" + detailsRowId + "', null)\"><span class=\"security-scroll\">" + escapeHtml(row.securityDisplayName) + "</span></button>";
            ReportTemplateHelper.writeHtmlRowWithClassAndAttributes(writer, rowClass, rowAttributes,
                    tickerToggle,
                    securityToggle,
                    dayChangeCell,
                    dayChangeValueCell,
                    dayChartCell,
                    fiftyTwoWeekRangeCell,
                    HtmlFormatter.formatUnits(row.units),
                    HtmlFormatter.formatMoney(row.averageCost, row.currencyCode, 2),
                    "<span class=\"js-row-last-price\">" + (row.latestPrice > 0 ? HtmlFormatter.formatMoney(row.latestPrice, row.currencyCode, 2) : "-") + "</span>",
                    HtmlFormatter.formatMoney(row.positionCostBasis, row.currencyCode, 2),
                    "<span class=\"js-row-market-value\">" + (row.latestPrice > 0 ? HtmlFormatter.formatMoney(row.marketValue, row.currencyCode, 2) : "-") + "</span>");

                    writer.write("<tr id=\"" + detailsRowId + "\" class=\"details-row\" data-group=\"overview-details\">\n");
                    writer.write("    <td class=\"details-cell\" colspan=\"11\">\n");
                    writer.write(buildHoldingDetailsTableHtml(security, row));
                    writer.write("    </td>\n");
                    writer.write("</tr>\n");

            previousAssetType = row.assetType;
                    detailsIndex++;
        }

        LinkedHashMap<String, Double> totalReturnBuckets = sumCurrencyBuckets(totalUnrealizedBuckets, totalRealizedBuckets, totalDividendsBuckets);

        double totalReturnForPct = convertBucketsToTarget(totalReturnBuckets, DEFAULT_TOTAL_CURRENCY, ratesToNok);
        double totalHistoricalCostForPct = convertBucketsToTarget(totalHistoricalCostBuckets, DEFAULT_TOTAL_CURRENCY, ratesToNok);
        double totalCostBasisForPct = convertBucketsToTarget(totalCostBasisBuckets, DEFAULT_TOTAL_CURRENCY, ratesToNok);
        double totalUnrealizedForPct = convertBucketsToTarget(totalUnrealizedBuckets, DEFAULT_TOTAL_CURRENCY, ratesToNok);
        double totalRealizedForPct = convertBucketsToTarget(totalRealizedBuckets, DEFAULT_TOTAL_CURRENCY, ratesToNok);

        double totalReturnPct = totalHistoricalCostForPct > 0 ? (totalReturnForPct / totalHistoricalCostForPct) * 100.0 : 0.0;
        double totalUnrealizedPct = totalCostBasisForPct > 0 ? (totalUnrealizedForPct / totalCostBasisForPct) * 100.0 : 0.0;
        double totalRealizedPct = totalCostBasisForPct > 0 ? (totalRealizedForPct / totalCostBasisForPct) * 100.0 : 0.0;
        double totalDayChangeForPct = convertBucketsToTarget(totalDayChangeBuckets, DEFAULT_TOTAL_CURRENCY, ratesToNok);
        double totalPrevCloseForPct = convertBucketsToTarget(totalPrevCloseValueBuckets, DEFAULT_TOTAL_CURRENCY, ratesToNok);
        double totalDayChangePct = totalPrevCloseForPct > 0.0 ? (totalDayChangeForPct / totalPrevCloseForPct) * 100.0 : 0.0;

        String totalDayChangePctCell = "-";
        String totalDayChangePctClass = "";
        if (!totalDayChangeBuckets.isEmpty() && totalPrevCloseForPct > 0.0) {
            totalDayChangePctClass = totalDayChangePct > 0.0
                ? "positive"
                : (totalDayChangePct < 0.0 ? "negative" : "");
            String pctText = HtmlFormatter.formatPercent(totalDayChangePct, 2);
            totalDayChangePctCell = escapeHtml(pctText);
        }
        String totalDayChangeValueCell = totalDayChangeBuckets.isEmpty()
            ? "<span id=\"holdings-total-day-change-value\" class=\"js-convert-money\" data-buckets=\"{}\" data-decimals=\"2\">-</span>"
            : renderConvertibleMoneyCellWithId("holdings-total-day-change-value", signedClass(totalDayChangeForPct), totalDayChangeBuckets, 2, ratesToNok);

        writer.write("<tr class=\"total-row\">\n");
        writer.write("    <td></td><td><strong>TOTAL</strong></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td>\n");
        writer.write("    <td>" + renderConvertibleMoneyCellWithId("overview-total-cost-basis", totalCostBasisBuckets, 2, ratesToNok) + "</td>\n");
        writer.write("    <td>" + renderConvertibleMoneyCellWithId("overview-total-market-value", totalMarketValueBuckets, 2, ratesToNok) + "</td>\n");
        writer.write("</tr>\n");

        writer.write("</table>\n</div>\n");

        writer.write("<div class=\"table-wrap overview-table-wrap js-overview-mode-panel\" data-overview-mode-panel=\"holdings\" hidden>\n<table class=\"overview-table overview-holdings-table\">\n");
        ReportTemplateHelper.writeHtmlRow(writer, true,
            "Ticker", "Security", "Change %", "Change", "Shares", "Avg Cost", "Last Price",
                "Cost Basis", "Market Value", "Unrealized", "Realized", "Dividends", "Total Return");

        previousAssetType = null;
        int holdingsDetailsIndex = 0;
        for (OverviewRow row : rows) {
            String rowClass = isStockFundBoundary(previousAssetType, row.assetType) ? "asset-split" : null;
            Security security = securityByKey.get(row.securityKey);
            String detailsRowId = "holdings-details-" + holdingsDetailsIndex;
                String unrealizedText = row.hasPrice
                    ? HtmlFormatter.formatMoney(row.unrealized, row.currencyCode, 2) + " (" + HtmlFormatter.formatPercent(row.unrealizedPct, 2) + ")"
                    : "-";
                String realizedText = HtmlFormatter.formatMoney(row.realized, row.currencyCode, 2)
                    + " (" + row.realizedReturnPctText + "%)";
                String totalReturnText = HtmlFormatter.formatMoney(row.totalReturn, row.currencyCode, 2)
                    + " (" + HtmlFormatter.formatPercent(row.totalReturnPct, 2) + ")";
                String unrealizedCell = row.hasPrice
                    ? "<span class=\"js-row-unrealized" + signedClassAttr(row.unrealized) + "\">" + escapeHtml(unrealizedText) + "</span>"
                    : "<span class=\"js-row-unrealized\">-</span>";
                String realizedCell = "<span class=\"" + signedClass(row.realized) + "\">" + escapeHtml(realizedText) + "</span>";
                if (signedClass(row.realized).isBlank()) {
                realizedCell = "<span>" + escapeHtml(realizedText) + "</span>";
                }
                String totalReturnCell = "<span class=\"js-row-total-return" + signedClassAttr(row.totalReturn) + "\">" + escapeHtml(totalReturnText) + "</span>";
            String dayChangeCell = formatDayChangeCell(row.dayChangePct, row.hasDayChangePct);
            String holdingsDayChangeValueCell = formatHoldingDayChangeValueCell(row);
                    String rowAttributes = "data-overview-security-key=\"" + escapeHtml(row.securityKey) + "\""
                        + " data-asset-group=\"" + escapeHtml(normalizeAssetBoundaryGroup(row.assetType)) + "\""
                        + " data-currency=\"" + escapeHtml(normalizeCurrencyCode(row.currencyCode)) + "\""
                        + " data-units=\"" + String.format(Locale.US, "%.8f", row.units) + "\""
                        + " data-position-cost-basis=\"" + String.format(Locale.US, "%.8f", row.positionCostBasis) + "\""
                        + " data-realized=\"" + String.format(Locale.US, "%.8f", row.realized) + "\""
                        + " data-dividends=\"" + String.format(Locale.US, "%.8f", row.dividends) + "\""
                        + " data-historical-cost-basis=\"" + String.format(Locale.US, "%.8f", row.historicalCostBasis) + "\""
                        + " data-latest-price=\"" + String.format(Locale.US, "%.8f", Math.max(0.0, row.latestPrice)) + "\""
                        + " data-previous-close=\"" + String.format(Locale.US, "%.8f", Math.max(0.0, row.previousClose)) + "\"";
                    String tickerToggle = "<button class=\"details-link-btn\" data-target=\"" + detailsRowId + "\" onclick=\"toggleOverviewDetails('" + detailsRowId + "', null)\"><span class=\"ticker-scroll\">" + escapeHtml(row.tickerText) + "</span></button>";
                    String securityToggle = "<button class=\"details-link-btn\" data-target=\"" + detailsRowId + "\" onclick=\"toggleOverviewDetails('" + detailsRowId + "', null)\"><span class=\"security-scroll\">" + escapeHtml(row.securityDisplayName) + "</span></button>";

                    ReportTemplateHelper.writeHtmlRowWithClassAndAttributes(writer, rowClass, rowAttributes,
                        tickerToggle,
                        securityToggle,
                    dayChangeCell,
                    holdingsDayChangeValueCell,
                    HtmlFormatter.formatUnits(row.units),
                    HtmlFormatter.formatMoney(row.averageCost, row.currencyCode, 2),
                    "<span class=\"js-row-last-price\">" + (row.latestPrice > 0 ? HtmlFormatter.formatMoney(row.latestPrice, row.currencyCode, 2) : "-") + "</span>",
                    HtmlFormatter.formatMoney(row.positionCostBasis, row.currencyCode, 2),
                    "<span class=\"js-row-market-value\">" + (row.latestPrice > 0 ? HtmlFormatter.formatMoney(row.marketValue, row.currencyCode, 2) : "-") + "</span>",
                    unrealizedCell,
                    realizedCell,
                    HtmlFormatter.formatMoney(row.dividends, row.currencyCode, 2),
                    totalReturnCell);

                    writer.write("<tr id=\"" + detailsRowId + "\" class=\"details-row\" data-group=\"holdings-details\">\n");
                    writer.write("    <td class=\"details-cell\" colspan=\"13\">\n");
                    writer.write(buildHoldingDetailsTableHtml(security, row));
                    writer.write("    </td>\n");
                    writer.write("</tr>\n");

            previousAssetType = row.assetType;
                    holdingsDetailsIndex++;
        }

        writer.write("<tr class=\"total-row\">\n");
        String totalDayChangePctClassAttr = totalDayChangePctClass.isBlank() ? "" : " class=\"" + totalDayChangePctClass + "\"";
        writer.write("    <td></td><td><strong>TOTAL</strong></td><td><span id=\"holdings-total-day-change-pct\"" + totalDayChangePctClassAttr + ">" + totalDayChangePctCell + "</span></td><td>" + totalDayChangeValueCell + "</td><td></td><td></td><td></td>\n");
        writer.write("    <td>" + renderConvertibleMoneyCellWithId("holdings-total-cost-basis", totalCostBasisBuckets, 2, ratesToNok) + "</td>\n");
        writer.write("    <td>" + renderConvertibleMoneyCellWithId("holdings-total-market-value", totalMarketValueBuckets, 2, ratesToNok) + "</td>\n");
        writer.write("    <td>" + renderConvertibleMoneyCellWithId("holdings-total-unrealized-value", signedClass(totalUnrealizedForPct), totalUnrealizedBuckets, 2, ratesToNok)
            + " <span id=\"holdings-total-unrealized-pct-wrap\" class=\"" + signedClass(totalUnrealizedForPct) + "\">(<span id=\"holdings-total-unrealized-pct\" class=\"" + signedClass(totalUnrealizedForPct) + "\">" + HtmlFormatter.formatPercent(totalUnrealizedPct, 2) + "</span>)</span></td>\n");
        writer.write("    <td>" + renderConvertibleMoneyCellWithId("holdings-total-realized-value", signedClass(totalRealizedForPct), totalRealizedBuckets, 2, ratesToNok)
            + " <span id=\"holdings-total-realized-pct-wrap\" class=\"" + signedClass(totalRealizedForPct) + "\">(<span id=\"holdings-total-realized-pct\" class=\"" + signedClass(totalRealizedForPct) + "\">" + HtmlFormatter.formatPercent(totalRealizedPct, 2) + "</span>)</span></td>\n");
        writer.write("    <td>" + renderConvertibleMoneyCellWithId("holdings-total-dividends-value", totalDividendsBuckets, 2, ratesToNok) + "</td>\n");
        writer.write("    <td>" + renderConvertibleMoneyCellWithId("holdings-total-total-return-value", signedClass(totalReturnForPct), totalReturnBuckets, 2, ratesToNok)
            + " <span id=\"holdings-total-total-return-pct-wrap\" class=\"" + signedClass(totalReturnForPct) + "\">(<span id=\"holdings-total-total-return-pct\" class=\"" + signedClass(totalReturnForPct) + "\">" + HtmlFormatter.formatPercent(totalReturnPct, 2) + "</span>)</span></td>\n");
        writer.write("</tr>\n");
        writer.write("</table>\n</div>\n");

        writer.write("<div class=\"table-wrap overview-table-wrap js-overview-mode-panel\" data-overview-mode-panel=\"fundamentals\" hidden>\n<table class=\"overview-table overview-fundamentals-table\">\n");
        ReportTemplateHelper.writeHtmlRow(writer, true,
            "Ticker", "Security", "Last<br>Price", "Market<br>Cap", "Avg Vol<br>(3M)", "EPS Est.<br>Next Yr", "Forward<br>P/E",
            "Div Payment<br>Date", "Ex-Div<br>Date", "Div/<br>Share", "Fwd Ann Div<br>Rate", "Fwd Ann Div<br>Yield",
            "Trl Ann Div<br>Rate", "Trl Ann Div<br>Yield", "Price /<br>Book");

        previousAssetType = null;
        for (OverviewRow row : rows) {
            String rowClass = isStockFundBoundary(previousAssetType, row.assetType) ? "asset-split" : null;
            Security.FundamentalsSnapshot fundamentals = fundamentalsByKey.get(row.securityKey);
            double fundamentalsLastPrice = fundamentals != null ? fundamentals.lastPrice : 0.0;
            String lastPriceCell = fundamentalsLastPrice > EPSILON
                ? HtmlFormatter.formatMoney(fundamentalsLastPrice, row.currencyCode, 2)
                : (row.latestPrice > EPSILON ? HtmlFormatter.formatMoney(row.latestPrice, row.currencyCode, 2) : "-");
            String fundamentalsRowAttributes = "data-overview-security-key=\"" + escapeHtml(row.securityKey) + "\""
                + " data-currency=\"" + escapeHtml(normalizeCurrencyCode(row.currencyCode)) + "\""
                + " data-latest-price=\"" + String.format(Locale.US, "%.8f", Math.max(0.0, fundamentalsLastPrice > EPSILON ? fundamentalsLastPrice : row.latestPrice)) + "\"";
            ReportTemplateHelper.writeHtmlRowWithClassAndAttributes(writer, rowClass, fundamentalsRowAttributes,
                    "<span class=\"ticker-scroll\">" + escapeHtml(row.tickerText) + "</span>",
                    "<span class=\"security-scroll\">" + escapeHtml(row.securityDisplayName) + "</span>",
                "<span class=\"js-row-fundamentals-last-price\">" + lastPriceCell + "</span>",
                formatCompactMetric(fundamentals != null ? fundamentals.marketCap : 0.0),
                formatCompactMetric(fundamentals != null ? fundamentals.averageVolume3Month : 0.0),
                formatFundamentalsDecimal(fundamentals != null ? fundamentals.epsEstimateNextYear : 0.0, 2),
                formatFundamentalsDecimal(fundamentals != null ? fundamentals.forwardPe : 0.0, 2),
                formatFundamentalsDate(fundamentals != null ? fundamentals.dividendPaymentDateEpochSeconds : 0L),
                formatFundamentalsDate(fundamentals != null ? fundamentals.exDividendDateEpochSeconds : 0L),
                formatFundamentalsMoney(fundamentals != null ? fundamentals.dividendPerShare : 0.0, row.currencyCode, 2),
                formatFundamentalsMoney(fundamentals != null ? fundamentals.forwardAnnualDividendRate : 0.0, row.currencyCode, 2),
                formatFundamentalsPercent(fundamentals != null ? fundamentals.forwardAnnualDividendYield : 0.0),
                formatFundamentalsMoney(fundamentals != null ? fundamentals.trailingAnnualDividendRate : 0.0, row.currencyCode, 2),
                formatFundamentalsPercent(fundamentals != null ? fundamentals.trailingAnnualDividendYield : 0.0),
                formatFundamentalsDecimal(fundamentals != null ? fundamentals.priceToBook : 0.0, 2));
            previousAssetType = row.assetType;
        }

        writer.write("</table>\n</div>\n\n");

        writer.write("<section class=\"allocation-card\">\n");
        writer.write("<h3>Market Value Allocation</h3>\n");
        writer.write("<div class=\"allocation-visuals\">\n");
        writer.write("<div class=\"allocation-row allocation-row-top\">\n");
        writer.write("<div class=\"allocation-panel asset-type-panel\"><h4 class=\"allocation-panel-title\">By Asset Type</h4>\n");
        writer.write(ChartBuilder.buildAssetTypeAllocationSvg(rows, store.getCurrentCashHoldings(), ratesToNok));
        writer.write("</div>\n");
        writer.write("<div class=\"allocation-panel sector-panel\"><h4 class=\"allocation-panel-title\">By Sector</h4>\n");
        writer.write(ChartBuilder.buildSectorAllocationSvg(rows, ratesToNok));
        writer.write("</div>\n");
        writer.write("<div class=\"allocation-panel region-panel\"><h4 class=\"allocation-panel-title\">By Region</h4>\n");
        writer.write(ChartBuilder.buildRegionAllocationSvg(rows, ratesToNok));
        writer.write("</div>\n");
        writer.write("</div>\n");

        writer.write("<div class=\"allocation-row allocation-row-bottom\">\n");
        writer.write("<div class=\"allocation-panel security-pie-panel\"><h4 class=\"allocation-panel-title\">By Security (Pie)</h4>\n");
        writer.write(ChartBuilder.buildMarketValueAllocationSvg(rows, ratesToNok));
        writer.write("</div>\n");
        writer.write("<div class=\"allocation-panel security-bar-panel\"><h4 class=\"allocation-panel-title\">By Security (Bar)</h4>\n");
        writer.write(ChartBuilder.buildMarketValueBarChartSvg(rows, ratesToNok));
        writer.write("</div>\n");
        writer.write("</div>\n");
        writer.write("</div>\n");
        writer.write("</section>\n");
    }

    private static String formatDayChangeCell(double dayChangePct, boolean hasDayChangePct) {
        if (!hasDayChangePct || !Double.isFinite(dayChangePct)) {
            return "<span class=\"js-row-day-change\">-</span>";
        }

        String cssClass = dayChangePct > 0.0
                ? "positive"
                : (dayChangePct < 0.0 ? "negative" : "");

        String valueText = HtmlFormatter.formatPercent(dayChangePct, 2);
        if (!cssClass.isBlank()) {
            return "<span class=\"js-row-day-change " + cssClass + "\">" + escapeHtml(valueText) + "</span>";
        }
        return "<span class=\"js-row-day-change\">" + escapeHtml(valueText) + "</span>";
    }

    private static String formatDayChangeValueCell(OverviewRow row) {
        if (row == null || row.latestPrice <= 0.0 || row.previousClose <= 0.0) {
            return "<span class=\"js-row-day-change-value\">-</span>";
        }

        double changeValue = row.latestPrice - row.previousClose;
        String cssClass = changeValue > 0.0
                ? "positive"
                : (changeValue < 0.0 ? "negative" : "");
        String valueText = HtmlFormatter.formatMoney(changeValue, row.currencyCode, 2);
        if (!cssClass.isBlank()) {
            return "<span class=\"js-row-day-change-value " + cssClass + "\">" + escapeHtml(valueText) + "</span>";
        }
        return "<span class=\"js-row-day-change-value\">" + escapeHtml(valueText) + "</span>";
    }

    private static String formatHoldingDayChangeValueCell(OverviewRow row) {
        if (row == null || row.latestPrice <= 0.0 || row.previousClose <= 0.0 || row.units <= 0.0) {
            return "<span class=\"js-row-day-change-value-position\">-</span>";
        }

        double changeValue = (row.latestPrice - row.previousClose) * row.units;
        String cssClass = changeValue > 0.0
                ? "positive"
                : (changeValue < 0.0 ? "negative" : "");
        String valueText = HtmlFormatter.formatMoney(changeValue, row.currencyCode, 2);
        if (!cssClass.isBlank()) {
            return "<span class=\"js-row-day-change-value-position " + cssClass + "\">" + escapeHtml(valueText) + "</span>";
        }
        return "<span class=\"js-row-day-change-value-position\">" + escapeHtml(valueText) + "</span>";
    }

    private static String formatDayChartCell(OverviewRow row) {
        if (row == null || row.latestPrice <= 0.0) {
            return "<span class=\"js-row-day-chart\" data-ticker=\"\">-</span>";
        }

        return "<span class=\"js-row-day-chart\" data-ticker=\"" + escapeHtml(row.tickerText)
            + "\">-</span>";
    }

    private static String format52WeekRangeCell(OverviewRow row, Security.FundamentalsSnapshot fundamentals) {
        if (row == null || fundamentals == null || !fundamentals.hasRangeData()) {
            return "-";
        }

        double low = fundamentals.fiftyTwoWeekLow;
        double high = fundamentals.fiftyTwoWeekHigh;
        double current = fundamentals.lastPrice > EPSILON ? fundamentals.lastPrice : row.latestPrice;
        if (current <= EPSILON) {
            current = low;
        }

        double ratio = (current - low) / Math.max(EPSILON, high - low);
        double clamped = Math.max(0.0, Math.min(1.0, ratio));
        String markerPct = String.format(Locale.US, "%.2f", clamped * 100.0);
        String lowLabel = HtmlFormatter.formatMoney(low, row.currencyCode, 2);
        String highLabel = HtmlFormatter.formatMoney(high, row.currencyCode, 2);
        String currentLabel = HtmlFormatter.formatMoney(current, row.currencyCode, 2);

        return "<div class=\"wk-range-cell\" title=\"" + escapeHtml(currentLabel) + "\">"
                + "<div class=\"wk-range-track\"><span class=\"wk-range-marker\" style=\"left:" + markerPct + "%;\"></span></div>"
                + "<div class=\"wk-range-labels\"><span>" + escapeHtml(lowLabel) + "</span><span>" + escapeHtml(highLabel) + "</span></div>"
                + "</div>";
    }

    private static String formatFundamentalsMoney(double value, String currencyCode, int decimals) {
        if (!Double.isFinite(value) || Math.abs(value) <= EPSILON) {
            return "-";
        }
        return HtmlFormatter.formatMoney(value, currencyCode, decimals);
    }

    private static String formatFundamentalsPercent(double value) {
        if (!Double.isFinite(value) || Math.abs(value) <= EPSILON) {
            return "-";
        }

        double percent = Math.abs(value) <= 1.0 ? value * 100.0 : value;
        return HtmlFormatter.formatPercent(percent, 2);
    }

    private static String formatFundamentalsDecimal(double value, int decimals) {
        if (!Double.isFinite(value) || Math.abs(value) <= EPSILON) {
            return "-";
        }

        String pattern = "%1$." + Math.max(0, decimals) + "f";
        String formatted = String.format(Locale.US, pattern, value);
        return trimTrailingZeros(formatted);
    }

    private static String formatFundamentalsDate(long epochSeconds) {
        if (epochSeconds <= 0L) {
            return "-";
        }

        try {
            LocalDate date = Instant.ofEpochSecond(epochSeconds).atZone(ZoneOffset.UTC).toLocalDate();
            return date.format(DETAIL_DATE_FORMATTER);
        } catch (Exception ignored) {
            return "-";
        }
    }

    private static String formatCompactMetric(double value) {
        if (!Double.isFinite(value) || value <= 0.0) {
            return "-";
        }

        double abs = Math.abs(value);
        double divisor = 1.0;
        String suffix = "";
        if (abs >= 1_000_000_000_000.0) {
            divisor = 1_000_000_000_000.0;
            suffix = "T";
        } else if (abs >= 1_000_000_000.0) {
            divisor = 1_000_000_000.0;
            suffix = "B";
        } else if (abs >= 1_000_000.0) {
            divisor = 1_000_000.0;
            suffix = "M";
        } else if (abs >= 1_000.0) {
            divisor = 1_000.0;
            suffix = "K";
        }

        double scaled = value / divisor;
        int decimals = Math.abs(scaled) >= 100.0 ? 0 : (Math.abs(scaled) >= 10.0 ? 1 : 2);
        String pattern = "%1$." + decimals + "f";
        return trimTrailingZeros(String.format(Locale.US, pattern, scaled)) + suffix;
    }

    private static String trimTrailingZeros(String value) {
        if (value == null) {
            return "";
        }
        if (!value.contains(".")) {
            return value;
        }
        return value.replaceAll("0+$", "").replaceAll("\\.$", "");
    }

    private static String signedClass(double value) {
        if (!Double.isFinite(value) || Math.abs(value) <= EPSILON) {
            return "";
        }
        return value > 0.0 ? "positive" : "negative";
    }

    private static String signedClassAttr(double value) {
        String cssClass = signedClass(value);
        return cssClass.isBlank() ? "" : " " + cssClass;
    }

    private static String signedSpan(String text, double value) {
        String safeText = escapeHtml(text);
        String cssClass = signedClass(value);
        if (cssClass.isBlank()) {
            return safeText;
        }
        return "<span class=\"" + cssClass + "\">" + safeText + "</span>";
    }

    private static String signedWrapHtml(String html, double value) {
        String cssClass = signedClass(value);
        if (cssClass.isBlank()) {
            return html;
        }
        return "<span class=\"" + cssClass + "\">" + html + "</span>";
    }

    private static void writeRealizedSummaryTableHtml(FileWriter writer, TransactionStore store, Map<String, Double> ratesToNok) throws IOException {
        writer.write("<h2>REALIZED OVERVIEW - ALL SALES</h2>\n");
        writer.write("<div class=\"table-wrap\">\n<table class=\"realized-table\">\n");
        ReportTemplateHelper.writeHtmlRow(writer, true, ReportTemplateHelper.buildTickerHeaderCell("realized-details"), "Security", "Cost Basis", "Sales Value", "Gain/Loss", "Dividends", "Total Return");

        ArrayList<Security> soldSecurities = getSortedSoldSecurities(store);
        LinkedHashMap<String, Double> totalSalesValueBuckets = new LinkedHashMap<>();
        LinkedHashMap<String, Double> totalCostBasisBuckets = new LinkedHashMap<>();
        LinkedHashMap<String, Double> totalRealizedGainBuckets = new LinkedHashMap<>();
        LinkedHashMap<String, Double> totalRealizedDividendsBuckets = new LinkedHashMap<>();
        String previousAssetType = null;
        int detailsIndex = 0;

        for (Security security : soldSecurities) {
            String currency = security.getCurrencyCode();
            double salesValue = security.getRealizedSalesValue();
            double costBasis = security.getRealizedCostBasis();
            double gain = security.getRealizedGain();
            double realizedDividends = security.isFullyRealized() ? security.getDividends() : 0.0;
            double totalReturnValue = gain + realizedDividends;
            double rowTotalReturnPct = costBasis > 0 ? (totalReturnValue / costBasis) * 100.0 : 0.0;
            String currentAssetType = security.getAssetType().name();
            String rowClass = isStockFundBoundary(previousAssetType, currentAssetType) ? "asset-split" : null;
            String totalReturnCombined = signedSpan(
                HtmlFormatter.formatMoney(totalReturnValue, currency, 2)
                    + " (" + HtmlFormatter.formatPercent(rowTotalReturnPct, 2) + ")",
                totalReturnValue);

            addToCurrencyBuckets(totalSalesValueBuckets, currency, salesValue);
            addToCurrencyBuckets(totalCostBasisBuckets, currency, costBasis);
            addToCurrencyBuckets(totalRealizedGainBuckets, currency, gain);
            addToCurrencyBuckets(totalRealizedDividendsBuckets, currency, realizedDividends);
                String detailsRowId = "realized-details-" + detailsIndex;
                String rowAttributes = "data-asset-group=\"" + escapeHtml(normalizeAssetBoundaryGroup(currentAssetType)) + "\"";

                String tickerToggle = "<button class=\"details-link-btn\" data-target=\"" + detailsRowId + "\" onclick=\"toggleOverviewDetails('" + detailsRowId + "', null)\"><span class=\"ticker-scroll\">" + escapeHtml(security.getTicker()) + "</span></button>";
                String securityToggle = "<button class=\"details-link-btn\" data-target=\"" + detailsRowId + "\" onclick=\"toggleOverviewDetails('" + detailsRowId + "', null)\"><span class=\"security-scroll\">" + escapeHtml(security.getDisplayName()) + "</span></button>";
                ReportTemplateHelper.writeHtmlRowWithClassAndAttributes(writer, rowClass, rowAttributes,
                    tickerToggle,
                    securityToggle,
                    HtmlFormatter.formatMoney(costBasis, currency, 2),
                    HtmlFormatter.formatMoney(salesValue, currency, 2),
                    signedSpan(HtmlFormatter.formatMoney(gain, currency, 2), gain),
                    signedSpan(HtmlFormatter.formatMoney(realizedDividends, currency, 2), realizedDividends),
                    totalReturnCombined);

                writer.write("<tr id=\"" + detailsRowId + "\" class=\"details-row\" data-group=\"realized-details\">\n");
                writer.write("    <td class=\"details-cell\" colspan=\"7\">\n");
                writer.write(buildRealizedSaleTradesDetailsHtml(security));
                writer.write("    </td>\n");
                writer.write("</tr>\n");

            previousAssetType = currentAssetType;
                detailsIndex++;
        }

        double totalCostBasisForPct = convertBucketsToTarget(totalCostBasisBuckets, DEFAULT_TOTAL_CURRENCY, ratesToNok);
        double totalRealizedGainForPct = convertBucketsToTarget(totalRealizedGainBuckets, DEFAULT_TOTAL_CURRENCY, ratesToNok);
        double totalRealizedDividendsForPct = convertBucketsToTarget(totalRealizedDividendsBuckets, DEFAULT_TOTAL_CURRENCY, ratesToNok);
        LinkedHashMap<String, Double> totalRealizedReturnBuckets = sumCurrencyBuckets(totalRealizedGainBuckets, totalRealizedDividendsBuckets);
        double totalRealizedReturnForPct = totalRealizedGainForPct + totalRealizedDividendsForPct;
        double totalReturnPct = totalCostBasisForPct > 0
            ? (totalRealizedReturnForPct / totalCostBasisForPct) * 100.0
            : 0.0;
        writer.write("<tr class=\"total-row\">\n");
        writer.write("    <td></td><td><strong>TOTAL</strong></td>\n");
        writer.write("    <td>" + renderConvertibleMoneyCell(totalCostBasisBuckets, 2, ratesToNok) + "</td>\n");
        writer.write("    <td>" + renderConvertibleMoneyCell(totalSalesValueBuckets, 2, ratesToNok) + "</td>\n");
        writer.write("    <td>" + signedWrapHtml(renderConvertibleMoneyCell(totalRealizedGainBuckets, 2, ratesToNok), totalRealizedGainForPct) + "</td>\n");
        writer.write("    <td>" + signedWrapHtml(renderConvertibleMoneyCell(totalRealizedDividendsBuckets, 2, ratesToNok), totalRealizedDividendsForPct) + "</td>\n");
        writer.write("    <td>"
            + signedWrapHtml(renderConvertibleMoneyCell(totalRealizedReturnBuckets, 2, ratesToNok), totalRealizedReturnForPct)
            + " "
            + signedSpan("(" + HtmlFormatter.formatPercent(totalReturnPct, 2) + ")", totalRealizedReturnForPct)
            + "</td>\n");
        writer.write("</tr>\n");

        writer.write("</table>\n</div>\n\n");
    }

    private static String buildRealizedSaleTradesDetailsHtml(Security security) {
        return buildRealizedSaleTradesDetailsHtml(security, null);
    }

    private static String buildRealizedSaleTradesDetailsHtml(Security security, Integer filterYear) {
        StringBuilder html = new StringBuilder();
        html.append("<div class=\"details-wrap\">\n");
        html.append("<h4>Sale Trades - ").append(escapeHtml(security.getDisplayName())).append("</h4>\n");

        List<Security.SaleTrade> saleTrades = security.getSaleTradesSortedByDate();
        if (filterYear != null) {
            int safeYear = filterYear;
            saleTrades = saleTrades.stream()
                    .filter(trade -> trade != null && trade.getTradeDate() != null && trade.getTradeDate().getYear() == safeYear)
                    .toList();
        }
        String currency = security.getCurrencyCode();
        if (saleTrades.isEmpty()) {
            html.append("<div class=\"app-shell-note\">No sale trades available.</div>\n");
        } else {
            html.append("<table class=\"details-table\">\n");
            html.append("<tr><th>Sale Date</th><th>Shares</th><th>Price/Share</th><th>Cost Basis</th><th>Sale Value</th><th>Gain/Loss</th><th>Return (%)</th></tr>\n");

            double totalUnits = 0.0;
            double totalSaleValue = 0.0;
            double totalCostBasis = 0.0;
            double totalGainLoss = 0.0;
            for (Security.SaleTrade trade : saleTrades) {
                totalUnits += trade.getUnits();
                totalSaleValue += trade.getSaleValue();
                totalCostBasis += trade.getCostBasis();
                totalGainLoss += trade.getGainLoss();

                html.append("<tr>");
                html.append("<td>").append(escapeHtml(trade.getTradeDateAsCsv())).append("</td>");
                html.append("<td>").append(escapeHtml(HtmlFormatter.formatUnits(trade.getUnits()))).append("</td>");
                html.append("<td>").append(escapeHtml(HtmlFormatter.formatMoney(trade.getUnitPrice(), currency, 2))).append("</td>");
                html.append("<td>").append(escapeHtml(HtmlFormatter.formatMoney(trade.getCostBasis(), currency, 0))).append("</td>");
                html.append("<td>").append(escapeHtml(HtmlFormatter.formatMoney(trade.getSaleValue(), currency, 0))).append("</td>");
                html.append("<td class=\"").append(signedClass(trade.getGainLoss())).append("\">").append(escapeHtml(HtmlFormatter.formatMoney(trade.getGainLoss(), currency, 0))).append("</td>");
                html.append("<td class=\"").append(signedClass(trade.getReturnPct())).append("\">").append(escapeHtml(HtmlFormatter.formatPercent(trade.getReturnPct(), 2))).append("</td>");
                html.append("</tr>\n");
            }

            double totalReturnPct = totalCostBasis > 0.0 ? (totalGainLoss / totalCostBasis) * 100.0 : 0.0;
            html.append("<tr class=\"total-row\">");
            html.append("<td><strong>TOTAL</strong></td>");
            html.append("<td>").append(escapeHtml(HtmlFormatter.formatUnits(totalUnits))).append("</td>");
            html.append("<td></td>");
            html.append("<td>").append(escapeHtml(HtmlFormatter.formatMoney(totalCostBasis, currency, 0))).append("</td>");
            html.append("<td>").append(escapeHtml(HtmlFormatter.formatMoney(totalSaleValue, currency, 0))).append("</td>");
            html.append("<td class=\"").append(signedClass(totalGainLoss)).append("\">").append(escapeHtml(HtmlFormatter.formatMoney(totalGainLoss, currency, 0))).append("</td>");
            html.append("<td class=\"").append(signedClass(totalReturnPct)).append("\">").append(escapeHtml(HtmlFormatter.formatPercent(totalReturnPct, 2))).append("</td>");
            html.append("</tr>\n");

            html.append("</table>\n");
        }

        List<Security.DividendEvent> dividendEvents = security.getAllDividendEventsSortedByDate();
        if (filterYear != null) {
            int safeYear = filterYear;
            dividendEvents = dividendEvents.stream()
                    .filter(event -> event != null && event.getTradeDate() != null && event.getTradeDate().getYear() == safeYear)
                    .toList();
        }
        if (!dividendEvents.isEmpty()) {
            html.append("<h4 style=\"margin-top:10px;\">Dividend Events</h4>\n");
            html.append("<table class=\"details-table\">\n");
            html.append("<tr><th>Date</th><th>Shares</th><th>Dividend</th></tr>\n");
            double totalDividendAmount = 0.0;
            for (Security.DividendEvent event : dividendEvents) {
                totalDividendAmount += event.getAmount();
                String unitsText = event.getUnits() > 0.0 ? HtmlFormatter.formatUnits(event.getUnits()) : "-";
                html.append("<tr>");
                html.append("<td>").append(escapeHtml(formatDetailDate(event.getTradeDate()))).append("</td>");
                html.append("<td>").append(escapeHtml(unitsText)).append("</td>");
                html.append("<td>").append(escapeHtml(HtmlFormatter.formatMoney(event.getAmount(), currency, 2))).append("</td>");
                html.append("</tr>\n");
            }
            html.append("<tr class=\"total-row\">");
            html.append("<td><strong>TOTAL</strong></td><td></td>");
            html.append("<td>").append(escapeHtml(HtmlFormatter.formatMoney(totalDividendAmount, currency, 2))).append("</td>");
            html.append("</tr>\n");
            html.append("</table>\n");
        }

        html.append("</div>\n");
        return html.toString();
    }

    private static Map<String, Security> buildSecurityLookupByKey(TransactionStore store) {
        Map<String, Security> byKey = new HashMap<>();
        for (Security security : store.getSecurities()) {
            byKey.put(getTrackingSecurityKey(security), security);
        }
        return byKey;
    }

    private static String getTrackingSecurityKey(Security security) {
        if (security == null) {
            return "";
        }

        String isin = security.getIsin();
        if (isin != null && !isin.isBlank()) {
            return isin.trim().toUpperCase(Locale.ROOT);
        }

        String name = security.getName();
        if (name == null || name.isBlank()) {
            return "";
        }

        return "NAME:" + name.trim().toUpperCase(Locale.ROOT);
    }

    private static boolean isTemporaryRightsSecurity(Security security) {
        if (security == null) {
            return false;
        }

        String name = security.getName() == null ? "" : security.getName();
        String displayName = security.getDisplayName() == null ? "" : security.getDisplayName();
        String ticker = security.getTicker() == null ? "" : security.getTicker();
        String upper = (name + " " + displayName + " " + ticker).toUpperCase(Locale.ROOT);
        return upper.contains("T-RETT") || upper.contains("TEGNINGSRETT");
    }

    private static String buildHoldingDetailsTableHtml(Security security, OverviewRow row) {
        StringBuilder html = new StringBuilder();
        html.append("<div class=\"details-wrap\">\n");
        html.append("<h4>Transaction Details - ").append(escapeHtml(row.securityDisplayName)).append("</h4>\n");

        if (security == null) {
            html.append("<div class=\"app-shell-note\">Details are not available for this security.</div>\n");
            html.append("</div>\n");
            return html.toString();
        }

        class DetailEntry {
            final LocalDate date;
            final int order;
            final String type;
            final String units;
            final String price;
            final String amount;
            final String unrealized;
            final String unrealizedPct;
            final String unrealizedClass;
            final String unrealizedPctClass;

            DetailEntry(LocalDate date, int order, String type, String units, String price,
                        String amount, String unrealized, String unrealizedPct,
                        String unrealizedClass, String unrealizedPctClass) {
                this.date = date;
                this.order = order;
                this.type = type;
                this.units = units;
                this.price = price;
                this.amount = amount;
                this.unrealized = unrealized;
                this.unrealizedPct = unrealizedPct;
                this.unrealizedClass = unrealizedClass;
                this.unrealizedPctClass = unrealizedPctClass;
            }
        }

        ArrayList<DetailEntry> entries = new ArrayList<>();
        for (Security.CurrentHoldingLot lot : security.getCurrentHoldingLotsSortedByDate()) {
            double lotCostBasis = lot.getCostBasis();
            String unrealizedText = "-";
            String unrealizedPctText = "-";
            String unrealizedClass = "";
            String unrealizedPctClass = "";
            if (row.latestPrice > 0.0 && lotCostBasis > 0.0) {
                double currentValue = lot.getUnits() * row.latestPrice;
                double unrealized = currentValue - lotCostBasis;
                double unrealizedPct = (unrealized / lotCostBasis) * 100.0;
                unrealizedText = HtmlFormatter.formatMoney(unrealized, row.currencyCode, 2);
                unrealizedPctText = HtmlFormatter.formatPercent(unrealizedPct, 2);
                unrealizedClass = signedClass(unrealized);
                unrealizedPctClass = signedClass(unrealizedPct);
            }

            entries.add(new DetailEntry(
                    lot.getTradeDate(),
                    0,
                    "<span class=\"details-buy\">BUY</span>",
                    HtmlFormatter.formatUnits(lot.getUnits()),
                    HtmlFormatter.formatMoney(lot.getUnitCost(), row.currencyCode, 2),
                    HtmlFormatter.formatMoney(lotCostBasis, row.currencyCode, 2),
                    unrealizedText,
                    unrealizedPctText,
                    unrealizedClass,
                    unrealizedPctClass
            ));
        }

        for (Security.DividendEvent event : security.getCurrentDividendEventsSortedByDate()) {
            String unitsText = event.getUnits() > 0.0 ? HtmlFormatter.formatUnits(event.getUnits()) : "-";
            entries.add(new DetailEntry(
                    event.getTradeDate(),
                    1,
                    "<span class=\"details-dividend\">DIVIDEND</span>",
                    unitsText,
                    "-",
                    HtmlFormatter.formatMoney(event.getAmount(), row.currencyCode, 2),
                    "-",
                    "-",
                    "",
                    ""
            ));
        }

        entries.sort(Comparator
                .comparing((DetailEntry e) -> e.date == null ? LocalDate.MIN : e.date)
                .thenComparingInt(e -> e.order));

        if (entries.isEmpty()) {
            html.append("<div class=\"app-shell-note\">No active buy/dividend entries for current holdings.</div>\n");
            html.append("</div>\n");
            return html.toString();
        }

        html.append("<table class=\"details-table\">\n");
        html.append("<tr><th>Date</th><th>Type</th><th>Shares</th><th>Price/Share</th><th>Amount</th><th>Unrealized</th><th>Unrealized (%)</th></tr>\n");
        for (DetailEntry entry : entries) {
            html.append("<tr>");
            html.append("<td>").append(escapeHtml(formatDetailDate(entry.date))).append("</td>");
            html.append("<td>").append(entry.type).append("</td>");
            html.append("<td>").append(escapeHtml(entry.units)).append("</td>");
            html.append("<td>").append(escapeHtml(entry.price)).append("</td>");
            html.append("<td>").append(escapeHtml(entry.amount)).append("</td>");
            html.append("<td class=\"").append(escapeHtml(entry.unrealizedClass)).append("\">").append(escapeHtml(entry.unrealized)).append("</td>");
            html.append("<td class=\"").append(escapeHtml(entry.unrealizedPctClass)).append("\">").append(escapeHtml(entry.unrealizedPct)).append("</td>");
            html.append("</tr>\n");
        }
        html.append("</table>\n</div>\n");
        return html.toString();
    }

    private static String formatDetailDate(LocalDate date) {
        if (date == null || date.equals(LocalDate.MIN)) {
            return "-";
        }
        return date.format(DETAIL_DATE_FORMATTER);
    }

    private static ArrayList<Security> getSortedSoldSecurities(TransactionStore store) {
        ArrayList<Security> sold = new ArrayList<>();
        for (Security s : store.getSecurities()) {
            if (s.hasSales()) sold.add(s);
        }
        sold.sort(Comparator
                .comparingInt((Security s) -> getAssetPriority(s.getAssetType().name()))
                .thenComparing(Security::getRealizedSalesValue, Comparator.reverseOrder())
                .thenComparing(Security::getName, String.CASE_INSENSITIVE_ORDER));
        return sold;
    }

    private static boolean isStockFundBoundary(String previousAssetType, String currentAssetType) {
        String previousGroup = normalizeAssetBoundaryGroup(previousAssetType);
        String currentGroup = normalizeAssetBoundaryGroup(currentAssetType);

        if (previousGroup == null || currentGroup == null || previousGroup.equals(currentGroup)) {
            return false;
        }

        return ("STOCK".equals(previousGroup) && "FUND".equals(currentGroup))
                || ("FUND".equals(previousGroup) && "STOCK".equals(currentGroup));
    }

    private static String normalizeAssetBoundaryGroup(String assetType) {
        if (assetType == null || assetType.isBlank()) {
            return null;
        }

        String normalized = assetType.trim().toUpperCase(Locale.ROOT);
        if ("FUND".equals(normalized)) {
            return "FUND";
        }

        // Treat UNKNOWN and all non-fund classes as STOCK for a stable single boundary.
        return "STOCK";
    }

    private static int getAssetPriority(String assetType) {
        if (assetType == null) {
            return 0;
        }
        return switch (assetType.toUpperCase(Locale.ROOT)) {
            case "STOCK", "UNKNOWN" -> 0;
            case "FUND" -> 1;
            default -> 2;
        };
    }

    private static Set<String> collectCurrencies(TransactionStore store, List<OverviewRow> overviewRows) {
        LinkedHashSet<String> currencies = new LinkedHashSet<>();
        currencies.add(DEFAULT_TOTAL_CURRENCY);

        for (OverviewRow row : overviewRows) {
            currencies.add(normalizeCurrencyCode(row.currencyCode));
        }

        for (Security security : store.getSecurities()) {
            currencies.add(normalizeCurrencyCode(security.getCurrencyCode()));
        }

        return currencies;
    }

    private static String normalizeCurrencyCode(String currencyCode) {
        if (currencyCode == null || currencyCode.isBlank()) {
            return DEFAULT_TOTAL_CURRENCY;
        }

        String normalized = currencyCode.trim().toUpperCase(Locale.ROOT);
        if (!normalized.matches("[A-Z]{3}")) {
            return DEFAULT_TOTAL_CURRENCY;
        }
        return normalized;
    }

    private static void addToCurrencyBuckets(Map<String, Double> buckets, String currencyCode, double amount) {
        String code = normalizeCurrencyCode(currencyCode);
        buckets.merge(code, amount, Double::sum);
    }

    private static LinkedHashMap<String, Double> singleCurrencyBuckets(String currencyCode, double amount) {
        LinkedHashMap<String, Double> buckets = new LinkedHashMap<>();
        addToCurrencyBuckets(buckets, currencyCode, amount);
        return buckets;
    }

    @SafeVarargs
    private static LinkedHashMap<String, Double> sumCurrencyBuckets(Map<String, Double>... bucketSets) {
        LinkedHashMap<String, Double> merged = new LinkedHashMap<>();
        if (bucketSets == null) {
            return merged;
        }

        for (Map<String, Double> bucketSet : bucketSets) {
            if (bucketSet == null) {
                continue;
            }

            for (Map.Entry<String, Double> entry : bucketSet.entrySet()) {
                if (entry.getValue() == null) {
                    continue;
                }
                addToCurrencyBuckets(merged, entry.getKey(), entry.getValue());
            }
        }

        return merged;
    }

    private static String toBucketsJson(Map<String, Double> buckets) {
        StringBuilder json = new StringBuilder();
        json.append("{");

        boolean first = true;
        for (Map.Entry<String, Double> entry : buckets.entrySet()) {
            String code = normalizeCurrencyCode(entry.getKey());
            double amount = entry.getValue() == null ? 0.0 : entry.getValue();

            if (!first) {
                json.append(",");
            }
            first = false;
            json.append("\"").append(code).append("\":")
                    .append(String.format(Locale.US, "%.8f", amount));
        }

        json.append("}");
        return json.toString();
    }

    private static double convertBucketsToTarget(Map<String, Double> buckets, String targetCurrency, Map<String, Double> ratesToNok) {
        if (buckets == null || buckets.isEmpty()) {
            return 0.0;
        }

        String target = normalizeCurrencyCode(targetCurrency);
        double targetRateToNok = ratesToNok.getOrDefault(target, 0.0);
        if (targetRateToNok <= 0.0) {
            return 0.0;
        }

        double totalInNok = 0.0;
        for (Map.Entry<String, Double> entry : buckets.entrySet()) {
            String source = normalizeCurrencyCode(entry.getKey());
            double amount = entry.getValue() == null ? 0.0 : entry.getValue();

            double sourceRateToNok = ratesToNok.getOrDefault(source, 0.0);
            if (sourceRateToNok <= 0.0) {
                continue;
            }
            totalInNok += amount * sourceRateToNok;
        }

        return totalInNok / targetRateToNok;
    }

    private static String formatBucketsInTarget(Map<String, Double> buckets, String targetCurrency, int decimals, Map<String, Double> ratesToNok) {
        String target = normalizeCurrencyCode(targetCurrency);
        double amount = convertBucketsToTarget(buckets, target, ratesToNok);
        return HtmlFormatter.formatMoney(amount, target, decimals);
    }

    private static String renderConvertibleMoneyCell(Map<String, Double> buckets, int decimals, Map<String, Double> ratesToNok) {
        return "<span class=\"js-convert-money\" data-buckets=\""
                + escapeHtml(toBucketsJson(buckets))
                + "\" data-decimals=\""
                + decimals
                + "\">"
                + formatBucketsInTarget(buckets, DEFAULT_TOTAL_CURRENCY, decimals, ratesToNok)
                + "</span>";
    }

    private static String renderConvertibleMoneyCellWithId(String id, Map<String, Double> buckets, int decimals, Map<String, Double> ratesToNok) {
        return "<span id=\"" + escapeHtml(id) + "\" class=\"js-convert-money\" data-buckets=\""
                + escapeHtml(toBucketsJson(buckets))
                + "\" data-decimals=\""
                + decimals
                + "\">"
                + formatBucketsInTarget(buckets, DEFAULT_TOTAL_CURRENCY, decimals, ratesToNok)
                + "</span>";
    }

    private static String renderConvertibleMoneyCellWithId(String id, String extraClass, Map<String, Double> buckets, int decimals, Map<String, Double> ratesToNok) {
        String classes = "js-convert-money";
        if (extraClass != null && !extraClass.isBlank()) {
            classes += " " + extraClass.trim();
        }
        return "<span id=\"" + escapeHtml(id) + "\" class=\"" + classes + "\" data-buckets=\""
                + escapeHtml(toBucketsJson(buckets))
                + "\" data-decimals=\""
                + decimals
                + "\">"
                + formatBucketsInTarget(buckets, DEFAULT_TOTAL_CURRENCY, decimals, ratesToNok)
                + "</span>";
    }

    private static void writeCurrencyConversionScript(FileWriter writer, Map<String, Double> ratesToNok) throws IOException {
        writer.write("const REPORT_RATES_TO_NOK = " + CurrencyConversionService.toJson(ratesToNok) + ";\n");
        writer.write("function normalizeCurrencyCodeInput(value) {\n");
        writer.write("  return String(value || '').trim().toUpperCase();\n");
        writer.write("}\n");
        writer.write("function formatGroupedNumber(value, decimals) {\n");
        writer.write("  var fixed = Number(value || 0).toFixed(decimals);\n");
        writer.write("  var parts = fixed.split('.');\n");
        writer.write("  var whole = parts[0].replace(/\\B(?=(\\d{3})+(?!\\d))/g, ' ');\n");
        writer.write("  return decimals > 0 ? whole + '.' + parts[1] : whole;\n");
        writer.write("}\n");
        writer.write("function formatMoneyValue(amount, currency, decimals) {\n");
        writer.write("  return formatGroupedNumber(amount, decimals) + ' ' + currency;\n");
        writer.write("}\n");
        writer.write("function formatCompactMoney(amount, currency) {\n");
        writer.write("  var absValue = Math.abs(Number(amount || 0));\n");
        writer.write("  var prefix = Number(amount || 0) < 0 ? '-' : '';\n");
        writer.write("  if (absValue >= 1000000000) return prefix + Number(absValue / 1000000000.0).toFixed(1) + 'B ' + currency;\n");
        writer.write("  if (absValue >= 1000000) return prefix + Number(absValue / 1000000.0).toFixed(1) + 'M ' + currency;\n");
        writer.write("  if (absValue >= 1000) return prefix + Number(absValue / 1000.0).toFixed(0) + 'k ' + currency;\n");
        writer.write("  return prefix + Number(absValue).toFixed(0) + ' ' + currency;\n");
        writer.write("}\n");
        writer.write("function convertBucketsToCurrency(buckets, targetCurrency) {\n");
        writer.write("  var target = normalizeCurrencyCodeInput(targetCurrency);\n");
        writer.write("  var targetRate = REPORT_RATES_TO_NOK[target];\n");
        writer.write("  if (!targetRate || targetRate <= 0) return null;\n");
        writer.write("  var totalNok = 0;\n");
        writer.write("  for (var code in buckets) {\n");
        writer.write("    if (!Object.prototype.hasOwnProperty.call(buckets, code)) continue;\n");
        writer.write("    var sourceRate = REPORT_RATES_TO_NOK[normalizeCurrencyCodeInput(code)];\n");
        writer.write("    if (!sourceRate || sourceRate <= 0) continue;\n");
        writer.write("    totalNok += Number(buckets[code] || 0) * sourceRate;\n");
        writer.write("  }\n");
        writer.write("  return totalNok / targetRate;\n");
        writer.write("}\n");
        writer.write("function refreshReportTotalsCurrency(targetCurrency) {\n");
        writer.write("  var target = normalizeCurrencyCodeInput(targetCurrency);\n");
        writer.write("  if (!REPORT_RATES_TO_NOK[target]) return false;\n");
        writer.write("  var fields = document.querySelectorAll('.js-convert-money');\n");
        writer.write("  fields.forEach(function (field) {\n");
        writer.write("    var raw = field.getAttribute('data-buckets');\n");
        writer.write("    if (!raw) return;\n");
        writer.write("    try {\n");
        writer.write("      var buckets = JSON.parse(raw);\n");
        writer.write("      var decimals = Number(field.getAttribute('data-decimals') || '2');\n");
        writer.write("      var converted = convertBucketsToCurrency(buckets, target);\n");
        writer.write("      if (converted == null) return;\n");
        writer.write("      field.textContent = formatMoneyValue(converted, target, decimals);\n");
        writer.write("    } catch (e) {\n");
        writer.write("      // Keep existing content if parsing fails.\n");
        writer.write("    }\n");
        writer.write("  });\n");
        writer.write("  document.querySelectorAll('.js-report-currency-code').forEach(function(node) {\n");
        writer.write("    node.textContent = target;\n");
        writer.write("  });\n");
        writer.write("  return true;\n");
        writer.write("}\n");
        writer.write("function refreshReportChartsCurrency(targetCurrency) {\n");
        writer.write("  var target = normalizeCurrencyCodeInput(targetCurrency);\n");
        writer.write("  var targetRate = REPORT_RATES_TO_NOK[target];\n");
        writer.write("  if (!targetRate || targetRate <= 0) return false;\n");
        writer.write("  document.querySelectorAll('.js-total-return-money-title').forEach(function (node) {\n");
        writer.write("    node.textContent = 'Total Return (' + target + ')';\n");
        writer.write("  });\n");
        writer.write("  document.querySelectorAll('.js-return-amount-label').forEach(function (node) {\n");
        writer.write("    node.textContent = 'Return (' + target + ')';\n");
        writer.write("  });\n");
        writer.write("  document.querySelectorAll('.js-chart-money').forEach(function (node) {\n");
        writer.write("    var valueNok = Number(node.getAttribute('data-value-nok') || '0');\n");
        writer.write("    var decimals = Number(node.getAttribute('data-decimals') || '0');\n");
        writer.write("    var prefix = node.getAttribute('data-prefix') || '';\n");
        writer.write("    var suffix = node.getAttribute('data-suffix') || '';\n");
        writer.write("    var mode = node.getAttribute('data-format') || 'money';\n");
        writer.write("    var converted = valueNok / targetRate;\n");
        writer.write("    var text = mode === 'compact'\n");
        writer.write("      ? prefix + formatCompactMoney(converted, target)\n");
        writer.write("      : prefix + formatMoneyValue(converted, target, decimals);\n");
        writer.write("    text += suffix;\n");
        writer.write("    node.textContent = text;\n");
        writer.write("  });\n");
        writer.write("  refreshSparklineSummaries();\n");
        writer.write("  return true;\n");
        writer.write("}\n");
        writer.write("function parseBucketsJson(raw) {\n");
        writer.write("  if (!raw) return {};\n");
        writer.write("  try {\n");
        writer.write("    var parsed = JSON.parse(raw);\n");
        writer.write("    return parsed && typeof parsed === 'object' ? parsed : {};\n");
        writer.write("  } catch (_) {\n");
        writer.write("    return {};\n");
        writer.write("  }\n");
        writer.write("}\n");
        writer.write("function addBucketValue(target, currency, amount) {\n");
        writer.write("  var code = normalizeCurrencyCodeInput(currency || 'NOK');\n");
        writer.write("  if (!code) code = 'NOK';\n");
        writer.write("  var value = Number(amount || 0);\n");
        writer.write("  target[code] = Number(target[code] || 0) + value;\n");
        writer.write("}\n");
        writer.write("function mergeBuckets(base, extra) {\n");
        writer.write("  var merged = {};\n");
        writer.write("  Object.keys(base || {}).forEach(function(code) {\n");
        writer.write("    merged[code] = Number(base[code] || 0);\n");
        writer.write("  });\n");
        writer.write("  Object.keys(extra || {}).forEach(function(code) {\n");
        writer.write("    merged[code] = Number(merged[code] || 0) + Number(extra[code] || 0);\n");
        writer.write("  });\n");
        writer.write("  return merged;\n");
        writer.write("}\n");
        writer.write("function refreshPortfolioValueBuckets() {\n");
        writer.write("  var portfolioNode = document.getElementById('hero-portfolio-value');\n");
        writer.write("  if (!portfolioNode) return;\n");
        writer.write("  var marketNode = document.getElementById('hero-total-market-value');\n");
        writer.write("  var cashNode = document.getElementById('cash-holdings-total');\n");
        writer.write("  var marketBuckets = parseBucketsJson(marketNode ? marketNode.getAttribute('data-buckets') : '{}');\n");
        writer.write("  var cashBuckets = parseBucketsJson(cashNode ? cashNode.getAttribute('data-buckets') : '{}');\n");
        writer.write("  portfolioNode.setAttribute('data-buckets', JSON.stringify(mergeBuckets(marketBuckets, cashBuckets)));\n");
        writer.write("}\n");
        writer.write("function getActiveReportCurrency() {\n");
        writer.write("  var input = document.getElementById('portfolio-currency-input');\n");
        writer.write("  var target = normalizeCurrencyCodeInput(input && input.value ? input.value : 'NOK');\n");
        writer.write("  if (!target || !REPORT_RATES_TO_NOK[target]) return 'NOK';\n");
        writer.write("  return target;\n");
        writer.write("}\n");
        writer.write("function formatPercentValue(value, decimals) {\n");
        writer.write("  return formatGroupedNumber(Number(value || 0), decimals) + '%';\n");
        writer.write("}\n");
        writer.write("function resolvePriceRefreshApiUrl() {\n");
        writer.write("  if (window.__portfolioReportPriceApiUrl) return String(window.__portfolioReportPriceApiUrl);\n");
        writer.write("  if (window.REPORT_PRICE_API_URL) return String(window.REPORT_PRICE_API_URL);\n");
        writer.write("  try {\n");
        writer.write("    if (window.parent && window.parent !== window) {\n");
        writer.write("      if (window.parent.__portfolioReportPriceApiUrl) return String(window.parent.__portfolioReportPriceApiUrl);\n");
        writer.write("      if (window.parent.REPORT_PRICE_API_URL) return String(window.parent.REPORT_PRICE_API_URL);\n");
        writer.write("    }\n");
        writer.write("  } catch (_) {\n");
        writer.write("    // Ignore parent access issues.\n");
        writer.write("  }\n");
        writer.write("  return '';\n");
        writer.write("}\n");
        writer.write("async function fetchLatestPricesFromApi(tickers) {\n");
        writer.write("  var apiUrl = resolvePriceRefreshApiUrl();\n");
        writer.write("  if (!apiUrl) throw new Error('No price API configured.');\n");
        writer.write("  var response = await fetch(apiUrl, {\n");
        writer.write("    method: 'POST',\n");
        writer.write("    headers: { 'Content-Type': 'application/json' },\n");
        writer.write("    body: JSON.stringify({ tickers: tickers })\n");
        writer.write("  });\n");
        writer.write("  if (!response.ok) {\n");
        writer.write("    var message = 'Price refresh failed with status ' + response.status;\n");
        writer.write("    try {\n");
        writer.write("      var payload = await response.json();\n");
        writer.write("      if (payload && payload.error) message = payload.error;\n");
        writer.write("    } catch (_) {\n");
        writer.write("      // Keep status-based message when no payload is available.\n");
        writer.write("    }\n");
        writer.write("    throw new Error(message);\n");
        writer.write("  }\n");
        writer.write("  var data = await response.json();\n");
        writer.write("  return data && data.prices && typeof data.prices === 'object' ? data.prices : {};\n");
        writer.write("}\n");
        writer.write("async function fetchLatestDaySeriesFromApi(tickers) {\n");
        writer.write("  var apiUrl = resolvePriceRefreshApiUrl();\n");
        writer.write("  if (!apiUrl) throw new Error('No price API configured.');\n");
        writer.write("  var response = await fetch(apiUrl, {\n");
        writer.write("    method: 'POST',\n");
        writer.write("    headers: { 'Content-Type': 'application/json' },\n");
        writer.write("    body: JSON.stringify({ tickers: tickers, includeDaySeries: true })\n");
        writer.write("  });\n");
        writer.write("  if (!response.ok) {\n");
        writer.write("    throw new Error('Day series fetch failed with status ' + response.status);\n");
        writer.write("  }\n");
        writer.write("  var data = await response.json();\n");
        writer.write("  var payload = data && data.daySeries && typeof data.daySeries === 'object' ? data.daySeries : {};\n");
        writer.write("  var normalized = {};\n");
        writer.write("  Object.keys(payload).forEach(function(ticker) {\n");
        writer.write("    var series = payload[ticker];\n");
        writer.write("    normalized[String(ticker || '').trim().toUpperCase()] = Array.isArray(series)\n");
        writer.write("      ? series.map(function(value) { return Number(value); }).filter(function(value) { return Number.isFinite(value) && value > 0; })\n");
        writer.write("      : [];\n");
        writer.write("  });\n");
        writer.write("  return normalized;\n");
        writer.write("}\n");
        writer.write("async function fetchLatestPriceFromYahooDirect(ticker) {\n");
        writer.write("  var symbol = String(ticker || '').trim().toUpperCase();\n");
        writer.write("  if (!symbol) return 0;\n");
        writer.write("  var url = 'https://query2.finance.yahoo.com/v8/finance/chart/' + encodeURIComponent(symbol) + '?interval=1d&range=5d';\n");
        writer.write("  var response = await fetch(url, { method: 'GET', headers: { 'Accept': 'application/json' } });\n");
        writer.write("  if (!response.ok) return 0;\n");
        writer.write("  var data = await response.json();\n");
        writer.write("  var result = data && data.chart && data.chart.result && data.chart.result[0];\n");
        writer.write("  if (!result) return 0;\n");
        writer.write("  var marketPrice = Number(result.meta && result.meta.regularMarketPrice || 0);\n");
        writer.write("  if (Number.isFinite(marketPrice) && marketPrice > 0) return marketPrice;\n");
        writer.write("  var closes = result.indicators && result.indicators.quote && result.indicators.quote[0] && result.indicators.quote[0].close;\n");
        writer.write("  if (!Array.isArray(closes)) return 0;\n");
        writer.write("  for (var i = closes.length - 1; i >= 0; i -= 1) {\n");
        writer.write("    var value = Number(closes[i]);\n");
        writer.write("    if (Number.isFinite(value) && value > 0) return value;\n");
        writer.write("  }\n");
        writer.write("  return 0;\n");
        writer.write("}\n");
        writer.write("async function fetchYahooChartResult(symbol, interval, range) {\n");
        writer.write("  var url = 'https://query2.finance.yahoo.com/v8/finance/chart/' + encodeURIComponent(symbol)\n");
        writer.write("    + '?interval=' + encodeURIComponent(interval)\n");
        writer.write("    + '&range=' + encodeURIComponent(range)\n");
        writer.write("    + '&includePrePost=false';\n");
        writer.write("  var response = await fetch(url, { method: 'GET', headers: { 'Accept': 'application/json' } });\n");
        writer.write("  if (!response.ok) return null;\n");
        writer.write("  var data = await response.json();\n");
        writer.write("  return data && data.chart && data.chart.result && data.chart.result[0] ? data.chart.result[0] : null;\n");
        writer.write("}\n");
        writer.write("function extractLatestTradingSessionSeries(result) {\n");
        writer.write("  if (!result) return [];\n");
        writer.write("  var timestamps = Array.isArray(result.timestamp) ? result.timestamp : [];\n");
        writer.write("  var closeSeries = result.indicators && result.indicators.quote && result.indicators.quote[0] && result.indicators.quote[0].close;\n");
        writer.write("  if (!Array.isArray(closeSeries) || !timestamps.length) return [];\n");
        writer.write("  var grouped = new Map();\n");
        writer.write("  var length = Math.min(timestamps.length, closeSeries.length);\n");
        writer.write("  for (var i = 0; i < length; i += 1) {\n");
        writer.write("    var ts = Number(timestamps[i]);\n");
        writer.write("    var closeValue = Number(closeSeries[i]);\n");
        writer.write("    if (!Number.isFinite(ts) || !Number.isFinite(closeValue) || closeValue <= 0) continue;\n");
        writer.write("    var dayKey = new Date(ts * 1000).toISOString().slice(0, 10);\n");
        writer.write("    if (!grouped.has(dayKey)) grouped.set(dayKey, []);\n");
        writer.write("    grouped.get(dayKey).push(closeValue);\n");
        writer.write("  }\n");
        writer.write("  if (!grouped.size) return [];\n");
        writer.write("  var orderedDays = Array.from(grouped.keys()).sort();\n");
        writer.write("  for (var j = orderedDays.length - 1; j >= 0; j -= 1) {\n");
        writer.write("    var daySeries = grouped.get(orderedDays[j]) || [];\n");
        writer.write("    if (daySeries.length >= 6) return daySeries;\n");
        writer.write("  }\n");
        writer.write("  return [];\n");
        writer.write("}\n");
        writer.write("async function fetchDaySeriesFromYahooDirect(ticker) {\n");
        writer.write("  var symbol = String(ticker || '').trim().toUpperCase();\n");
        writer.write("  if (!symbol) return [];\n");
        writer.write("  var oneDay = await fetchYahooChartResult(symbol, '5m', '1d');\n");
        writer.write("  var oneDaySeries = extractLatestTradingSessionSeries(oneDay);\n");
        writer.write("  if (oneDaySeries.length >= 6) return oneDaySeries;\n");
        writer.write("  var fiveDay = await fetchYahooChartResult(symbol, '5m', '5d');\n");
        writer.write("  var fiveDaySeries = extractLatestTradingSessionSeries(fiveDay);\n");
        writer.write("  return fiveDaySeries.length >= 6 ? fiveDaySeries : [];\n");
        writer.write("}\n");
        writer.write("function buildMiniDayChartSvg(prices) {\n");
        writer.write("  var chartPrices = Array.isArray(prices)\n");
        writer.write("    ? prices.map(function(value) { return Number(value); }).filter(function(value) { return Number.isFinite(value) && value > 0; })\n");
        writer.write("    : [];\n");
        writer.write("  if (chartPrices.length < 6) return '';\n");
        writer.write("  var width = 96;\n");
        writer.write("  var height = 30;\n");
        writer.write("  var left = 1.5;\n");
        writer.write("  var right = width - 1.5;\n");
        writer.write("  var top = 2.0;\n");
        writer.write("  var bottom = height - 2.0;\n");
        writer.write("  var min = Math.min.apply(null, chartPrices);\n");
        writer.write("  var max = Math.max.apply(null, chartPrices);\n");
        writer.write("  var span = Math.max(0.000001, max - min);\n");
        writer.write("  var pointsArray = chartPrices.map(function(price, index) {\n");
        writer.write("    var ratioX = chartPrices.length <= 1 ? 0 : index / (chartPrices.length - 1);\n");
        writer.write("    var ratioY = (price - min) / span;\n");
        writer.write("    var x = left + ratioX * (right - left);\n");
        writer.write("    var y = bottom - ratioY * (bottom - top);\n");
        writer.write("    return { x: x, y: y };\n");
        writer.write("  });\n");
        writer.write("  var points = pointsArray.map(function(point) { return point.x.toFixed(2) + ',' + point.y.toFixed(2); }).join(' ');\n");
        writer.write("  var start = Number(chartPrices[0]);\n");
        writer.write("  var end = Number(chartPrices[chartPrices.length - 1]);\n");
        writer.write("  var lineClass = end >= start ? 'positive' : 'negative';\n");
        writer.write("  var openRatio = (start - min) / span;\n");
        writer.write("  var openY = bottom - openRatio * (bottom - top);\n");
        writer.write("  openY = Math.max(top, Math.min(bottom, openY));\n");
        writer.write("  var first = pointsArray[0];\n");
        writer.write("  var last = pointsArray[pointsArray.length - 1];\n");
        writer.write("  var areaPoints = first.x.toFixed(2) + ',' + openY.toFixed(2) + ' ' + points + ' ' + last.x.toFixed(2) + ',' + openY.toFixed(2);\n");
        writer.write("  return '<svg class=\"mini-day-chart\" viewBox=\"0 0 ' + width + ' ' + height + '\" xmlns=\"http://www.w3.org/2000/svg\" aria-hidden=\"true\">'\n");
        writer.write("    + '<line class=\"mini-day-chart-open\" x1=\"0\" y1=\"' + openY.toFixed(2) + '\" x2=\"' + width + '\" y2=\"' + openY.toFixed(2) + '\"></line>'\n");
        writer.write("    + '<polygon class=\"mini-day-chart-area ' + lineClass + '\" points=\"' + areaPoints + '\"></polygon>'\n");
        writer.write("    + '<polyline class=\"mini-day-chart-line ' + lineClass + '\" points=\"' + points + '\"></polyline>'\n");
        writer.write("    + '<circle class=\"mini-day-chart-end ' + lineClass + '\" cx=\"' + last.x.toFixed(2) + '\" cy=\"' + last.y.toFixed(2) + '\" r=\"2.3\"></circle>'\n");
        writer.write("    + '</svg>';\n");
        writer.write("}\n");
        writer.write("function initOverviewDayCharts() {\n");
        writer.write("  var nodes = Array.prototype.slice.call(document.querySelectorAll('.js-row-day-chart[data-ticker]'));\n");
        writer.write("  if (!nodes.length) return;\n");
        writer.write("  nodes.forEach(function(node) { node.textContent = '-'; });\n");
        writer.write("  var byTicker = new Map();\n");
        writer.write("  nodes.forEach(function(node) {\n");
        writer.write("    var ticker = String(node.getAttribute('data-ticker') || '').trim().toUpperCase();\n");
        writer.write("    if (!ticker) return;\n");
        writer.write("    if (!byTicker.has(ticker)) byTicker.set(ticker, []);\n");
        writer.write("    byTicker.get(ticker).push(node);\n");
        writer.write("  });\n");
        writer.write("  var tickers = Array.from(byTicker.keys());\n");
        writer.write("  if (!tickers.length) return;\n");
        writer.write("  fetchLatestDaySeriesFromApi(tickers).then(function(seriesMap) {\n");
        writer.write("    byTicker.forEach(function(targets, ticker) {\n");
        writer.write("      var series = seriesMap[ticker] || [];\n");
        writer.write("      var svg = buildMiniDayChartSvg(series);\n");
        writer.write("      targets.forEach(function(node) {\n");
        writer.write("        if (svg) node.innerHTML = svg;\n");
        writer.write("        else node.textContent = '-';\n");
        writer.write("      });\n");
        writer.write("    });\n");
        writer.write("  }).catch(function() {\n");
        writer.write("    byTicker.forEach(function(targets) {\n");
        writer.write("      targets.forEach(function(node) { node.textContent = '-'; });\n");
        writer.write("    });\n");
        writer.write("  });\n");
        writer.write("}\n");
        writer.write("async function fetchLatestPricesDirect(tickers) {\n");
        writer.write("  var prices = {};\n");
        writer.write("  for (var i = 0; i < tickers.length; i += 1) {\n");
        writer.write("    var symbol = String(tickers[i] || '').trim().toUpperCase();\n");
        writer.write("    if (!symbol) continue;\n");
        writer.write("    try {\n");
        writer.write("      prices[symbol] = await fetchLatestPriceFromYahooDirect(symbol);\n");
        writer.write("    } catch (_) {\n");
        writer.write("      prices[symbol] = 0;\n");
        writer.write("    }\n");
        writer.write("  }\n");
        writer.write("  return prices;\n");
        writer.write("}\n");
        writer.write("async function fetchLatestPrices(tickers) {\n");
        writer.write("  try {\n");
        writer.write("    return await fetchLatestPricesFromApi(tickers);\n");
        writer.write("  } catch (_) {\n");
        writer.write("    return fetchLatestPricesDirect(tickers);\n");
        writer.write("  }\n");
        writer.write("}\n");
        writer.write("function findOverviewRowsBySecurityKey(securityKey) {\n");
        writer.write("  var key = String(securityKey || '').trim();\n");
        writer.write("  if (!key) return [];\n");
        writer.write("  return Array.prototype.slice.call(document.querySelectorAll('tr[data-overview-security-key]')).filter(function(row) {\n");
        writer.write("    return String(row.getAttribute('data-overview-security-key') || '').trim() === key;\n");
        writer.write("  });\n");
        writer.write("}\n");
        writer.write("function applyComputedOverviewRowValues(row, nextPrice) {\n");
        writer.write("  if (!row || !Number.isFinite(nextPrice) || nextPrice <= 0) return false;\n");
        writer.write("  var currency = normalizeCurrencyCodeInput(row.getAttribute('data-currency') || 'NOK');\n");
        writer.write("  var units = Number(row.getAttribute('data-units') || 0);\n");
        writer.write("  var positionCostBasis = Number(row.getAttribute('data-position-cost-basis') || 0);\n");
        writer.write("  var realized = Number(row.getAttribute('data-realized') || 0);\n");
        writer.write("  var dividends = Number(row.getAttribute('data-dividends') || 0);\n");
        writer.write("  var historicalCostBasis = Number(row.getAttribute('data-historical-cost-basis') || 0);\n");
        writer.write("  var previousClose = Number(row.getAttribute('data-previous-close') || 0);\n");
        writer.write("  var marketValue = units * nextPrice;\n");
        writer.write("  var unrealized = marketValue - positionCostBasis;\n");
        writer.write("  var unrealizedPct = positionCostBasis > 0 ? (unrealized / positionCostBasis) * 100 : 0;\n");
        writer.write("  var totalReturn = unrealized + realized + dividends;\n");
        writer.write("  var totalReturnPct = historicalCostBasis > 0 ? (totalReturn / historicalCostBasis) * 100 : 0;\n");
        writer.write("  var hasDayChange = Number.isFinite(previousClose) && previousClose > 0;\n");
        writer.write("  var dayChangeValue = hasDayChange ? (nextPrice - previousClose) : Number.NaN;\n");
        writer.write("  var dayChangePct = hasDayChange ? ((nextPrice / previousClose) - 1) * 100 : Number.NaN;\n");
        writer.write("  row.setAttribute('data-latest-price', String(nextPrice));\n");
        writer.write("  if (Number.isFinite(marketValue)) row.setAttribute('data-market-value', String(marketValue));\n");
        writer.write("  if (Number.isFinite(unrealized)) row.setAttribute('data-unrealized-value', String(unrealized));\n");
        writer.write("  if (Number.isFinite(totalReturn)) row.setAttribute('data-total-return-value', String(totalReturn));\n");
        writer.write("  var priceCell = row.querySelector('.js-row-last-price');\n");
        writer.write("  var marketCell = row.querySelector('.js-row-market-value');\n");
        writer.write("  var unrealizedCell = row.querySelector('.js-row-unrealized');\n");
        writer.write("  var totalReturnCell = row.querySelector('.js-row-total-return');\n");
        writer.write("  var dayChangeCell = row.querySelector('.js-row-day-change');\n");
        writer.write("  var dayChangeValueCell = row.querySelector('.js-row-day-change-value');\n");
        writer.write("  var dayChangePositionValueCell = row.querySelector('.js-row-day-change-value-position');\n");
        writer.write("  var fundamentalsPriceCell = row.querySelector('.js-row-fundamentals-last-price');\n");
        writer.write("  var dayChartCell = row.querySelector('.js-row-day-chart');\n");
        writer.write("  if (priceCell) priceCell.textContent = formatMoneyValue(nextPrice, currency, 2);\n");
        writer.write("  if (marketCell) marketCell.textContent = formatMoneyValue(marketValue, currency, 2);\n");
        writer.write("  if (unrealizedCell) {\n");
        writer.write("    unrealizedCell.textContent = formatMoneyValue(unrealized, currency, 2) + ' (' + formatPercentValue(unrealizedPct, 2) + ')';\n");
        writer.write("    unrealizedCell.classList.remove('positive', 'negative');\n");
        writer.write("    if (unrealized > 0) unrealizedCell.classList.add('positive');\n");
        writer.write("    else if (unrealized < 0) unrealizedCell.classList.add('negative');\n");
        writer.write("  }\n");
        writer.write("  if (totalReturnCell) {\n");
        writer.write("    totalReturnCell.textContent = formatMoneyValue(totalReturn, currency, 2) + ' (' + formatPercentValue(totalReturnPct, 2) + ')';\n");
        writer.write("    totalReturnCell.classList.remove('positive', 'negative');\n");
        writer.write("    if (totalReturn > 0) totalReturnCell.classList.add('positive');\n");
        writer.write("    else if (totalReturn < 0) totalReturnCell.classList.add('negative');\n");
        writer.write("  }\n");
        writer.write("  if (fundamentalsPriceCell) fundamentalsPriceCell.textContent = formatMoneyValue(nextPrice, currency, 2);\n");
        writer.write("  if (dayChangeCell) {\n");
        writer.write("    if (hasDayChange && Number.isFinite(dayChangePct)) {\n");
        writer.write("      dayChangeCell.textContent = formatPercentValue(dayChangePct, 2);\n");
        writer.write("      dayChangeCell.classList.remove('positive', 'negative');\n");
        writer.write("      if (dayChangePct > 0) dayChangeCell.classList.add('positive');\n");
        writer.write("      else if (dayChangePct < 0) dayChangeCell.classList.add('negative');\n");
        writer.write("    } else {\n");
        writer.write("      dayChangeCell.textContent = '-';\n");
        writer.write("      dayChangeCell.classList.remove('positive', 'negative');\n");
        writer.write("    }\n");
        writer.write("  }\n");
        writer.write("  if (dayChangeValueCell) {\n");
        writer.write("    if (hasDayChange && Number.isFinite(dayChangeValue)) {\n");
        writer.write("      dayChangeValueCell.textContent = formatMoneyValue(dayChangeValue, currency, 2);\n");
        writer.write("      dayChangeValueCell.classList.remove('positive', 'negative');\n");
        writer.write("      if (dayChangeValue > 0) dayChangeValueCell.classList.add('positive');\n");
        writer.write("      else if (dayChangeValue < 0) dayChangeValueCell.classList.add('negative');\n");
        writer.write("    } else {\n");
        writer.write("      dayChangeValueCell.textContent = '-';\n");
        writer.write("      dayChangeValueCell.classList.remove('positive', 'negative');\n");
        writer.write("    }\n");
        writer.write("  }\n");
        writer.write("  if (dayChangePositionValueCell) {\n");
        writer.write("    var positionDayChangeValue = hasDayChange ? (dayChangeValue * units) : Number.NaN;\n");
        writer.write("    if (Number.isFinite(positionDayChangeValue)) {\n");
        writer.write("      dayChangePositionValueCell.textContent = formatMoneyValue(positionDayChangeValue, currency, 2);\n");
        writer.write("      dayChangePositionValueCell.classList.remove('positive', 'negative');\n");
        writer.write("      if (positionDayChangeValue > 0) dayChangePositionValueCell.classList.add('positive');\n");
        writer.write("      else if (positionDayChangeValue < 0) dayChangePositionValueCell.classList.add('negative');\n");
        writer.write("    } else {\n");
        writer.write("      dayChangePositionValueCell.textContent = '-';\n");
        writer.write("      dayChangePositionValueCell.classList.remove('positive', 'negative');\n");
        writer.write("    }\n");
        writer.write("  }\n");
        writer.write("  if (dayChartCell) {\n");
        writer.write("    dayChartCell.textContent = '-';\n");
        writer.write("  }\n");
        writer.write("  return true;\n");
        writer.write("}\n");
        writer.write("function applyOverviewRowPrice(row, nextPrice) {\n");
        writer.write("  if (!row || !Number.isFinite(nextPrice) || nextPrice <= 0) return false;\n");
        writer.write("  var securityKey = String(row.getAttribute('data-overview-security-key') || '').trim();\n");
        writer.write("  var linkedRows = securityKey ? findOverviewRowsBySecurityKey(securityKey) : [];\n");
        writer.write("  if (!linkedRows.length) linkedRows = [row];\n");
        writer.write("  return linkedRows.some(function(targetRow) {\n");
        writer.write("    return applyComputedOverviewRowValues(targetRow, nextPrice);\n");
        writer.write("  });\n");
        writer.write("}\n");
        writer.write("function recalculateOverviewAndHeaderTotalsAfterPriceRefresh() {\n");
        writer.write("  var rows = Array.prototype.slice.call(document.querySelectorAll('tr[data-overview-row=\"1\"]'));\n");
        writer.write("  var totalMarketBuckets = {};\n");
        writer.write("  var totalCostBasisBuckets = {};\n");
        writer.write("  var totalUnrealizedBuckets = {};\n");
        writer.write("  var totalRealizedBuckets = {};\n");
        writer.write("  var totalDividendsBuckets = {};\n");
        writer.write("  var totalHistoricalBuckets = {};\n");
        writer.write("  var activeTotalReturnBuckets = {};\n");
        writer.write("  var dayChangeBuckets = {};\n");
        writer.write("  var previousDayValueBuckets = {};\n");
        writer.write("  rows.forEach(function(row) {\n");
        writer.write("    var currency = normalizeCurrencyCodeInput(row.getAttribute('data-currency') || 'NOK');\n");
        writer.write("    var latestPrice = Number(row.getAttribute('data-latest-price') || 0);\n");
        writer.write("    var previousClose = Number(row.getAttribute('data-previous-close') || 0);\n");
        writer.write("    var units = Number(row.getAttribute('data-units') || 0);\n");
        writer.write("    var positionCostBasis = Number(row.getAttribute('data-position-cost-basis') || 0);\n");
        writer.write("    var realized = Number(row.getAttribute('data-realized') || 0);\n");
        writer.write("    var dividends = Number(row.getAttribute('data-dividends') || 0);\n");
        writer.write("    var historicalCostBasis = Number(row.getAttribute('data-historical-cost-basis') || 0);\n");
        writer.write("    var hasPrice = Number.isFinite(latestPrice) && latestPrice > 0;\n");
        writer.write("    var marketValue = hasPrice ? (units * latestPrice) : 0;\n");
        writer.write("    var unrealized = hasPrice ? (marketValue - positionCostBasis) : 0;\n");
        writer.write("    var totalReturn = unrealized + realized + dividends;\n");
        writer.write("    addBucketValue(totalMarketBuckets, currency, marketValue);\n");
        writer.write("    addBucketValue(totalCostBasisBuckets, currency, positionCostBasis);\n");
        writer.write("    addBucketValue(totalUnrealizedBuckets, currency, unrealized);\n");
        writer.write("    addBucketValue(totalRealizedBuckets, currency, realized);\n");
        writer.write("    addBucketValue(totalDividendsBuckets, currency, dividends);\n");
        writer.write("    addBucketValue(totalHistoricalBuckets, currency, historicalCostBasis);\n");
        writer.write("    addBucketValue(activeTotalReturnBuckets, currency, totalReturn);\n");
        writer.write("    if (hasPrice && Number.isFinite(previousClose) && previousClose > 0) {\n");
        writer.write("      addBucketValue(dayChangeBuckets, currency, units * (latestPrice - previousClose));\n");
        writer.write("      addBucketValue(previousDayValueBuckets, currency, units * previousClose);\n");
        writer.write("    }\n");
        writer.write("  });\n");
        writer.write("  var totalReturnBuckets = mergeBuckets(totalUnrealizedBuckets, mergeBuckets(totalRealizedBuckets, totalDividendsBuckets));\n");
        writer.write("  var totalCostBasisNok = Number(convertBucketsToCurrency(totalCostBasisBuckets, 'NOK') || 0);\n");
        writer.write("  var totalUnrealizedNok = Number(convertBucketsToCurrency(totalUnrealizedBuckets, 'NOK') || 0);\n");
        writer.write("  var totalRealizedNok = Number(convertBucketsToCurrency(totalRealizedBuckets, 'NOK') || 0);\n");
        writer.write("  var totalReturnNok = Number(convertBucketsToCurrency(totalReturnBuckets, 'NOK') || 0);\n");
        writer.write("  var totalHistoricalNok = Number(convertBucketsToCurrency(totalHistoricalBuckets, 'NOK') || 0);\n");
        writer.write("  var dayChangeNok = Number(convertBucketsToCurrency(dayChangeBuckets, 'NOK') || 0);\n");
        writer.write("  var previousDayValueNok = Number(convertBucketsToCurrency(previousDayValueBuckets, 'NOK') || 0);\n");
        writer.write("  var unrealizedPct = totalCostBasisNok > 0 ? (totalUnrealizedNok / totalCostBasisNok) * 100 : 0;\n");
        writer.write("  var realizedPct = totalCostBasisNok > 0 ? (totalRealizedNok / totalCostBasisNok) * 100 : 0;\n");
        writer.write("  var totalReturnPct = totalHistoricalNok > 0 ? (totalReturnNok / totalHistoricalNok) * 100 : 0;\n");
        writer.write("  var dayChangePct = previousDayValueNok > 0 ? (dayChangeNok / previousDayValueNok) * 100 : 0;\n");
        writer.write("  var mapping = [\n");
        writer.write("    ['overview-total-cost-basis', totalCostBasisBuckets],\n");
        writer.write("    ['overview-total-market-value', totalMarketBuckets],\n");
        writer.write("    ['overview-total-unrealized', totalUnrealizedBuckets],\n");
        writer.write("    ['overview-total-realized', totalRealizedBuckets],\n");
        writer.write("    ['overview-total-dividends', totalDividendsBuckets],\n");
        writer.write("    ['overview-total-return', totalReturnBuckets],\n");
        writer.write("    ['holdings-total-cost-basis', totalCostBasisBuckets],\n");
        writer.write("    ['holdings-total-market-value', totalMarketBuckets],\n");
        writer.write("    ['holdings-total-unrealized-value', totalUnrealizedBuckets],\n");
        writer.write("    ['holdings-total-realized-value', totalRealizedBuckets],\n");
        writer.write("    ['holdings-total-dividends-value', totalDividendsBuckets],\n");
        writer.write("    ['holdings-total-total-return-value', totalReturnBuckets],\n");
        writer.write("    ['holdings-total-day-change-value', dayChangeBuckets],\n");
        writer.write("    ['hero-total-market-value', totalMarketBuckets],\n");
        writer.write("    ['hero-unrealized-value', totalUnrealizedBuckets],\n");
        writer.write("    ['hero-realized-value', totalRealizedBuckets],\n");
        writer.write("    ['hero-day-change-value', dayChangeBuckets]\n");
        writer.write("  ];\n");
        writer.write("  mapping.forEach(function(entry) {\n");
        writer.write("    var node = document.getElementById(entry[0]);\n");
        writer.write("    if (!node) return;\n");
        writer.write("    node.setAttribute('data-buckets', JSON.stringify(entry[1]));\n");
        writer.write("  });\n");
        writer.write("  var overviewUnrealizedPct = document.getElementById('overview-total-unrealized-pct');\n");
        writer.write("  var overviewRealizedPct = document.getElementById('overview-total-realized-pct');\n");
        writer.write("  var overviewTotalPct = document.getElementById('overview-total-return-pct');\n");
        writer.write("  if (overviewUnrealizedPct) overviewUnrealizedPct.textContent = formatPercentValue(unrealizedPct, 2);\n");
        writer.write("  if (overviewRealizedPct) overviewRealizedPct.textContent = formatPercentValue(realizedPct, 2);\n");
        writer.write("  if (overviewTotalPct) overviewTotalPct.textContent = formatPercentValue(totalReturnPct, 2);\n");
        writer.write("  var holdingsDayChangePct = document.getElementById('holdings-total-day-change-pct');\n");
        writer.write("  var holdingsDayChangeValue = document.getElementById('holdings-total-day-change-value');\n");
        writer.write("  if (holdingsDayChangePct) {\n");
        writer.write("    holdingsDayChangePct.classList.remove('positive', 'negative');\n");
        writer.write("    if (previousDayValueNok > 0 && Number.isFinite(dayChangePct)) {\n");
        writer.write("      holdingsDayChangePct.textContent = formatPercentValue(dayChangePct, 2);\n");
        writer.write("      if (dayChangePct > 0) holdingsDayChangePct.classList.add('positive');\n");
        writer.write("      else if (dayChangePct < 0) holdingsDayChangePct.classList.add('negative');\n");
        writer.write("    } else {\n");
        writer.write("      holdingsDayChangePct.textContent = '-';\n");
        writer.write("    }\n");
        writer.write("  }\n");
        writer.write("  if (holdingsDayChangeValue) {\n");
        writer.write("    holdingsDayChangeValue.classList.remove('positive', 'negative');\n");
        writer.write("    if (dayChangeNok > 0) holdingsDayChangeValue.classList.add('positive');\n");
        writer.write("    else if (dayChangeNok < 0) holdingsDayChangeValue.classList.add('negative');\n");
        writer.write("  }\n");
        writer.write("  var holdingsUnrealizedPct = document.getElementById('holdings-total-unrealized-pct');\n");
        writer.write("  var holdingsUnrealizedPctWrap = document.getElementById('holdings-total-unrealized-pct-wrap');\n");
        writer.write("  var holdingsUnrealizedValue = document.getElementById('holdings-total-unrealized-value');\n");
        writer.write("  if (holdingsUnrealizedPct) {\n");
        writer.write("    holdingsUnrealizedPct.textContent = formatPercentValue(unrealizedPct, 2);\n");
        writer.write("    holdingsUnrealizedPct.classList.remove('positive', 'negative');\n");
        writer.write("    if (totalUnrealizedNok > 0) holdingsUnrealizedPct.classList.add('positive');\n");
        writer.write("    else if (totalUnrealizedNok < 0) holdingsUnrealizedPct.classList.add('negative');\n");
        writer.write("  }\n");
        writer.write("  if (holdingsUnrealizedPctWrap) {\n");
        writer.write("    holdingsUnrealizedPctWrap.classList.remove('positive', 'negative');\n");
        writer.write("    if (totalUnrealizedNok > 0) holdingsUnrealizedPctWrap.classList.add('positive');\n");
        writer.write("    else if (totalUnrealizedNok < 0) holdingsUnrealizedPctWrap.classList.add('negative');\n");
        writer.write("  }\n");
        writer.write("  if (holdingsUnrealizedValue) {\n");
        writer.write("    holdingsUnrealizedValue.classList.remove('positive', 'negative');\n");
        writer.write("    if (totalUnrealizedNok > 0) holdingsUnrealizedValue.classList.add('positive');\n");
        writer.write("    else if (totalUnrealizedNok < 0) holdingsUnrealizedValue.classList.add('negative');\n");
        writer.write("  }\n");
        writer.write("  var holdingsRealizedPct = document.getElementById('holdings-total-realized-pct');\n");
        writer.write("  var holdingsRealizedPctWrap = document.getElementById('holdings-total-realized-pct-wrap');\n");
        writer.write("  var holdingsRealizedValue = document.getElementById('holdings-total-realized-value');\n");
        writer.write("  if (holdingsRealizedPct) {\n");
        writer.write("    holdingsRealizedPct.textContent = formatPercentValue(realizedPct, 2);\n");
        writer.write("    holdingsRealizedPct.classList.remove('positive', 'negative');\n");
        writer.write("    if (totalRealizedNok > 0) holdingsRealizedPct.classList.add('positive');\n");
        writer.write("    else if (totalRealizedNok < 0) holdingsRealizedPct.classList.add('negative');\n");
        writer.write("  }\n");
        writer.write("  if (holdingsRealizedPctWrap) {\n");
        writer.write("    holdingsRealizedPctWrap.classList.remove('positive', 'negative');\n");
        writer.write("    if (totalRealizedNok > 0) holdingsRealizedPctWrap.classList.add('positive');\n");
        writer.write("    else if (totalRealizedNok < 0) holdingsRealizedPctWrap.classList.add('negative');\n");
        writer.write("  }\n");
        writer.write("  if (holdingsRealizedValue) {\n");
        writer.write("    holdingsRealizedValue.classList.remove('positive', 'negative');\n");
        writer.write("    if (totalRealizedNok > 0) holdingsRealizedValue.classList.add('positive');\n");
        writer.write("    else if (totalRealizedNok < 0) holdingsRealizedValue.classList.add('negative');\n");
        writer.write("  }\n");
        writer.write("  var holdingsTotalReturnPct = document.getElementById('holdings-total-total-return-pct');\n");
        writer.write("  var holdingsTotalReturnPctWrap = document.getElementById('holdings-total-total-return-pct-wrap');\n");
        writer.write("  var holdingsTotalReturnValue = document.getElementById('holdings-total-total-return-value');\n");
        writer.write("  if (holdingsTotalReturnPct) {\n");
        writer.write("    holdingsTotalReturnPct.textContent = formatPercentValue(totalReturnPct, 2);\n");
        writer.write("    holdingsTotalReturnPct.classList.remove('positive', 'negative');\n");
        writer.write("    if (totalReturnNok > 0) holdingsTotalReturnPct.classList.add('positive');\n");
        writer.write("    else if (totalReturnNok < 0) holdingsTotalReturnPct.classList.add('negative');\n");
        writer.write("  }\n");
        writer.write("  if (holdingsTotalReturnPctWrap) {\n");
        writer.write("    holdingsTotalReturnPctWrap.classList.remove('positive', 'negative');\n");
        writer.write("    if (totalReturnNok > 0) holdingsTotalReturnPctWrap.classList.add('positive');\n");
        writer.write("    else if (totalReturnNok < 0) holdingsTotalReturnPctWrap.classList.add('negative');\n");
        writer.write("  }\n");
        writer.write("  if (holdingsTotalReturnValue) {\n");
        writer.write("    holdingsTotalReturnValue.classList.remove('positive', 'negative');\n");
        writer.write("    if (totalReturnNok > 0) holdingsTotalReturnValue.classList.add('positive');\n");
        writer.write("    else if (totalReturnNok < 0) holdingsTotalReturnValue.classList.add('negative');\n");
        writer.write("  }\n");
        writer.write("  var heroTotalReturn = document.getElementById('hero-total-return-value');\n");
        writer.write("  var heroTotalReturnPct = document.getElementById('hero-total-return-pct');\n");
        writer.write("  if (heroTotalReturn) {\n");
        writer.write("    var soldOnlyReturnBuckets = parseBucketsJson(heroTotalReturn.getAttribute('data-sold-only-return-buckets'));\n");
        writer.write("    var soldOnlyHistoricalBuckets = parseBucketsJson(heroTotalReturn.getAttribute('data-sold-only-historical-buckets'));\n");
        writer.write("    var fullReturnBuckets = mergeBuckets(soldOnlyReturnBuckets, activeTotalReturnBuckets);\n");
        writer.write("    var fullHistoricalBuckets = mergeBuckets(soldOnlyHistoricalBuckets, totalHistoricalBuckets);\n");
        writer.write("    heroTotalReturn.setAttribute('data-buckets', JSON.stringify(fullReturnBuckets));\n");
        writer.write("    var fullReturnNok = Number(convertBucketsToCurrency(fullReturnBuckets, 'NOK') || 0);\n");
        writer.write("    var fullHistoricalNok = Number(convertBucketsToCurrency(fullHistoricalBuckets, 'NOK') || 0);\n");
        writer.write("    var fullReturnPct = fullHistoricalNok > 0 ? (fullReturnNok / fullHistoricalNok) * 100 : 0;\n");
        writer.write("    heroTotalReturn.classList.remove('positive', 'negative');\n");
        writer.write("    if (fullReturnNok > 0) heroTotalReturn.classList.add('positive');\n");
        writer.write("    else if (fullReturnNok < 0) heroTotalReturn.classList.add('negative');\n");
        writer.write("    if (heroTotalReturnPct) {\n");
        writer.write("      heroTotalReturnPct.textContent = formatPercentValue(fullReturnPct, 2);\n");
        writer.write("      heroTotalReturnPct.classList.remove('positive', 'negative');\n");
        writer.write("      if (fullReturnNok > 0) heroTotalReturnPct.classList.add('positive');\n");
        writer.write("      else if (fullReturnNok < 0) heroTotalReturnPct.classList.add('negative');\n");
        writer.write("    }\n");
        writer.write("  }\n");
        writer.write("  var heroDividends = document.getElementById('hero-dividends-value');\n");
        writer.write("  if (heroDividends) {\n");
        writer.write("    var soldOnlyDividendsBuckets = parseBucketsJson(heroDividends.getAttribute('data-sold-only-dividends-buckets'));\n");
        writer.write("    var fullDividendsBuckets = mergeBuckets(soldOnlyDividendsBuckets, totalDividendsBuckets);\n");
        writer.write("    heroDividends.setAttribute('data-buckets', JSON.stringify(fullDividendsBuckets));\n");
        writer.write("  }\n");
        writer.write("  var heroDayChangeValue = document.getElementById('hero-day-change-value');\n");
        writer.write("  var heroDayChangePct = document.getElementById('hero-day-change-pct');\n");
        writer.write("  if (heroDayChangeValue) {\n");
        writer.write("    heroDayChangeValue.classList.remove('positive', 'negative');\n");
        writer.write("    if (dayChangeNok > 0) heroDayChangeValue.classList.add('positive');\n");
        writer.write("    else if (dayChangeNok < 0) heroDayChangeValue.classList.add('negative');\n");
        writer.write("  }\n");
        writer.write("  if (heroDayChangePct) {\n");
        writer.write("    heroDayChangePct.textContent = formatPercentValue(dayChangePct, 2);\n");
        writer.write("    heroDayChangePct.classList.remove('positive', 'negative');\n");
        writer.write("    if (dayChangeNok > 0) heroDayChangePct.classList.add('positive');\n");
        writer.write("    else if (dayChangeNok < 0) heroDayChangePct.classList.add('negative');\n");
        writer.write("  }\n");
        writer.write("  var heroUnrealizedValue = document.getElementById('hero-unrealized-value');\n");
        writer.write("  var heroUnrealizedPct = document.getElementById('hero-unrealized-pct');\n");
        writer.write("  if (heroUnrealizedValue) {\n");
        writer.write("    heroUnrealizedValue.classList.remove('positive', 'negative');\n");
        writer.write("    if (totalUnrealizedNok > 0) heroUnrealizedValue.classList.add('positive');\n");
        writer.write("    else if (totalUnrealizedNok < 0) heroUnrealizedValue.classList.add('negative');\n");
        writer.write("  }\n");
        writer.write("  if (heroUnrealizedPct) {\n");
        writer.write("    heroUnrealizedPct.textContent = formatPercentValue(unrealizedPct, 2);\n");
        writer.write("    heroUnrealizedPct.classList.remove('positive', 'negative');\n");
        writer.write("    if (totalUnrealizedNok > 0) heroUnrealizedPct.classList.add('positive');\n");
        writer.write("    else if (totalUnrealizedNok < 0) heroUnrealizedPct.classList.add('negative');\n");
        writer.write("  }\n");
        writer.write("  var heroRealizedValue = document.getElementById('hero-realized-value');\n");
        writer.write("  var heroRealizedPct = document.getElementById('hero-realized-pct');\n");
        writer.write("  if (heroRealizedValue) {\n");
        writer.write("    heroRealizedValue.classList.remove('positive', 'negative');\n");
        writer.write("    if (totalRealizedNok > 0) heroRealizedValue.classList.add('positive');\n");
        writer.write("    else if (totalRealizedNok < 0) heroRealizedValue.classList.add('negative');\n");
        writer.write("  }\n");
        writer.write("  if (heroRealizedPct) {\n");
        writer.write("    heroRealizedPct.textContent = formatPercentValue(realizedPct, 2);\n");
        writer.write("    heroRealizedPct.classList.remove('positive', 'negative');\n");
        writer.write("    if (totalRealizedNok > 0) heroRealizedPct.classList.add('positive');\n");
        writer.write("    else if (totalRealizedNok < 0) heroRealizedPct.classList.add('negative');\n");
        writer.write("  }\n");
        writer.write("  refreshPortfolioValueBuckets();\n");
        writer.write("  var activeCurrency = getActiveReportCurrency();\n");
        writer.write("  refreshReportTotalsCurrency(activeCurrency);\n");
        writer.write("  refreshReportChartsCurrency(activeCurrency);\n");
        writer.write("}\n");
        writer.write("function formatReportRefreshTimestamp(dateValue) {\n");
        writer.write("  var date = dateValue instanceof Date ? dateValue : new Date();\n");
        writer.write("  if (!date || Number.isNaN(date.getTime())) return '-';\n");
        writer.write("  return date.toLocaleDateString() + ' ' + date.toLocaleTimeString();\n");
        writer.write("}\n");
        writer.write("function updateReportDateChip(dateValue) {\n");
        writer.write("  var node = document.getElementById('report-date-value');\n");
        writer.write("  if (!node) return;\n");
        writer.write("  node.textContent = formatReportRefreshTimestamp(dateValue);\n");
        writer.write("}\n");
        writer.write("function initPriceRefreshButton() {\n");
        writer.write("  var button = document.getElementById('refresh-prices-btn');\n");
        writer.write("  var status = document.getElementById('refresh-prices-status');\n");
        writer.write("  if (!button) return;\n");
        writer.write("  button.addEventListener('click', async function() {\n");
        writer.write("    var rows = Array.prototype.slice.call(document.querySelectorAll('tr[data-overview-row=\"1\"]'));\n");
        writer.write("    var tickers = Array.from(new Set(rows.map(function(row) {\n");
        writer.write("      return String(row.getAttribute('data-ticker') || '').trim().toUpperCase();\n");
        writer.write("    }).filter(function(value) { return value.length > 0; })));\n");
        writer.write("    if (!tickers.length) {\n");
        writer.write("      if (status) status.textContent = 'No holdings available for update.';\n");
        writer.write("      return;\n");
        writer.write("    }\n");
        writer.write("    button.disabled = true;\n");
        writer.write("    if (status) status.textContent = 'Updating portfolio data...';\n");
        writer.write("    try {\n");
        writer.write("      var prices = await fetchLatestPrices(tickers);\n");
        writer.write("      var updatedRows = 0;\n");
        writer.write("      rows.forEach(function(row) {\n");
        writer.write("        var ticker = String(row.getAttribute('data-ticker') || '').trim().toUpperCase();\n");
        writer.write("        var nextPrice = Number(prices[ticker] || 0);\n");
        writer.write("        if (applyOverviewRowPrice(row, nextPrice)) {\n");
        writer.write("          updatedRows += 1;\n");
        writer.write("        }\n");
        writer.write("      });\n");
        writer.write("      recalculateOverviewAndHeaderTotalsAfterPriceRefresh();\n");
        writer.write("      initOverviewDayCharts();\n");
        writer.write("      var refreshedAt = new Date();\n");
        writer.write("      updateReportDateChip(refreshedAt);\n");
        writer.write("      if (status) status.textContent = 'Updated portfolio for ' + updatedRows + ' holdings at ' + formatReportRefreshTimestamp(refreshedAt) + '.';\n");
        writer.write("    } catch (error) {\n");
        writer.write("      if (status) status.textContent = 'Could not update portfolio: ' + (error && error.message ? error.message : 'Unknown error');\n");
        writer.write("    } finally {\n");
        writer.write("      button.disabled = false;\n");
        writer.write("    }\n");
        writer.write("  });\n");
        writer.write("}\n");
        writer.write("function parseSortableNumber(value) {\n");
        writer.write("  var text = String(value || '').trim();\n");
        writer.write("  if (!text || text === '-') return Number.NaN;\n");
        writer.write("  var normalized = text.replace(/\u00A0/g, ' ').replace(/,/g, '.');\n");
        writer.write("  var match = normalized.match(/-?\\d[\\d\\s]*(?:\\.\\d+)?/);\n");
        writer.write("  if (!match) return Number.NaN;\n");
        writer.write("  return Number(match[0].replace(/\s+/g, ''));\n");
        writer.write("}\n");
        writer.write("function parseSortableDate(value) {\n");
        writer.write("  var text = String(value || '').trim();\n");
        writer.write("  if (!text || text === '-') return Number.NaN;\n");
        writer.write("  if (/^\\d{2}\\.\\d{2}\\.\\d{4}$/.test(text)) {\n");
        writer.write("    var parts = text.split('.');\n");
        writer.write("    var day = Number(parts[0]);\n");
        writer.write("    var month = Number(parts[1]);\n");
        writer.write("    var year = Number(parts[2]);\n");
        writer.write("    var date = new Date(year, month - 1, day);\n");
        writer.write("    var timestamp = date.getTime();\n");
        writer.write("    return Number.isFinite(timestamp) ? timestamp : Number.NaN;\n");
        writer.write("  }\n");
        writer.write("  var parsed = Date.parse(text);\n");
        writer.write("  return Number.isFinite(parsed) ? parsed : Number.NaN;\n");
        writer.write("}\n");
        writer.write("function detectSortMode(headerLabel) {\n");
        writer.write("  var text = String(headerLabel || '').trim().toLowerCase();\n");
        writer.write("  if (!text || text.indexOf('details') === 0) return 'none';\n");
        writer.write("  if (text.indexOf('ticker') >= 0 || text.indexOf('security') >= 0 || text.indexOf('type') >= 0) return 'text';\n");
        writer.write("  if (text.indexOf('date') >= 0) return 'date';\n");
        writer.write("  return 'number';\n");
        writer.write("}\n");
        writer.write("function compareValues(a, b, direction) {\n");
        writer.write("  if (typeof a === 'string' || typeof b === 'string') {\n");
        writer.write("    var left = String(a || '');\n");
        writer.write("    var right = String(b || '');\n");
        writer.write("    return direction === 'asc'\n");
        writer.write("      ? left.localeCompare(right, undefined, { sensitivity: 'base' })\n");
        writer.write("      : right.localeCompare(left, undefined, { sensitivity: 'base' });\n");
        writer.write("  }\n");
        writer.write("  var leftNumber = Number(a);\n");
        writer.write("  var rightNumber = Number(b);\n");
        writer.write("  if (!Number.isFinite(leftNumber) && !Number.isFinite(rightNumber)) return 0;\n");
        writer.write("  if (!Number.isFinite(leftNumber)) return 1;\n");
        writer.write("  if (!Number.isFinite(rightNumber)) return -1;\n");
        writer.write("  return direction === 'asc' ? leftNumber - rightNumber : rightNumber - leftNumber;\n");
        writer.write("}\n");
        writer.write("function extractSortableValue(cell, mode) {\n");
        writer.write("  if (!cell) return mode === 'text' ? '' : Number.NaN;\n");
        writer.write("  var text = String(cell.textContent || '').trim();\n");
        writer.write("  if (mode === 'text') return text.toLowerCase();\n");
        writer.write("  if (mode === 'date') return parseSortableDate(text);\n");
        writer.write("  return parseSortableNumber(text);\n");
        writer.write("}\n");
        writer.write("function applyHeaderSortState(headers, activeHeader, direction) {\n");
        writer.write("  headers.forEach(function(header) {\n");
        writer.write("    header.classList.remove('sort-asc', 'sort-desc');\n");
        writer.write("    if (header !== activeHeader) {\n");
        writer.write("      header.dataset.sortDirection = '';\n");
        writer.write("    }\n");
        writer.write("  });\n");
        writer.write("  activeHeader.classList.add(direction === 'asc' ? 'sort-asc' : 'sort-desc');\n");
        writer.write("  activeHeader.dataset.sortDirection = direction;\n");
        writer.write("}\n");
        writer.write("function getDirectTableRows(table) {\n");
        writer.write("  if (!table) return [];\n");
        writer.write("  var body = table.tBodies && table.tBodies.length ? table.tBodies[0] : table;\n");
        writer.write("  if (table.tBodies && table.tBodies.length > 1) {\n");
        writer.write("    for (var i = 1; i < table.tBodies.length; i += 1) {\n");
        writer.write("      var extraBody = table.tBodies[i];\n");
        writer.write("      while (extraBody && extraBody.firstElementChild) {\n");
        writer.write("        body.appendChild(extraBody.firstElementChild);\n");
        writer.write("      }\n");
        writer.write("    }\n");
        writer.write("  }\n");
        writer.write("  return Array.prototype.slice.call(body.children).filter(function(node) {\n");
        writer.write("    return node && node.tagName === 'TR';\n");
        writer.write("  });\n");
        writer.write("}\n");
        writer.write("function sortPrimaryTable(table, columnIndex, mode, direction) {\n");
        writer.write("  if (!table || mode === 'none') return;\n");
        writer.write("  var body = table.tBodies && table.tBodies.length ? table.tBodies[0] : table;\n");
        writer.write("  var allRows = getDirectTableRows(table);\n");
        writer.write("  if (allRows.length <= 2) return;\n");
        writer.write("  var groups = [];\n");
        writer.write("  for (var i = 1; i < allRows.length; i += 1) {\n");
        writer.write("    var row = allRows[i];\n");
        writer.write("    if (row.classList.contains('details-row') || row.classList.contains('total-row')) continue;\n");
        writer.write("    var assetGroup = String(row.getAttribute('data-asset-group') || 'STOCK').toUpperCase() === 'FUND' ? 'FUND' : 'STOCK';\n");
        writer.write("    var details = row.nextElementSibling;\n");
        writer.write("    if (details && details.classList && details.classList.contains('details-row')) {\n");
        writer.write("      groups.push({ row: row, details: details, value: extractSortableValue(row.cells[columnIndex], mode), assetGroup: assetGroup });\n");
        writer.write("    } else {\n");
        writer.write("      groups.push({ row: row, details: null, value: extractSortableValue(row.cells[columnIndex], mode), assetGroup: assetGroup });\n");
        writer.write("    }\n");
        writer.write("  }\n");
        writer.write("  var byAssetGroup = { STOCK: [], FUND: [] };\n");
        writer.write("  groups.forEach(function(group) {\n");
        writer.write("    byAssetGroup[group.assetGroup].push(group);\n");
        writer.write("  });\n");
        writer.write("  byAssetGroup.STOCK.sort(function(a, b) { return compareValues(a.value, b.value, direction); });\n");
        writer.write("  byAssetGroup.FUND.sort(function(a, b) { return compareValues(a.value, b.value, direction); });\n");
        writer.write("  var orderedGroups = byAssetGroup.STOCK.concat(byAssetGroup.FUND);\n");
        writer.write("  var totalRow = allRows.find(function(row) { return row.classList && row.classList.contains('total-row'); }) || null;\n");
        writer.write("  var previousGroup = '';\n");
        writer.write("  orderedGroups.forEach(function(group, index) {\n");
        writer.write("    group.row.classList.remove('asset-split');\n");
        writer.write("    if (index > 0 && group.assetGroup !== previousGroup) {\n");
        writer.write("      group.row.classList.add('asset-split');\n");
        writer.write("    }\n");
        writer.write("    previousGroup = group.assetGroup;\n");
        writer.write("    body.appendChild(group.row);\n");
        writer.write("    if (group.details) body.appendChild(group.details);\n");
        writer.write("  });\n");
        writer.write("  if (totalRow) body.appendChild(totalRow);\n");
        writer.write("}\n");
        writer.write("function sortSimpleTable(table, columnIndex, mode, direction) {\n");
        writer.write("  if (!table || mode === 'none') return;\n");
        writer.write("  var body = table.tBodies && table.tBodies.length ? table.tBodies[0] : table;\n");
        writer.write("  var allRows = getDirectTableRows(table);\n");
        writer.write("  if (allRows.length <= 2) return;\n");
        writer.write("  var dataRows = allRows.slice(1).filter(function(row) { return !row.classList.contains('total-row'); });\n");
        writer.write("  dataRows.sort(function(a, b) {\n");
        writer.write("    var aValue = extractSortableValue(a.cells[columnIndex], mode);\n");
        writer.write("    var bValue = extractSortableValue(b.cells[columnIndex], mode);\n");
        writer.write("    return compareValues(aValue, bValue, direction);\n");
        writer.write("  });\n");
        writer.write("  var totalRow = allRows.find(function(row) { return row.classList && row.classList.contains('total-row'); }) || null;\n");
        writer.write("  dataRows.forEach(function(row) { body.appendChild(row); });\n");
        writer.write("  if (totalRow) body.appendChild(totalRow);\n");
        writer.write("}\n");
        writer.write("function initSortableTables() {\n");
        writer.write("  function wireTable(table, sorter) {\n");
        writer.write("    if (!table || table.dataset.sortableReady === '1') return;\n");
        writer.write("    var headerRow = getDirectTableRows(table)[0];\n");
        writer.write("    if (!headerRow) return;\n");
        writer.write("    var headers = Array.prototype.slice.call(headerRow.querySelectorAll('th'));\n");
        writer.write("    headers.forEach(function(header, index) {\n");
        writer.write("      var mode = detectSortMode(header.textContent);\n");
        writer.write("      if (mode === 'none') return;\n");
        writer.write("      header.classList.add('sortable-header');\n");
        writer.write("      header.addEventListener('click', function() {\n");
        writer.write("        var nextDirection = header.dataset.sortDirection === 'asc' ? 'desc' : 'asc';\n");
        writer.write("        applyHeaderSortState(headers, header, nextDirection);\n");
        writer.write("        sorter(table, index, mode, nextDirection);\n");
        writer.write("      });\n");
        writer.write("    });\n");
        writer.write("    table.dataset.sortableReady = '1';\n");
        writer.write("  }\n");
        writer.write("  document.querySelectorAll('.overview-table, .realized-table').forEach(function(table) {\n");
        writer.write("    wireTable(table, sortPrimaryTable);\n");
        writer.write("  });\n");
        writer.write("  document.querySelectorAll('.details-table').forEach(function(table) {\n");
        writer.write("    wireTable(table, sortSimpleTable);\n");
        writer.write("  });\n");
        writer.write("}\n");
        writer.write("function initOverviewModeSwitcher() {\n");
        writer.write("  var buttons = Array.prototype.slice.call(document.querySelectorAll('.overview-mode-btn[data-overview-mode]'));\n");
        writer.write("  var panels = Array.prototype.slice.call(document.querySelectorAll('.js-overview-mode-panel[data-overview-mode-panel]'));\n");
        writer.write("  var detailsToggleBtn = document.getElementById('overview-details-toggle');\n");
        writer.write("  if (!buttons.length || !panels.length) return;\n");
        writer.write("  function activate(mode) {\n");
        writer.write("    buttons.forEach(function(btn) {\n");
        writer.write("      var active = btn.getAttribute('data-overview-mode') === mode;\n");
        writer.write("      btn.classList.toggle('is-active', active);\n");
        writer.write("      btn.setAttribute('aria-selected', active ? 'true' : 'false');\n");
        writer.write("    });\n");
        writer.write("    panels.forEach(function(panel) {\n");
        writer.write("      var visible = panel.getAttribute('data-overview-mode-panel') === mode;\n");
        writer.write("      panel.hidden = !visible;\n");
        writer.write("    });\n");
        writer.write("    if (detailsToggleBtn) {\n");
        writer.write("      var groupName = mode === 'summary' ? 'overview-details' : (mode === 'holdings' ? 'holdings-details' : '');\n");
        writer.write("      var baseLabel = detailsToggleBtn.getAttribute('data-detail-label') || 'Open all details';\n");
        writer.write("      detailsToggleBtn.textContent = baseLabel + ' ▸';\n");
        writer.write("      detailsToggleBtn.setAttribute('data-detail-group', groupName);\n");
        writer.write("      detailsToggleBtn.disabled = !groupName;\n");
        writer.write("      detailsToggleBtn.setAttribute('aria-disabled', groupName ? 'false' : 'true');\n");
        writer.write("    }\n");
        writer.write("  }\n");
        writer.write("  if (detailsToggleBtn && detailsToggleBtn.dataset.bound !== '1') {\n");
        writer.write("    detailsToggleBtn.dataset.bound = '1';\n");
        writer.write("    detailsToggleBtn.addEventListener('click', function() {\n");
        writer.write("      var groupName = detailsToggleBtn.getAttribute('data-detail-group') || '';\n");
        writer.write("      if (!groupName) return;\n");
        writer.write("      toggleDetailGroup(groupName, detailsToggleBtn);\n");
        writer.write("    });\n");
        writer.write("  }\n");
        writer.write("  buttons.forEach(function(btn) {\n");
        writer.write("    btn.addEventListener('click', function() {\n");
        writer.write("      activate(btn.getAttribute('data-overview-mode') || 'summary');\n");
        writer.write("    });\n");
        writer.write("  });\n");
        writer.write("  activate('summary');\n");
        writer.write("}\n");
        writer.write("function resolveCashHoldingsBridge() {\n");
        writer.write("  try {\n");
        writer.write("    if (window.parent && window.parent !== window && window.parent.__portfolioCashHoldingsBridge) {\n");
        writer.write("      return window.parent.__portfolioCashHoldingsBridge;\n");
        writer.write("    }\n");
        writer.write("  } catch (_) {\n");
        writer.write("    // Parent bridge is optional.\n");
        writer.write("  }\n");
        writer.write("  return null;\n");
        writer.write("}\n");
        writer.write("function detectLoggedInFromParentContext() {\n");
        writer.write("  try {\n");
        writer.write("    if (!window.parent || window.parent === window || !window.parent.document) return false;\n");
        writer.write("    var logoutButton = window.parent.document.getElementById('logout-btn');\n");
        writer.write("    return !!(logoutButton && logoutButton.hidden === false);\n");
        writer.write("  } catch (_) {\n");
        writer.write("    return false;\n");
        writer.write("  }\n");
        writer.write("}\n");
        writer.write("function resolveCashHoldingsStorageKey() {\n");
        writer.write("  var identity = 'anonymous';\n");
        writer.write("  try {\n");
        writer.write("    if (window.parent && window.parent !== window && window.parent.document) {\n");
        writer.write("      var accountLabel = window.parent.document.getElementById('account-email');\n");
        writer.write("      if (accountLabel && accountLabel.textContent) {\n");
        writer.write("        identity = String(accountLabel.textContent).replace(/^\\s*User:\\s*/i, '').trim() || identity;\n");
        writer.write("      }\n");
        writer.write("    }\n");
        writer.write("  } catch (_) {\n");
        writer.write("    // Keep anonymous identity.\n");
        writer.write("  }\n");
        writer.write("  return 'portfolio-manual-cash-holdings::' + identity;\n");
        writer.write("}\n");
        writer.write("function sanitizeManualCashState(raw) {\n");
        writer.write("  var state = raw && typeof raw === 'object' ? raw : {};\n");
        writer.write("  var accounts = Array.isArray(state.accounts) ? state.accounts : [];\n");
        writer.write("  return {\n");
        writer.write("    accounts: accounts.map(function(account) {\n");
        writer.write("      var transactions = Array.isArray(account && account.transactions) ? account.transactions : [];\n");
        writer.write("      return {\n");
        writer.write("        id: String(account && account.id || 'acc-' + Math.random().toString(36).slice(2, 10)),\n");
        writer.write("        name: String(account && account.name || '').trim(),\n");
        writer.write("        transactions: transactions.map(function(tx) {\n");
        writer.write("          return {\n");
        writer.write("            id: String(tx && tx.id || 'tx-' + Math.random().toString(36).slice(2, 10)),\n");
        writer.write("            amount: Number(tx && tx.amount || 0),\n");
        writer.write("            currency: normalizeCurrencyCodeInput(tx && tx.currency || 'NOK') || 'NOK'\n");
        writer.write("          };\n");
        writer.write("        }).filter(function(tx) { return Number.isFinite(tx.amount); })\n");
        writer.write("      };\n");
        writer.write("    }).filter(function(account) { return account.name.length > 0; })\n");
        writer.write("  };\n");
        writer.write("}\n");
        writer.write("async function loadManualCashState() {\n");
        writer.write("  var bridge = resolveCashHoldingsBridge();\n");
        writer.write("  if (bridge && typeof bridge.loadManualCashHoldings === 'function') {\n");
        writer.write("    try {\n");
        writer.write("      return sanitizeManualCashState(await bridge.loadManualCashHoldings());\n");
        writer.write("    } catch (_) {\n");
        writer.write("      // Fall back to local storage below.\n");
        writer.write("    }\n");
        writer.write("  }\n");
        writer.write("  try {\n");
        writer.write("    var raw = window.localStorage.getItem(resolveCashHoldingsStorageKey());\n");
        writer.write("    return sanitizeManualCashState(raw ? JSON.parse(raw) : {});\n");
        writer.write("  } catch (_) {\n");
        writer.write("    return sanitizeManualCashState({});\n");
        writer.write("  }\n");
        writer.write("}\n");
        writer.write("async function saveManualCashState(state) {\n");
        writer.write("  var safeState = sanitizeManualCashState(state);\n");
        writer.write("  var bridge = resolveCashHoldingsBridge();\n");
        writer.write("  if (bridge && typeof bridge.saveManualCashHoldings === 'function') {\n");
        writer.write("    try {\n");
        writer.write("      await bridge.saveManualCashHoldings(safeState);\n");
        writer.write("      return safeState;\n");
        writer.write("    } catch (_) {\n");
        writer.write("      // Fall back to local storage below.\n");
        writer.write("    }\n");
        writer.write("  }\n");
        writer.write("  try {\n");
        writer.write("    window.localStorage.setItem(resolveCashHoldingsStorageKey(), JSON.stringify(safeState));\n");
        writer.write("  } catch (_) {\n");
        writer.write("    // Ignore local persistence failures.\n");
        writer.write("  }\n");
        writer.write("  return safeState;\n");
        writer.write("}\n");
        writer.write("function summarizeManualCashAccount(account) {\n");
        writer.write("  var totals = {};\n");
        writer.write("  (account.transactions || []).forEach(function(tx) {\n");
        writer.write("    var currency = normalizeCurrencyCodeInput(tx.currency || 'NOK') || 'NOK';\n");
        writer.write("    totals[currency] = Number(totals[currency] || 0) + Number(tx.amount || 0);\n");
        writer.write("  });\n");
        writer.write("  var parts = Object.keys(totals).sort().map(function(currency) {\n");
        writer.write("    return formatMoneyValue(totals[currency], currency, 0);\n");
        writer.write("  });\n");
        writer.write("  var accountName = String(account && account.name || '').trim() || 'Unnamed account';\n");
        writer.write("  return accountName + ': ' + (parts.length ? parts.join(' + ') : '0 NOK');\n");
        writer.write("}\n");
        writer.write("function formatCashBucketsSummary(buckets) {\n");
        writer.write("  var parts = Object.keys(buckets || {}).sort().map(function(currency) {\n");
        writer.write("    return formatMoneyValue(Number(buckets[currency] || 0), currency, 0);\n");
        writer.write("  });\n");
        writer.write("  return parts.length ? parts.join(' + ') : '0 NOK';\n");
        writer.write("}\n");
        writer.write("function resolvePortfolioBaseCashBuckets() {\n");
        writer.write("  var totalNode = document.getElementById('cash-holdings-total');\n");
        writer.write("  if (!totalNode) return {};\n");
        writer.write("  var raw = totalNode.getAttribute('data-base-buckets') || totalNode.getAttribute('data-buckets') || '{}';\n");
        writer.write("  return parseBucketsJson(raw);\n");
        writer.write("}\n");
        writer.write("function renderManualCashSummaryLines(state) {\n");
        writer.write("  var list = document.getElementById('manual-cash-holdings-list');\n");
        writer.write("  if (!list) return;\n");
        writer.write("  list.innerHTML = '';\n");
        writer.write("  var portfolioLine = document.createElement('div');\n");
        writer.write("  portfolioLine.className = 'manual-cash-holding-line is-portfolio';\n");
        writer.write("  portfolioLine.textContent = 'Portfolio: ' + formatCashBucketsSummary(resolvePortfolioBaseCashBuckets());\n");
        writer.write("  list.appendChild(portfolioLine);\n");
        writer.write("  var accounts = state.accounts || [];\n");
        writer.write("  if (!accounts.length) {\n");
        writer.write("    return;\n");
        writer.write("  }\n");
        writer.write("  accounts.forEach(function(account) {\n");
        writer.write("    var line = document.createElement('div');\n");
        writer.write("    line.className = 'manual-cash-holding-line';\n");
        writer.write("    line.textContent = summarizeManualCashAccount(account);\n");
        writer.write("    list.appendChild(line);\n");
        writer.write("  });\n");
        writer.write("}\n");
        writer.write("function buildManualCashBuckets(state) {\n");
        writer.write("  var buckets = {};\n");
        writer.write("  (state.accounts || []).forEach(function(account) {\n");
        writer.write("    (account.transactions || []).forEach(function(tx) {\n");
        writer.write("      addBucketValue(buckets, tx.currency || 'NOK', Number(tx.amount || 0));\n");
        writer.write("    });\n");
        writer.write("  });\n");
        writer.write("  return buckets;\n");
        writer.write("}\n");
        writer.write("function refreshCashHoldingsTotalFromState(state) {\n");
        writer.write("  var totalNode = document.getElementById('cash-holdings-total');\n");
        writer.write("  if (!totalNode) return;\n");
        writer.write("  var baseRaw = totalNode.getAttribute('data-base-buckets');\n");
        writer.write("  if (!baseRaw) {\n");
        writer.write("    baseRaw = totalNode.getAttribute('data-buckets') || '{}';\n");
        writer.write("    totalNode.setAttribute('data-base-buckets', baseRaw);\n");
        writer.write("  }\n");
        writer.write("  var baseBuckets = parseBucketsJson(baseRaw);\n");
        writer.write("  var manualBuckets = buildManualCashBuckets(state);\n");
        writer.write("  totalNode.setAttribute('data-buckets', JSON.stringify(mergeBuckets(baseBuckets, manualBuckets)));\n");
        writer.write("  refreshPortfolioValueBuckets();\n");
        writer.write("  refreshReportTotalsCurrency(getActiveReportCurrency());\n");
        writer.write("}\n");
        writer.write("function parseManualCashAmountInput(value) {\n");
        writer.write("  var normalized = String(value == null ? '' : value).replace(/\\s+/g, '').replace(',', '.');\n");
        writer.write("  if (!normalized) return Number.NaN;\n");
        writer.write("  var amount = Number(normalized);\n");
        writer.write("  return Number.isFinite(amount) ? amount : Number.NaN;\n");
        writer.write("}\n");
        writer.write("function setManualCashManagerMessage(ui, text, isError) {\n");
        writer.write("  if (!ui || !ui.message) return;\n");
        writer.write("  ui.message.textContent = String(text || '');\n");
        writer.write("  ui.message.classList.toggle('is-error', !!isError);\n");
        writer.write("}\n");
        writer.write("function ensureManualCashManagerUi() {\n");
        writer.write("  if (window.__manualCashManagerUi) return window.__manualCashManagerUi;\n");
        writer.write("  var overlay = document.createElement('div');\n");
        writer.write("  overlay.className = 'cash-manager-overlay';\n");
        writer.write("  overlay.hidden = true;\n");
        writer.write("  overlay.innerHTML = '<div class=\"cash-manager-dialog\" role=\"dialog\" aria-modal=\"true\" aria-label=\"Manage cash holdings\"><div class=\"cash-manager-header\"><h4>Manage Cash Holdings</h4><button type=\"button\" class=\"cash-manager-close\" aria-label=\"Close\">×</button></div><div class=\"cash-manager-form-row\"><input type=\"text\" class=\"cash-manager-account-name\" placeholder=\"New account name\"><button type=\"button\" class=\"cash-manager-btn cash-manager-create-account\">Create account</button></div><div class=\"cash-manager-form-row\"><select class=\"cash-manager-account-select\" aria-label=\"Select account\"></select><input type=\"text\" class=\"cash-manager-amount-input\" placeholder=\"Amount (+/-)\" inputmode=\"decimal\"><input type=\"text\" class=\"cash-manager-currency-input\" value=\"NOK\" maxlength=\"5\" aria-label=\"Currency\"><button type=\"button\" class=\"cash-manager-btn cash-manager-add-transaction\">Add transaction</button></div><div class=\"cash-manager-message\"></div><div class=\"cash-manager-accounts\"></div></div>';\n");
        writer.write("  document.body.appendChild(overlay);\n");
        writer.write("  var closeButton = overlay.querySelector('.cash-manager-close');\n");
        writer.write("  var accountNameInput = overlay.querySelector('.cash-manager-account-name');\n");
        writer.write("  var accountSelect = overlay.querySelector('.cash-manager-account-select');\n");
        writer.write("  var amountInput = overlay.querySelector('.cash-manager-amount-input');\n");
        writer.write("  var currencyInput = overlay.querySelector('.cash-manager-currency-input');\n");
        writer.write("  var createAccountButton = overlay.querySelector('.cash-manager-create-account');\n");
        writer.write("  var addTransactionButton = overlay.querySelector('.cash-manager-add-transaction');\n");
        writer.write("  var accountsContainer = overlay.querySelector('.cash-manager-accounts');\n");
        writer.write("  var message = overlay.querySelector('.cash-manager-message');\n");
        writer.write("  function close() { overlay.hidden = true; }\n");
        writer.write("  function open() { overlay.hidden = false; if (accountNameInput) accountNameInput.focus(); }\n");
        writer.write("  closeButton.addEventListener('click', close);\n");
        writer.write("  overlay.addEventListener('click', function(event) { if (event.target === overlay) close(); });\n");
        writer.write("  document.addEventListener('keydown', function(event) { if (event.key === 'Escape' && !overlay.hidden) close(); });\n");
        writer.write("  window.__manualCashManagerUi = {\n");
        writer.write("    overlay: overlay,\n");
        writer.write("    open: open,\n");
        writer.write("    message: message,\n");
        writer.write("    accountNameInput: accountNameInput,\n");
        writer.write("    accountSelect: accountSelect,\n");
        writer.write("    amountInput: amountInput,\n");
        writer.write("    currencyInput: currencyInput,\n");
        writer.write("    createAccountButton: createAccountButton,\n");
        writer.write("    addTransactionButton: addTransactionButton,\n");
        writer.write("    accountsContainer: accountsContainer\n");
        writer.write("  };\n");
        writer.write("  return window.__manualCashManagerUi;\n");
        writer.write("}\n");
        writer.write("function renderManualCashManager(state, ui, commitStateChange) {\n");
        writer.write("  if (!ui) return;\n");
        writer.write("  var accounts = state.accounts || [];\n");
        writer.write("  var previousSelection = ui.accountSelect.value;\n");
        writer.write("  ui.accountSelect.innerHTML = '';\n");
        writer.write("  if (accounts.length) {\n");
        writer.write("    accounts.forEach(function(account) {\n");
        writer.write("      var option = document.createElement('option');\n");
        writer.write("      option.value = account.id;\n");
        writer.write("      option.textContent = account.name;\n");
        writer.write("      ui.accountSelect.appendChild(option);\n");
        writer.write("    });\n");
        writer.write("    ui.accountSelect.disabled = false;\n");
        writer.write("    if (previousSelection && accounts.some(function(account) { return account.id === previousSelection; })) ui.accountSelect.value = previousSelection;\n");
        writer.write("  } else {\n");
        writer.write("    var emptyOption = document.createElement('option');\n");
        writer.write("    emptyOption.value = '';\n");
        writer.write("    emptyOption.textContent = 'No account available';\n");
        writer.write("    ui.accountSelect.appendChild(emptyOption);\n");
        writer.write("    ui.accountSelect.disabled = true;\n");
        writer.write("  }\n");
        writer.write("  ui.accountsContainer.innerHTML = '';\n");
        writer.write("  if (!accounts.length) {\n");
        writer.write("    var empty = document.createElement('p');\n");
        writer.write("    empty.className = 'cash-manager-empty';\n");
        writer.write("    empty.textContent = 'No manual cash accounts created yet.';\n");
        writer.write("    ui.accountsContainer.appendChild(empty);\n");
        writer.write("    return;\n");
        writer.write("  }\n");
        writer.write("  accounts.forEach(function(account) {\n");
        writer.write("    var block = document.createElement('section');\n");
        writer.write("    block.className = 'cash-account-block';\n");
        writer.write("    var head = document.createElement('div');\n");
        writer.write("    head.className = 'cash-account-head';\n");
        writer.write("    var title = document.createElement('p');\n");
        writer.write("    title.className = 'cash-account-title';\n");
        writer.write("    title.textContent = account.name;\n");
        writer.write("    var deleteAccountButton = document.createElement('button');\n");
        writer.write("    deleteAccountButton.type = 'button';\n");
        writer.write("    deleteAccountButton.className = 'cash-manager-btn danger';\n");
        writer.write("    deleteAccountButton.textContent = 'Delete account';\n");
        writer.write("    deleteAccountButton.addEventListener('click', function() {\n");
        writer.write("      commitStateChange(function(nextState) {\n");
        writer.write("        nextState.accounts = (nextState.accounts || []).filter(function(candidate) { return candidate.id !== account.id; });\n");
        writer.write("      }, 'Account deleted.');\n");
        writer.write("    });\n");
        writer.write("    head.appendChild(title);\n");
        writer.write("    head.appendChild(deleteAccountButton);\n");
        writer.write("    block.appendChild(head);\n");
        writer.write("    var summary = document.createElement('p');\n");
        writer.write("    summary.className = 'cash-account-summary';\n");
        writer.write("    summary.textContent = summarizeManualCashAccount(account);\n");
        writer.write("    block.appendChild(summary);\n");
        writer.write("    var txList = document.createElement('div');\n");
        writer.write("    txList.className = 'cash-transaction-list';\n");
        writer.write("    if (!account.transactions.length) {\n");
        writer.write("      var noTx = document.createElement('div');\n");
        writer.write("      noTx.className = 'cash-manager-empty';\n");
        writer.write("      noTx.textContent = 'No transactions yet.';\n");
        writer.write("      txList.appendChild(noTx);\n");
        writer.write("    } else {\n");
        writer.write("      account.transactions.slice().reverse().forEach(function(tx) {\n");
        writer.write("        var txItem = document.createElement('div');\n");
        writer.write("        txItem.className = 'cash-transaction-item';\n");
        writer.write("        var txText = document.createElement('span');\n");
        writer.write("        txText.textContent = formatMoneyValue(tx.amount, tx.currency, 0);\n");
        writer.write("        var deleteTxButton = document.createElement('button');\n");
        writer.write("        deleteTxButton.type = 'button';\n");
        writer.write("        deleteTxButton.className = 'cash-manager-btn danger';\n");
        writer.write("        deleteTxButton.textContent = 'Delete';\n");
        writer.write("        deleteTxButton.addEventListener('click', function() {\n");
        writer.write("          commitStateChange(function(nextState) {\n");
        writer.write("            var targetAccount = (nextState.accounts || []).find(function(candidate) { return candidate.id === account.id; });\n");
        writer.write("            if (!targetAccount || !Array.isArray(targetAccount.transactions)) return;\n");
        writer.write("            targetAccount.transactions = targetAccount.transactions.filter(function(candidate) { return candidate.id !== tx.id; });\n");
        writer.write("          }, 'Transaction deleted.');\n");
        writer.write("        });\n");
        writer.write("        txItem.appendChild(txText);\n");
        writer.write("        txItem.appendChild(deleteTxButton);\n");
        writer.write("        txList.appendChild(txItem);\n");
        writer.write("      });\n");
        writer.write("    }\n");
        writer.write("    block.appendChild(txList);\n");
        writer.write("    ui.accountsContainer.appendChild(block);\n");
        writer.write("  });\n");
        writer.write("}\n");
        writer.write("async function initManualCashHoldings() {\n");
        writer.write("  var addButton = document.getElementById('cash-holdings-add-btn');\n");
        writer.write("  if (!addButton) return;\n");
        writer.write("  var isLoggedIn = detectLoggedInFromParentContext();\n");
        writer.write("  if (!isLoggedIn) {\n");
        writer.write("    addButton.style.display = 'none';\n");
        writer.write("    return;\n");
        writer.write("  }\n");
        writer.write("  addButton.style.display = 'inline-flex';\n");
        writer.write("  addButton.textContent = 'Manage';\n");
        writer.write("  var managerUi = ensureManualCashManagerUi();\n");
        writer.write("  var state = await loadManualCashState();\n");
        writer.write("  renderManualCashSummaryLines(state);\n");
        writer.write("  refreshCashHoldingsTotalFromState(state);\n");
        writer.write("  async function commitStateChange(mutator, successMessage) {\n");
        writer.write("    var nextState = sanitizeManualCashState(state);\n");
        writer.write("    if (typeof mutator === 'function') mutator(nextState);\n");
        writer.write("    state = sanitizeManualCashState(await saveManualCashState(nextState));\n");
        writer.write("    renderManualCashSummaryLines(state);\n");
        writer.write("    refreshCashHoldingsTotalFromState(state);\n");
        writer.write("    renderManualCashManager(state, managerUi, commitStateChange);\n");
        writer.write("    if (successMessage) setManualCashManagerMessage(managerUi, successMessage, false);\n");
        writer.write("  }\n");
        writer.write("  renderManualCashManager(state, managerUi, commitStateChange);\n");
        writer.write("  if (!managerUi.overlay.dataset.boundEvents) {\n");
        writer.write("    managerUi.overlay.dataset.boundEvents = '1';\n");
        writer.write("    managerUi.createAccountButton.addEventListener('click', function() {\n");
        writer.write("      var name = String(managerUi.accountNameInput.value || '').trim();\n");
        writer.write("      if (!name) {\n");
        writer.write("        setManualCashManagerMessage(managerUi, 'Account name is required.', true);\n");
        writer.write("        managerUi.accountNameInput.focus();\n");
        writer.write("        return;\n");
        writer.write("      }\n");
        writer.write("      managerUi.accountNameInput.value = '';\n");
        writer.write("      commitStateChange(function(nextState) {\n");
        writer.write("        nextState.accounts.push({ id: 'acc-' + Math.random().toString(36).slice(2, 10), name: name, transactions: [] });\n");
        writer.write("      }, 'Account created.');\n");
        writer.write("    });\n");
        writer.write("    managerUi.addTransactionButton.addEventListener('click', function() {\n");
        writer.write("      var accountId = String(managerUi.accountSelect.value || '').trim();\n");
        writer.write("      if (!accountId) {\n");
        writer.write("        setManualCashManagerMessage(managerUi, 'Create an account before adding transactions.', true);\n");
        writer.write("        return;\n");
        writer.write("      }\n");
        writer.write("      var amount = parseManualCashAmountInput(managerUi.amountInput.value);\n");
        writer.write("      if (!Number.isFinite(amount) || amount === 0) {\n");
        writer.write("        setManualCashManagerMessage(managerUi, 'Amount must be a non-zero number.', true);\n");
        writer.write("        managerUi.amountInput.focus();\n");
        writer.write("        return;\n");
        writer.write("      }\n");
        writer.write("      var currency = normalizeCurrencyCodeInput(managerUi.currencyInput.value || 'NOK') || 'NOK';\n");
        writer.write("      managerUi.amountInput.value = '';\n");
        writer.write("      managerUi.currencyInput.value = currency;\n");
        writer.write("      commitStateChange(function(nextState) {\n");
        writer.write("        var account = (nextState.accounts || []).find(function(candidate) { return candidate.id === accountId; });\n");
        writer.write("        if (!account) return;\n");
        writer.write("        if (!Array.isArray(account.transactions)) account.transactions = [];\n");
        writer.write("        account.transactions.push({ id: 'tx-' + Math.random().toString(36).slice(2, 10), amount: amount, currency: currency });\n");
        writer.write("      }, 'Transaction added.');\n");
        writer.write("    });\n");
        writer.write("    managerUi.accountNameInput.addEventListener('keydown', function(event) {\n");
        writer.write("      if (event.key !== 'Enter') return;\n");
        writer.write("      event.preventDefault();\n");
        writer.write("      managerUi.createAccountButton.click();\n");
        writer.write("    });\n");
        writer.write("    managerUi.amountInput.addEventListener('keydown', function(event) {\n");
        writer.write("      if (event.key !== 'Enter') return;\n");
        writer.write("      event.preventDefault();\n");
        writer.write("      managerUi.addTransactionButton.click();\n");
        writer.write("    });\n");
        writer.write("    managerUi.currencyInput.addEventListener('keydown', function(event) {\n");
        writer.write("      if (event.key !== 'Enter') return;\n");
        writer.write("      event.preventDefault();\n");
        writer.write("      managerUi.addTransactionButton.click();\n");
        writer.write("    });\n");
        writer.write("  }\n");
        writer.write("  addButton.addEventListener('click', function() {\n");
        writer.write("    state = sanitizeManualCashState(state);\n");
        writer.write("    renderManualCashManager(state, managerUi, commitStateChange);\n");
        writer.write("    setManualCashManagerMessage(managerUi, '', false);\n");
        writer.write("    managerUi.open();\n");
        writer.write("  });\n");
        writer.write("}\n");
        writer.write("function initChartHoverEffects() {\n");
        writer.write("  var tooltip = document.createElement('div');\n");
        writer.write("  tooltip.className = 'chart-tooltip';\n");
        writer.write("  document.body.appendChild(tooltip);\n");
        writer.write("  var activeTarget = null;\n");
        writer.write("  function getActiveCurrencyCode() {\n");
        writer.write("    var input = document.getElementById('portfolio-currency-input');\n");
        writer.write("    var fallback = 'NOK';\n");
        writer.write("    var code = normalizeCurrencyCodeInput(input && input.value ? input.value : fallback);\n");
        writer.write("    return REPORT_RATES_TO_NOK[code] ? code : fallback;\n");
        writer.write("  }\n");
        writer.write("  function formatMoneyTooltip(target) {\n");
        writer.write("    var currency = getActiveCurrencyCode();\n");
        writer.write("    var targetRate = REPORT_RATES_TO_NOK[currency];\n");
        writer.write("    if (!targetRate || targetRate <= 0) return '';\n");
        writer.write("    var valueNok = Number(target.getAttribute('data-tooltip-value-nok') || '0');\n");
        writer.write("    var decimals = Number(target.getAttribute('data-tooltip-decimals') || '0');\n");
        writer.write("    var prefix = target.getAttribute('data-tooltip-prefix') || '';\n");
        writer.write("    var suffix = target.getAttribute('data-tooltip-suffix') || '';\n");
        writer.write("    var mode = target.getAttribute('data-tooltip-format') || 'money';\n");
        writer.write("    var converted = valueNok / targetRate;\n");
        writer.write("    var text = mode === 'compact'\n");
        writer.write("      ? prefix + formatCompactMoney(converted, currency)\n");
        writer.write("      : prefix + formatMoneyValue(converted, currency, decimals);\n");
        writer.write("    return text + suffix;\n");
        writer.write("  }\n");
        writer.write("  function readTooltipText(target) {\n");
        writer.write("    if (!target) return '';\n");
        writer.write("    if (target.getAttribute('data-tooltip-kind') === 'money') {\n");
        writer.write("      var moneyText = formatMoneyTooltip(target);\n");
        writer.write("      if (moneyText) return moneyText;\n");
        writer.write("      return String(target.getAttribute('data-tooltip-fallback') || '').trim();\n");
        writer.write("    }\n");
        writer.write("    return String(target.getAttribute('data-tooltip') || '').trim();\n");
        writer.write("  }\n");
        writer.write("  function placeTooltip(event) {\n");
        writer.write("    var offsetX = 14;\n");
        writer.write("    var offsetY = 14;\n");
        writer.write("    var x = event.clientX + offsetX;\n");
        writer.write("    var y = event.clientY + offsetY;\n");
        writer.write("    var rect = tooltip.getBoundingClientRect();\n");
        writer.write("    if (x + rect.width + 10 > window.innerWidth) x = event.clientX - rect.width - 12;\n");
        writer.write("    if (y + rect.height + 10 > window.innerHeight) y = event.clientY - rect.height - 12;\n");
        writer.write("    tooltip.style.left = Math.max(8, x) + 'px';\n");
        writer.write("    tooltip.style.top = Math.max(8, y) + 'px';\n");
        writer.write("  }\n");
        writer.write("  function hideTooltip() {\n");
        writer.write("    tooltip.classList.remove('visible');\n");
        writer.write("    activeTarget = null;\n");
        writer.write("  }\n");
        writer.write("  document.querySelectorAll('.chart-hover-target').forEach(function (target) {\n");
        writer.write("    var nativeTitle = target.querySelector('title');\n");
        writer.write("    if (nativeTitle) {\n");
        writer.write("      var isMoneyTitle = (nativeTitle.classList && nativeTitle.classList.contains('js-chart-money')) || nativeTitle.hasAttribute('data-value-nok');\n");
        writer.write("      if (isMoneyTitle) {\n");
        writer.write("        target.setAttribute('data-tooltip-kind', 'money');\n");
        writer.write("        target.setAttribute('data-tooltip-value-nok', nativeTitle.getAttribute('data-value-nok') || '0');\n");
        writer.write("        target.setAttribute('data-tooltip-decimals', nativeTitle.getAttribute('data-decimals') || '0');\n");
        writer.write("        target.setAttribute('data-tooltip-prefix', nativeTitle.getAttribute('data-prefix') || '');\n");
        writer.write("        target.setAttribute('data-tooltip-suffix', nativeTitle.getAttribute('data-suffix') || '');\n");
        writer.write("        target.setAttribute('data-tooltip-format', nativeTitle.getAttribute('data-format') || 'money');\n");
        writer.write("        target.setAttribute('data-tooltip-fallback', String(nativeTitle.textContent || '').trim());\n");
        writer.write("      } else {\n");
        writer.write("        target.setAttribute('data-tooltip-kind', 'text');\n");
        writer.write("        target.setAttribute('data-tooltip', String(nativeTitle.textContent || '').trim());\n");
        writer.write("      }\n");
        writer.write("      var key = String(nativeTitle.textContent || '').toLowerCase();\n");
        writer.write("      if (key) target.setAttribute('data-filter-key', key);\n");
        writer.write("      nativeTitle.remove();\n");
        writer.write("    }\n");
        writer.write("    target.addEventListener('mouseenter', function (event) {\n");
        writer.write("      var text = readTooltipText(target);\n");
        writer.write("      if (!text) return;\n");
        writer.write("      activeTarget = target;\n");
        writer.write("      target.classList.add('is-hovered');\n");
        writer.write("      tooltip.textContent = text;\n");
        writer.write("      tooltip.classList.add('visible');\n");
        writer.write("      placeTooltip(event);\n");
        writer.write("    });\n");
        writer.write("    target.addEventListener('mousemove', function (event) {\n");
        writer.write("      if (activeTarget !== target) return;\n");
        writer.write("      var text = readTooltipText(target);\n");
        writer.write("      if (text) tooltip.textContent = text;\n");
        writer.write("      placeTooltip(event);\n");
        writer.write("    });\n");
        writer.write("    target.addEventListener('mouseleave', function () {\n");
        writer.write("      target.classList.remove('is-hovered');\n");
        writer.write("      hideTooltip();\n");
        writer.write("    });\n");
        writer.write("  });\n");
        writer.write("  window.addEventListener('scroll', hideTooltip, true);\n");
        writer.write("}\n");
        writer.write("function initThemeToggle() {\n");
        writer.write("  var button = document.getElementById('report-theme-toggle');\n");
        writer.write("  var storageKey = 'portfolioReportTheme';\n");
        writer.write("  var savedTheme = '';\n");
        writer.write("  try { savedTheme = window.localStorage.getItem(storageKey) || ''; } catch (_) { savedTheme = ''; }\n");
        writer.write("  if (savedTheme === 'dark') {\n");
        writer.write("    document.body.classList.add('theme-dark');\n");
        writer.write("  }\n");
        writer.write("  function refreshLabel() {\n");
        writer.write("    if (!button) return;\n");
        writer.write("    button.textContent = document.body.classList.contains('theme-dark') ? 'Light mode' : 'Dark mode';\n");
        writer.write("  }\n");
        writer.write("  refreshLabel();\n");
        writer.write("  if (!button) return;\n");
        writer.write("  button.addEventListener('click', function () {\n");
        writer.write("    var dark = document.body.classList.toggle('theme-dark');\n");
        writer.write("    try { window.localStorage.setItem(storageKey, dark ? 'dark' : 'light'); } catch (_) { }\n");
        writer.write("    refreshLabel();\n");
        writer.write("  });\n");
        writer.write("}\n");
        writer.write("function initInteractiveChartControls() {\n");
        writer.write("  function ensureToolbar(container, row) {\n");
        writer.write("    var existing = container.querySelector('.chart-toolbar');\n");
        writer.write("    if (existing) return existing;\n");
        writer.write("    var toolbar = document.createElement('div');\n");
        writer.write("    toolbar.className = 'chart-toolbar';\n");
        writer.write("    row.insertAdjacentElement('afterend', toolbar);\n");
        writer.write("    return toolbar;\n");
        writer.write("  }\n");
        writer.write("  function attachControls(container) {\n");
        writer.write("    var svg = container.querySelector('svg.chart-svg');\n");
        writer.write("    if (!svg || svg.dataset.interactiveReady === '1') return;\n");
        writer.write("    svg.dataset.interactiveReady = '1';\n");
        writer.write("    var row = container.querySelector('.chart-title-row');\n");
        writer.write("    if (!row) return;\n");
        writer.write("    var toolbar = ensureToolbar(container, row);\n");
        writer.write("    var viewport = svg.closest('.chart-viewport');\n");
        writer.write("    if (!viewport) {\n");
        writer.write("      viewport = document.createElement('div');\n");
        writer.write("      viewport.className = 'chart-viewport';\n");
        writer.write("      svg.parentNode.insertBefore(viewport, svg);\n");
        writer.write("      viewport.appendChild(svg);\n");
        writer.write("    }\n");
        writer.write("    var zoomInBtn = document.createElement('button');\n");
        writer.write("    zoomInBtn.type = 'button';\n");
        writer.write("    zoomInBtn.className = 'chart-tool-btn';\n");
        writer.write("    zoomInBtn.textContent = '+';\n");
        writer.write("    var zoomOutBtn = document.createElement('button');\n");
        writer.write("    zoomOutBtn.type = 'button';\n");
        writer.write("    zoomOutBtn.className = 'chart-tool-btn';\n");
        writer.write("    zoomOutBtn.textContent = '-';\n");
        writer.write("    var resetBtn = document.createElement('button');\n");
        writer.write("    resetBtn.type = 'button';\n");
        writer.write("    resetBtn.className = 'chart-tool-btn';\n");
        writer.write("    resetBtn.textContent = 'Reset';\n");
        writer.write("    var filterInput = document.createElement('input');\n");
        writer.write("    filterInput.className = 'chart-filter-input';\n");
        writer.write("    filterInput.type = 'text';\n");
        writer.write("    filterInput.placeholder = 'Filter';\n");
        writer.write("    toolbar.appendChild(zoomInBtn);\n");
        writer.write("    toolbar.appendChild(zoomOutBtn);\n");
        writer.write("    toolbar.appendChild(resetBtn);\n");
        writer.write("    toolbar.appendChild(filterInput);\n");
        writer.write("    var state = { scale: 1, tx: 0, ty: 0, dragging: false, startX: 0, startY: 0, originTx: 0, originTy: 0 };\n");
        writer.write("    function applyTransform() {\n");
        writer.write("      svg.style.transform = 'translate(' + state.tx + 'px,' + state.ty + 'px) scale(' + state.scale + ')';\n");
        writer.write("    }\n");
        writer.write("    function zoomTo(nextScale, anchorClientX, anchorClientY) {\n");
        writer.write("      var clamped = Math.max(1, Math.min(4, nextScale));\n");
        writer.write("      if (Math.abs(clamped - state.scale) < 0.0001) return;\n");
        writer.write("      var rect = svg.getBoundingClientRect();\n");
        writer.write("      var anchorX = rect.width / 2;\n");
        writer.write("      var anchorY = rect.height / 2;\n");
        writer.write("      if (typeof anchorClientX === 'number' && typeof anchorClientY === 'number') {\n");
        writer.write("        anchorX = Math.max(0, Math.min(rect.width, anchorClientX - rect.left));\n");
        writer.write("        anchorY = Math.max(0, Math.min(rect.height, anchorClientY - rect.top));\n");
        writer.write("      }\n");
        writer.write("      var contentX = (anchorX - state.tx) / state.scale;\n");
        writer.write("      var contentY = (anchorY - state.ty) / state.scale;\n");
        writer.write("      state.scale = clamped;\n");
        writer.write("      state.tx = anchorX - contentX * state.scale;\n");
        writer.write("      state.ty = anchorY - contentY * state.scale;\n");
        writer.write("      if (state.scale <= 1) { state.tx = 0; state.ty = 0; }\n");
        writer.write("      applyTransform();\n");
        writer.write("    }\n");
        writer.write("    zoomInBtn.addEventListener('click', function() { zoomTo(state.scale + 0.2); });\n");
        writer.write("    zoomOutBtn.addEventListener('click', function() { zoomTo(state.scale - 0.2); });\n");
        writer.write("    resetBtn.addEventListener('click', function() { state.scale = 1; state.tx = 0; state.ty = 0; applyTransform(); filterInput.value = ''; applyFilter(''); });\n");
        writer.write("    svg.addEventListener('wheel', function(event) {\n");
        writer.write("      if (!(event.ctrlKey || event.metaKey)) return;\n");
        writer.write("      event.preventDefault();\n");
        writer.write("      var zoomFactor = Math.exp(-event.deltaY * 0.0015);\n");
        writer.write("      zoomTo(state.scale * zoomFactor, event.clientX, event.clientY);\n");
        writer.write("    }, { passive: false });\n");
        writer.write("    svg.addEventListener('mousedown', function(event) {\n");
        writer.write("      if (event.button !== 0 || state.scale <= 1) return;\n");
        writer.write("      state.dragging = true;\n");
        writer.write("      state.startX = event.clientX;\n");
        writer.write("      state.startY = event.clientY;\n");
        writer.write("      state.originTx = state.tx;\n");
        writer.write("      state.originTy = state.ty;\n");
        writer.write("      svg.classList.add('is-panning');\n");
        writer.write("      event.preventDefault();\n");
        writer.write("    });\n");
        writer.write("    window.addEventListener('mousemove', function(event) {\n");
        writer.write("      if (!state.dragging) return;\n");
        writer.write("      state.tx = state.originTx + (event.clientX - state.startX);\n");
        writer.write("      state.ty = state.originTy + (event.clientY - state.startY);\n");
        writer.write("      applyTransform();\n");
        writer.write("    });\n");
        writer.write("    window.addEventListener('mouseup', function() {\n");
        writer.write("      state.dragging = false;\n");
        writer.write("      svg.classList.remove('is-panning');\n");
        writer.write("    });\n");
        writer.write("    function applyFilter(term) {\n");
        writer.write("      var normalized = String(term || '').trim().toLowerCase();\n");
        writer.write("      var targets = svg.querySelectorAll('.chart-hover-target');\n");
        writer.write("      targets.forEach(function(target) {\n");
        writer.write("        var key = String(target.getAttribute('data-filter-key') || '').toLowerCase();\n");
        writer.write("        var visible = !normalized || key.indexOf(normalized) >= 0;\n");
        writer.write("        target.style.opacity = visible ? '' : '0.08';\n");
        writer.write("        target.style.pointerEvents = visible ? '' : 'none';\n");
        writer.write("      });\n");
        writer.write("    }\n");
        writer.write("    filterInput.addEventListener('input', function() { applyFilter(filterInput.value); });\n");
        writer.write("  }\n");
        writer.write("  document.querySelectorAll('.overview-chart, .allocation-panel, .hero-side').forEach(attachControls);\n");
        writer.write("}\n");
                writer.write("function formatSparklineRangeLabel(rangeKey) {\n");
                writer.write("  var key = String(rangeKey || '').toUpperCase();\n");
                writer.write("  if (key === '1M') return '1-month';\n");
                writer.write("  if (key === '3M') return '3-month';\n");
                writer.write("  if (key === '6M') return '6-month';\n");
                writer.write("  if (key === '1Y') return '1-year';\n");
                writer.write("  if (key === 'YTD') return 'year-to-date';\n");
                writer.write("  if (key === '3Y') return '3-year';\n");
                writer.write("  if (key === '5Y') return '5-year';\n");
                writer.write("  if (key === 'YEAR') return 'year';\n");
                writer.write("  return key || 'selected';\n");
                writer.write("}\n");
                writer.write("function updateSparklineReturnSummary(widget) {\n");
                writer.write("  if (!widget) return;\n");
                writer.write("  var summaryNode = widget.querySelector('.js-sparkline-return-summary');\n");
                writer.write("  if (!summaryNode) return;\n");
                writer.write("  var activePanel = widget.querySelector('.sparkline-panel.is-active');\n");
                writer.write("  if (!activePanel) { summaryNode.textContent = ''; return; }\n");
                writer.write("  var returnNok = Number(activePanel.getAttribute('data-return-nok') || 0);\n");
                writer.write("  var returnPct = Number(activePanel.getAttribute('data-return-pct') || 0);\n");
                writer.write("  var rangeLabel = formatSparklineRangeLabel(activePanel.getAttribute('data-range'));\n");
                writer.write("  var targetCurrency = getActiveReportCurrency();\n");
                writer.write("  var targetRate = REPORT_RATES_TO_NOK[targetCurrency] || REPORT_RATES_TO_NOK.NOK || 1;\n");
                writer.write("  var converted = Number.isFinite(targetRate) && targetRate > 0 ? (returnNok / targetRate) : returnNok;\n");
                writer.write("  var direction = returnNok > 0 ? 'gain' : (returnNok < 0 ? 'loss' : 'change');\n");
                writer.write("  summaryNode.classList.remove('positive', 'negative');\n");
                writer.write("  if (returnNok > 0) summaryNode.classList.add('positive');\n");
                writer.write("  else if (returnNok < 0) summaryNode.classList.add('negative');\n");
                writer.write("  summaryNode.textContent = formatMoneyValue(converted, targetCurrency, 2) + ' (' + formatPercentValue(returnPct, 2) + ') ' + rangeLabel + ' ' + direction;\n");
                writer.write("}\n");
                writer.write("function updateSparklineGrowthSummary(widget) {\n");
                writer.write("  if (!widget) return;\n");
                writer.write("  var summaryNode = widget.querySelector('.js-sparkline-growth-summary');\n");
                writer.write("  if (!summaryNode) return;\n");
                writer.write("  var activePanel = widget.querySelector('.sparkline-panel.is-active');\n");
                writer.write("  if (!activePanel) { summaryNode.textContent = ''; return; }\n");
                writer.write("  var growthNok = Number(activePanel.getAttribute('data-growth-nok') || 0);\n");
                writer.write("  var growthPct = Number(activePanel.getAttribute('data-growth-pct') || 0);\n");
                writer.write("  var rangeLabel = formatSparklineRangeLabel(activePanel.getAttribute('data-range'));\n");
                writer.write("  var targetCurrency = getActiveReportCurrency();\n");
                writer.write("  var targetRate = REPORT_RATES_TO_NOK[targetCurrency] || REPORT_RATES_TO_NOK.NOK || 1;\n");
                writer.write("  var converted = Number.isFinite(targetRate) && targetRate > 0 ? (growthNok / targetRate) : growthNok;\n");
                writer.write("  var direction = growthNok > 0 ? 'growth' : (growthNok < 0 ? 'decline' : 'change');\n");
                writer.write("  summaryNode.classList.remove('positive', 'negative');\n");
                writer.write("  if (growthNok > 0) summaryNode.classList.add('positive');\n");
                writer.write("  else if (growthNok < 0) summaryNode.classList.add('negative');\n");
                writer.write("  summaryNode.textContent = formatMoneyValue(converted, targetCurrency, 2) + ' (' + formatPercentValue(growthPct, 2) + ') ' + rangeLabel + ' ' + direction;\n");
                writer.write("}\n");
                writer.write("function refreshSparklineSummaries() {\n");
                writer.write("  document.querySelectorAll('.sparkline-widget').forEach(function(widget) {\n");
                writer.write("    updateSparklineReturnSummary(widget);\n");
                writer.write("    updateSparklineGrowthSummary(widget);\n");
                writer.write("  });\n");
                writer.write("}\n");
                writer.write("function initSparklineRangeControls() {\n");
                writer.write("  document.querySelectorAll('.sparkline-widget').forEach(function(widget) {\n");
                writer.write("    var rangeButtons = widget.querySelectorAll('.sparkline-range-btn');\n");
                writer.write("    var metricButtons = widget.querySelectorAll('.sparkline-metric-btn');\n");
                writer.write("    var panels = widget.querySelectorAll('.sparkline-panel');\n");
                writer.write("    if (!panels.length) return;\n");
                writer.write("    var defaultRangeSource = widget.querySelector('.sparkline-range-btn.is-active') || rangeButtons[0] || panels[0];\n");
                writer.write("    var activeRange = defaultRangeSource ? defaultRangeSource.getAttribute('data-range') : '';\n");
                writer.write("    var metricSource = widget.querySelector('.sparkline-metric-btn.is-active') || metricButtons[0] || panels[0];\n");
                writer.write("    var activeMetric = metricSource ? String(metricSource.getAttribute('data-metric') || '') : '';\n");
                writer.write("    function refresh() {\n");
                writer.write("      rangeButtons.forEach(function(btn) {\n");
                writer.write("        btn.classList.toggle('is-active', btn.getAttribute('data-range') === activeRange);\n");
                writer.write("      });\n");
                writer.write("      metricButtons.forEach(function(btn) {\n");
                writer.write("        btn.classList.toggle('is-active', btn.getAttribute('data-metric') === activeMetric);\n");
                writer.write("      });\n");
                writer.write("      panels.forEach(function(panel) {\n");
                writer.write("        var matchRange = panel.getAttribute('data-range') === activeRange;\n");
                writer.write("        var panelMetric = String(panel.getAttribute('data-metric') || '');\n");
                writer.write("        var matchMetric = !activeMetric || panelMetric === activeMetric;\n");
                writer.write("        panel.classList.toggle('is-active', matchRange && matchMetric);\n");
                writer.write("      });\n");
                writer.write("      updateSparklineReturnSummary(widget);\n");
                writer.write("      updateSparklineGrowthSummary(widget);\n");
                writer.write("    }\n");
                writer.write("    rangeButtons.forEach(function(btn) {\n");
                writer.write("      btn.addEventListener('click', function() {\n");
                writer.write("        activeRange = btn.getAttribute('data-range');\n");
                writer.write("        refresh();\n");
                writer.write("      });\n");
                writer.write("    });\n");
                writer.write("    metricButtons.forEach(function(btn) {\n");
                writer.write("      btn.addEventListener('click', function() {\n");
                writer.write("        activeMetric = btn.getAttribute('data-metric');\n");
                writer.write("        refresh();\n");
                writer.write("      });\n");
                writer.write("    });\n");
                writer.write("    refresh();\n");
                writer.write("  });\n");
                writer.write("}\n");
        writer.write("function initTimelineInfoPopup() {\n");
        writer.write("  document.querySelectorAll('.timeline-info-btn').forEach(function(openBtn) {\n");
        writer.write("    if (openBtn.dataset.timelineInfoBound === '1') return;\n");
        writer.write("    openBtn.dataset.timelineInfoBound = '1';\n");
        writer.write("    var host = openBtn.closest('.hero-side, .annual-graphs-section');\n");
        writer.write("    var overlay = host ? host.querySelector('.timeline-info-overlay') : null;\n");
        writer.write("    if (!overlay) return;\n");
        writer.write("    var closeBtn = overlay.querySelector('.timeline-info-close');\n");
        writer.write("    function closeDialog() { overlay.hidden = true; }\n");
        writer.write("    function openDialog() { overlay.hidden = false; }\n");
        writer.write("    openBtn.addEventListener('click', openDialog);\n");
        writer.write("    if (closeBtn) closeBtn.addEventListener('click', closeDialog);\n");
        writer.write("    overlay.addEventListener('click', function(event) {\n");
        writer.write("      if (event.target === overlay) closeDialog();\n");
        writer.write("    });\n");
        writer.write("    window.addEventListener('keydown', function(event) {\n");
        writer.write("      if (!overlay.hidden && event.key === 'Escape') closeDialog();\n");
        writer.write("    });\n");
        writer.write("  });\n");
        writer.write("}\n");
        writer.write("function initInlineCellDragHandles() {\n");
        writer.write("  var active = null;\n");
        writer.write("  var startX = 0;\n");
        writer.write("  var startScrollLeft = 0;\n");
        writer.write("  function stopDrag() {\n");
        writer.write("    if (!active) return;\n");
        writer.write("    active.classList.remove('is-dragging');\n");
        writer.write("    document.body.classList.remove('inline-cell-dragging');\n");
        writer.write("    active = null;\n");
        writer.write("  }\n");
        writer.write("  window.addEventListener('mousemove', function (event) {\n");
        writer.write("    if (!active) return;\n");
        writer.write("    var deltaX = event.clientX - startX;\n");
        writer.write("    active.scrollLeft = startScrollLeft - deltaX;\n");
        writer.write("  });\n");
        writer.write("  window.addEventListener('mouseup', stopDrag);\n");
        writer.write("  window.addEventListener('mouseleave', stopDrag);\n");
        writer.write("  document.querySelectorAll('.security-scroll, .ticker-scroll').forEach(function (node) {\n");
        writer.write("    if (node.dataset.dragHandleInit === '1') return;\n");
        writer.write("    node.dataset.dragHandleInit = '1';\n");
        writer.write("    node.addEventListener('mousedown', function (event) {\n");
        writer.write("      if (event.button !== 0) return;\n");
        writer.write("      active = node;\n");
        writer.write("      startX = event.clientX;\n");
        writer.write("      startScrollLeft = node.scrollLeft;\n");
        writer.write("      node.classList.add('is-dragging');\n");
        writer.write("      document.body.classList.add('inline-cell-dragging');\n");
        writer.write("      event.preventDefault();\n");
        writer.write("    });\n");
        writer.write("  });\n");
        writer.write("}\n");
        writer.write("function toDownloadSafeName(value) {\n");
        writer.write("  return String(value || 'chart')\n");
        writer.write("    .trim()\n");
        writer.write("    .toLowerCase()\n");
        writer.write("    .replace(/[^a-z0-9]+/g, '-')\n");
        writer.write("    .replace(/^-+|-+$/g, '') || 'chart';\n");
        writer.write("}\n");
        writer.write("function resolveChartName(container, fallback) {\n");
        writer.write("  if (!container) return fallback;\n");
        writer.write("  var heading = container.querySelector('h3, h4, .hero-side-title');\n");
        writer.write("  var text = heading ? String(heading.textContent || '').trim() : '';\n");
        writer.write("  return text || fallback;\n");
        writer.write("}\n");
        writer.write("function downloadSvgImage(svgElement, fileName) {\n");
        writer.write("  if (!svgElement) return;\n");
        writer.write("  var cloned = svgElement.cloneNode(true);\n");
        writer.write("  if (!cloned.getAttribute('xmlns')) cloned.setAttribute('xmlns', 'http://www.w3.org/2000/svg');\n");
        writer.write("  if (!cloned.getAttribute('xmlns:xlink')) cloned.setAttribute('xmlns:xlink', 'http://www.w3.org/1999/xlink');\n");
        writer.write("  var serializer = new XMLSerializer();\n");
        writer.write("  var source = serializer.serializeToString(cloned);\n");
        writer.write("  var encoded = window.btoa(unescape(encodeURIComponent(source)));\n");
        writer.write("  var dataUrl = 'data:image/svg+xml;base64,' + encoded;\n");
        writer.write("  var img = new Image();\n");
        writer.write("  img.onload = function() {\n");
        writer.write("    var width = Number(svgElement.viewBox && svgElement.viewBox.baseVal && svgElement.viewBox.baseVal.width) || svgElement.clientWidth || 1200;\n");
        writer.write("    var height = Number(svgElement.viewBox && svgElement.viewBox.baseVal && svgElement.viewBox.baseVal.height) || svgElement.clientHeight || 800;\n");
        writer.write("    width = Math.max(1, Math.round(width));\n");
        writer.write("    height = Math.max(1, Math.round(height));\n");
        writer.write("    var canvas = document.createElement('canvas');\n");
        writer.write("    canvas.width = width;\n");
        writer.write("    canvas.height = height;\n");
        writer.write("    var context = canvas.getContext('2d');\n");
        writer.write("    if (!context) return;\n");
        writer.write("    context.fillStyle = '#ffffff';\n");
        writer.write("    context.fillRect(0, 0, width, height);\n");
        writer.write("    context.drawImage(img, 0, 0, width, height);\n");
        writer.write("    canvas.toBlob(function(blob) {\n");
        writer.write("      if (!blob) return;\n");
        writer.write("      var url = URL.createObjectURL(blob);\n");
        writer.write("      var anchor = document.createElement('a');\n");
        writer.write("      anchor.href = url;\n");
        writer.write("      anchor.download = toDownloadSafeName(fileName) + '.png';\n");
        writer.write("      document.body.appendChild(anchor);\n");
        writer.write("      anchor.click();\n");
        writer.write("      anchor.remove();\n");
        writer.write("      URL.revokeObjectURL(url);\n");
        writer.write("    }, 'image/png');\n");
        writer.write("  };\n");
        writer.write("  img.src = dataUrl;\n");
        writer.write("}\n");
        writer.write("function appendChartDownloadButton(container, findSvg, fallbackName) {\n");
        writer.write("  if (!container || container.querySelector('.chart-download-btn')) return;\n");
        writer.write("  var heading = container.querySelector(':scope > h3, :scope > h4, :scope > .hero-side-title');\n");
        writer.write("  if (!heading) return;\n");
        writer.write("  var row = document.createElement('div');\n");
        writer.write("  row.className = 'chart-title-row';\n");
        writer.write("  var button = document.createElement('button');\n");
        writer.write("  button.type = 'button';\n");
        writer.write("  button.className = 'chart-download-btn';\n");
        writer.write("  button.setAttribute('title', 'Download PNG');\n");
        writer.write("  button.setAttribute('aria-label', 'Download PNG');\n");
        writer.write("  button.innerHTML = '<svg viewBox=\"0 0 24 24\" aria-hidden=\"true\"><path d=\"M12 3v11\"></path><path d=\"M7.5 10.5L12 15l4.5-4.5\"></path><path d=\"M4 18h16\"></path></svg>';\n");
        writer.write("  button.addEventListener('click', function() {\n");
        writer.write("    var svg = typeof findSvg === 'function' ? findSvg() : null;\n");
        writer.write("    if (!svg) return;\n");
        writer.write("    var chartName = resolveChartName(container, fallbackName);\n");
        writer.write("    downloadSvgImage(svg, chartName);\n");
        writer.write("  });\n");
        writer.write("  heading.parentNode.insertBefore(row, heading);\n");
        writer.write("  row.appendChild(heading);\n");
        writer.write("  row.appendChild(button);\n");
        writer.write("}\n");
        writer.write("function initChartDownloadButtons() {\n");
        writer.write("  document.querySelectorAll('.overview-chart, .allocation-panel').forEach(function(container) {\n");
        writer.write("    appendChartDownloadButton(container, function() {\n");
        writer.write("      return container.querySelector('svg.chart-svg');\n");
        writer.write("    }, 'chart');\n");
        writer.write("  });\n");
        writer.write("  document.querySelectorAll('.hero-side').forEach(function(side) {\n");
        writer.write("    appendChartDownloadButton(side, function() {\n");
        writer.write("      var widget = side.querySelector('.sparkline-widget');\n");
        writer.write("      if (!widget) return null;\n");
        writer.write("      var activePanel = widget.querySelector('.sparkline-panel.is-active');\n");
        writer.write("      var fallbackPanel = widget.querySelector('.sparkline-panel');\n");
        writer.write("      var panel = activePanel || fallbackPanel;\n");
        writer.write("      return panel ? panel.querySelector('svg') : null;\n");
        writer.write("    }, 'portfolio-value-timeline');\n");
        writer.write("  });\n");
        writer.write("}\n");
        writer.write("(function initReportCurrencyInput() {\n");
        writer.write("  var input = document.getElementById('portfolio-currency-input');\n");
        writer.write("  if (!input) return;\n");
        writer.write("  input.addEventListener('keydown', function (event) {\n");
        writer.write("    if (event.key !== 'Enter') return;\n");
        writer.write("    event.preventDefault();\n");
        writer.write("    var target = normalizeCurrencyCodeInput(input.value);\n");
        writer.write("    if (!target) { target = '" + DEFAULT_TOTAL_CURRENCY + "'; }\n");
        writer.write("    if (!refreshReportTotalsCurrency(target) || !refreshReportChartsCurrency(target)) {\n");
        writer.write("      window.alert('No exchange rate available for ' + target + '. Try one of: ' + Object.keys(REPORT_RATES_TO_NOK).join(', '));\n");
        writer.write("      input.value = '" + DEFAULT_TOTAL_CURRENCY + "';\n");
        writer.write("      refreshReportTotalsCurrency('" + DEFAULT_TOTAL_CURRENCY + "');\n");
        writer.write("      refreshReportChartsCurrency('" + DEFAULT_TOTAL_CURRENCY + "');\n");
        writer.write("      return;\n");
        writer.write("    }\n");
        writer.write("    input.value = target;\n");
        writer.write("  });\n");
        writer.write("  refreshReportTotalsCurrency('" + DEFAULT_TOTAL_CURRENCY + "');\n");
        writer.write("  refreshReportChartsCurrency('" + DEFAULT_TOTAL_CURRENCY + "');\n");
        writer.write("  initSparklineRangeControls();\n");
        writer.write("  initTimelineInfoPopup();\n");
        writer.write("  initInlineCellDragHandles();\n");
        writer.write("  initSortableTables();\n");
        writer.write("  initOverviewModeSwitcher();\n");
        writer.write("  initOverviewDayCharts();\n");
        writer.write("  initChartDownloadButtons();\n");
        writer.write("  initChartHoverEffects();\n");
        writer.write("  initInteractiveChartControls();\n");
        writer.write("  initThemeToggle();\n");
        writer.write("  initPriceRefreshButton();\n");
        writer.write("  initManualCashHoldings();\n");
        writer.write("})();\n");
    }

    private static String escapeHtml(String text) {
        if (text == null) return "";
        return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\"", "&quot;");
    }
}
