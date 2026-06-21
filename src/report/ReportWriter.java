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
    static final String DEFAULT_TOTAL_CURRENCY = "NOK";
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
            annualSnapshotRows = buildAnnualSnapshotRows(store, resolveYearSnapshotDate(snapshotYear));
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
            writer.write("        .realized-highlights { display:grid; grid-template-columns:repeat(6,minmax(0,1fr)); gap:10px; margin:10px 0 14px; }\n");
            writer.write("        @media (max-width:1060px) { .realized-highlights { grid-template-columns:repeat(3,minmax(0,1fr)); } }\n");
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
            writer.write("        .currency-select { border:1px solid rgba(235,245,255,.45); border-radius:6px; background-color:rgba(255,255,255,.18); color:#fff; font-weight:700; text-transform:uppercase; font-size:.84rem; padding:2px 20px 2px 6px; outline:none; cursor:pointer; -webkit-appearance:none; appearance:none; background-image:url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 6'%3E%3Cpath d='M1 1 5 5 9 1' fill='none' stroke='%23ffffff' stroke-width='1.5' stroke-linecap='round' stroke-linejoin='round'/%3E%3C/svg%3E\"); background-repeat:no-repeat; background-position:right 6px center; background-size:9px 6px; }\n");
            writer.write("        .currency-select:focus { border-color:#fff; background-color:rgba(255,255,255,.28); }\n");
            writer.write("        .currency-select option { color:#16202a; background:#ffffff; text-transform:none; font-weight:600; }\n");
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
            writer.write("        .report-standard .kpi-help { color:#7a96ae; }\n");
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
            writer.write("        .cash-account-block.is-hidden { background:#f4f7fa; border-color:#c8d8e8; }\n");
            writer.write("        .cash-account-block.is-hidden .cash-account-summary,.cash-account-block.is-hidden .cash-transaction-list { opacity:.45; pointer-events:none; }\n");
            writer.write("        .cash-account-head { display:flex; justify-content:space-between; align-items:center; gap:8px; margin-bottom:8px; }\n");
            writer.write("        .cash-account-head-buttons { display:flex; gap:5px; flex-shrink:0; }\n");
            writer.write("        .cash-account-title { margin:0; font-size:.84rem; font-weight:700; color:#1e3951; }\n");
            writer.write("        .cash-account-title.is-hidden-label::after { content:' (hidden)'; font-weight:400; color:#7a96ae; font-size:.78rem; }\n");
            writer.write("        .cash-account-summary { margin:0 0 8px; font-size:.78rem; color:#355572; }\n");
            writer.write("        .cash-tx-edit-row { display:flex; align-items:center; gap:5px; flex-wrap:wrap; }\n");
            writer.write("        .cash-tx-edit-row input { border:1px solid #a8bfd4; border-radius:6px; height:26px; padding:0 6px; font-size:.78rem; }\n");
            writer.write("        .cash-tx-edit-amount { width:88px; }\n");
            writer.write("        .cash-tx-edit-currency { width:58px; text-transform:uppercase; }\n");
            writer.write("        .cash-manager-empty { margin:0 0 8px; font-size:.78rem; color:#56708a; }\n");
            writer.write("        .cash-transaction-list { margin:0 0 8px; padding-left:16px; display:grid; gap:3px; }\n");
            writer.write("        .cash-transaction-item { display:flex; align-items:center; justify-content:space-between; gap:6px; font-size:.76rem; }\n");
            writer.write("        .cash-transaction-item .cash-manager-btn { height:24px; padding:0 8px; font-size:.72rem; border-radius:6px; }\n");
            writer.write("        .cash-account-caret { color:#5c7795; font-weight:400; }\n");
            writer.write("        .annual-headline-grid .kpi-card { min-height:116px; }\n");
            writer.write("        .kpi-label { color:#c8d9eb; font-size:.8rem; text-transform:uppercase; }\n");
            writer.write("        .kpi-value { margin-top:2px; font-size:1.02rem; font-weight:700; color:#fff; }\n");
            writer.write("        .kpi-help { margin-top:4px; font-size:.72rem; line-height:1.35; color:#b6c9dc; text-transform:none; letter-spacing:0; }\n");
            writer.write("        .performer { margin-top:6px; font-size:.84rem; color:#dce8f3; }\n");
            writer.write("        .performer strong { display:block; font-size:.9rem; margin-bottom:2px; }\n");
            writer.write("        .report-standard .kpi-card-bestworst .performer strong { white-space:nowrap; overflow-x:auto; overflow-y:hidden; text-overflow:clip; scrollbar-width:none; -ms-overflow-style:none; }\n");
            writer.write("        .report-standard .kpi-card-bestworst .performer strong::-webkit-scrollbar { display:none; width:0; height:0; }\n");
            writer.write("        .performer-metrics { display:block; }\n");
            writer.write("        .hero-side { position:relative; background:rgba(255,255,255,.06); border:1px solid rgba(235,245,255,.22); border-radius:12px; padding:10px; min-height:172px; }\n");
            writer.write("        .timeline-title-row { display:flex; align-items:center; gap:6px; margin-bottom:8px; }\n");
            writer.write("        .timeline-title-row .annual-kpi-deck-title { margin:0; }\n");
            writer.write("        .hero-side-title { color:#d4e3f0; font-size:.86rem; text-transform:uppercase; margin:0; }\n");
            writer.write("        .timeline-info-btn { width:18px; height:18px; border-radius:999px; border:1px solid rgba(235,245,255,.55); background:rgba(255,255,255,.14); color:#e8f2fb; font-size:.72rem; font-weight:800; line-height:1; cursor:pointer; display:inline-flex; align-items:center; justify-content:center; padding:0; }\n");
            writer.write("        .timeline-info-btn:hover { background:rgba(255,255,255,.24); }\n");
            writer.write("        .annual-graphs-section .timeline-info-btn,.annual-kpi-deck .timeline-info-btn { border-color:#8fa9c2; background:#eaf2fb; color:#24415a; }\n");
            writer.write("        .annual-graphs-section .timeline-info-btn:hover,.annual-kpi-deck .timeline-info-btn:hover { background:#ddeaf7; }\n");
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
            writer.write("        .allocation-drilldown-list { margin:4px 2% 0; border:1px solid #d5e1ec; border-radius:8px; background:#fff; max-height:200px; overflow:auto; padding:8px 10px; }\n");
            writer.write("        .allocation-drilldown-list[hidden] { display:none !important; }\n");
            writer.write("        .allocation-drilldown-grid { display:grid; grid-template-columns:1fr auto; gap:4px 12px; align-items:start; }\n");
            writer.write("        .allocation-drilldown-selected { margin:0 0 8px; font-size:.78rem; font-weight:700; color:#2e4963; }\n");
            writer.write("        .allocation-drilldown-name { min-width:0; display:flex; align-items:flex-start; gap:6px; color:#2f2f2f; font-size:.8rem; line-height:1.25; word-break:break-word; overflow-wrap:anywhere; }\n");
            writer.write("        .allocation-drilldown-dot { flex:0 0 auto; width:7px; height:7px; border-radius:999px; background:#3b5978; margin-top:4px; }\n");
            writer.write("        .allocation-drilldown-pct { color:#4a4a4a; font-size:.8rem; line-height:1.25; text-align:right; white-space:nowrap; }\n");
            writer.write("        .chart-svg { width:100%; height:auto; background:var(--card); border:1px solid var(--line); border-radius:8px; }\n");
            writer.write("        .allocation-panel .chart-svg { width:96%; margin:6px auto 10px; display:block; }\n");
            writer.write("        .security-pie-panel .chart-svg, .security-bar-panel .chart-svg { height:340px; width:100%; }\n");
            writer.write("        .chart-hover-target { cursor:pointer; transform-box:fill-box; transform-origin:center; transition:transform .14s ease, filter .14s ease, opacity .14s ease; }\n");
            writer.write("        .chart-hover-target.is-hovered { filter:brightness(1.08); opacity:.96; }\n");
            writer.write("        .chart-hover-bar.is-hovered { transform:translateY(-2px); }\n");
            writer.write("        .chart-hover-slice.is-hovered { transform:scale(1.03); }\n");
            writer.write("        .chart-hover-slice.is-selected { transform:scale(1.03); filter:brightness(1.12); opacity:1; stroke:rgba(10,24,40,.35); stroke-width:1.4; }\n");
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
            writer.write("        .chart-hover-legend { cursor:pointer; }\n");
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
            writer.write("        body.theme-dark .allocation-drilldown-list { border-color:#32485f; background:#102033; }\n");
            writer.write("        body.theme-dark .allocation-drilldown-selected { color:#c8d9ea; }\n");
            writer.write("        body.theme-dark .allocation-drilldown-name { color:#d4e1ee; }\n");
            writer.write("        body.theme-dark .allocation-drilldown-pct { color:#c2d3e4; }\n");
            writer.write("        body.theme-dark .chart-hover-legend { opacity:1; }\n");
            writer.write("        body.theme-dark .chart-download-btn { background:#1f3347; border-color:#3a5878; color:#d8e7f5; }\n");
            writer.write("        body.theme-dark .chart-download-btn:hover { background:#2a4663; }\n");
            writer.write("        body.theme-dark .chart-tool-btn { background:#1f3347; border-color:#3a5878; color:#d8e7f5; }\n");
            writer.write("        body.theme-dark .chart-tool-btn:hover { background:#2a4663; }\n");
            writer.write("        body.theme-dark .chart-filter-input { background:#152535; border-color:#3a5878; color:#d8e7f5; }\n");
            writer.write("        body.theme-dark .chart-filter-input::placeholder { color:#6a8daa; }\n");
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
            writer.write("        body.theme-dark .annual-graphs-section .timeline-info-btn,body.theme-dark .annual-kpi-deck .timeline-info-btn { border-color:#4d6a87; background:#21374e; color:#d7e8f8; }\n");
            writer.write("        body.theme-dark .annual-graphs-section .timeline-info-btn:hover,body.theme-dark .annual-kpi-deck .timeline-info-btn:hover { background:#2a4663; }\n");
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
            writer.write("        body.theme-dark .kpi-help { color:#b6c9dc; }\n");
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
            writer.write("        @media (max-width:760px) { .annual-graph-card, .total-return-chart, .report-standard .total-return-chart, .report-standard .annual-graph-card { min-height:0; } }\n");
            writer.write("        @media (max-width:760px) { .total-return-chart .chart-viewport, .security-bar-panel .chart-viewport, .security-pie-panel .chart-viewport { overflow-x:auto; overflow-y:hidden; -webkit-overflow-scrolling:touch; } }\n");
            writer.write("        @media (max-width:760px) { .total-return-chart .chart-svg { min-width:640px; height:auto; } }\n");
            writer.write("        @media (max-width:760px) { .security-bar-panel .chart-svg, .security-pie-panel .chart-svg { width:100%; min-width:520px; height:auto; margin:6px 0 10px; } }\n");
            writer.write("        @media (max-width:760px) { .annual-graph-content .sparkline-panel > svg { min-height:0; height:auto; } }\n");
            writer.write("        @media (max-width:1200px) { .report-standard .annual-summary-grid{grid-template-columns:repeat(4,minmax(0,1fr));} }\n");
            writer.write("        @media (max-width:760px) { .report-standard .annual-summary-grid{grid-template-columns:repeat(2,minmax(0,1fr));} }\n");
            writer.write("        @media (max-width:560px) { .report-standard .annual-summary-grid,.annual-summary-grid{grid-template-columns:repeat(2,minmax(0,1fr));} .realized-highlights{grid-template-columns:repeat(2,minmax(0,1fr));} .page{padding:12px 6px 20px;} .hero-title h1{font-size:1.45rem;} .annual-kpi-deck-title{font-size:.96rem;} }\n");
            writer.write("        @media (max-width:760px) { .report-standard .overview-summary-table{table-layout:auto; width:max-content; min-width:100%;} .report-standard .overview-summary-table th, .report-standard .overview-summary-table td{overflow:visible !important; text-overflow:clip !important;} .report-standard .overview-summary-table tr > *:nth-child(n+3){overflow:visible !important; text-overflow:clip !important; white-space:nowrap !important;} }\n");
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
            ReportClientScript.write(writer, ratesToNok);
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

    // Options for the header display-currency dropdown: the default currency first,
    // then every currency we have an FX rate for (so all options are convertible).
    private static String buildCurrencyOptionsHtml(Map<String, Double> ratesToNok) {
        java.util.TreeSet<String> codes = new java.util.TreeSet<>();
        if (ratesToNok != null) {
            for (String code : ratesToNok.keySet()) {
                if (code != null && !code.isBlank()) codes.add(code.trim().toUpperCase(Locale.ROOT));
            }
        }
        codes.add(DEFAULT_TOTAL_CURRENCY);
        StringBuilder sb = new StringBuilder();
        sb.append("<option value=\"").append(escapeHtml(DEFAULT_TOTAL_CURRENCY)).append("\" selected>")
          .append(escapeHtml(DEFAULT_TOTAL_CURRENCY)).append("</option>");
        for (String code : codes) {
            if (code.equals(DEFAULT_TOTAL_CURRENCY)) continue;
            sb.append("<option value=\"").append(escapeHtml(code)).append("\">")
              .append(escapeHtml(code)).append("</option>");
        }
        return sb.toString();
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
        writer.write("<span class=\"meta-chip\">Currency: <strong><select id=\"portfolio-currency-input\" class=\"currency-select\" aria-label=\"Display currency\" title=\"Choose display currency\">" + buildCurrencyOptionsHtml(ratesToNok) + "</select></strong></span>\n");
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
        writer.write("<p class=\"annual-graph-note\">Month-end portfolio value in selected year (latest available date for current month).</p>\n");
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
        writer.write("<p class=\"annual-graph-note\">Month-end portfolio return in selected year (latest available date for current month).</p>\n");
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

        LocalDate snapshotDate = resolveYearSnapshotDate(reportYear);
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

    private static LocalDate resolveYearSnapshotDate(int year) {
        int safeYear = Math.max(2000, Math.min(2100, year));
        LocalDate yearEnd = LocalDate.of(safeYear, 12, 31);
        LocalDate today = LocalDate.now();
        if (today.getYear() == safeYear && today.isBefore(yearEnd)) {
            return today;
        }
        return yearEnd;
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
        writer.write("<div class=\"overview-mode-shell\" role=\"group\" aria-label=\"Realized details controls\">\n");
        writer.write("<button type=\"button\" class=\"overview-mode-btn overview-details-toggle-btn\" data-detail-label=\"Open all details\" onclick=\"toggleDetailGroup('realized-details-year', this)\">Open all details ▸</button>\n");
        writer.write("</div>\n");
        writer.write("<div class=\"table-wrap\">\n<table class=\"realized-table\">\n");
        ReportTemplateHelper.writeHtmlRow(writer, true, "Ticker", "Security", "Cost Basis", "Sales Value", "Gain/Loss", "Dividends", "Total Return");

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
        int currentYear = LocalDate.now().getYear();
        reportYear = Math.max(2000, Math.min(currentYear, reportYear));

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
        writer.write("<span class=\"meta-chip\">Currency: <strong><select id=\"portfolio-currency-input\" class=\"currency-select\" aria-label=\"Display currency\" title=\"Choose display currency\">" + buildCurrencyOptionsHtml(ratesToNok) + "</select></strong></span>\n");
        writer.write("<button id=\"refresh-prices-btn\" class=\"hero-refresh-btn\" type=\"button\">Update</button>\n");
        writer.write("<button id=\"report-theme-toggle\" class=\"hero-theme-btn\" type=\"button\">Dark mode</button>\n");
        writer.write("</div>\n");
        writer.write("<div id=\"refresh-prices-status\" class=\"price-refresh-status\"></div>\n");
        writer.write("</div>\n");
        writer.write("</section>\n");
        writer.write("<section class=\"annual-kpi-deck\">\n");
        writer.write("<div class=\"timeline-title-row\"><h2 class=\"annual-kpi-deck-title\">Portfolio Highlights</h2><button type=\"button\" class=\"timeline-info-btn\" aria-label=\"Show info about portfolio highlights\" title=\"What is shown here?\">i</button></div>\n");
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
            + formatBucketsInTarget(portfolioValueBuckets, DEFAULT_TOTAL_CURRENCY, 0, ratesToNok) + "</div><div class=\"kpi-help\">Market value + cash holdings</div></article>\n");

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
        writer.write("<div class=\"timeline-info-overlay\" hidden><div class=\"timeline-info-dialog\" role=\"dialog\" aria-modal=\"true\" aria-label=\"Portfolio highlights info\"><div class=\"timeline-info-header\"><h4>Portfolio Highlights — What is shown?</h4><button type=\"button\" class=\"timeline-info-close\" aria-label=\"Close\">×</button></div><div class=\"timeline-info-body\"><ul><li><strong>Total Return &amp; 1-Year Change:</strong> Include <em>everything</em> — both current holdings and all realized positions. This gives your true all-in return across the full portfolio history.</li><li><strong>Unrealized Return:</strong> Only the open positions you currently hold. Fully sold securities are not counted here.</li><li><strong>Realized Return &amp; Dividends:</strong> Only positions still in the portfolio (fully realized positions that you no longer hold appear exclusively in the Realized Overview below).</li><li><strong>Day Change, Market Value, Cost Basis:</strong> Current portfolio only — what you hold right now.</li><li><strong>Fully realized positions:</strong> If you have sold all shares in a security and no longer hold it, it does <em>not</em> appear here. See the <strong>Realized Overview</strong> section for those.</li><li><strong>Partially realized or repurchased:</strong> If you sold some shares but still hold others (or bought back in), the position <em>is</em> included here and the realized gain on the sold portion is reflected in the return figures.</li></ul></div></div></div>\n");
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
        writer.write("<div class=\"timeline-title-row\"><h2 class=\"annual-kpi-deck-title\">Risk Analytics</h2><button type=\"button\" class=\"timeline-info-btn\" aria-label=\"Show info about risk analytics\" title=\"How is this calculated?\">i</button></div>\n");
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
        writer.write("<div class=\"timeline-info-overlay\" hidden><div class=\"timeline-info-dialog\" role=\"dialog\" aria-modal=\"true\" aria-label=\"Risk analytics info\"><div class=\"timeline-info-header\"><h4>Risk Analytics — How it is calculated</h4><button type=\"button\" class=\"timeline-info-close\" aria-label=\"Close\">×</button></div><div class=\"timeline-info-body\"><ul>"
            + "<li><strong>Volatility (Ann.):</strong> Standard deviation of monthly portfolio returns, multiplied by √12 to annualize. Measures how much your returns vary from month to month. A higher number means wider swings.</li>"
            + "<li><strong>Sharpe Ratio:</strong> Average monthly return divided by the monthly return standard deviation, then annualized (×√12). Uses 0 % as the risk-free rate. A ratio above 1.0 is generally considered good — it means you are earning more return per unit of risk taken.</li>"
            + "<li><strong>Beta:</strong> Covariance of your portfolio's monthly returns with the benchmark's monthly returns, divided by the benchmark's monthly return variance. Beta &gt; 1 means your portfolio tends to move more than the benchmark; Beta &lt; 1 means it moves less. Requires at least 6 months of overlapping benchmark data.</li>"
            + "<li><strong>Monte Carlo:</strong> Simulates future portfolio value by randomly sampling monthly returns based on your historical mean return and volatility, repeated across many iterations. Shows the <strong>median</strong> projected terminal value, along with the 10th percentile (P10 — pessimistic) and 90th percentile (P90 — optimistic) outcomes. This is a statistical estimate, not a forecast.</li>"
            + "</ul></div></div></div>\n");
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

    private static String buildRealizedLeaderCard(String label, Security security, double value, String valueClass, Map<String, Double> ratesToNok) {
        if (security == null) {
            return "<article class=\"kpi-card\"><div class=\"kpi-label\">" + escapeHtml(label) + "</div><div class=\"kpi-value\">-</div></article>\n";
        }
        String currency = security.getCurrencyCode();
        LinkedHashMap<String, Double> buckets = singleCurrencyBuckets(currency, value);
        String formatted = formatBucketsInTarget(buckets, DEFAULT_TOTAL_CURRENCY, 0, ratesToNok);
        String ticker = escapeHtml(security.getTicker());
        String name = escapeHtml(security.getDisplayName());
        String cls = valueClass.isEmpty() ? "" : " " + valueClass;
        return "<article class=\"kpi-card\">"
            + "<div class=\"kpi-label\">" + escapeHtml(label) + "</div>"
            + "<div class=\"kpi-value js-convert-money" + cls + "\" data-buckets=\"" + escapeHtml(toBucketsJson(buckets)) + "\" data-decimals=\"0\">" + formatted + "</div>"
            + "<div class=\"kpi-help\" style=\"overflow:hidden;text-overflow:ellipsis;white-space:nowrap;\"><strong>" + ticker + "</strong> · " + name + "</div>"
            + "</article>\n";
    }

    /** Leader card showing a percentage value (gain/loss %) rather than a money amount. */
    private static String buildRealizedLeaderPctCard(String label, Security security, double pctValue) {
        if (security == null || pctValue == Double.NEGATIVE_INFINITY) {
            return "<article class=\"kpi-card\"><div class=\"kpi-label\">" + escapeHtml(label) + "</div><div class=\"kpi-value\">-</div></article>\n";
        }
        String ticker = escapeHtml(security.getTicker());
        String name = escapeHtml(security.getDisplayName());
        String cls = pctValue > 0 ? " positive" : pctValue < 0 ? " negative" : "";
        return "<article class=\"kpi-card\">"
            + "<div class=\"kpi-label\">" + escapeHtml(label) + "</div>"
            + "<div class=\"kpi-value" + cls + "\">" + HtmlFormatter.formatPercent(pctValue, 2) + "</div>"
            + "<div class=\"kpi-help\" style=\"overflow:hidden;text-overflow:ellipsis;white-space:nowrap;\"><strong>" + ticker + "</strong> · " + name + "</div>"
            + "</article>\n";
    }

    private static void writeRealizedSummaryTableHtml(FileWriter writer, TransactionStore store, Map<String, Double> ratesToNok) throws IOException {
        ArrayList<Security> soldSecurities = getSortedSoldSecurities(store);

        // Pre-pass: accumulate totals and track per-category leaders
        LinkedHashMap<String, Double> totalSalesValueBuckets = new LinkedHashMap<>();
        LinkedHashMap<String, Double> totalCostBasisBuckets = new LinkedHashMap<>();
        LinkedHashMap<String, Double> totalRealizedGainBuckets = new LinkedHashMap<>();
        LinkedHashMap<String, Double> totalRealizedDividendsBuckets = new LinkedHashMap<>();
        Security leadCostBasis = null, leadSalesValue = null, leadGain = null, leadGainPct = null, leadDividends = null, leadTotalReturn = null;
        double leadCostBasisNok = Double.NEGATIVE_INFINITY, leadSalesValueNok = Double.NEGATIVE_INFINITY;
        double leadGainNok = Double.NEGATIVE_INFINITY, leadGainPctVal = Double.NEGATIVE_INFINITY;
        double leadDividendsNok = Double.NEGATIVE_INFINITY, leadTotalReturnNok = Double.NEGATIVE_INFINITY;
        double leadCostBasisVal = 0, leadSalesValueVal = 0, leadGainVal = 0, leadDividendsVal = 0, leadTotalReturnVal = 0;
        for (Security security : soldSecurities) {
            String currency = security.getCurrencyCode();
            double costBasis = security.getRealizedCostBasis();
            double salesValue = security.getRealizedSalesValue();
            double gain = security.getRealizedGain();
            double realizedDividends = security.isFullyRealized() ? security.getDividends() : 0.0;
            double totalReturn = gain + realizedDividends;
            addToCurrencyBuckets(totalSalesValueBuckets, currency, salesValue);
            addToCurrencyBuckets(totalCostBasisBuckets, currency, costBasis);
            addToCurrencyBuckets(totalRealizedGainBuckets, currency, gain);
            addToCurrencyBuckets(totalRealizedDividendsBuckets, currency, realizedDividends);
            double costBasisNok = convertBucketsToTarget(singleCurrencyBuckets(currency, costBasis), DEFAULT_TOTAL_CURRENCY, ratesToNok);
            double salesValueNok = convertBucketsToTarget(singleCurrencyBuckets(currency, salesValue), DEFAULT_TOTAL_CURRENCY, ratesToNok);
            double gainNok = convertBucketsToTarget(singleCurrencyBuckets(currency, gain), DEFAULT_TOTAL_CURRENCY, ratesToNok);
            double dividendsNok = convertBucketsToTarget(singleCurrencyBuckets(currency, realizedDividends), DEFAULT_TOTAL_CURRENCY, ratesToNok);
            double totalReturnNok = convertBucketsToTarget(singleCurrencyBuckets(currency, totalReturn), DEFAULT_TOTAL_CURRENCY, ratesToNok);
            double secGainPct = costBasis > 0.0 ? (gain / costBasis) * 100.0 : Double.NEGATIVE_INFINITY;
            if (costBasisNok > leadCostBasisNok) { leadCostBasisNok = costBasisNok; leadCostBasis = security; leadCostBasisVal = costBasis; }
            if (salesValueNok > leadSalesValueNok) { leadSalesValueNok = salesValueNok; leadSalesValue = security; leadSalesValueVal = salesValue; }
            if (gainNok > leadGainNok) { leadGainNok = gainNok; leadGain = security; leadGainVal = gain; }
            if (secGainPct > leadGainPctVal) { leadGainPctVal = secGainPct; leadGainPct = security; }
            if (dividendsNok > leadDividendsNok) { leadDividendsNok = dividendsNok; leadDividends = security; leadDividendsVal = realizedDividends; }
            if (totalReturnNok > leadTotalReturnNok) { leadTotalReturnNok = totalReturnNok; leadTotalReturn = security; leadTotalReturnVal = totalReturn; }
        }
        double totalCostBasisForPct = convertBucketsToTarget(totalCostBasisBuckets, DEFAULT_TOTAL_CURRENCY, ratesToNok);
        double totalRealizedGainForPct = convertBucketsToTarget(totalRealizedGainBuckets, DEFAULT_TOTAL_CURRENCY, ratesToNok);
        double totalRealizedDividendsForPct = convertBucketsToTarget(totalRealizedDividendsBuckets, DEFAULT_TOTAL_CURRENCY, ratesToNok);
        LinkedHashMap<String, Double> totalRealizedReturnBuckets = sumCurrencyBuckets(totalRealizedGainBuckets, totalRealizedDividendsBuckets);
        double totalRealizedReturnForPct = totalRealizedGainForPct + totalRealizedDividendsForPct;
        double totalGainPct = totalCostBasisForPct > 0.0 ? (totalRealizedGainForPct / totalCostBasisForPct) * 100.0 : 0.0;
        double totalReturnPct = totalCostBasisForPct > 0
            ? (totalRealizedReturnForPct / totalCostBasisForPct) * 100.0
            : 0.0;
        String gainClass = signedClass(totalRealizedGainForPct);
        String dividendsClass = signedClass(totalRealizedDividendsForPct);
        String returnClass = signedClass(totalRealizedReturnForPct);

        writer.write("<h2>REALIZED OVERVIEW - ALL SALES</h2>\n");

        // Highlights grid
        writer.write("<section class=\"annual-kpi-deck\">\n");
        writer.write("<h2 class=\"annual-kpi-deck-title\">Realized Highlights</h2>\n");
        writer.write("<div class=\"realized-highlights\">\n");
        writer.write("<article class=\"kpi-card\"><div class=\"kpi-label\">Cost Basis</div>"
            + "<div class=\"kpi-value js-convert-money\" data-buckets=\"" + escapeHtml(toBucketsJson(totalCostBasisBuckets)) + "\" data-decimals=\"0\">"
            + formatBucketsInTarget(totalCostBasisBuckets, DEFAULT_TOTAL_CURRENCY, 0, ratesToNok)
            + "</div></article>\n");
        writer.write("<article class=\"kpi-card\"><div class=\"kpi-label\">Sales Value</div>"
            + "<div class=\"kpi-value js-convert-money\" data-buckets=\"" + escapeHtml(toBucketsJson(totalSalesValueBuckets)) + "\" data-decimals=\"0\">"
            + formatBucketsInTarget(totalSalesValueBuckets, DEFAULT_TOTAL_CURRENCY, 0, ratesToNok)
            + "</div></article>\n");
        writer.write("<article class=\"kpi-card\"><div class=\"kpi-label\">Gain / Loss</div>"
            + "<div class=\"kpi-value js-convert-money " + gainClass + "\" data-buckets=\"" + escapeHtml(toBucketsJson(totalRealizedGainBuckets)) + "\" data-decimals=\"0\">"
            + formatBucketsInTarget(totalRealizedGainBuckets, DEFAULT_TOTAL_CURRENCY, 0, ratesToNok)
            + "</div></article>\n");
        writer.write("<article class=\"kpi-card\"><div class=\"kpi-label\">Gain / Loss %</div>"
            + "<div class=\"kpi-value " + gainClass + "\">" + HtmlFormatter.formatPercent(totalGainPct, 2)
            + "</div></article>\n");
        writer.write("<article class=\"kpi-card\"><div class=\"kpi-label\">Dividends</div>"
            + "<div class=\"kpi-value js-convert-money " + dividendsClass + "\" data-buckets=\"" + escapeHtml(toBucketsJson(totalRealizedDividendsBuckets)) + "\" data-decimals=\"0\">"
            + formatBucketsInTarget(totalRealizedDividendsBuckets, DEFAULT_TOTAL_CURRENCY, 0, ratesToNok)
            + "</div></article>\n");
        writer.write("<article class=\"kpi-card\"><div class=\"kpi-label\">Total Return</div>"
            + "<div class=\"kpi-value js-convert-money " + returnClass + "\" data-buckets=\"" + escapeHtml(toBucketsJson(totalRealizedReturnBuckets)) + "\" data-decimals=\"0\">"
            + formatBucketsInTarget(totalRealizedReturnBuckets, DEFAULT_TOTAL_CURRENCY, 0, ratesToNok)
            + "</div><div class=\"kpi-label " + returnClass + "\">" + HtmlFormatter.formatPercent(totalReturnPct, 2) + "</div></article>\n");
        writer.write(buildRealizedLeaderCard("Highest Cost Basis", leadCostBasis, leadCostBasisVal, "", ratesToNok));
        writer.write(buildRealizedLeaderCard("Highest Sales Value", leadSalesValue, leadSalesValueVal, "", ratesToNok));
        writer.write(buildRealizedLeaderCard("Highest Gain / Loss", leadGain, leadGainVal, signedClass(leadGainNok), ratesToNok));
        writer.write(buildRealizedLeaderPctCard("Highest %", leadGainPct, leadGainPctVal));
        writer.write(buildRealizedLeaderCard("Highest Dividends", leadDividends, leadDividendsVal, "", ratesToNok));
        writer.write(buildRealizedLeaderCard("Highest Total Return", leadTotalReturn, leadTotalReturnVal, signedClass(leadTotalReturnNok), ratesToNok));
        writer.write("</div>\n");
        writer.write("</section>\n");

        // Table
        writer.write("<div class=\"overview-mode-shell\" role=\"group\" aria-label=\"Realized details controls\">\n");
        writer.write("<button type=\"button\" class=\"overview-mode-btn overview-details-toggle-btn\" data-detail-label=\"Open all details\" onclick=\"toggleDetailGroup('realized-details', this)\">Open all details ▸</button>\n");
        writer.write("</div>\n");
        writer.write("<div class=\"table-wrap\">\n<table class=\"realized-table\">\n");
        ReportTemplateHelper.writeHtmlRow(writer, true, "Ticker", "Security", "Cost Basis", "Sales Value", "Gain/Loss", "Dividends", "Total Return");

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

        writer.write("<tr class=\"total-row\">\n");
        writer.write("    <td></td><td><strong>TOTAL</strong></td>\n");
        writer.write("    <td>" + renderConvertibleMoneyCell(totalCostBasisBuckets, 2, ratesToNok) + "</td>\n");
        writer.write("    <td>" + renderConvertibleMoneyCell(totalSalesValueBuckets, 2, ratesToNok) + "</td>\n");
        writer.write("    <td>"
            + signedWrapHtml(renderConvertibleMoneyCell(totalRealizedGainBuckets, 2, ratesToNok), totalRealizedGainForPct)
            + " "
            + signedSpan("(" + HtmlFormatter.formatPercent(totalGainPct, 2) + ")", totalRealizedGainForPct)
            + "</td>\n");
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
            // Raw numeric values for BUY rows so JS can refresh unrealized on price update.
            // Zero for non-BUY entries (DIVIDEND etc.).
            final double rawLotUnits;
            final double rawLotCostBasis;

            DetailEntry(LocalDate date, int order, String type, String units, String price,
                        String amount, String unrealized, String unrealizedPct,
                        String unrealizedClass, String unrealizedPctClass,
                        double rawLotUnits, double rawLotCostBasis) {
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
                this.rawLotUnits = rawLotUnits;
                this.rawLotCostBasis = rawLotCostBasis;
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
                    unrealizedPctClass,
                    lot.getUnits(),   // raw lot units for JS refresh
                    lotCostBasis      // raw cost basis for JS refresh
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
                    "",
                    0.0, // not a lot row — JS won't update
                    0.0
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
            // BUY rows carry raw numeric attributes so JS can refresh unrealized on price update.
            boolean isBuyLot = entry.rawLotUnits > 0.0 && entry.rawLotCostBasis > 0.0;
            if (isBuyLot) {
                html.append(String.format(Locale.US,
                    "<tr data-lot-units=\"%.8f\" data-lot-cost-basis=\"%.8f\" data-overview-security-key=\"%s\" data-currency=\"%s\">",
                    entry.rawLotUnits, entry.rawLotCostBasis,
                    escapeHtml(row.securityKey), escapeHtml(row.currencyCode)));
            } else {
                html.append("<tr>");
            }
            html.append("<td>").append(escapeHtml(formatDetailDate(entry.date))).append("</td>");
            html.append("<td>").append(entry.type).append("</td>");
            html.append("<td>").append(escapeHtml(entry.units)).append("</td>");
            html.append("<td>").append(escapeHtml(entry.price)).append("</td>");
            html.append("<td>").append(escapeHtml(entry.amount)).append("</td>");
            // js-lot-unrealized / js-lot-unrealized-pct let refreshOpenDetailPanels() update these cells
            if (isBuyLot) {
                html.append("<td class=\"js-lot-unrealized").append(entry.unrealizedClass.isEmpty() ? "" : " " + escapeHtml(entry.unrealizedClass)).append("\">").append(escapeHtml(entry.unrealized)).append("</td>");
                html.append("<td class=\"js-lot-unrealized-pct").append(entry.unrealizedPctClass.isEmpty() ? "" : " " + escapeHtml(entry.unrealizedPctClass)).append("\">").append(escapeHtml(entry.unrealizedPct)).append("</td>");
            } else {
                html.append("<td class=\"").append(escapeHtml(entry.unrealizedClass)).append("\">").append(escapeHtml(entry.unrealized)).append("</td>");
                html.append("<td class=\"").append(escapeHtml(entry.unrealizedPctClass)).append("\">").append(escapeHtml(entry.unrealizedPct)).append("</td>");
            }
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


    private static String escapeHtml(String text) {
        if (text == null) return "";
        return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\"", "&quot;");
    }
}
