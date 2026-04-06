package report;

import csv.TransactionStore;
import model.Events;
import model.Security;

import java.io.FileWriter;
import java.io.IOException;
import java.time.LocalDate;
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
            writer.write("        :root { --bg:#eef3f7; --line:#d8e0e9; --card:#ffffff; --ink:#16202a; --muted:#5a6877; --good:#1f8b4d; --bad:#b23a31; --spark-text:#5c7187; --spark-axis:#9eb1c3; --spark-axis-soft:#b8c7d6; --spark-grid:#cfdbe6; --spark-line:#223c55; --spark-point:#223c55; }\n");
            writer.write("        body.theme-dark { --bg:#0f1722; --line:#253245; --card:#162231; --ink:#e5edf7; --muted:#aebdce; --good:#59c887; --bad:#f07f7f; --spark-text:#d5e1ef; --spark-axis:#7f95ab; --spark-axis-soft:#9ab0c6; --spark-grid:#8ea4ba; --spark-line:#edf4fc; --spark-point:#edf4fc; }\n");
            writer.write("        * { box-sizing: border-box; }\n");
            writer.write("        body { font-family: 'Segoe UI','Avenir Next','Helvetica Neue',Arial,sans-serif; margin:0; background: radial-gradient(circle at top,#f8fbfe 0%,var(--bg) 58%); color:var(--ink); }\n");
            writer.write("        body.theme-dark { background: radial-gradient(circle at top,#1c2b3f 0%, var(--bg) 62%); }\n");
            writer.write("        .page { width:100vw; max-width:none; margin:0; padding:24px 8px 32px; }\n");
            writer.write("        h2 { margin:26px 2px 12px; font-size:1.14rem; color:var(--ink); }\n");
            writer.write("        table { width:100%; border-collapse:collapse; min-width:0; table-layout:fixed; background:var(--card); }\n");
            writer.write("        th, td { padding:5px 5px; border-bottom:1px solid #edf2f7; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }\n");
            writer.write("        th { background:#f5f8fb; text-align:left; font-size:.72rem; text-transform:uppercase; letter-spacing:.2px; color:#374556; border-bottom:1px solid var(--line); }\n");
            writer.write("        td { font-size:.72rem; }\n");
            writer.write("        td.num, th.num { text-align:right; }\n");
            writer.write("        .table-wrap { background:var(--card); border:1px solid var(--line); border-radius:14px; overflow-x:auto; overflow-y:hidden; -webkit-overflow-scrolling:touch; scrollbar-gutter:stable both-edges; box-shadow:0 5px 14px rgba(15,23,33,.06); }\n");
            writer.write("        .table-wrap::-webkit-scrollbar { height:12px; }\n");
            writer.write("        .table-wrap::-webkit-scrollbar-track { background:#e7eef5; border-radius:999px; }\n");
            writer.write("        .table-wrap::-webkit-scrollbar-thumb { background:#9db0c3; border-radius:999px; border:2px solid #e7eef5; }\n");
            writer.write("        .overview-table { table-layout:auto; min-width:1020px; }\n");
            writer.write("        .overview-table tr > *:nth-child(1) { width:120px; max-width:120px; min-width:120px; overflow:hidden; text-overflow:ellipsis; }\n");
            writer.write("        .overview-table tr > *:nth-child(2) { width:auto; min-width:24ch; max-width:none; }\n");
            writer.write("        .overview-table tr > *:nth-child(3) { width:90px; min-width:90px; max-width:90px; }\n");
            writer.write("        .overview-table tr > *:nth-child(n+4) { min-width:13ch; white-space:nowrap; overflow:visible; text-overflow:clip; }\n");
            writer.write("        .realized-table { table-layout:auto; }\n");
            writer.write("        .realized-table tr > *:nth-child(1) { width:106px; max-width:106px; min-width:106px; overflow:visible; text-overflow:clip; }\n");
            writer.write("        .realized-table tr > *:nth-child(2) { width:auto; min-width:9ch; max-width:none; overflow:visible; text-overflow:clip; }\n");
            writer.write("        .realized-table tr > *:nth-child(3) { width:auto; min-width:14ch; max-width:none; overflow:visible; text-overflow:clip; }\n");
            writer.write("        .realized-table tr > *:nth-child(8) { width:160px; max-width:160px; }\n");
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
            writer.write("        .annual-headline-grid .kpi-card { min-height:116px; }\n");
            writer.write("        .kpi-label { color:#c8d9eb; font-size:.8rem; text-transform:uppercase; }\n");
            writer.write("        .kpi-value { margin-top:2px; font-size:1.02rem; font-weight:700; color:#fff; }\n");
            writer.write("        .performer { margin-top:6px; font-size:.84rem; color:#dce8f3; }\n");
            writer.write("        .performer strong { display:block; font-size:.9rem; margin-bottom:2px; }\n");
            writer.write("        .performer-metrics { display:block; }\n");
            writer.write("        .hero-side { position:relative; background:rgba(255,255,255,.06); border:1px solid rgba(235,245,255,.22); border-radius:12px; padding:10px; min-height:172px; }\n");
            writer.write("        .timeline-title-row { display:flex; align-items:center; gap:6px; margin-bottom:8px; }\n");
            writer.write("        .hero-side-title { color:#d4e3f0; font-size:.86rem; text-transform:uppercase; margin:0; }\n");
            writer.write("        .timeline-info-btn { width:18px; height:18px; border-radius:999px; border:1px solid rgba(235,245,255,.55); background:rgba(255,255,255,.14); color:#e8f2fb; font-size:.72rem; font-weight:800; line-height:1; cursor:pointer; display:inline-flex; align-items:center; justify-content:center; padding:0; }\n");
            writer.write("        .timeline-info-btn:hover { background:rgba(255,255,255,.24); }\n");
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
            writer.write("        .sparkline-range-btn { border:1px solid #b7c7d7; background:#f2f7fc; color:#27415a; border-radius:999px; padding:3px 9px; font-size:.72rem; font-weight:700; letter-spacing:.2px; cursor:pointer; }\n");
            writer.write("        .sparkline-range-btn:hover { background:#e8f0f8; }\n");
            writer.write("        .sparkline-range-btn.is-active { background:#dbe9f8; color:#1f3f5b; border-color:#9eb9d5; }\n");
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
            writer.write("        .annual-kpi-deck-title { margin:0 0 9px; font-size:.9rem; font-weight:700; letter-spacing:.12px; color:var(--muted); text-transform:uppercase; }\n");
            writer.write("        .annual-summary-grid { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:10px; }\n");
            writer.write("        .annual-summary-card { border:1px solid #d4dfeb; border-radius:11px; padding:11px; background:linear-gradient(180deg,#f9fcff 0%,#f2f8fd 100%); }\n");
            writer.write("        .annual-summary-card h4 { margin:0 0 4px; font-size:.82rem; color:#40576c; text-transform:uppercase; }\n");
            writer.write("        .annual-summary-value { font-size:1.05rem; font-weight:700; }\n");
            writer.write("        .annual-summary-sub { margin-top:4px; font-size:.78rem; color:#5f7488; }\n");
            writer.write("        .annual-value-warning { margin-top:6px; padding:6px 7px; font-size:.74rem; line-height:1.35; border:1px solid #f0d8a8; border-radius:8px; background:#fff5df; color:#7b4a00; }\n");
            writer.write("        .annual-summary-card .performer { margin-top:5px; color:#253d53; font-size:.82rem; }\n");
            writer.write("        .annual-summary-card .performer strong { margin-bottom:1px; font-size:.88rem; color:#1f3345; }\n");
            writer.write("        .annual-graphs-section { margin:0 0 18px; padding:12px; border:1px solid var(--line); border-radius:14px; background:var(--card); box-shadow:0 5px 14px rgba(15,23,33,.06); }\n");
            writer.write("        .annual-graphs-heading { display:flex; flex-wrap:wrap; align-items:baseline; justify-content:space-between; gap:8px; margin:0 0 10px; }\n");
            writer.write("        .annual-graphs-heading h2 { margin:0; font-size:1.02rem; color:var(--ink); }\n");
            writer.write("        .annual-graphs-heading p { margin:0; font-size:.8rem; color:var(--muted); }\n");
            writer.write("        .annual-graphs-row { display:grid; grid-template-columns:1fr 1fr; gap:14px; margin:0; }\n");
            writer.write("        .annual-graph-card { display:flex; flex-direction:column; min-height:332px; padding:14px; border:1px solid #d4dfeb; border-radius:13px; background:linear-gradient(180deg,#f9fcff 0%,#f2f8fd 100%); box-shadow:0 2px 8px rgba(19,35,51,.06); overflow:hidden; }\n");
            writer.write("        .annual-graph-card h3 { margin:0 0 8px; font-size:1rem; color:#1e3448; }\n");
            writer.write("        .annual-graph-note { margin:0 0 10px; font-size:.78rem; color:#5f7488; }\n");
            writer.write("        .annual-graph-content { flex:1; display:flex; flex-direction:column; justify-content:flex-end; min-height:0; }\n");
            writer.write("        .annual-graph-content > svg { display:block; width:100%; margin-top:auto; }\n");
            writer.write("        .annual-graph-content .sparkline-widget { display:flex; flex-direction:column; min-height:100%; }\n");
            writer.write("        .annual-graph-content .sparkline-panel { flex:1; }\n");
            writer.write("        .annual-graph-content .sparkline-panel.is-active { display:flex; align-items:flex-end; }\n");
            writer.write("        .annual-graph-content .sparkline-panel > svg { display:block; width:100%; }\n");
            writer.write("        .allocation-visuals { display:grid; gap:10px; }\n");
            writer.write("        .allocation-row { display:grid; gap:10px; }\n");
            writer.write("        .allocation-row-top { grid-template-columns:repeat(3,minmax(0,1fr)); }\n");
            writer.write("        .allocation-row-bottom { grid-template-columns:repeat(2,minmax(0,1fr)); }\n");
            writer.write("        .allocation-panel { border:1px solid var(--line); border-radius:10px; padding:16px; background:#fafcfe; overflow:hidden; }\n");
            writer.write("        .allocation-panel-title { margin:0 0 6px; font-size:.84rem; text-transform:uppercase; color:#41576d; letter-spacing:.3px; white-space:nowrap; }\n");
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
            writer.write("        body.theme-dark .annual-kpi-deck-title { color:#bdd1e4; }\n");
            writer.write("        body.theme-dark .annual-summary-card { border-color:#2e4258; background:#1a2d42; }\n");
            writer.write("        body.theme-dark .annual-summary-card h4 { color:#c8d9eb; }\n");
            writer.write("        body.theme-dark .annual-summary-sub { color:#d6e4f1; }\n");
            writer.write("        body.theme-dark .annual-value-warning { background:#3d2e19; border-color:#8e6a33; color:#ffdca8; }\n");
            writer.write("        body.theme-dark .annual-summary-card .performer { color:#d4e3f2; }\n");
            writer.write("        body.theme-dark .annual-summary-card .performer strong { color:#e9f2fc; }\n");
            writer.write("        body.theme-dark .annual-graphs-heading h2 { color:#e5edf7; }\n");
            writer.write("        body.theme-dark .annual-graphs-heading p { color:#bad0e5; }\n");
            writer.write("        body.theme-dark .annual-graph-card { border-color:#2e4258; background:#1a2d42; box-shadow:none; }\n");
            writer.write("        body.theme-dark .annual-graph-card h3 { color:#e5edf7; }\n");
            writer.write("        body.theme-dark .annual-graph-note { color:#bad0e5; }\n");
            writer.write("        body.theme-dark .sparkline-metric-btn, body.theme-dark .sparkline-range-btn { border-color:#45627f; background:#22374d; color:#cfe0f2; }\n");
            writer.write("        body.theme-dark .sparkline-metric-btn:hover, body.theme-dark .sparkline-range-btn:hover { background:#2b4560; }\n");
            writer.write("        body.theme-dark .sparkline-metric-btn.is-active { background:#dceafb; color:#173047; border-color:#dceafb; }\n");
            writer.write("        body.theme-dark .sparkline-range-btn.is-active { background:#c3d7ed; color:#173047; border-color:#9bb7d3; }\n");
            writer.write("        body.theme-dark .chart-title-row > h3, body.theme-dark .chart-title-row > h4, body.theme-dark .chart-title-row > .hero-side-title { color:#dce8f5; }\n");
            writer.write("        body.theme-dark .chart-svg { background:#162231; border-color:#2b3a4d; }\n");
            writer.write("        body.theme-dark .chart-svg text { fill:#d4e1ee !important; }\n");
            writer.write("        body.theme-dark .chart-total-return-label, body.theme-dark .chart-security-bar-label { fill:#e7f0fa !important; stroke:#0b1624 !important; stroke-width:1.0; }\n");
            writer.write("        body.theme-dark .chart-security-label { fill:#e7f0fa !important; stroke:#0b1624 !important; stroke-width:1.05; }\n");
            writer.write("        body.theme-dark .market-value-bar-chart line[stroke='#495057'] { stroke:#dce8f4 !important; }\n");
            writer.write("        body.theme-dark .market-value-bar-chart text[fill='#495057'] { fill:#dce8f4 !important; }\n");
            writer.write("        body.theme-dark .app-shell-note, body.theme-dark .hero-side-note { color:#d1e0ef; }\n");
            writer.write("        .expand-btn { border:1px solid #86a4bf; background:#f3f8fd; color:#1e3951; border-radius:7px; min-width:62px; padding:2px 8px; font-size:.66rem; font-weight:700; cursor:pointer; text-align:center; }\n");
            writer.write("        .expand-btn:hover { background:#e6f1fb; }\n");
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
            writer.write("        @media (max-width:1060px) { .report-hero{grid-template-columns:1fr;} .hero-kpis,.annual-headline-grid{grid-template-columns:1fr;} .annual-summary-grid{grid-template-columns:repeat(2,minmax(0,1fr));} .annual-graphs-row{grid-template-columns:1fr;} .overview-charts{grid-template-columns:1fr;} .allocation-row-top,.allocation-row-bottom{grid-template-columns:1fr;} .page{width:100vw; padding:16px 8px 22px;} .table-wrap{overflow-x:auto;} .table-wrap table{min-width:980px;} .ticker-scroll,.security-scroll{max-width:none;} }\n");
            writer.write("        @media (max-width:760px) { .annual-summary-grid{grid-template-columns:1fr;} .annual-graphs-heading{flex-direction:column; align-items:flex-start;} }\n");
            writer.write("    </style>\n");
            writer.write("</head>\n");
            writer.write("<body>\n");
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
            writer.write("function toggleOverviewDetails(rowId, button) {\n");
            writer.write("  var row = document.getElementById(rowId);\n");
            writer.write("  if (!row) return;\n");
            writer.write("  var isOpen = row.style.display === 'table-row';\n");
            writer.write("  row.style.display = isOpen ? 'none' : 'table-row';\n");
            writer.write("  if (button) button.textContent = isOpen ? 'Show' : 'Hide';\n");
            writer.write("}\n");
            writer.write("function toggleDetailGroup(groupName, button) {\n");
            writer.write("  var rows = document.querySelectorAll('tr.details-row[data-group=\\\"' + groupName + '\\\"]');\n");
            writer.write("  if (!rows.length) return;\n");
            writer.write("  window.__detailGroupNextAction = window.__detailGroupNextAction || {};\n");
            writer.write("  var action = window.__detailGroupNextAction[groupName] || 'open';\n");
            writer.write("  var open = action === 'open';\n");
            writer.write("  rows.forEach(function(row) {\n");
            writer.write("    row.style.display = open ? 'table-row' : 'none';\n");
            writer.write("    var rowId = row.id;\n");
            writer.write("    if (!rowId) return;\n");
            writer.write("    var rowButton = document.querySelector('button.expand-btn[data-target=\\\"' + rowId + '\\\"]');\n");
            writer.write("    if (rowButton) rowButton.textContent = open ? 'Hide' : 'Show';\n");
            writer.write("  });\n");
            writer.write("  window.__detailGroupNextAction[groupName] = open ? 'close' : 'open';\n");
            writer.write("  if (button) button.textContent = open ? '▾' : '▸';\n");
            writer.write("}\n");
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

        String portfolioClass = summary.portfolioReturnNok >= 0 ? "positive" : "negative";
        double benchmarkDelta = summary.hasBenchmarkData ? (summary.portfolioReturnPct - summary.benchmarkReturnPct) : 0.0;
        String deltaClass = benchmarkDelta >= 0 ? "positive" : "negative";
        String bestClass = metrics.best.returnNok >= 0.0 ? "positive" : "negative";
        String worstClass = metrics.worst.returnNok >= 0.0 ? "positive" : "negative";
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

        String realizedGainClass = summary.realizedGainNok >= 0 ? "positive" : "negative";
        writer.write("<article class=\"kpi-card annual-summary-card\"><h4>Realized Gain/Loss</h4><div class=\"annual-summary-value " + realizedGainClass + "\">"
            + renderConvertibleMoneyCell(realizedGainBuckets, 2, ratesToNok)
            + "</div><div class=\"annual-summary-sub\">Closed sales in selected year</div></article>\n");

        writer.write("<article class=\"kpi-card annual-summary-card\"><h4>Dividends</h4><div class=\"annual-summary-value\">"
            + renderConvertibleMoneyCell(dividendsBuckets, 2, ratesToNok)
            + "</div><div class=\"annual-summary-sub\">Dividend cash flows in selected year</div></article>\n");

        String realizedTotalClass = summary.realizedTotalNok >= 0 ? "positive" : "negative";
        writer.write("<article class=\"kpi-card annual-summary-card\"><h4>Total Realized</h4><div class=\"annual-summary-value " + realizedTotalClass + "\">"
            + renderConvertibleMoneyCell(realizedTotalBuckets, 2, ratesToNok)
            + "</div><div class=\"annual-summary-sub\">Realized gain/loss plus dividends</div></article>\n");

        if (summary.hasBenchmarkData) {
            writer.write("<article class=\"kpi-card annual-summary-card\"><h4>Benchmark (" + escapeHtml(summary.benchmarkTicker) + ")</h4><div class=\"annual-summary-value\">"
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
        writeHtmlRow(writer, true,
                "Ticker", "Security", "Units", "Avg Cost", "Price", "Cost Basis", "Market Value", "Unrealized");

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
                    ? HtmlFormatter.formatMoney(row.unrealized, row.currencyCode, 2) + " (" + HtmlFormatter.formatPercent(row.unrealizedPct, 2) + ")"
                    : "-";

            writeHtmlRowWithClass(writer, rowClass,
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
        writer.write("    <td>" + renderConvertibleMoneyCell(totalUnrealizedBuckets, 2, ratesToNok) + " (" + HtmlFormatter.formatPercent(totalUnrealizedPct, 2) + ")</td>\n");
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
        writeHtmlRow(writer, true, buildDetailsHeaderCell("realized-details-year"), "Ticker", "Security", "Cost Basis", "Sales Value", "Gain/Loss", "Dividends", "Total Return");

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
            String totalReturnCombined = HtmlFormatter.formatMoney(totalReturnValue, currency, 2)
                    + " (" + HtmlFormatter.formatPercent(rowTotalReturnPct, 2) + ")";

            addToCurrencyBuckets(totalSalesValueBuckets, currency, salesValue);
            addToCurrencyBuckets(totalCostBasisBuckets, currency, costBasis);
            addToCurrencyBuckets(totalRealizedGainBuckets, currency, gain);
            addToCurrencyBuckets(totalRealizedDividendsBuckets, currency, realizedDividends);

            String detailsRowId = "realized-year-details-" + detailsIndex;
            writeHtmlRowWithClass(writer, rowClass,
                    "<button class=\"expand-btn\" data-target=\"" + detailsRowId + "\" onclick=\"toggleOverviewDetails('" + detailsRowId + "', this)\">Show</button>",
                    "<span class=\"ticker-scroll\">" + escapeHtml(security.getTicker()) + "</span>",
                    "<span class=\"security-scroll\">" + escapeHtml(security.getDisplayName()) + "</span>",
                    HtmlFormatter.formatMoney(costBasis, currency, 2),
                    HtmlFormatter.formatMoney(salesValue, currency, 2),
                    HtmlFormatter.formatMoney(gain, currency, 2),
                    HtmlFormatter.formatMoney(realizedDividends, currency, 2),
                    totalReturnCombined);

            writer.write("<tr id=\"" + detailsRowId + "\" class=\"details-row\" data-group=\"realized-details-year\">\n");
            writer.write("    <td class=\"details-cell\" colspan=\"8\">\n");
            writer.write(buildRealizedSaleTradesDetailsHtml(security, safeYear));
            writer.write("    </td>\n");
            writer.write("</tr>\n");

            previousAssetType = currentAssetType;
            detailsIndex++;
            includedRows++;
        }

        if (includedRows == 0) {
            writer.write("<tr><td colspan=\"8\" class=\"app-shell-note\">No sales or dividends were recorded for " + safeYear + ".</td></tr>\n");
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
        writer.write("    <td></td><td></td><td><strong>TOTAL</strong></td>\n");
        writer.write("    <td>" + renderConvertibleMoneyCell(totalCostBasisBuckets, 2, ratesToNok) + "</td>\n");
        writer.write("    <td>" + renderConvertibleMoneyCell(totalSalesValueBuckets, 2, ratesToNok) + "</td>\n");
        writer.write("    <td>" + renderConvertibleMoneyCell(totalRealizedGainBuckets, 2, ratesToNok) + "</td>\n");
        writer.write("    <td>" + renderConvertibleMoneyCell(totalRealizedDividendsBuckets, 2, ratesToNok) + "</td>\n");
        writer.write("    <td>" + renderConvertibleMoneyCell(totalRealizedReturnBuckets, 2, ratesToNok) + " (" + HtmlFormatter.formatPercent(totalReturnPct, 2) + ")</td>\n");
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
        String bestClass = s.bestReturn >= 0 ? "positive" : "negative";
        String worstClass = s.worstReturn >= 0 ? "positive" : "negative";

        writer.write("<section class=\"report-hero\">\n");
        writer.write("<div class=\"hero-title\">\n");
        writer.write("<h1>Portfolio Report</h1>\n");
        writer.write("<div class=\"hero-meta\">\n");
        writer.write("<span class=\"meta-chip\">Date: <strong>" + escapeHtml(s.generatedDate) + "</strong></span>\n");
        writer.write("<span class=\"meta-chip\">Files: <strong>" + s.fileCount + "</strong></span>\n");
        writer.write("<span class=\"meta-chip\">Transactions: <strong>" + s.transactionCount + "</strong></span>\n");
        writer.write("<span class=\"meta-chip\">Holdings: <strong>" + s.holdingsCount + "</strong></span>\n");
        writer.write("<span class=\"meta-chip\">Currency: <strong><input id=\"portfolio-currency-input\" class=\"currency-input\" type=\"text\" value=\"" + DEFAULT_TOTAL_CURRENCY + "\" maxlength=\"3\" autocomplete=\"off\" spellcheck=\"false\" title=\"Skriv valutakode (f.eks. NOK, USD) og trykk Enter\"></strong></span>\n");
        writer.write("<button id=\"report-theme-toggle\" class=\"hero-theme-btn\" type=\"button\">Dark mode</button>\n");
        writer.write("</div>\n");
        writer.write("<div class=\"hero-kpis\">\n");
        LinkedHashMap<String, Double> totalMarketBuckets = new LinkedHashMap<>();
        LinkedHashMap<String, Double> totalReturnBuckets = new LinkedHashMap<>();
        LinkedHashMap<String, Double> totalHistoricalCostBuckets = new LinkedHashMap<>();
        Set<String> activeSecurityKeys = new LinkedHashSet<>();
        for (OverviewRow row : overviewRows) {
            activeSecurityKeys.add(row.securityKey);
            addToCurrencyBuckets(totalMarketBuckets, row.currencyCode, row.marketValue);
            addToCurrencyBuckets(totalReturnBuckets, row.currencyCode, row.totalReturn);
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
            addToCurrencyBuckets(totalHistoricalCostBuckets, currency, realizedCostBasis);
        }

        LinkedHashMap<String, Double> cashBuckets = new LinkedHashMap<>();
        cashBuckets.put(DEFAULT_TOTAL_CURRENCY, store.getCurrentCashHoldings());

        double totalReturnInDefaultCurrency = convertBucketsToTarget(totalReturnBuckets, DEFAULT_TOTAL_CURRENCY, ratesToNok);
        double totalHistoricalCostInDefaultCurrency = convertBucketsToTarget(totalHistoricalCostBuckets, DEFAULT_TOTAL_CURRENCY, ratesToNok);
        double totalReturnPct = totalHistoricalCostInDefaultCurrency > 0.0
            ? (totalReturnInDefaultCurrency / totalHistoricalCostInDefaultCurrency) * 100.0
            : 0.0;
        String totalClass = totalReturnInDefaultCurrency >= 0 ? "positive" : "negative";

        writer.write("<article class=\"kpi-card\"><div class=\"kpi-label\">Total Market Value</div><div class=\"kpi-value js-convert-money\" data-buckets=\""
            + escapeHtml(toBucketsJson(totalMarketBuckets)) + "\" data-decimals=\"0\">"
            + formatBucketsInTarget(totalMarketBuckets, DEFAULT_TOTAL_CURRENCY, 0, ratesToNok) + "</div></article>\n");

        writer.write("<article class=\"kpi-card\"><div class=\"kpi-label\">Cash Holdings</div><div class=\"kpi-value js-convert-money\" data-buckets=\""
            + escapeHtml(toBucketsJson(cashBuckets)) + "\" data-decimals=\"0\">"
            + formatBucketsInTarget(cashBuckets, DEFAULT_TOTAL_CURRENCY, 0, ratesToNok) + "</div></article>\n");

        writer.write("<article class=\"kpi-card\"><div class=\"kpi-label\">Total Return</div><div class=\"kpi-value js-convert-money " + totalClass
            + "\" data-buckets=\"" + escapeHtml(toBucketsJson(totalReturnBuckets)) + "\" data-decimals=\"0\">"
            + formatBucketsInTarget(totalReturnBuckets, DEFAULT_TOTAL_CURRENCY, 0, ratesToNok)
            + "</div><div class=\"kpi-label " + totalClass + "\">" + HtmlFormatter.formatPercent(totalReturnPct) + "</div></article>\n");

        writer.write("<article class=\"kpi-card\"><div class=\"kpi-label\">Best / Worst</div><div class=\"performer " + bestClass + "\"><strong>"
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
        writer.write("</div></div>\n");
        writer.write("<aside class=\"hero-side\"><div class=\"timeline-title-row\"><div class=\"hero-side-title\">Portfolio Value Timeline</div><button type=\"button\" class=\"timeline-info-btn\" aria-label=\"Show calculation info\" title=\"Show calculation info\">i</button></div>");
        if (s.sparklineSvg != null && !s.sparklineSvg.isBlank()) {
            writer.write(s.sparklineSvg);
        } else {
            writer.write("<div class=\"hero-side-note\">Timeline data not available yet for this dataset.</div>");
        }
        writer.write("<div class=\"timeline-info-overlay\" hidden><div class=\"timeline-info-dialog\" role=\"dialog\" aria-modal=\"true\" aria-label=\"Portfolio timeline info\"><div class=\"timeline-info-header\"><h4>Portfolio Value Timeline - Info</h4><button type=\"button\" class=\"timeline-info-close\" aria-label=\"Close\">×</button></div><div class=\"timeline-info-body\"><p>This chart is an indicative estimate based on imported transactions, cash snapshots, and historical prices.</p><ul><li><strong>Value:</strong> Estimated portfolio value at each month-end in the selected display currency.</li><li><strong>Return (NOK):</strong> Cumulative cashflow-adjusted return (TWR-based) for the selected range, expressed in NOK from the range start value.</li><li><strong>Return (%):</strong> Cumulative time-weighted return (TWR) from the selected range start.</li><li><strong>External cash flows:</strong> Deposits, withdrawals, and transfers are neutralized in return calculations so contributions/withdrawals do not count as performance.</li><li><strong>Pricing:</strong> Historical close prices are primarily fetched from Yahoo Finance. If data points are missing, transaction-derived fallback pricing is used.</li><li><strong>Disclaimer:</strong> Values are for analysis and may differ from official broker reporting.</li></ul></div></div></div>");
        writer.write("</aside>\n");
        writer.write("</section>\n");
    }

    private static void writeOverviewTableHtml(FileWriter writer, List<OverviewRow> rows, TransactionStore store, Map<String, Double> ratesToNok) throws IOException {
        writer.write("<h2>PORTFOLIO OVERVIEW - CURRENT HOLDINGS</h2>\n");
        Map<String, Security> securityByKey = buildSecurityLookupByKey(store);

        writer.write("<div class=\"overview-charts\">\n");
        writer.write("<section class=\"overview-chart total-return-chart\"><h3 class=\"js-total-return-money-title\">Total Return (" + DEFAULT_TOTAL_CURRENCY + ")</h3>\n");
        writer.write(ChartBuilder.buildOverviewBarChartSvg(rows, false, ratesToNok));
        writer.write("</section>\n");
        writer.write("<section class=\"overview-chart total-return-chart\"><h3>Total Return (%)</h3>\n");
        writer.write(ChartBuilder.buildOverviewBarChartSvg(rows, true, ratesToNok));
        writer.write("</section>\n");
        writer.write("</div>\n");

        writer.write("<div class=\"table-wrap\">\n<table class=\"overview-table\">\n");
        writeHtmlRow(writer, true,
            buildDetailsHeaderCell("overview-details"), "Ticker", "Security", "Units", "Avg Cost", "Last Price",
                "Cost Basis", "Market Value", "Unrealized", "Realized", "Dividends", "Total Return");

        LinkedHashMap<String, Double> totalMarketValueBuckets = new LinkedHashMap<>();
        LinkedHashMap<String, Double> totalCostBasisBuckets = new LinkedHashMap<>();
        LinkedHashMap<String, Double> totalUnrealizedBuckets = new LinkedHashMap<>();
        LinkedHashMap<String, Double> totalRealizedBuckets = new LinkedHashMap<>();
        LinkedHashMap<String, Double> totalDividendsBuckets = new LinkedHashMap<>();
        LinkedHashMap<String, Double> totalHistoricalCostBuckets = new LinkedHashMap<>();
        String previousAssetType = null;

        int detailsIndex = 0;
        for (OverviewRow row : rows) {
            addToCurrencyBuckets(totalMarketValueBuckets, row.currencyCode, row.marketValue);
            addToCurrencyBuckets(totalCostBasisBuckets, row.currencyCode, row.positionCostBasis);
            addToCurrencyBuckets(totalUnrealizedBuckets, row.currencyCode, row.unrealized);
            addToCurrencyBuckets(totalRealizedBuckets, row.currencyCode, row.realized);
            addToCurrencyBuckets(totalDividendsBuckets, row.currencyCode, row.dividends);
            addToCurrencyBuckets(totalHistoricalCostBuckets, row.currencyCode, row.historicalCostBasis);
            String rowClass = isStockFundBoundary(previousAssetType, row.assetType) ? "asset-split" : null;
            String detailsRowId = "overview-details-" + detailsIndex;
            Security security = securityByKey.get(row.securityKey);
            String unrealizedCombined = row.hasPrice
                ? HtmlFormatter.formatMoney(row.unrealized, row.currencyCode, 2) + " (" + HtmlFormatter.formatPercent(row.unrealizedPct, 2) + ")"
                : "-";
            String realizedCombined = HtmlFormatter.formatMoney(row.realized, row.currencyCode, 2)
                + " (" + row.realizedReturnPctText + "%)";
            String totalReturnCombined = HtmlFormatter.formatMoney(row.totalReturn, row.currencyCode, 2)
                + " (" + HtmlFormatter.formatPercent(row.totalReturnPct, 2) + ")";

            writeHtmlRowWithClass(writer, rowClass,
                "<button class=\"expand-btn\" data-target=\"" + detailsRowId + "\" onclick=\"toggleOverviewDetails('" + detailsRowId + "', this)\">Show</button>",
                    "<span class=\"ticker-scroll\">" + escapeHtml(row.tickerText) + "</span>",
                    "<span class=\"security-scroll\">" + escapeHtml(row.securityDisplayName) + "</span>",
                    HtmlFormatter.formatUnits(row.units),
                    HtmlFormatter.formatMoney(row.averageCost, row.currencyCode, 2),
                    row.latestPrice > 0 ? HtmlFormatter.formatMoney(row.latestPrice, row.currencyCode, 2) : "-",
                    HtmlFormatter.formatMoney(row.positionCostBasis, row.currencyCode, 2),
                    row.latestPrice > 0 ? HtmlFormatter.formatMoney(row.marketValue, row.currencyCode, 2) : "-",
                    unrealizedCombined,
                    realizedCombined,
                    HtmlFormatter.formatMoney(row.dividends, row.currencyCode, 2),
                    totalReturnCombined);

                    writer.write("<tr id=\"" + detailsRowId + "\" class=\"details-row\" data-group=\"overview-details\">\n");
                    writer.write("    <td class=\"details-cell\" colspan=\"12\">\n");
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

        writer.write("<tr class=\"total-row\">\n");
        writer.write("    <td></td><td></td><td><strong>TOTAL</strong></td><td></td><td></td><td></td>\n");
        writer.write("    <td>" + renderConvertibleMoneyCell(totalCostBasisBuckets, 2, ratesToNok) + "</td>\n");
        writer.write("    <td>" + renderConvertibleMoneyCell(totalMarketValueBuckets, 2, ratesToNok) + "</td>\n");
        writer.write("    <td>" + renderConvertibleMoneyCell(totalUnrealizedBuckets, 2, ratesToNok) + " (" + HtmlFormatter.formatPercent(totalUnrealizedPct, 2) + ")</td>\n");
        writer.write("    <td>" + renderConvertibleMoneyCell(totalRealizedBuckets, 2, ratesToNok) + " (" + HtmlFormatter.formatPercent(totalRealizedPct, 2) + ")</td>\n");
        writer.write("    <td>" + renderConvertibleMoneyCell(totalDividendsBuckets, 2, ratesToNok) + "</td>\n");
        writer.write("    <td>" + renderConvertibleMoneyCell(totalReturnBuckets, 2, ratesToNok) + " (" + HtmlFormatter.formatPercent(totalReturnPct, 2) + ")</td>\n");
        writer.write("</tr>\n");

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

    private static void writeRealizedSummaryTableHtml(FileWriter writer, TransactionStore store, Map<String, Double> ratesToNok) throws IOException {
        writer.write("<h2>REALIZED OVERVIEW - ALL SALES</h2>\n");
        writer.write("<div class=\"table-wrap\">\n<table class=\"realized-table\">\n");
        writeHtmlRow(writer, true, buildDetailsHeaderCell("realized-details"), "Ticker", "Security", "Cost Basis", "Sales Value", "Gain/Loss", "Dividends", "Total Return");

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
            String totalReturnCombined = HtmlFormatter.formatMoney(totalReturnValue, currency, 2)
                + " (" + HtmlFormatter.formatPercent(rowTotalReturnPct, 2) + ")";

            addToCurrencyBuckets(totalSalesValueBuckets, currency, salesValue);
            addToCurrencyBuckets(totalCostBasisBuckets, currency, costBasis);
            addToCurrencyBuckets(totalRealizedGainBuckets, currency, gain);
            addToCurrencyBuckets(totalRealizedDividendsBuckets, currency, realizedDividends);
                String detailsRowId = "realized-details-" + detailsIndex;

                writeHtmlRowWithClass(writer, rowClass,
                    "<button class=\"expand-btn\" data-target=\"" + detailsRowId + "\" onclick=\"toggleOverviewDetails('" + detailsRowId + "', this)\">Show</button>",
                    "<span class=\"ticker-scroll\">" + escapeHtml(security.getTicker()) + "</span>",
                    "<span class=\"security-scroll\">" + escapeHtml(security.getDisplayName()) + "</span>",
                    HtmlFormatter.formatMoney(costBasis, currency, 2),
                    HtmlFormatter.formatMoney(salesValue, currency, 2),
                    HtmlFormatter.formatMoney(gain, currency, 2),
                    HtmlFormatter.formatMoney(realizedDividends, currency, 2),
                    totalReturnCombined);

                writer.write("<tr id=\"" + detailsRowId + "\" class=\"details-row\" data-group=\"realized-details\">\n");
                writer.write("    <td class=\"details-cell\" colspan=\"8\">\n");
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
        writer.write("    <td></td><td></td><td><strong>TOTAL</strong></td>\n");
        writer.write("    <td>" + renderConvertibleMoneyCell(totalCostBasisBuckets, 2, ratesToNok) + "</td>\n");
        writer.write("    <td>" + renderConvertibleMoneyCell(totalSalesValueBuckets, 2, ratesToNok) + "</td>\n");
        writer.write("    <td>" + renderConvertibleMoneyCell(totalRealizedGainBuckets, 2, ratesToNok) + "</td>\n");
        writer.write("    <td>" + renderConvertibleMoneyCell(totalRealizedDividendsBuckets, 2, ratesToNok) + "</td>\n");
        writer.write("    <td>" + renderConvertibleMoneyCell(totalRealizedReturnBuckets, 2, ratesToNok) + " (" + HtmlFormatter.formatPercent(totalReturnPct, 2) + ")</td>\n");
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
            html.append("<tr><th>Sale Date</th><th>Units</th><th>Price/Unit</th><th>Cost Basis</th><th>Sale Value</th><th>Gain/Loss</th><th>Return (%)</th></tr>\n");

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
                html.append("<td>").append(escapeHtml(HtmlFormatter.formatMoney(trade.getGainLoss(), currency, 0))).append("</td>");
                html.append("<td>").append(escapeHtml(HtmlFormatter.formatPercent(trade.getReturnPct(), 2))).append("</td>");
                html.append("</tr>\n");
            }

            double totalReturnPct = totalCostBasis > 0.0 ? (totalGainLoss / totalCostBasis) * 100.0 : 0.0;
            html.append("<tr class=\"total-row\">");
            html.append("<td><strong>TOTAL</strong></td>");
            html.append("<td>").append(escapeHtml(HtmlFormatter.formatUnits(totalUnits))).append("</td>");
            html.append("<td></td>");
            html.append("<td>").append(escapeHtml(HtmlFormatter.formatMoney(totalCostBasis, currency, 0))).append("</td>");
            html.append("<td>").append(escapeHtml(HtmlFormatter.formatMoney(totalSaleValue, currency, 0))).append("</td>");
            html.append("<td>").append(escapeHtml(HtmlFormatter.formatMoney(totalGainLoss, currency, 0))).append("</td>");
            html.append("<td>").append(escapeHtml(HtmlFormatter.formatPercent(totalReturnPct, 2))).append("</td>");
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
            html.append("<tr><th>Date</th><th>Units</th><th>Dividend</th></tr>\n");
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

            DetailEntry(LocalDate date, int order, String type, String units, String price,
                        String amount, String unrealized, String unrealizedPct) {
                this.date = date;
                this.order = order;
                this.type = type;
                this.units = units;
                this.price = price;
                this.amount = amount;
                this.unrealized = unrealized;
                this.unrealizedPct = unrealizedPct;
            }
        }

        ArrayList<DetailEntry> entries = new ArrayList<>();
        for (Security.CurrentHoldingLot lot : security.getCurrentHoldingLotsSortedByDate()) {
            double lotCostBasis = lot.getCostBasis();
            String unrealizedText = "-";
            String unrealizedPctText = "-";
            if (row.latestPrice > 0.0 && lotCostBasis > 0.0) {
                double currentValue = lot.getUnits() * row.latestPrice;
                double unrealized = currentValue - lotCostBasis;
                double unrealizedPct = (unrealized / lotCostBasis) * 100.0;
                unrealizedText = HtmlFormatter.formatMoney(unrealized, row.currencyCode, 2);
                unrealizedPctText = HtmlFormatter.formatPercent(unrealizedPct, 2);
            }

            entries.add(new DetailEntry(
                    lot.getTradeDate(),
                    0,
                    "<span class=\"details-buy\">BUY</span>",
                    HtmlFormatter.formatUnits(lot.getUnits()),
                    HtmlFormatter.formatMoney(lot.getUnitCost(), row.currencyCode, 2),
                    HtmlFormatter.formatMoney(lotCostBasis, row.currencyCode, 2),
                    unrealizedText,
                    unrealizedPctText
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
                    "-"
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
        html.append("<tr><th>Date</th><th>Type</th><th>Units</th><th>Price</th><th>Amount</th><th>Unrealized</th><th>Unrealized (%)</th></tr>\n");
        for (DetailEntry entry : entries) {
            html.append("<tr>");
            html.append("<td>").append(escapeHtml(formatDetailDate(entry.date))).append("</td>");
            html.append("<td>").append(entry.type).append("</td>");
            html.append("<td>").append(escapeHtml(entry.units)).append("</td>");
            html.append("<td>").append(escapeHtml(entry.price)).append("</td>");
            html.append("<td>").append(escapeHtml(entry.amount)).append("</td>");
            html.append("<td>").append(escapeHtml(entry.unrealized)).append("</td>");
            html.append("<td>").append(escapeHtml(entry.unrealizedPct)).append("</td>");
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

    private static String buildDetailsHeaderCell(String groupName) {
        return "<span class=\"details-head\">Details"
            + "<button class=\"detail-group-toggle\" onclick=\"toggleDetailGroup('" + groupName + "', this)\" title=\"Expand/collapse all details\">▸</button>"
            + "</span>";
    }

    private static void writeHtmlRow(FileWriter writer, boolean isHeader, String... cells) throws IOException {
        writer.write("<tr>\n");
        for (String cell : cells) {
            if (isHeader) {
                writer.write("    <th>" + (cell != null ? cell : "") + "</th>\n");
            } else {
                writer.write("    <td>" + (cell != null ? cell : "") + "</td>\n");
            }
        }
        writer.write("</tr>\n");
    }

    private static void writeHtmlRowWithClass(FileWriter writer, String rowClass, String... cells) throws IOException {
        String classAttribute = (rowClass == null || rowClass.isBlank()) ? "" : " class=\"" + escapeHtml(rowClass) + "\"";
        writer.write("<tr" + classAttribute + ">\n");
        for (String cell : cells) {
            writer.write("    <td>" + (cell != null ? cell : "") + "</td>\n");
        }
        writer.write("</tr>\n");
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
        writer.write("  return true;\n");
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
        writer.write("function initSparklineRangeControls() {\n");
        writer.write("  document.querySelectorAll('.sparkline-widget').forEach(function(widget) {\n");
        writer.write("    var rangeButtons = widget.querySelectorAll('.sparkline-range-btn');\n");
        writer.write("    var metricButtons = widget.querySelectorAll('.sparkline-metric-btn');\n");
        writer.write("    var panels = widget.querySelectorAll('.sparkline-panel');\n");
        writer.write("    if (!metricButtons.length || !panels.length) return;\n");
        writer.write("    var defaultRangeSource = widget.querySelector('.sparkline-range-btn.is-active') || rangeButtons[0] || panels[0];\n");
        writer.write("    var activeRange = defaultRangeSource ? defaultRangeSource.getAttribute('data-range') : '';\n");
        writer.write("    var activeMetric = (widget.querySelector('.sparkline-metric-btn.is-active') || metricButtons[0]).getAttribute('data-metric');\n");
        writer.write("    function refresh() {\n");
        writer.write("      rangeButtons.forEach(function(btn) {\n");
        writer.write("        btn.classList.toggle('is-active', btn.getAttribute('data-range') === activeRange);\n");
        writer.write("      });\n");
        writer.write("      metricButtons.forEach(function(btn) {\n");
        writer.write("        btn.classList.toggle('is-active', btn.getAttribute('data-metric') === activeMetric);\n");
        writer.write("      });\n");
        writer.write("      panels.forEach(function(panel) {\n");
        writer.write("        var matchRange = panel.getAttribute('data-range') === activeRange;\n");
        writer.write("        var matchMetric = panel.getAttribute('data-metric') === activeMetric;\n");
        writer.write("        panel.classList.toggle('is-active', matchRange && matchMetric);\n");
        writer.write("      });\n");
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
        writer.write("  document.querySelectorAll('.hero-side').forEach(function(side) {\n");
        writer.write("    var openBtn = side.querySelector('.timeline-info-btn');\n");
        writer.write("    var overlay = side.querySelector('.timeline-info-overlay');\n");
        writer.write("    if (!openBtn || !overlay) return;\n");
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
        writer.write("  document.querySelectorAll('.security-scroll').forEach(function (node) {\n");
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
        writer.write("  initChartDownloadButtons();\n");
        writer.write("  initChartHoverEffects();\n");
        writer.write("  initInteractiveChartControls();\n");
        writer.write("  initThemeToggle();\n");
        writer.write("})();\n");
    }

    private static String escapeHtml(String text) {
        if (text == null) return "";
        return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\"", "&quot;");
    }
}