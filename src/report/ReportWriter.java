package report;

import csv.TransactionStore;
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
import java.util.Set;

public class ReportWriter {

    private static final DateTimeFormatter DETAIL_DATE_FORMATTER = DateTimeFormatter.ofPattern("dd.MM.yyyy");
    private static final String DEFAULT_TOTAL_CURRENCY = "NOK";

    public static void writeHtmlReport(TransactionStore store, String outputFile) throws IOException {
        List<OverviewRow> overviewRows = PortfolioCalculator.buildOverviewRows(store);
        Map<String, Double> ratesToNok = CurrencyConversionService.loadRatesToNok(collectCurrencies(store, overviewRows));
        HeaderSummary headerSummary = PortfolioCalculator.buildHeaderSummary(store, overviewRows, ratesToNok);

        try (FileWriter writer = new FileWriter(outputFile)) {
            writer.write("<!DOCTYPE html>\n");
            writer.write("<html lang=\"no\">\n");
            writer.write("<head>\n");
            writer.write("    <meta charset=\"UTF-8\">\n");
            writer.write("    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">\n");
            writer.write("    <title>Portfolio Report</title>\n");
            writer.write("    <style>\n");
            writer.write("        :root { --bg:#eef3f7; --line:#d8e0e9; --card:#ffffff; --ink:#16202a; --muted:#5a6877; --good:#1f8b4d; --bad:#b23a31; }\n");
            writer.write("        * { box-sizing: border-box; }\n");
            writer.write("        body { font-family: 'Segoe UI','Avenir Next','Helvetica Neue',Arial,sans-serif; margin:0; background: radial-gradient(circle at top,#f8fbfe 0%,var(--bg) 58%); color:var(--ink); }\n");
            writer.write("        .page { width:100vw; max-width:none; margin:0; padding:24px 8px 32px; }\n");
            writer.write("        h2 { margin:26px 2px 12px; font-size:1.14rem; color:var(--ink); }\n");
            writer.write("        table { width:100%; border-collapse:collapse; min-width:1120px; background:var(--card); }\n");
            writer.write("        th, td { padding:8px 8px; border-bottom:1px solid #edf2f7; white-space:nowrap; }\n");
            writer.write("        th { background:#f5f8fb; text-align:left; font-size:.8rem; text-transform:uppercase; letter-spacing:.2px; color:#374556; border-bottom:1px solid var(--line); }\n");
            writer.write("        td { font-size:.88rem; }\n");
            writer.write("        td.num, th.num { text-align:right; }\n");
            writer.write("        .table-wrap { background:var(--card); border:1px solid var(--line); border-radius:14px; overflow:auto; box-shadow:0 5px 14px rgba(15,23,33,.06); }\n");
            writer.write("        .total-row { font-weight:700; background:#f3f7fb; }\n");
            writer.write("        .asset-split td { border-top:3px solid #8a9eb3 !important; }\n");
            writer.write("        .positive { color:var(--good); } .negative { color:var(--bad); }\n");
            writer.write("        .report-hero { display:grid; grid-template-columns:1.25fr 1fr; gap:16px; background:linear-gradient(120deg,#0f2238 0%,#18344f 60%,#164663 100%); border-radius:18px; padding:22px; color:#f4f8fc; box-shadow:0 14px 26px rgba(10,24,38,.2); margin-bottom:18px; }\n");
            writer.write("        .hero-title h1 { margin:0; font-size:1.75rem; letter-spacing:.4px; }\n");
            writer.write("        .hero-meta { margin-top:10px; display:flex; flex-wrap:wrap; gap:8px; }\n");
            writer.write("        .meta-chip { display:inline-flex; align-items:center; gap:6px; padding:6px 11px; border-radius:999px; border:1px solid rgba(235,245,255,.28); background:rgba(255,255,255,.1); color:#d7e6f4; font-size:.84rem; font-weight:600; }\n");
            writer.write("        .meta-chip strong { color:#ffffff; font-weight:700; }\n");
            writer.write("        .currency-input { width:56px; border:1px solid rgba(235,245,255,.45); border-radius:6px; background:rgba(255,255,255,.18); color:#fff; font-weight:700; text-transform:uppercase; padding:2px 6px; outline:none; }\n");
            writer.write("        .currency-input:focus { border-color:#fff; background:rgba(255,255,255,.26); }\n");
            writer.write("        .hero-kpis { margin-top:14px; display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:10px; }\n");
            writer.write("        .kpi-card { background:rgba(255,255,255,.08); border:1px solid rgba(235,245,255,.2); border-radius:10px; padding:10px 11px; }\n");
            writer.write("        .kpi-label { color:#c8d9eb; font-size:.8rem; text-transform:uppercase; }\n");
            writer.write("        .kpi-value { margin-top:2px; font-size:1.02rem; font-weight:700; color:#fff; }\n");
            writer.write("        .performer { margin-top:6px; font-size:.84rem; color:#dce8f3; }\n");
            writer.write("        .performer strong { display:block; font-size:.9rem; margin-bottom:2px; }\n");
            writer.write("        .performer-metrics { display:block; }\n");
            writer.write("        .hero-side { background:rgba(255,255,255,.06); border:1px solid rgba(235,245,255,.22); border-radius:12px; padding:10px; min-height:172px; }\n");
            writer.write("        .hero-side-title { color:#d4e3f0; font-size:.86rem; text-transform:uppercase; margin-bottom:8px; }\n");
            writer.write("        .hero-side-note { color:#d4e3f0; font-size:.92rem; }\n");
            writer.write("        .sparkline-widget { display:block; }\n");
            writer.write("        .sparkline-controls { display:flex; flex-wrap:wrap; gap:6px; margin:0 0 8px; }\n");
            writer.write("        .sparkline-range-btn { border:1px solid rgba(235,245,255,.35); background:rgba(255,255,255,.12); color:#e4eef8; border-radius:999px; padding:3px 8px; font-size:.72rem; font-weight:700; letter-spacing:.2px; cursor:pointer; }\n");
            writer.write("        .sparkline-range-btn:hover { background:rgba(255,255,255,.2); }\n");
            writer.write("        .sparkline-range-btn.is-active { background:#eaf4ff; color:#16344d; border-color:#ffffff; }\n");
            writer.write("        .sparkline-panel { display:none; }\n");
            writer.write("        .sparkline-panel.is-active { display:block; }\n");
            writer.write("        .overview-charts { display:grid; grid-template-columns:1fr 1fr; gap:14px; margin:12px 0 14px; }\n");
            writer.write("        .overview-chart { padding:14px; border:1px solid var(--line); border-radius:14px; background:var(--card); box-shadow:0 5px 14px rgba(15,23,33,.06); overflow:hidden; }\n");
            writer.write("        .overview-chart h3 { margin:0 0 10px; font-size:1rem; }\n");
            writer.write("        .overview-chart .chart-svg { display:block; width:100%; margin:0 auto 12px; }\n");
            writer.write("        .allocation-card { margin:16px 0 18px; padding:14px; border:1px solid var(--line); border-radius:14px; background:var(--card); box-shadow:0 5px 14px rgba(15,23,33,.06); }\n");
            writer.write("        .allocation-card h3 { margin:0 0 10px; font-size:1rem; }\n");
            writer.write("        .allocation-visuals { display:grid; gap:10px; }\n");
            writer.write("        .allocation-row { display:grid; gap:10px; }\n");
            writer.write("        .allocation-row-top { grid-template-columns:repeat(3,minmax(0,1fr)); }\n");
            writer.write("        .allocation-row-bottom { grid-template-columns:repeat(2,minmax(0,1fr)); }\n");
            writer.write("        .allocation-panel { border:1px solid var(--line); border-radius:10px; padding:16px; background:#fafcfe; overflow:hidden; }\n");
            writer.write("        .allocation-panel-title { margin:0 0 6px; font-size:.84rem; text-transform:uppercase; color:#41576d; letter-spacing:.3px; }\n");
            writer.write("        .chart-svg { width:100%; height:auto; background:var(--card); border:1px solid var(--line); border-radius:8px; }\n");
            writer.write("        .allocation-panel .chart-svg { width:96%; margin:6px auto 10px; display:block; }\n");
            writer.write("        .security-pie-panel .chart-svg, .security-bar-panel .chart-svg { height:340px; width:100%; }\n");
            writer.write("        .chart-hover-target { cursor:pointer; transform-box:fill-box; transform-origin:center; transition:transform .14s ease, filter .14s ease, opacity .14s ease; }\n");
            writer.write("        .chart-hover-target.is-hovered { filter:brightness(1.08); opacity:.96; }\n");
            writer.write("        .chart-hover-bar.is-hovered { transform:translateY(-2px); }\n");
            writer.write("        .chart-hover-slice.is-hovered { transform:scale(1.03); }\n");
            writer.write("        .chart-hover-point.is-hovered { transform:scale(1.75); stroke:#ffffff; stroke-width:1; }\n");
            writer.write("        .chart-tooltip { position:fixed; pointer-events:none; z-index:10000; max-width:340px; padding:7px 10px; border-radius:8px; background:rgba(16,28,40,.94); color:#f6fbff; font-size:.8rem; font-weight:600; line-height:1.3; box-shadow:0 8px 18px rgba(7,16,26,.28); border:1px solid rgba(255,255,255,.14); opacity:0; transform:translateY(4px); transition:opacity .1s ease, transform .1s ease; }\n");
            writer.write("        .chart-tooltip.visible { opacity:1; transform:translateY(0); }\n");
            writer.write("        .expand-btn { border:1px solid #86a4bf; background:#f3f8fd; color:#1e3951; border-radius:7px; padding:4px 8px; font-size:.78rem; font-weight:600; cursor:pointer; }\n");
            writer.write("        .expand-btn:hover { background:#e6f1fb; }\n");
            writer.write("        .details-head { display:inline-flex; align-items:center; gap:6px; }\n");
            writer.write("        .detail-group-toggle { border:1px solid #86a4bf; background:#f3f8fd; color:#1e3951; border-radius:50%; width:18px; height:18px; padding:0; line-height:16px; font-size:.72rem; font-weight:700; cursor:pointer; display:inline-flex; align-items:center; justify-content:center; }\n");
            writer.write("        .detail-group-toggle:hover { background:#e6f1fb; }\n");
            writer.write("        .details-row { display:none; }\n");
            writer.write("        .details-cell { padding:0 !important; background:#f9fcff; }\n");
            writer.write("        .details-wrap { padding:10px 12px 12px; }\n");
            writer.write("        .details-wrap h4 { margin:0 0 8px; font-size:.85rem; color:#31495f; text-transform:uppercase; letter-spacing:.25px; }\n");
            writer.write("        .details-table { width:100%; border-collapse:collapse; min-width:0; background:#fff; border:1px solid #dfe7ef; }\n");
            writer.write("        .details-table th, .details-table td { padding:6px 7px; border-bottom:1px solid #edf2f7; font-size:.82rem; white-space:nowrap; }\n");
            writer.write("        .details-table th { background:#f4f8fc; color:#405a70; }\n");
            writer.write("        .details-buy { color:#1d5d92; font-weight:600; }\n");
            writer.write("        .details-dividend { color:#1f8b4d; font-weight:600; }\n");
            writer.write("        @media (max-width:1200px) { .allocation-row-top{grid-template-columns:1fr 1fr;} }\n");
            writer.write("        @media (max-width:1060px) { .report-hero{grid-template-columns:1fr;} .hero-kpis{grid-template-columns:1fr;} .overview-charts{grid-template-columns:1fr;} .allocation-row-top,.allocation-row-bottom{grid-template-columns:1fr;} .page{width:100vw; padding:16px 8px 22px;} }\n");
            writer.write("    </style>\n");
            writer.write("</head>\n");
            writer.write("<body>\n");
            writer.write("<main class=\"page\">\n");

            // Header Summary
            writeHeaderSummaryHtml(writer, headerSummary, overviewRows, store, ratesToNok);

            // Overview Section
            writeOverviewTableHtml(writer, overviewRows, store, ratesToNok);

            // Realized Overview
            writeRealizedSummaryTableHtml(writer, store, ratesToNok);

            writer.write("</main>\n");
            writer.write("<script>\n");
            writer.write("function toggleOverviewDetails(rowId, button) {\n");
            writer.write("  var row = document.getElementById(rowId);\n");
            writer.write("  if (!row) return;\n");
            writer.write("  var isOpen = row.style.display === 'table-row';\n");
            writer.write("  row.style.display = isOpen ? 'none' : 'table-row';\n");
            writer.write("  if (button) button.textContent = isOpen ? 'Show details' : 'Hide details';\n");
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
            writer.write("    if (rowButton) rowButton.textContent = open ? 'Hide details' : 'Show details';\n");
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
        writer.write("</div>\n");
        writer.write("<div class=\"hero-kpis\">\n");
        LinkedHashMap<String, Double> totalMarketBuckets = new LinkedHashMap<>();
        LinkedHashMap<String, Double> totalReturnBuckets = new LinkedHashMap<>();
        LinkedHashMap<String, Double> totalHistoricalCostBuckets = new LinkedHashMap<>();
        for (OverviewRow row : overviewRows) {
            addToCurrencyBuckets(totalMarketBuckets, row.currencyCode, row.marketValue);
            addToCurrencyBuckets(totalReturnBuckets, row.currencyCode, row.totalReturn);
            addToCurrencyBuckets(totalHistoricalCostBuckets, row.currencyCode, row.historicalCostBasis);
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
        writer.write("<aside class=\"hero-side\"><div class=\"hero-side-title\">Portfolio Value Timeline</div>");
        if (s.sparklineSvg != null && !s.sparklineSvg.isBlank()) {
            writer.write(s.sparklineSvg);
        } else {
            writer.write("<div class=\"hero-side-note\">Timeline data not available yet for this dataset.</div>");
        }
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

        writer.write("<div class=\"table-wrap\">\n<table>\n");
        writeHtmlRow(writer, true,
            buildDetailsHeaderCell("overview-details"), "Ticker", "Security", "Units", "Avg Cost", "Last Price",
                "Market Value", "Cost Basis", "Unrealized", "Unrealized (%)",
                "Realized (%)", "Realized", "Dividends", "Total Return", "Total Return (%)");

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

            writeHtmlRowWithClass(writer, rowClass,
                "<button class=\"expand-btn\" data-target=\"" + detailsRowId + "\" onclick=\"toggleOverviewDetails('" + detailsRowId + "', this)\">Show details</button>",
                    row.tickerText,
                    row.securityDisplayName,
                    HtmlFormatter.formatUnits(row.units),
                    HtmlFormatter.formatMoney(row.averageCost, row.currencyCode, 2),
                    row.latestPrice > 0 ? HtmlFormatter.formatMoney(row.latestPrice, row.currencyCode, 2) : "-",
                    row.latestPrice > 0 ? HtmlFormatter.formatMoney(row.marketValue, row.currencyCode, 2) : "-",
                    HtmlFormatter.formatMoney(row.positionCostBasis, row.currencyCode, 2),
                    row.hasPrice ? HtmlFormatter.formatMoney(row.unrealized, row.currencyCode, 2) : "-",
                    row.hasPrice ? HtmlFormatter.formatPercent(row.unrealizedPct, 2) : "-",
                    row.realizedReturnPctText + "%",
                    HtmlFormatter.formatMoney(row.realized, row.currencyCode, 2),
                    HtmlFormatter.formatMoney(row.dividends, row.currencyCode, 2),
                    HtmlFormatter.formatMoney(row.totalReturn, row.currencyCode, 2),
                    HtmlFormatter.formatPercent(row.totalReturnPct, 2));

                    writer.write("<tr id=\"" + detailsRowId + "\" class=\"details-row\" data-group=\"overview-details\">\n");
                    writer.write("    <td class=\"details-cell\" colspan=\"15\">\n");
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
        writer.write("    <td>" + renderConvertibleMoneyCell(totalMarketValueBuckets, 2, ratesToNok) + "</td>\n");
        writer.write("    <td>" + renderConvertibleMoneyCell(totalCostBasisBuckets, 2, ratesToNok) + "</td>\n");
        writer.write("    <td>" + renderConvertibleMoneyCell(totalUnrealizedBuckets, 2, ratesToNok) + "</td>\n");
        writer.write("    <td>" + HtmlFormatter.formatPercent(totalUnrealizedPct, 2) + "</td>\n");
        writer.write("    <td>" + HtmlFormatter.formatPercent(totalRealizedPct, 2) + "</td>\n");
        writer.write("    <td>" + renderConvertibleMoneyCell(totalRealizedBuckets, 2, ratesToNok) + "</td>\n");
        writer.write("    <td>" + renderConvertibleMoneyCell(totalDividendsBuckets, 2, ratesToNok) + "</td>\n");
        writer.write("    <td>" + renderConvertibleMoneyCell(totalReturnBuckets, 2, ratesToNok) + "</td>\n");
        writer.write("    <td>" + HtmlFormatter.formatPercent(totalReturnPct, 2) + "</td>\n");
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
        writer.write("<div class=\"table-wrap\">\n<table>\n");
        writeHtmlRow(writer, true, buildDetailsHeaderCell("realized-details"), "Ticker", "Security", "Sales Value", "Cost Basis", "Realized Gain/Loss", "Dividends", "Return (%)");

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
            double returnPct = costBasis > 0 ? (gain / costBasis) * 100.0 : (gain > 0 ? 100.0 : 0.0);
            String currentAssetType = security.getAssetType().name();
            String rowClass = isStockFundBoundary(previousAssetType, currentAssetType) ? "asset-split" : null;

            addToCurrencyBuckets(totalSalesValueBuckets, currency, salesValue);
            addToCurrencyBuckets(totalCostBasisBuckets, currency, costBasis);
            addToCurrencyBuckets(totalRealizedGainBuckets, currency, gain);
            addToCurrencyBuckets(totalRealizedDividendsBuckets, currency, realizedDividends);
                String detailsRowId = "realized-details-" + detailsIndex;

                writeHtmlRowWithClass(writer, rowClass,
                    "<button class=\"expand-btn\" data-target=\"" + detailsRowId + "\" onclick=\"toggleOverviewDetails('" + detailsRowId + "', this)\">Show details</button>",
                    security.getTicker(),
                    security.getDisplayName(),
                    HtmlFormatter.formatMoney(salesValue, currency, 2),
                    HtmlFormatter.formatMoney(costBasis, currency, 2),
                    HtmlFormatter.formatMoney(gain, currency, 2),
                    HtmlFormatter.formatMoney(realizedDividends, currency, 2),
                    HtmlFormatter.formatPercent(returnPct, 2));

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
        double totalReturnPct = totalCostBasisForPct > 0
            ? (totalRealizedGainForPct / totalCostBasisForPct) * 100.0
            : (totalRealizedGainForPct > 0 ? 100.0 : 0.0);
        writer.write("<tr class=\"total-row\">\n");
        writer.write("    <td></td><td></td><td><strong>TOTAL</strong></td>\n");
        writer.write("    <td>" + renderConvertibleMoneyCell(totalSalesValueBuckets, 2, ratesToNok) + "</td>\n");
        writer.write("    <td>" + renderConvertibleMoneyCell(totalCostBasisBuckets, 2, ratesToNok) + "</td>\n");
        writer.write("    <td>" + renderConvertibleMoneyCell(totalRealizedGainBuckets, 2, ratesToNok) + "</td>\n");
        writer.write("    <td>" + renderConvertibleMoneyCell(totalRealizedDividendsBuckets, 2, ratesToNok) + "</td>\n");
        writer.write("    <td>" + HtmlFormatter.formatPercent(totalReturnPct, 2) + "</td>\n");
        writer.write("</tr>\n");

        writer.write("</table>\n</div>\n\n");
    }

    private static String buildRealizedSaleTradesDetailsHtml(Security security) {
        StringBuilder html = new StringBuilder();
        html.append("<div class=\"details-wrap\">\n");
        html.append("<h4>Sale Trades - ").append(escapeHtml(security.getDisplayName())).append("</h4>\n");

        List<Security.SaleTrade> saleTrades = security.getSaleTradesSortedByDate();
        String currency = security.getCurrencyCode();
        if (saleTrades.isEmpty()) {
            html.append("<div class=\"hero-side-note\">No sale trades available.</div>\n");
            html.append("</div>\n");
            return html.toString();
        }

        html.append("<table class=\"details-table\">\n");
        html.append("<tr><th>Sale Date</th><th>Units</th><th>Price/Unit</th><th>Sale Value</th><th>Cost Basis</th><th>Gain/Loss</th><th>Return (%)</th></tr>\n");

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
            html.append("<td>").append(escapeHtml(HtmlFormatter.formatMoney(trade.getSaleValue(), currency, 0))).append("</td>");
            html.append("<td>").append(escapeHtml(HtmlFormatter.formatMoney(trade.getCostBasis(), currency, 0))).append("</td>");
            html.append("<td>").append(escapeHtml(HtmlFormatter.formatMoney(trade.getGainLoss(), currency, 0))).append("</td>");
            html.append("<td>").append(escapeHtml(HtmlFormatter.formatPercent(trade.getReturnPct(), 2))).append("</td>");
            html.append("</tr>\n");
        }

        double totalReturnPct = totalCostBasis > 0.0 ? (totalGainLoss / totalCostBasis) * 100.0 : 0.0;
        html.append("<tr class=\"total-row\">");
        html.append("<td><strong>TOTAL</strong></td>");
        html.append("<td>").append(escapeHtml(HtmlFormatter.formatUnits(totalUnits))).append("</td>");
        html.append("<td></td>");
        html.append("<td>").append(escapeHtml(HtmlFormatter.formatMoney(totalSaleValue, currency, 0))).append("</td>");
        html.append("<td>").append(escapeHtml(HtmlFormatter.formatMoney(totalCostBasis, currency, 0))).append("</td>");
        html.append("<td>").append(escapeHtml(HtmlFormatter.formatMoney(totalGainLoss, currency, 0))).append("</td>");
        html.append("<td>").append(escapeHtml(HtmlFormatter.formatPercent(totalReturnPct, 2))).append("</td>");
        html.append("</tr>\n");

        html.append("</table>\n</div>\n");
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

    private static String buildHoldingDetailsTableHtml(Security security, OverviewRow row) {
        StringBuilder html = new StringBuilder();
        html.append("<div class=\"details-wrap\">\n");
        html.append("<h4>Transaction Details - ").append(escapeHtml(row.securityDisplayName)).append("</h4>\n");

        if (security == null) {
            html.append("<div class=\"hero-side-note\">Details are not available for this security.</div>\n");
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
            html.append("<div class=\"hero-side-note\">No active buy/dividend entries for current holdings.</div>\n");
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
        if (previousAssetType == null || currentAssetType == null || previousAssetType.equals(currentAssetType)) {
            return false;
        }

        return ("STOCK".equals(previousAssetType) && "FUND".equals(currentAssetType))
                || ("FUND".equals(previousAssetType) && "STOCK".equals(currentAssetType));
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
        writer.write("  function readTooltipText(target) {\n");
        writer.write("    if (!target) return '';\n");
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
        writer.write("      target.setAttribute('data-tooltip', String(nativeTitle.textContent || '').trim());\n");
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
        writer.write("function initSparklineRangeControls() {\n");
        writer.write("  document.querySelectorAll('.sparkline-widget').forEach(function(widget) {\n");
        writer.write("    var buttons = widget.querySelectorAll('.sparkline-range-btn');\n");
        writer.write("    var panels = widget.querySelectorAll('.sparkline-panel');\n");
        writer.write("    if (!buttons.length || !panels.length) return;\n");
        writer.write("    function activate(range) {\n");
        writer.write("      buttons.forEach(function(btn) {\n");
        writer.write("        var on = btn.getAttribute('data-range') === range;\n");
        writer.write("        btn.classList.toggle('is-active', on);\n");
        writer.write("      });\n");
        writer.write("      panels.forEach(function(panel) {\n");
        writer.write("        var on = panel.getAttribute('data-range') === range;\n");
        writer.write("        panel.classList.toggle('is-active', on);\n");
        writer.write("      });\n");
        writer.write("    }\n");
        writer.write("    buttons.forEach(function(btn) {\n");
        writer.write("      btn.addEventListener('click', function() {\n");
        writer.write("        activate(btn.getAttribute('data-range'));\n");
        writer.write("      });\n");
        writer.write("    });\n");
        writer.write("    var active = widget.querySelector('.sparkline-range-btn.is-active');\n");
        writer.write("    activate(active ? active.getAttribute('data-range') : buttons[0].getAttribute('data-range'));\n");
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
        writer.write("  initChartHoverEffects();\n");
        writer.write("})();\n");
    }

    private static String escapeHtml(String text) {
        if (text == null) return "";
        return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\"", "&quot;");
    }
}