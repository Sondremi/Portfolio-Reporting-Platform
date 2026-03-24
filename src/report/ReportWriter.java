package report;

import csv.TransactionStore;
import model.Security;

import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;

public class ReportWriter {

    public static void writeHtmlReport(TransactionStore store, String outputFile) throws IOException {
        List<OverviewRow> overviewRows = PortfolioCalculator.buildOverviewRows(store);
        HeaderSummary headerSummary = PortfolioCalculator.buildHeaderSummary(store, overviewRows);

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
            writer.write("        .security-pie-panel .chart-svg, .security-bar-panel .chart-svg { height:340px; width:98%; }\n");
            writer.write("        @media (max-width:1200px) { .allocation-row-top{grid-template-columns:1fr 1fr;} }\n");
            writer.write("        @media (max-width:1060px) { .report-hero{grid-template-columns:1fr;} .hero-kpis{grid-template-columns:1fr;} .overview-charts{grid-template-columns:1fr;} .allocation-row-top,.allocation-row-bottom{grid-template-columns:1fr;} .page{width:100vw; padding:16px 8px 22px;} }\n");
            writer.write("    </style>\n");
            writer.write("</head>\n");
            writer.write("<body>\n");
            writer.write("<main class=\"page\">\n");

            // Header Summary
            writeHeaderSummaryHtml(writer, headerSummary);

            // Overview Section
            writeOverviewTableHtml(writer, overviewRows, store);

            // Realized Overview
            writeRealizedSummaryTableHtml(writer, store);

            // Sale Trades Details
            writeSaleTradesTablesHtml(writer, store);

            writer.write("</main>\n");
            writer.write("</body>\n");
            writer.write("</html>\n");
        }
    }

    private static void writeHeaderSummaryHtml(FileWriter writer, HeaderSummary s) throws IOException {
        String totalClass = s.totalReturn >= 0 ? "positive" : "negative";
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
        writer.write("</div>\n");
        writer.write("<div class=\"hero-kpis\">\n");
        writer.write("<article class=\"kpi-card\"><div class=\"kpi-label\">Total Market Value</div><div class=\"kpi-value\">" + HtmlFormatter.formatMoney(s.totalMarketValue, s.totalCurrencyCode, 0) + "</div></article>\n");
        writer.write("<article class=\"kpi-card\"><div class=\"kpi-label\">Cash Holdings</div><div class=\"kpi-value\">" + HtmlFormatter.formatMoney(s.cashHoldings, s.totalCurrencyCode, 0) + "</div></article>\n");
        writer.write("<article class=\"kpi-card\"><div class=\"kpi-label\">Total Return</div><div class=\"kpi-value " + totalClass + "\">" + HtmlFormatter.formatMoney(s.totalReturn, s.totalCurrencyCode, 0) + "</div><div class=\"kpi-label " + totalClass + "\">" + HtmlFormatter.formatPercent(s.totalReturnPct) + "</div></article>\n");
        writer.write("<article class=\"kpi-card\"><div class=\"kpi-label\">Best / Worst</div><div class=\"performer " + bestClass + "\"><strong>" + escapeHtml(s.bestLabel) + "</strong><span class=\"performer-metrics\">" + HtmlFormatter.formatMoney(s.bestReturn, s.totalCurrencyCode, 0) + " | " + HtmlFormatter.formatPercent(s.bestReturnPct) + "</span></div><div class=\"performer " + worstClass + "\"><strong>" + escapeHtml(s.worstLabel) + "</strong><span class=\"performer-metrics\">" + HtmlFormatter.formatMoney(s.worstReturn, s.totalCurrencyCode, 0) + " | " + HtmlFormatter.formatPercent(s.worstReturnPct) + "</span></div></article>\n");
        writer.write("</div></div>\n");
        writer.write("<aside class=\"hero-side\"><div class=\"hero-side-title\">Portfolio Value Last 12 Months</div>");
        if (s.sparklineSvg != null && !s.sparklineSvg.isBlank()) {
            writer.write(s.sparklineSvg);
        } else {
            writer.write("<div class=\"hero-side-note\">Timeline data not available yet for this dataset.</div>");
        }
        writer.write("</aside>\n");
        writer.write("</section>\n");
    }

    private static void writeOverviewTableHtml(FileWriter writer, List<OverviewRow> rows, TransactionStore store) throws IOException {
        writer.write("<h2>PORTFOLIO OVERVIEW - CURRENT HOLDINGS</h2>\n");

        writer.write("<div class=\"overview-charts\">\n");
        writer.write("<section class=\"overview-chart total-return-chart\"><h3>Total Return (NOK)</h3>\n");
        writer.write(ChartBuilder.buildOverviewBarChartSvg(rows, false));
        writer.write("</section>\n");
        writer.write("<section class=\"overview-chart total-return-chart\"><h3>Total Return (%)</h3>\n");
        writer.write(ChartBuilder.buildOverviewBarChartSvg(rows, true));
        writer.write("</section>\n");
        writer.write("</div>\n");

        writer.write("<div class=\"table-wrap\">\n<table>\n");
        writeHtmlRow(writer, true,
            "Ticker", "Security", "Units", "Avg Cost", "Last Price",
                "Market Value", "Cost Basis", "Unrealized", "Unrealized (%)",
                "Realized (%)", "Realized", "Dividends", "Total Return", "Total Return (%)");

        double totalMarketValue = 0.0;
        double totalCostBasis = 0.0;
        double totalUnrealized = 0.0;
        double totalRealized = 0.0;
        double totalDividends = 0.0;
        double totalHistoricalCost = 0.0;
        String totalCurrencyCode = null;
        String previousAssetType = null;

        for (OverviewRow row : rows) {
            totalMarketValue += row.marketValue;
            totalCostBasis += row.positionCostBasis;
            totalUnrealized += row.unrealized;
            totalRealized += row.realized;
            totalDividends += row.dividends;
            totalHistoricalCost += row.historicalCostBasis;
            totalCurrencyCode = mergeCurrencyCodes(totalCurrencyCode, row.currencyCode);
            String rowClass = isStockFundBoundary(previousAssetType, row.assetType) ? "asset-split" : null;

            writeHtmlRowWithClass(writer, rowClass,
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

            previousAssetType = row.assetType;
        }

        double totalReturn = totalUnrealized + totalRealized + totalDividends;
        double totalReturnPct = totalHistoricalCost > 0 ? (totalReturn / totalHistoricalCost) * 100.0 : 0.0;
        double totalUnrealizedPct = totalCostBasis > 0 ? (totalUnrealized / totalCostBasis) * 100.0 : 0.0;
        double totalRealizedPct = totalCostBasis > 0 ? (totalRealized / totalCostBasis) * 100.0 : 0.0;

        writer.write("<tr class=\"total-row\">\n");
        writer.write("    <td></td><td><strong>TOTAL</strong></td><td></td><td></td><td></td>\n");
        writer.write("    <td>" + formatTotalMoney(totalMarketValue, totalCurrencyCode, 2) + "</td>\n");
        writer.write("    <td>" + formatTotalMoney(totalCostBasis, totalCurrencyCode, 2) + "</td>\n");
        writer.write("    <td>" + formatTotalMoney(totalUnrealized, totalCurrencyCode, 2) + "</td>\n");
        writer.write("    <td>" + HtmlFormatter.formatPercent(totalUnrealizedPct, 2) + "</td>\n");
        writer.write("    <td>" + HtmlFormatter.formatPercent(totalRealizedPct, 2) + "</td>\n");
        writer.write("    <td>" + formatTotalMoney(totalRealized, totalCurrencyCode, 2) + "</td>\n");
        writer.write("    <td>" + formatTotalMoney(totalDividends, totalCurrencyCode, 2) + "</td>\n");
        writer.write("    <td>" + formatTotalMoney(totalReturn, totalCurrencyCode, 2) + "</td>\n");
        writer.write("    <td>" + HtmlFormatter.formatPercent(totalReturnPct, 2) + "</td>\n");
        writer.write("</tr>\n");

        writer.write("</table>\n</div>\n\n");

        writer.write("<section class=\"allocation-card\">\n");
        writer.write("<h3>Market Value Allocation</h3>\n");
        writer.write("<div class=\"allocation-visuals\">\n");
        writer.write("<div class=\"allocation-row allocation-row-top\">\n");
        writer.write("<div class=\"allocation-panel asset-type-panel\"><h4 class=\"allocation-panel-title\">By Asset Type</h4>\n");
        writer.write(ChartBuilder.buildAssetTypeAllocationSvg(rows, store.getCurrentCashHoldings()));
        writer.write("</div>\n");
        writer.write("<div class=\"allocation-panel sector-panel\"><h4 class=\"allocation-panel-title\">By Sector</h4>\n");
        writer.write(ChartBuilder.buildSectorAllocationSvg(rows));
        writer.write("</div>\n");
        writer.write("<div class=\"allocation-panel region-panel\"><h4 class=\"allocation-panel-title\">By Region</h4>\n");
        writer.write(ChartBuilder.buildRegionAllocationSvg(rows));
        writer.write("</div>\n");
        writer.write("</div>\n");

        writer.write("<div class=\"allocation-row allocation-row-bottom\">\n");
        writer.write("<div class=\"allocation-panel security-pie-panel\"><h4 class=\"allocation-panel-title\">By Security (Pie)</h4>\n");
        writer.write(ChartBuilder.buildMarketValueAllocationSvg(rows));
        writer.write("</div>\n");
        writer.write("<div class=\"allocation-panel security-bar-panel\"><h4 class=\"allocation-panel-title\">By Security (Bar)</h4>\n");
        writer.write(ChartBuilder.buildMarketValueBarChartSvg(rows));
        writer.write("</div>\n");
        writer.write("</div>\n");
        writer.write("</div>\n");
        writer.write("</section>\n");
    }

    private static void writeRealizedSummaryTableHtml(FileWriter writer, TransactionStore store) throws IOException {
        writer.write("<h2>REALIZED OVERVIEW - ALL SALES</h2>\n");
        writer.write("<div class=\"table-wrap\">\n<table>\n");
        writeHtmlRow(writer, true, "Ticker", "Security", "Sales Value", "Cost Basis", "Realized Gain/Loss", "Dividends", "Return (%)");

        ArrayList<Security> soldSecurities = getSortedSoldSecurities(store);
        double totalSalesValue = 0.0;
        double totalCostBasis = 0.0;
        double totalRealizedGain = 0.0;
        double totalRealizedDividends = 0.0;
        String totalCurrencyCode = null;
        String previousAssetType = null;

        for (Security security : soldSecurities) {
            String currency = security.getCurrencyCode();
            double salesValue = security.getRealizedSalesValue();
            double costBasis = security.getRealizedCostBasis();
            double gain = security.getRealizedGain();
            double realizedDividends = security.isFullyRealized() ? security.getDividends() : 0.0;
            double returnPct = costBasis > 0 ? (gain / costBasis) * 100.0 : (gain > 0 ? 100.0 : 0.0);
            String currentAssetType = security.getAssetType().name();
            String rowClass = isStockFundBoundary(previousAssetType, currentAssetType) ? "asset-split" : null;

            totalSalesValue += salesValue;
            totalCostBasis += costBasis;
            totalRealizedGain += gain;
            totalRealizedDividends += realizedDividends;
            totalCurrencyCode = mergeCurrencyCodes(totalCurrencyCode, currency);

            writeHtmlRowWithClass(writer, rowClass,
                    security.getTicker(),
                    security.getDisplayName(),
                    HtmlFormatter.formatMoney(salesValue, currency, 2),
                    HtmlFormatter.formatMoney(costBasis, currency, 2),
                    HtmlFormatter.formatMoney(gain, currency, 2),
                    HtmlFormatter.formatMoney(realizedDividends, currency, 2),
                    HtmlFormatter.formatPercent(returnPct, 2));

            previousAssetType = currentAssetType;
        }

        double totalReturnPct = totalCostBasis > 0 ? (totalRealizedGain / totalCostBasis) * 100.0 : (totalRealizedGain > 0 ? 100.0 : 0.0);
        writer.write("<tr class=\"total-row\">\n");
        writer.write("    <td></td><td><strong>TOTAL</strong></td>\n");
        writer.write("    <td>" + formatTotalMoney(totalSalesValue, totalCurrencyCode, 2) + "</td>\n");
        writer.write("    <td>" + formatTotalMoney(totalCostBasis, totalCurrencyCode, 2) + "</td>\n");
        writer.write("    <td>" + formatTotalMoney(totalRealizedGain, totalCurrencyCode, 2) + "</td>\n");
        writer.write("    <td>" + formatTotalMoney(totalRealizedDividends, totalCurrencyCode, 2) + "</td>\n");
        writer.write("    <td>" + HtmlFormatter.formatPercent(totalReturnPct, 2) + "</td>\n");
        writer.write("</tr>\n");

        writer.write("</table>\n</div>\n\n");
    }

    private static void writeSaleTradesTablesHtml(FileWriter writer, TransactionStore store) throws IOException {
        ArrayList<Security> soldSecurities = getSortedSoldSecurities(store);

        for (Security security : soldSecurities) {
            List<Security.SaleTrade> saleTrades = security.getSaleTradesSortedByDate();
            String currency = security.getCurrencyCode();
            writer.write("<h2>SALE TRADES - " + escapeHtml(security.getDisplayName()) + "</h2>\n");
            writer.write("<div class=\"table-wrap\">\n<table>\n");
            writeHtmlRow(writer, true, "Sale Date", "Units", "Price/Unit", "Sale Value", "Cost Basis", "Gain/Loss", "Return (%)");

            for (Security.SaleTrade trade : saleTrades) {
                writeHtmlRow(writer, false,
                        trade.getTradeDateAsCsv(),
                        HtmlFormatter.formatUnits(trade.getUnits()),
                        HtmlFormatter.formatMoney(trade.getUnitPrice(), currency, 2),
                        HtmlFormatter.formatMoney(trade.getSaleValue(), currency, 0),
                        HtmlFormatter.formatMoney(trade.getCostBasis(), currency, 0),
                        HtmlFormatter.formatMoney(trade.getGainLoss(), currency, 0),
                        HtmlFormatter.formatPercent(trade.getReturnPct(), 2));
            }

                double totalUnits = saleTrades.stream().mapToDouble(Security.SaleTrade::getUnits).sum();
                double totalSaleValue = saleTrades.stream().mapToDouble(Security.SaleTrade::getSaleValue).sum();
                double totalCostBasis = saleTrades.stream().mapToDouble(Security.SaleTrade::getCostBasis).sum();
                double totalGainLoss = saleTrades.stream().mapToDouble(Security.SaleTrade::getGainLoss).sum();
                double totalReturnPct = totalCostBasis > 0.0 ? (totalGainLoss / totalCostBasis) * 100.0 : 0.0;

                writer.write("<tr class=\"total-row\">\n");
                writer.write("    <td><strong>TOTAL</strong></td>\n");
                writer.write("    <td>" + HtmlFormatter.formatUnits(totalUnits) + "</td>\n");
                writer.write("    <td></td>\n");
                writer.write("    <td>" + HtmlFormatter.formatMoney(totalSaleValue, currency, 0) + "</td>\n");
                writer.write("    <td>" + HtmlFormatter.formatMoney(totalCostBasis, currency, 0) + "</td>\n");
                writer.write("    <td>" + HtmlFormatter.formatMoney(totalGainLoss, currency, 0) + "</td>\n");
                writer.write("    <td>" + HtmlFormatter.formatPercent(totalReturnPct, 2) + "</td>\n");
                writer.write("</tr>\n");

            writer.write("</table>\n</div>\n\n");
        }
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

    private static String mergeCurrencyCodes(String currentCurrencyCode, String nextCurrencyCode) {
        String next = (nextCurrencyCode == null || nextCurrencyCode.isBlank())
                ? "NOK"
                : nextCurrencyCode.trim().toUpperCase(Locale.ROOT);

        if (currentCurrencyCode == null || currentCurrencyCode.isBlank()) {
            return next;
        }
        if ("MIXED".equals(currentCurrencyCode)) {
            return currentCurrencyCode;
        }
        if (currentCurrencyCode.equalsIgnoreCase(next)) {
            return currentCurrencyCode.toUpperCase(Locale.ROOT);
        }
        return "MIXED";
    }

    private static String formatTotalMoney(double value, String aggregateCurrencyCode, int decimals) {
        if (aggregateCurrencyCode == null || aggregateCurrencyCode.isBlank() || "MIXED".equals(aggregateCurrencyCode)) {
            return String.format(Locale.US, "%,." + decimals + "f mixed", value);
        }
        return HtmlFormatter.formatMoney(value, aggregateCurrencyCode, decimals);
    }

    private static String escapeHtml(String text) {
        if (text == null) return "";
        return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\"", "&quot;");
    }
}