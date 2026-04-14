package report;

import csv.TransactionStore;
import model.Events;
import model.Security;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.URI;
import java.net.URL;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.time.Instant;
import java.time.LocalDate;
import java.time.YearMonth;
import java.time.ZoneOffset;
import java.time.Duration;
import java.time.format.DateTimeFormatter;
import java.time.temporal.ChronoUnit;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.*;

public class PortfolioCalculator {

    private static final DateTimeFormatter DATE_FORMATTER = DateTimeFormatter.ofPattern("dd.MM.yyyy");
    private static final String DEFAULT_CURRENCY_CODE = "NOK";
    private static final Pattern YAHOO_TIMESTAMP_ARRAY = Pattern.compile("\\\"timestamp\\\"\\s*:\\s*\\[(.*?)\\]", Pattern.DOTALL);
    private static final Pattern YAHOO_ADJ_CLOSE_ARRAY = Pattern.compile("\\\"adjclose\\\"\\s*:\\s*\\[\\s*\\{\\s*\\\"adjclose\\\"\\s*:\\s*\\[(.*?)\\]", Pattern.DOTALL);
    private static final Pattern YAHOO_CLOSE_ARRAY = Pattern.compile("\\\"close\\\"\\s*:\\s*\\[(.*?)\\]", Pattern.DOTALL);
    private static final Path HISTORY_CACHE_DIR = Path.of(".cache", "history-series-adjclose");
    private static final Duration HISTORY_CACHE_TTL = Duration.ofHours(12);
    private static final int BENCHMARK_FETCH_BUFFER_DAYS = 45;
    private static final int BENCHMARK_BOUNDARY_TOLERANCE_DAYS = 21;
    private static final double DEFAULT_ANALYTICS_RISK_FREE_RATE = 0.02;
    private static final int DEFAULT_MONTE_CARLO_HORIZON_MONTHS = 12;
    private static final int DEFAULT_MONTE_CARLO_ITERATIONS = 2000;

    private static final class PortfolioValuePoint {
        private final LocalDate monthEnd;
        private final double value;
        private final double twrCumulativeReturnPct;

        private PortfolioValuePoint(
                LocalDate monthEnd,
                double value,
                double twrCumulativeReturnPct) {
            this.monthEnd = monthEnd;
            this.value = value;
            this.twrCumulativeReturnPct = twrCumulativeReturnPct;
        }
    }

    private static final class AnalyticsSummary {
        private final boolean hasAnalytics;
        private final double annualizedVolatilityPct;
        private final double sharpeRatio;
        private final boolean hasBeta;
        private final double beta;

        private AnalyticsSummary(
                boolean hasAnalytics,
                double annualizedVolatilityPct,
                double sharpeRatio,
                boolean hasBeta,
                double beta) {
            this.hasAnalytics = hasAnalytics;
            this.annualizedVolatilityPct = annualizedVolatilityPct;
            this.sharpeRatio = sharpeRatio;
            this.hasBeta = hasBeta;
            this.beta = beta;
        }
    }

    private static final class MonteCarloSummary {
        private final boolean hasMonteCarlo;
        private final int horizonMonths;
        private final int iterations;
        private final double startValueNok;
        private final double medianEndValueNok;
        private final double p10EndValueNok;
        private final double p90EndValueNok;
        private final double expectedEndValueNok;

        private MonteCarloSummary(
                boolean hasMonteCarlo,
                int horizonMonths,
                int iterations,
                double startValueNok,
                double medianEndValueNok,
                double p10EndValueNok,
                double p90EndValueNok,
                double expectedEndValueNok) {
            this.hasMonteCarlo = hasMonteCarlo;
            this.horizonMonths = horizonMonths;
            this.iterations = iterations;
            this.startValueNok = startValueNok;
            this.medianEndValueNok = medianEndValueNok;
            this.p10EndValueNok = p10EndValueNok;
            this.p90EndValueNok = p90EndValueNok;
            this.expectedEndValueNok = expectedEndValueNok;
        }
    }

    public static final class StandardAnalyticsSummary {
        public final boolean hasAnalytics;
        public final double annualizedVolatilityPct;
        public final double sharpeRatio;
        public final boolean hasBeta;
        public final double beta;
        public final String benchmarkTicker;
        public final boolean hasMonteCarlo;
        public final int monteCarloHorizonMonths;
        public final int monteCarloIterations;
        public final double monteCarloStartValueNok;
        public final double monteCarloMedianEndValueNok;
        public final double monteCarloP10EndValueNok;
        public final double monteCarloP90EndValueNok;
        public final double monteCarloExpectedEndValueNok;

        private StandardAnalyticsSummary(
                boolean hasAnalytics,
                double annualizedVolatilityPct,
                double sharpeRatio,
                boolean hasBeta,
                double beta,
                String benchmarkTicker,
                boolean hasMonteCarlo,
                int monteCarloHorizonMonths,
                int monteCarloIterations,
                double monteCarloStartValueNok,
                double monteCarloMedianEndValueNok,
                double monteCarloP10EndValueNok,
                double monteCarloP90EndValueNok,
                double monteCarloExpectedEndValueNok) {
            this.hasAnalytics = hasAnalytics;
            this.annualizedVolatilityPct = annualizedVolatilityPct;
            this.sharpeRatio = sharpeRatio;
            this.hasBeta = hasBeta;
            this.beta = beta;
            this.benchmarkTicker = benchmarkTicker;
            this.hasMonteCarlo = hasMonteCarlo;
            this.monteCarloHorizonMonths = monteCarloHorizonMonths;
            this.monteCarloIterations = monteCarloIterations;
            this.monteCarloStartValueNok = monteCarloStartValueNok;
            this.monteCarloMedianEndValueNok = monteCarloMedianEndValueNok;
            this.monteCarloP10EndValueNok = monteCarloP10EndValueNok;
            this.monteCarloP90EndValueNok = monteCarloP90EndValueNok;
            this.monteCarloExpectedEndValueNok = monteCarloExpectedEndValueNok;
        }
    }

    private enum SparklineMetric {
        VALUE,
        RETURN_NOK,
        RETURN_PCT
    }

    public static final class PriceResolution {
        private final double price;
        private final boolean estimated;

        private PriceResolution(double price, boolean estimated) {
            this.price = price;
            this.estimated = estimated;
        }

        public double getPrice() {
            return price;
        }

        public boolean isEstimated() {
            return estimated;
        }
    }

    public static List<OverviewRow> buildOverviewRows(TransactionStore store) {
        List<Security> securities = store.getSecurities();
        List<OverviewRow> rows = new ArrayList<>();

        for (Security security : securities) {
            if (security.getUnitsOwned() <= 0.0000001) continue;
            if (isReplacedSecurity(security, store)) continue;
            if (isTemporaryRightsSecurity(security)) continue;

            double units = security.getUnitsOwned();
            double avgCost = security.getAverageCost();
            double latestPrice = security.getLatestPrice();
            double previousClose = security.getPreviousClose();
            boolean hasDayChangePct = security.hasDayChangePct();
            double dayChangePct = hasDayChangePct ? security.getDayChangePct() : 0.0;
            double positionCostBasis = units * avgCost;
                double realizedCostBasis = security.getRealizedCostBasis();
                double historicalCostBasis = positionCostBasis + realizedCostBasis;
                boolean hasPrice = latestPrice > 0.0;
                double marketValue = hasPrice ? units * latestPrice : 0.0;
                double unrealized = hasPrice ? (marketValue - positionCostBasis) : 0.0;
                double unrealizedPct = hasPrice && positionCostBasis > 0 ? (unrealized / positionCostBasis) * 100.0 : 0.0;

            double realized = security.getRealizedGain();
            double dividends = security.getDividends();
            double totalReturn = unrealized + realized + dividends;
                double totalReturnPct = historicalCostBasis > 0 ? (totalReturn / historicalCostBasis) * 100.0 : 0.0;

            OverviewRow row = new OverviewRow(
                    getTrackingSecurityKey(security),
                    getTickerText(security),
                    getPreferredSecurityName(security),
                    dayChangePct,
                    previousClose,
                    hasDayChangePct,
                    security.getAssetType().name(),
                    security.getResolvedSector(),
                    security.getResolvedSectorWeights(),
                    security.getResolvedRegion(),
                    security.getResolvedRegionWeights(),
                    security.getCurrencyCode(),
                    security.getRealizedReturnPctAsText(),
                    units,
                    avgCost,
                    latestPrice,
                    positionCostBasis,
                    realizedCostBasis,
                    historicalCostBasis,
                    marketValue,
                    unrealized,
                    unrealizedPct,
                    realized,
                    dividends,
                    totalReturn,
                    totalReturnPct,
                        hasPrice
            );

            rows.add(row);
        }

        rows.sort(Comparator
            .comparingInt((OverviewRow r) -> getAssetPriority(r.assetType))
            .thenComparing(r -> r.securityDisplayName, String.CASE_INSENSITIVE_ORDER));

        return rows;
    }

    public static HeaderSummary buildHeaderSummary(TransactionStore store, List<OverviewRow> overviewRows, Map<String, Double> ratesToNok) {
        int holdingsCount = overviewRows.size();
        double totalMarketValue = 0.0;
        double cashHoldings = store.getCurrentCashHoldings();
        double totalReturn = 0.0;
        double totalHistoricalCostBasis = 0.0;
        String totalCurrencyCode = DEFAULT_CURRENCY_CODE;

        OverviewRow best = null;
        OverviewRow worst = null;
        String bestLabel = "-";
        String bestCurrencyCode = DEFAULT_CURRENCY_CODE;
        double bestReturn = 0.0;
        double bestReturnPct = 0.0;
        String worstLabel = "-";
        String worstCurrencyCode = DEFAULT_CURRENCY_CODE;
        double worstReturn = 0.0;
        double worstReturnPct = 0.0;
        double bestReturnNok = Double.NEGATIVE_INFINITY;
        double worstReturnNok = Double.POSITIVE_INFINITY;

        for (OverviewRow row : overviewRows) {
            double marketValueNok = convertAmountToNok(row.marketValue, row.currencyCode, ratesToNok);
            double totalReturnNok = convertAmountToNok(row.totalReturn, row.currencyCode, ratesToNok);
            double historicalCostBasisNok = convertAmountToNok(row.historicalCostBasis, row.currencyCode, ratesToNok);

            totalMarketValue += marketValueNok;
            totalReturn += totalReturnNok;
            totalHistoricalCostBasis += historicalCostBasisNok;

            if (best == null || totalReturnNok > bestReturnNok) {
                best = row;
                bestReturnNok = totalReturnNok;
            }
            if (worst == null || totalReturnNok < worstReturnNok) {
                worst = row;
                worstReturnNok = totalReturnNok;
            }
        }

        if (best != null) {
            bestLabel = getOverviewRowLabel(best);
            bestCurrencyCode = DEFAULT_CURRENCY_CODE;
            bestReturn = Double.isFinite(bestReturnNok) ? bestReturnNok : 0.0;
            bestReturnPct = best.totalReturnPct;
        }
        if (worst != null) {
            worstLabel = getOverviewRowLabel(worst);
            worstCurrencyCode = DEFAULT_CURRENCY_CODE;
            worstReturn = Double.isFinite(worstReturnNok) ? worstReturnNok : 0.0;
            worstReturnPct = worst.totalReturnPct;
        }

        double totalReturnPct = totalHistoricalCostBasis > 0 ? (totalReturn / totalHistoricalCostBasis) * 100.0 : 0.0;
        String sparklineSvg = buildPortfolioValueSparklineWidget(store, ratesToNok);

        return new HeaderSummary(
                LocalDate.now().format(DATE_FORMATTER),
                store.getLoadedCsvFileCount(),
                store.getLoadedTransactionRowCount(),
                holdingsCount,
                totalCurrencyCode,
                totalMarketValue,
                cashHoldings,
                totalReturn,
                totalReturnPct,
                bestLabel,
                bestCurrencyCode,
                bestReturn,
                bestReturnPct,
                worstLabel,
                worstCurrencyCode,
                worstReturn,
                worstReturnPct,
                sparklineSvg
        );
    }

    private static double convertAmountToNok(double amount, String currencyCode, Map<String, Double> ratesToNok) {
        if (!Double.isFinite(amount)) {
            return 0.0;
        }

        String normalized = normalizeCurrencyCode(currencyCode);
        double rateToNok = ratesToNok == null ? 0.0 : ratesToNok.getOrDefault(normalized, 0.0);
        if (!Double.isFinite(rateToNok) || rateToNok <= 0.0) {
            rateToNok = ratesToNok == null ? 1.0 : ratesToNok.getOrDefault(DEFAULT_CURRENCY_CODE, 1.0);
        }
        if (!Double.isFinite(rateToNok) || rateToNok <= 0.0) {
            rateToNok = 1.0;
        }

        return amount * rateToNok;
    }

    private static String buildPortfolioValueSparklineWidget(TransactionStore store, Map<String, Double> ratesToNok) {
        ArrayList<PortfolioValuePoint> allPoints = buildPortfolioValueTimeline(store, ratesToNok, 60);
        if (allPoints.isEmpty()) {
            return "";
        }

        LinkedHashMap<String, ArrayList<PortfolioValuePoint>> byRange = buildStandardSparklineRanges(allPoints);

        String defaultRange = byRange.containsKey("1Y") ? "1Y" : byRange.keySet().iterator().next();
        String defaultMetric = "value";

        StringBuilder html = new StringBuilder();
        html.append("<div class=\"sparkline-widget\">\n");
        html.append("<div class=\"sparkline-metric-controls\">\n");
        html.append("<button type=\"button\" class=\"sparkline-metric-btn is-active\" data-metric=\"value\">Value</button>\n");
        html.append("<button type=\"button\" class=\"sparkline-metric-btn js-return-amount-label\" data-metric=\"return-nok\">Return (NOK)</button>\n");
        html.append("<button type=\"button\" class=\"sparkline-metric-btn\" data-metric=\"return-pct\">Return (%)</button>\n");
        html.append("</div>\n");

        html.append("<div class=\"sparkline-controls\">\n");
        for (String range : byRange.keySet()) {
            String active = range.equals(defaultRange) ? " is-active" : "";
            html.append("<button type=\"button\" class=\"sparkline-range-btn")
                .append(active)
                .append("\" data-range=\"")
                .append(range)
                .append("\">")
                .append(range)
                .append("</button>\n");
        }
        html.append("</div>\n");

        for (Map.Entry<String, ArrayList<PortfolioValuePoint>> entry : byRange.entrySet()) {
            String range = entry.getKey();
            ArrayList<PortfolioValuePoint> points = entry.getValue();

            html.append("<div class=\"sparkline-panel")
                .append(range.equals(defaultRange) && "value".equals(defaultMetric) ? " is-active" : "")
                .append("\" data-range=\"")
                .append(range)
                .append("\" data-metric=\"value\">\n")
                .append(buildPortfolioValueSparkline(points, SparklineMetric.VALUE))
                .append("</div>\n");

            html.append("<div class=\"sparkline-panel")
                .append(range.equals(defaultRange) && "return-nok".equals(defaultMetric) ? " is-active" : "")
                .append("\" data-range=\"")
                .append(range)
                .append("\" data-metric=\"return-nok\">\n")
                .append(buildPortfolioValueSparkline(points, SparklineMetric.RETURN_NOK))
                .append("</div>\n");

            html.append("<div class=\"sparkline-panel")
                .append(range.equals(defaultRange) && "return-pct".equals(defaultMetric) ? " is-active" : "")
                .append("\" data-range=\"")
                .append(range)
                .append("\" data-metric=\"return-pct\">\n")
                .append(buildPortfolioValueSparkline(points, SparklineMetric.RETURN_PCT))
                .append("</div>\n");
        }

        html.append("</div>\n");
        return html.toString();
    }

    private static LinkedHashMap<String, ArrayList<PortfolioValuePoint>> buildStandardSparklineRanges(ArrayList<PortfolioValuePoint> allPoints) {
        LinkedHashMap<String, ArrayList<PortfolioValuePoint>> byRange = new LinkedHashMap<>();
        if (allPoints == null || allPoints.isEmpty()) {
            return byRange;
        }

        byRange.put("1M", takeLastPoints(allPoints, 2));
        byRange.put("3M", takeLastPoints(allPoints, 4));
        byRange.put("6M", takeLastPoints(allPoints, 7));
        byRange.put("1Y", takeLastPoints(allPoints, 12));
        byRange.put("YTD", takeYearToDatePoints(allPoints));
        byRange.put("3Y", takeLastPoints(allPoints, 36));
        byRange.put("5Y", takeLastPoints(allPoints, 60));

        byRange.entrySet().removeIf(entry -> entry.getValue() == null || entry.getValue().isEmpty());
        if (byRange.isEmpty()) {
            byRange.put("ALL", new ArrayList<>(allPoints));
        }
        return byRange;
    }

    private static String buildPortfolioValueSparkline(ArrayList<PortfolioValuePoint> points, SparklineMetric metric) {
        if (points == null || points.isEmpty()) {
            return "";
        }

        final String axisColor = "var(--spark-axis,#9eb1c3)";
        final String axisSoftColor = "var(--spark-axis-soft,#b8c7d6)";
        final String textColor = "var(--spark-text,#5c7187)";
        final String gridColor = "var(--spark-grid,#cfdbe6)";
        final String lineColor = "var(--spark-line,#223c55)";
        final String pointColor = "var(--spark-point,#223c55)";

        final double width = 500.0;
        final double height = 124.0;
        final double left = 46.0;
        final double right = 14.0;
        final double top = 8.0;
        final double bottom = 20.0;
        final double plotWidth = width - left - right;
        final double plotHeight = height - top - bottom;

        double[] displayValues = new double[points.size()];
        double baselineValue = points.get(0).value;
        double baselineTwrFactor = 1.0 + (points.get(0).twrCumulativeReturnPct / 100.0);
        if (!Double.isFinite(baselineTwrFactor) || Math.abs(baselineTwrFactor) < 1e-12) {
            baselineTwrFactor = 1.0;
        }
        for (int i = 0; i < points.size(); i++) {
            PortfolioValuePoint point = points.get(i);
            if (metric == SparklineMetric.RETURN_NOK) {
                double currentTwrFactor = 1.0 + (point.twrCumulativeReturnPct / 100.0);
                if (!Double.isFinite(currentTwrFactor)) {
                    currentTwrFactor = baselineTwrFactor;
                }
                displayValues[i] = baselineValue * ((currentTwrFactor / baselineTwrFactor) - 1.0);
            } else if (metric == SparklineMetric.RETURN_PCT) {
                double currentTwrFactor = 1.0 + (point.twrCumulativeReturnPct / 100.0);
                if (!Double.isFinite(currentTwrFactor)) {
                    currentTwrFactor = baselineTwrFactor;
                }
                displayValues[i] = ((currentTwrFactor / baselineTwrFactor) - 1.0) * 100.0;
            } else {
                displayValues[i] = point.value;
            }
        }

        double min = Double.POSITIVE_INFINITY;
        double max = Double.NEGATIVE_INFINITY;
        int n = points.size();
        for (double value : displayValues) {
            min = Math.min(min, value);
            max = Math.max(max, value);
        }
        if (metric != SparklineMetric.VALUE) {
            min = Math.min(min, 0.0);
            max = Math.max(max, 0.0);
        }
        if (!Double.isFinite(min) || !Double.isFinite(max) || Math.abs(max - min) < 1e-9) {
            max += 1.0;
            min -= 1.0;
        }

        double[] xValues = new double[n];
        double[] yValues = new double[n];
        for (int i = 0; i < n; i++) {
            double x = left + (plotWidth * i / Math.max(1.0, n - 1.0));
            double y = mapValueToY(displayValues[i], min, max, top, plotHeight);
            xValues[i] = x;
            yValues[i] = y;
        }

        StringBuilder svg = new StringBuilder();
        svg.append("<svg viewBox=\"0 0 ").append(String.format(Locale.US, "%.0f %.0f", width, height))
                .append("\" xmlns=\"http://www.w3.org/2000/svg\" role=\"img\">");

        double axisY = top + plotHeight;
        svg.append("<line x1=\"").append(svgNumber(left)).append("\" y1=\"").append(svgNumber(top))
                .append("\" x2=\"").append(svgNumber(left)).append("\" y2=\"").append(svgNumber(axisY))
            .append("\" stroke=\"").append(axisColor).append("\" stroke-width=\"1\"/>");
        svg.append("<line x1=\"").append(svgNumber(left)).append("\" y1=\"").append(svgNumber(axisY))
                .append("\" x2=\"").append(svgNumber(left + plotWidth)).append("\" y2=\"").append(svgNumber(axisY))
            .append("\" stroke=\"").append(axisColor).append("\" stroke-width=\"1\"/>");

        if (metric == SparklineMetric.RETURN_PCT) {
            svg.append("<text x=\"").append(svgNumber(left - 4.0)).append("\" y=\"").append(svgNumber(top + 1.0))
                .append("\" text-anchor=\"end\" dominant-baseline=\"hanging\" font-size=\"8\" fill=\"").append(textColor).append("\">")
                .append(formatCompactPercent(max))
                .append("</text>");
            svg.append("<text x=\"").append(svgNumber(left - 4.0)).append("\" y=\"").append(svgNumber(axisY))
                .append("\" text-anchor=\"end\" dominant-baseline=\"middle\" font-size=\"8\" fill=\"").append(textColor).append("\">")
                .append(formatCompactPercent(min))
                .append("</text>");
        } else {
            svg.append("<text x=\"").append(svgNumber(left - 4.0)).append("\" y=\"").append(svgNumber(top + 1.0))
                .append("\" class=\"js-chart-money\" data-value-nok=\"").append(svgNumber(max)).append("\" data-format=\"compact\" text-anchor=\"end\" dominant-baseline=\"hanging\" font-size=\"8\" fill=\"").append(textColor).append("\">")
                .append(formatCompactKroner(max))
                .append("</text>");
            svg.append("<text x=\"").append(svgNumber(left - 4.0)).append("\" y=\"").append(svgNumber(axisY))
                .append("\" class=\"js-chart-money\" data-value-nok=\"").append(svgNumber(min)).append("\" data-format=\"compact\" text-anchor=\"end\" dominant-baseline=\"middle\" font-size=\"8\" fill=\"").append(textColor).append("\">")
                .append(formatCompactKroner(min))
                .append("</text>");
        }

        svg.append("<line x1=\"").append(svgNumber(left)).append("\" y1=\"").append(svgNumber(top))
                .append("\" x2=\"").append(svgNumber(left + plotWidth)).append("\" y2=\"").append(svgNumber(top))
                .append("\" stroke=\"").append(gridColor).append("\" stroke-width=\"0.8\" opacity=\"0.7\" stroke-dasharray=\"3 3\"/>");

        StringBuilder path = new StringBuilder();
        for (int i = 0; i < n; i++) {
            path.append(i == 0 ? "M " : " L ")
                    .append(svgNumber(xValues[i])).append(" ").append(svgNumber(yValues[i]));
        }

        svg.append("<path d=\"").append(path).append("\" fill=\"none\" stroke=\"").append(lineColor).append("\" stroke-width=\"2.2\" stroke-linecap=\"round\"/>");

        DateTimeFormatter axisMonthFormat = DateTimeFormatter.ofPattern("MMM", Locale.ENGLISH);
        int labelStep = Math.max(1, (int) Math.ceil(n / 10.0));
        for (int i = 0; i < n; i++) {
            String monthLabel = points.get(i).monthEnd.format(axisMonthFormat);
            svg.append("<circle class=\"chart-hover-target chart-hover-point\" cx=\"").append(svgNumber(xValues[i])).append("\" cy=\"").append(svgNumber(yValues[i]))
                .append("\" r=\"2.2\" fill=\"").append(pointColor).append("\">");
            if (metric == SparklineMetric.RETURN_PCT) {
                svg.append("<title>")
                    .append(monthLabel).append(": ").append(formatSparklineTooltipValue(metric, displayValues[i]))
                    .append("</title>");
            } else {
                String prefix = escapeHtmlAttribute(monthLabel + ": ");
                svg.append("<title class=\"js-chart-money\" data-value-nok=\"").append(svgNumber(displayValues[i]))
                    .append("\" data-format=\"compact\" data-prefix=\"").append(prefix).append("\">")
                    .append(monthLabel).append(": ").append(formatCompactKroner(displayValues[i]))
                    .append("</title>");
            }
            svg.append("</circle>");

            boolean isFirst = i == 0;
            boolean isLast = i == n - 1;
            boolean showXAxisLabel = isFirst || isLast || (i % labelStep == 0);

            if (showXAxisLabel) {
                svg.append("<line x1=\"").append(svgNumber(xValues[i])).append("\" y1=\"").append(svgNumber(axisY))
                        .append("\" x2=\"").append(svgNumber(xValues[i])).append("\" y2=\"").append(svgNumber(axisY + 2.8))
                        .append("\" stroke=\"").append(axisSoftColor).append("\" stroke-width=\"0.8\"/>");
            }

            String tickAnchor = "middle";
            double tickLabelX = xValues[i];
            if (isFirst) {
                tickAnchor = "start";
                tickLabelX = Math.max(tickLabelX, left + 1.0);
            } else if (isLast) {
                tickAnchor = "end";
                tickLabelX = Math.min(tickLabelX, left + plotWidth - 1.0);
            }

            if (showXAxisLabel) {
                svg.append("<text x=\"").append(svgNumber(tickLabelX)).append("\" y=\"").append(svgNumber(axisY + 13.0))
                        .append("\" text-anchor=\"").append(tickAnchor).append("\" font-size=\"7\" fill=\"").append(textColor).append("\">")
                        .append(monthLabel)
                        .append("</text>");
            }
        }

        svg.append("</svg>");
        return svg.toString();
    }

    private static double mapValueToY(double value, double minValue, double maxValue, double chartTop, double chartHeight) {
        if (Math.abs(maxValue - minValue) < 1e-12) {
            return chartTop + chartHeight / 2.0;
        }
        return chartTop + ((maxValue - value) / (maxValue - minValue)) * chartHeight;
    }

    private static String svgNumber(double value) {
        return String.format(Locale.US, "%.2f", value);
    }

    private static String formatCompactKroner(double value) {
        double absValue = Math.abs(value);
        String prefix = value < 0.0 ? "-" : "";
        if (absValue >= 1_000_000_000.0) {
            return prefix + String.format(Locale.US, "%.1f", absValue / 1_000_000_000.0) + "B NOK";
        }
        if (absValue >= 1_000_000.0) {
            return prefix + String.format(Locale.US, "%.1f", absValue / 1_000_000.0) + "M NOK";
        }
        if (absValue >= 1_000.0) {
            return prefix + String.format(Locale.US, "%.0f", absValue / 1_000.0) + "k NOK";
        }
        return prefix + String.format(Locale.US, "%.0f", absValue) + " NOK";
    }

    private static String formatCompactPercent(double value) {
        return String.format(Locale.US, "%.1f%%", value);
    }

    private static String formatSparklineTooltipValue(SparklineMetric metric, double value) {
        if (metric == SparklineMetric.RETURN_PCT) {
            return String.format(Locale.US, "%.2f%%", value);
        }
        return formatCompactKroner(value);
    }

    private static String escapeHtmlAttribute(String text) {
        if (text == null) {
            return "";
        }
        return text
                .replace("&", "&amp;")
                .replace("\"", "&quot;")
                .replace("<", "&lt;")
                .replace(">", "&gt;");
    }

    private static ArrayList<PortfolioValuePoint> takeLastPoints(ArrayList<PortfolioValuePoint> allPoints, int count) {
        ArrayList<PortfolioValuePoint> selected = new ArrayList<>();
        if (allPoints == null || allPoints.isEmpty()) {
            return selected;
        }

        int safeCount = Math.max(2, count);
        int from = Math.max(0, allPoints.size() - safeCount);
        for (int i = from; i < allPoints.size(); i++) {
            selected.add(allPoints.get(i));
        }

        if (selected.size() == 1 && allPoints.size() > 1) {
            selected.add(0, allPoints.get(Math.max(0, from - 1)));
        }
        return selected;
    }

    private static ArrayList<PortfolioValuePoint> takeYearToDatePoints(ArrayList<PortfolioValuePoint> allPoints) {
        ArrayList<PortfolioValuePoint> selected = new ArrayList<>();
        if (allPoints == null || allPoints.isEmpty()) {
            return selected;
        }

        int currentYear = LocalDate.now().getYear();
        for (PortfolioValuePoint point : allPoints) {
            if (point.monthEnd.getYear() == currentYear) {
                selected.add(point);
            }
        }

        if (selected.size() < 2) {
            return takeLastPoints(allPoints, 2);
        }
        return selected;
    }

    private static ArrayList<PortfolioValuePoint> buildPortfolioValueTimeline(TransactionStore store, Map<String, Double> ratesToNok, int months) {
        ArrayList<PortfolioValuePoint> timeline = new ArrayList<>();
        List<Events.UnitEvent> unitEvents = store.getUnitEvents();
        List<Events.CashEvent> cashEvents = store.getCashEvents();
        List<Events.PortfolioCashSnapshot> portfolioCashSnapshots = store.getPortfolioCashSnapshots();
        if (unitEvents.isEmpty() && cashEvents.isEmpty() && portfolioCashSnapshots.isEmpty()) {
            return timeline;
        }

        boolean useAuthoritativeSnapshots = !portfolioCashSnapshots.isEmpty();

        ArrayList<Events.UnitEvent> sortedUnitEvents = new ArrayList<>(unitEvents);
        sortedUnitEvents.sort(Comparator.comparing(Events.UnitEvent::tradeDate)
                .thenComparing(Events.UnitEvent::securityKey));

        ArrayList<Events.CashEvent> sortedCashEvents = new ArrayList<>(cashEvents);
        sortedCashEvents.sort(Comparator.comparing(Events.CashEvent::tradeDate));

        ArrayList<Events.PortfolioCashSnapshot> sortedCashSnapshots = new ArrayList<>(portfolioCashSnapshots);
        sortedCashSnapshots.sort(Comparator.comparing(Events.PortfolioCashSnapshot::tradeDate)
                .thenComparingLong(Events.PortfolioCashSnapshot::sortId));

        Map<String, Security> securitiesByTrackingKey = new HashMap<>();
        for (Security security : store.getSecurities()) {
            securitiesByTrackingKey.put(getTrackingSecurityKey(security), security);
        }

        Map<String, Double> unitsBySecurity = new HashMap<>();
        Map<String, NavigableMap<LocalDate, Double>> priceSeriesCache = new HashMap<>();

        int monthCount = Math.max(2, months);
        YearMonth endMonth = YearMonth.now();
        YearMonth startMonth = endMonth.minusMonths(monthCount - 1L);

        ArrayList<LocalDate> monthEnds = new ArrayList<>(monthCount);
        for (int monthOffset = 0; monthOffset < monthCount; monthOffset++) {
            monthEnds.add(startMonth.plusMonths(monthOffset).atEndOfMonth());
        }

        LocalDate startDate = startMonth.atDay(1);
        LocalDate endDate = monthEnds.get(monthEnds.size() - 1);

        int unitEventIndex = 0;
        int cashEventIndex = 0;
        int[] cashSnapshotIndexRef = new int[]{0};
        double runningCash = 0.0;
        LinkedHashMap<String, Double> latestBalanceByPortfolio = new LinkedHashMap<>();

        while (unitEventIndex < sortedUnitEvents.size()
                && sortedUnitEvents.get(unitEventIndex).tradeDate().isBefore(startDate)) {
            Events.UnitEvent event = sortedUnitEvents.get(unitEventIndex);
            unitsBySecurity.merge(event.securityKey(), event.unitsDelta(), Double::sum);
            unitEventIndex++;
        }

        while (cashEventIndex < sortedCashEvents.size()
                && sortedCashEvents.get(cashEventIndex).tradeDate().isBefore(startDate)) {
            Events.CashEvent cashEvent = sortedCashEvents.get(cashEventIndex);
            if (!useAuthoritativeSnapshots) {
                runningCash += cashEvent.cashDelta();
            }
            cashEventIndex++;
        }

        if (useAuthoritativeSnapshots) {
            sumPortfolioBalancesOnOrBefore(
                    startDate.minusDays(1),
                    sortedCashSnapshots,
                    cashSnapshotIndexRef,
                    latestBalanceByPortfolio
            );
        }

        int monthEndIndex = 0;
        double previousDayValue = Double.NaN;
        double previousSnapshotCash = Double.NaN;
        double twrFactor = 1.0;

        for (LocalDate date = startDate; !date.isAfter(endDate); date = date.plusDays(1)) {
            double estimatedTradeCashDeltaToday = 0.0;
            while (unitEventIndex < sortedUnitEvents.size()
                    && !sortedUnitEvents.get(unitEventIndex).tradeDate().isAfter(date)) {
                Events.UnitEvent event = sortedUnitEvents.get(unitEventIndex);
                Security security = securitiesByTrackingKey.get(event.securityKey());
                if (security != null) {
                    estimatedTradeCashDeltaToday += estimateUnitEventCashDeltaNok(
                            event,
                            security,
                            ratesToNok,
                            priceSeriesCache
                    );
                }
                unitsBySecurity.merge(event.securityKey(), event.unitsDelta(), Double::sum);
                unitEventIndex++;
            }

            double externalFlowToday = 0.0;
            double cashEventDeltaToday = 0.0;
            double nonExternalCashEventDeltaToday = 0.0;
            while (cashEventIndex < sortedCashEvents.size()
                    && !sortedCashEvents.get(cashEventIndex).tradeDate().isAfter(date)) {
                Events.CashEvent cashEvent = sortedCashEvents.get(cashEventIndex);
                if (!useAuthoritativeSnapshots) {
                    runningCash += cashEvent.cashDelta();
                }
                cashEventDeltaToday += cashEvent.cashDelta();
                if (cashEvent.externalFlow()) {
                    externalFlowToday += cashEvent.cashDelta();
                } else {
                    nonExternalCashEventDeltaToday += cashEvent.cashDelta();
                }
                cashEventIndex++;
            }

            double uncoveredTradeCashDelta = estimatedTradeCashDeltaToday - nonExternalCashEventDeltaToday;
            double tradeInferenceThreshold = Math.max(250.0, Math.abs(estimatedTradeCashDeltaToday) * 0.02);
            if (Double.isFinite(uncoveredTradeCashDelta) && Math.abs(uncoveredTradeCashDelta) > tradeInferenceThreshold) {
                externalFlowToday += -uncoveredTradeCashDelta;
            }

            double balanceSnapshotCash = sumPortfolioBalancesOnOrBefore(
                    date,
                    sortedCashSnapshots,
                    cashSnapshotIndexRef,
                    latestBalanceByPortfolio
            );

            if (useAuthoritativeSnapshots && Double.isFinite(previousSnapshotCash)) {
                double snapshotDelta = balanceSnapshotCash - previousSnapshotCash;
                double inferredExternalFlow = snapshotDelta - cashEventDeltaToday;
                double inferenceThreshold = Math.max(100.0, Math.abs(balanceSnapshotCash) * 0.0005);
                if (Double.isFinite(inferredExternalFlow) && Math.abs(inferredExternalFlow) > inferenceThreshold) {
                    externalFlowToday += inferredExternalFlow;
                }
            }

                        double holdingsValue = calculateHoldingsMarketValueNok(
                        unitsBySecurity,
                        securitiesByTrackingKey,
                        ratesToNok,
                        date,
                        priceSeriesCache
                    );

                    double totalValue = useAuthoritativeSnapshots
                        ? balanceSnapshotCash
                        : runningCash + balanceSnapshotCash;
                    totalValue += holdingsValue;

            if (Double.isFinite(previousDayValue)) {
                if (previousDayValue > 1e-9) {
                    double periodReturn = (totalValue - previousDayValue - externalFlowToday) / previousDayValue;
                    if (Double.isFinite(periodReturn) && periodReturn > -0.999999999) {
                        twrFactor *= (1.0 + periodReturn);
                    }
                }
            }

            previousDayValue = totalValue;
            if (useAuthoritativeSnapshots) {
                previousSnapshotCash = balanceSnapshotCash;
            }

            if (monthEndIndex < monthEnds.size() && date.equals(monthEnds.get(monthEndIndex))) {
                double twrPct = (twrFactor - 1.0) * 100.0;
                timeline.add(new PortfolioValuePoint(date, totalValue, twrPct));
                monthEndIndex++;
            }
        }

        return timeline;
    }

    private static double calculateHoldingsMarketValueNok(
            Map<String, Double> unitsBySecurity,
            Map<String, Security> securitiesByTrackingKey,
            Map<String, Double> ratesToNok,
            LocalDate valuationDate,
            Map<String, NavigableMap<LocalDate, Double>> priceSeriesCache) {

        double totalValue = 0.0;
        for (Map.Entry<String, Double> entry : unitsBySecurity.entrySet()) {
            double units = entry.getValue();
            if (units <= 0.0000001) {
                continue;
            }

            Security security = securitiesByTrackingKey.get(entry.getKey());
            if (security == null) {
                continue;
            }

            double price = resolveHistoricalPrice(security, valuationDate, priceSeriesCache);
            if (price <= 0.0) {
                continue;
            }

            double positionValue = units * price;
            String securityCurrency = normalizeCurrencyCode(security.getCurrencyCode());
            double rateToNok = ratesToNok == null ? 0.0 : ratesToNok.getOrDefault(securityCurrency, 0.0);
            if (rateToNok <= 0.0) {
                rateToNok = ratesToNok == null ? 1.0 : ratesToNok.getOrDefault(DEFAULT_CURRENCY_CODE, 1.0);
            }
            totalValue += positionValue * rateToNok;
        }
        return totalValue;
    }

    private static double estimateUnitEventCashDeltaNok(
            Events.UnitEvent unitEvent,
            Security security,
            Map<String, Double> ratesToNok,
            Map<String, NavigableMap<LocalDate, Double>> priceSeriesCache) {

        if (unitEvent == null || security == null) {
            return 0.0;
        }

        double price = resolveHistoricalPrice(security, unitEvent.tradeDate(), priceSeriesCache);
        if (price <= 0.0 || !Double.isFinite(price)) {
            return 0.0;
        }

        String securityCurrency = normalizeCurrencyCode(security.getCurrencyCode());
        double rateToNok = ratesToNok == null ? 0.0 : ratesToNok.getOrDefault(securityCurrency, 0.0);
        if (rateToNok <= 0.0 || !Double.isFinite(rateToNok)) {
            rateToNok = ratesToNok == null ? 1.0 : ratesToNok.getOrDefault(DEFAULT_CURRENCY_CODE, 1.0);
        }

        double tradeNotionalNok = unitEvent.unitsDelta() * price * rateToNok;
        if (!Double.isFinite(tradeNotionalNok)) {
            return 0.0;
        }

        // Buying units should reduce cash, selling units should increase cash.
        return -tradeNotionalNok;
    }

    private static double sumPortfolioBalancesOnOrBefore(
            LocalDate monthEnd,
            ArrayList<Events.PortfolioCashSnapshot> sortedSnapshots,
            int[] snapshotIndexRef,
            LinkedHashMap<String, Double> latestBalanceByPortfolio) {

        int snapshotIndex = snapshotIndexRef[0];
        while (snapshotIndex < sortedSnapshots.size()
                && !sortedSnapshots.get(snapshotIndex).tradeDate().isAfter(monthEnd)) {
            Events.PortfolioCashSnapshot snapshot = sortedSnapshots.get(snapshotIndex);
            if (snapshot.portfolioId() != null && !snapshot.portfolioId().isBlank()) {
                latestBalanceByPortfolio.put(snapshot.portfolioId(), snapshot.balance());
            }
            snapshotIndex++;
        }
        snapshotIndexRef[0] = snapshotIndex;

        double cash = 0.0;
        for (double balance : latestBalanceByPortfolio.values()) {
            cash += balance;
        }
        return cash;
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

    private static double resolveHistoricalPrice(Security security, LocalDate monthEnd,
                                                 Map<String, NavigableMap<LocalDate, Double>> priceSeriesCache) {
        return resolveHistoricalPriceDetailed(security, monthEnd, priceSeriesCache).getPrice();
    }

    private static PriceResolution resolveHistoricalPriceDetailed(
            Security security,
            LocalDate monthEnd,
            Map<String, NavigableMap<LocalDate, Double>> priceSeriesCache) {

        if (security == null || monthEnd == null) {
            return new PriceResolution(0.0, true);
        }

        String ticker = security.getTicker();
        if (ticker == null || ticker.isBlank() || "-".equals(ticker)) {
            double latestPrice = Math.max(0.0, security.getLatestPrice());
            if (latestPrice > 0.0) {
                return new PriceResolution(latestPrice, true);
            }
            return new PriceResolution(0.0, true);
        }

        NavigableMap<LocalDate, Double> series = priceSeriesCache.computeIfAbsent(
                ticker,
            t -> fetchHistoricalCloseSeries(t, LocalDate.now().minusMonths(180), LocalDate.now())
        );

        if (series.isEmpty()) {
            return resolveFallbackPriceDetailed(security, monthEnd);
        }

        Map.Entry<LocalDate, Double> floor = series.floorEntry(monthEnd);
        if (floor != null && floor.getKey() != null && floor.getValue() != null && floor.getValue() > 0.0) {
            long daysGap = java.time.temporal.ChronoUnit.DAYS.between(floor.getKey(), monthEnd);
            if (daysGap >= 0 && daysGap <= 10) {
                return new PriceResolution(floor.getValue(), false);
            }
        }

        Map.Entry<LocalDate, Double> first = series.firstEntry();
        if (first != null && first.getKey() != null && first.getValue() != null && first.getValue() > 0.0) {
            long daysGap = java.time.temporal.ChronoUnit.DAYS.between(monthEnd, first.getKey());
            if (daysGap >= 0 && daysGap <= 10) {
                return new PriceResolution(first.getValue(), true);
            }
        }

        return resolveFallbackPriceDetailed(security, monthEnd);
    }

    private static PriceResolution resolveFallbackPriceDetailed(Security security, LocalDate monthEnd) {
        if (security == null || monthEnd == null) {
            return new PriceResolution(0.0, true);
        }

        double nearestObservedPrice = security.getClosestObservedTradePriceAround(monthEnd);
        if (nearestObservedPrice > 0.0) {
            return new PriceResolution(nearestObservedPrice, true);
        }

        double tradePriceOnOrBefore = security.getLatestObservedTradePriceOnOrBefore(monthEnd);
        if (tradePriceOnOrBefore > 0.0) {
            return new PriceResolution(tradePriceOnOrBefore, true);
        }

        double tradePriceAfter = security.getEarliestObservedTradePriceAfter(monthEnd);
        if (tradePriceAfter > 0.0) {
            return new PriceResolution(tradePriceAfter, true);
        }

        List<Security.SaleTrade> sales = security.getSaleTradesSortedByDate();
        double latestPastSalePrice = 0.0;
        LocalDate latestPastSaleDate = null;
        for (Security.SaleTrade sale : sales) {
            if (sale == null || sale.getUnitPrice() <= 0.0) {
                continue;
            }
            if (sale.getTradeDate() != null && !sale.getTradeDate().isAfter(monthEnd)) {
                if (latestPastSaleDate == null || sale.getTradeDate().isAfter(latestPastSaleDate)) {
                    latestPastSaleDate = sale.getTradeDate();
                    latestPastSalePrice = sale.getUnitPrice();
                }
            }
        }
        if (latestPastSalePrice > 0.0) {
            return new PriceResolution(latestPastSalePrice, true);
        }

        double averageCost = security.getAverageCost();
        if (averageCost > 0.0) {
            return new PriceResolution(averageCost, true);
        }

        double latestPrice = security.getLatestPrice();
        if (latestPrice > 0.0) {
            return new PriceResolution(latestPrice, true);
        }

        return new PriceResolution(0.0, true);
    }

    private static NavigableMap<LocalDate, Double> fetchHistoricalCloseSeries(String ticker, LocalDate fromDate, LocalDate toDate) {
        NavigableMap<LocalDate, Double> series = new TreeMap<>();
        if (ticker == null || ticker.isBlank() || fromDate == null || toDate == null) {
            return series;
        }

        LocalDate safeFrom = fromDate;
        LocalDate safeTo = toDate;
        if (safeFrom.isAfter(safeTo)) {
            LocalDate tmp = safeFrom;
            safeFrom = safeTo;
            safeTo = tmp;
        }

        NavigableMap<LocalDate, Double> cachedSeries = loadHistoricalSeriesFromCache(ticker);
        if (!cachedSeries.isEmpty() && hasSeriesCoverage(cachedSeries, safeFrom, safeTo)) {
            return cachedSeries;
        }

        NavigableMap<LocalDate, Double> fetchedSeries = fetchHistoricalCloseSeriesFromYahoo(ticker, safeFrom, safeTo);
        if (!fetchedSeries.isEmpty()) {
            if (!cachedSeries.isEmpty()) {
                cachedSeries.putAll(fetchedSeries);
                fetchedSeries = cachedSeries;
            }
            saveHistoricalSeriesToCache(ticker, fetchedSeries);
            return fetchedSeries;
        }

        if (!cachedSeries.isEmpty()) {
            return cachedSeries;
        }

        return series;
    }

    private static boolean hasSeriesCoverage(
            NavigableMap<LocalDate, Double> series,
            LocalDate fromDate,
            LocalDate toDate) {
        if (series == null || series.isEmpty() || fromDate == null || toDate == null) {
            return false;
        }
        LocalDate first = series.firstKey();
        LocalDate last = series.lastKey();
        return !first.isAfter(fromDate) && !last.isBefore(toDate);
    }

    private static NavigableMap<LocalDate, Double> fetchHistoricalCloseSeriesFromYahoo(
            String ticker,
            LocalDate fromDate,
            LocalDate toDate) {
        NavigableMap<LocalDate, Double> series = new TreeMap<>();
        String encodedTicker = URLEncoder.encode(ticker, StandardCharsets.UTF_8);
        long period1 = fromDate.atStartOfDay(ZoneOffset.UTC).toEpochSecond();
        long period2 = toDate.plusDays(1).atStartOfDay(ZoneOffset.UTC).toEpochSecond();
        String[] hosts = {"query1.finance.yahoo.com", "query2.finance.yahoo.com"};

        for (String host : hosts) {
            String urlText = "https://" + host + "/v8/finance/chart/" + encodedTicker
                    + "?period1=" + period1
                    + "&period2=" + period2
                    + "&interval=1d&events=history";
            series = fetchHistoricalCloseSeriesFromUrl(urlText);
            if (!series.isEmpty()) {
                return series;
            }
        }

        return series;
    }

    private static NavigableMap<LocalDate, Double> fetchHistoricalCloseSeriesFromUrl(String urlText) {
        NavigableMap<LocalDate, Double> series = new TreeMap<>();
        HttpURLConnection connection = null;
        try {
            URL url = URI.create(urlText).toURL();
            connection = (HttpURLConnection) url.openConnection();
            connection.setRequestMethod("GET");
            connection.setConnectTimeout(5000);
            connection.setReadTimeout(7000);
            connection.setRequestProperty("Accept", "application/json");
            connection.setRequestProperty("User-Agent", "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)");

            if (connection.getResponseCode() != 200) {
                return series;
            }

            StringBuilder body = new StringBuilder();
            try (BufferedReader reader = new BufferedReader(new InputStreamReader(connection.getInputStream(), StandardCharsets.UTF_8))) {
                String line;
                while ((line = reader.readLine()) != null) {
                    body.append(line);
                }
            }

            Matcher timestampMatcher = YAHOO_TIMESTAMP_ARRAY.matcher(body);
            if (!timestampMatcher.find()) {
                return series;
            }

            Matcher adjCloseMatcher = YAHOO_ADJ_CLOSE_ARRAY.matcher(body);
            Matcher closeMatcher = YAHOO_CLOSE_ARRAY.matcher(body);
            String closeValues;
            if (adjCloseMatcher.find()) {
                closeValues = adjCloseMatcher.group(1);
            } else if (closeMatcher.find()) {
                closeValues = closeMatcher.group(1);
            } else {
                return series;
            }

            String[] timestamps = timestampMatcher.group(1).split(",");
            String[] closes = closeValues.split(",");
            int length = Math.min(timestamps.length, closes.length);

            for (int i = 0; i < length; i++) {
                String tsText = timestamps[i].trim();
                String closeText = closes[i].trim();
                if (tsText.isEmpty() || closeText.isEmpty() || "null".equalsIgnoreCase(closeText)) {
                    continue;
                }

                long epochSeconds;
                double close;
                try {
                    epochSeconds = Long.parseLong(tsText);
                    close = Double.parseDouble(closeText);
                } catch (NumberFormatException ignored) {
                    continue;
                }

                if (!Double.isFinite(close) || close <= 0.0) {
                    continue;
                }

                LocalDate date = Instant.ofEpochSecond(epochSeconds).atZone(ZoneOffset.UTC).toLocalDate();
                series.put(date, close);
            }
        } catch (IOException ignored) {
            return series;
        } finally {
            if (connection != null) {
                connection.disconnect();
            }
        }

        return series;
    }

    private static List<String> buildBenchmarkTickerCandidates(String benchmarkTicker) {
        String baseTicker = (benchmarkTicker == null || benchmarkTicker.isBlank())
                ? "^OSEAX"
                : benchmarkTicker.trim();

        LinkedHashSet<String> candidates = new LinkedHashSet<>();
        candidates.add(baseTicker);

        String withoutCaret = baseTicker.startsWith("^") ? baseTicker.substring(1) : baseTicker;
        if (!withoutCaret.isBlank()) {
            candidates.add(withoutCaret);
            candidates.add("^" + withoutCaret);

            String upper = withoutCaret.toUpperCase(Locale.ROOT);
            if (!upper.endsWith(".OL")) {
                candidates.add(withoutCaret + ".OL");
            }
        }

        return new ArrayList<>(candidates);
    }

    private static NavigableMap<LocalDate, Double> loadHistoricalSeriesFromCache(String ticker) {
        NavigableMap<LocalDate, Double> series = new TreeMap<>();
        Path cacheFile = resolveHistoryCacheFile(ticker);
        if (cacheFile == null || !Files.exists(cacheFile)) {
            return series;
        }

        try {
            Instant lastModified = Files.getLastModifiedTime(cacheFile).toInstant();
            if (lastModified.plus(HISTORY_CACHE_TTL).isBefore(Instant.now())) {
                return series;
            }

            List<String> lines = Files.readAllLines(cacheFile, StandardCharsets.UTF_8);
            for (String line : lines) {
                String trimmed = line == null ? "" : line.trim();
                if (trimmed.isEmpty() || trimmed.startsWith("#")) {
                    continue;
                }

                String[] parts = trimmed.split(",");
                if (parts.length != 2) {
                    continue;
                }

                try {
                    LocalDate date = LocalDate.parse(parts[0]);
                    double close = Double.parseDouble(parts[1]);
                    if (Double.isFinite(close) && close > 0.0) {
                        series.put(date, close);
                    }
                } catch (Exception ignored) {
                    // Ignore malformed cache rows.
                }
            }
        } catch (IOException ignored) {
            return new TreeMap<>();
        }

        return series;
    }

    private static void saveHistoricalSeriesToCache(String ticker, NavigableMap<LocalDate, Double> series) {
        if (series == null || series.isEmpty()) {
            return;
        }

        Path cacheFile = resolveHistoryCacheFile(ticker);
        if (cacheFile == null) {
            return;
        }

        try {
            Files.createDirectories(HISTORY_CACHE_DIR);
            List<String> lines = new ArrayList<>();
            lines.add("# date,close");
            for (Map.Entry<LocalDate, Double> entry : series.entrySet()) {
                if (entry.getKey() == null || entry.getValue() == null || entry.getValue() <= 0.0) {
                    continue;
                }
                lines.add(entry.getKey().toString() + "," + String.format(Locale.US, "%.8f", entry.getValue()));
            }
            Files.write(cacheFile, lines, StandardCharsets.UTF_8,
                    StandardOpenOption.CREATE, StandardOpenOption.TRUNCATE_EXISTING, StandardOpenOption.WRITE);
        } catch (IOException ignored) {
            // Cache persistence should not fail report generation.
        }
    }

    private static Path resolveHistoryCacheFile(String ticker) {
        if (ticker == null || ticker.isBlank()) {
            return null;
        }
        String safe = ticker.trim().replaceAll("[^a-zA-Z0-9._-]", "_");
        if (safe.isBlank()) {
            return null;
        }
        return HISTORY_CACHE_DIR.resolve(safe + ".csv");
    }

    private static String getTickerText(Security security) {
        String t = security.getTicker();
        return (t == null || t.isBlank()) ? "-" : t;
    }

    private static String getPreferredSecurityName(Security security) {
        String d = security.getDisplayName();
        return (d == null || d.isBlank()) ? security.getName() : d;
    }

    private static boolean isReplacedSecurity(Security security, TransactionStore store) {
        String isin = security.getIsin();
        if (isin == null || isin.isBlank()) return false;
        String replacement = store.getRenamedSecurityIsin().get(isin);
        if (replacement == null) return false;

        for (Security s : store.getSecurities()) {
            if (replacement.equals(s.getIsin()) && s.getUnitsOwned() > 0.0000001) {
                return true;
            }
        }
        return false;
    }

    private static boolean isTemporaryRightsSecurity(Security security) {
        String name = security.getName();
        if (name == null) return false;
        String upper = name.toUpperCase(Locale.ROOT);
        return upper.contains("T-RETT") || upper.contains("TEGNINGSRETT");
    }

    private static String getOverviewRowLabel(OverviewRow row) {
        if (row.securityDisplayName == null || row.securityDisplayName.isBlank()) {
            return row.tickerText;
        }
        return row.securityDisplayName;
    }

    private static String normalizeCurrencyCode(String currencyCode) {
        if (currencyCode == null || currencyCode.isBlank()) {
            return DEFAULT_CURRENCY_CODE;
        }

        String normalized = currencyCode.trim().toUpperCase(Locale.ROOT);
        if (!normalized.matches("[A-Z]{3}")) {
            return DEFAULT_CURRENCY_CODE;
        }
        return normalized;
    }

    public static AnnualPerformanceSummary buildAnnualPerformanceSummary(
            TransactionStore store,
            Map<String, Double> ratesToNok,
            int year,
            String benchmarkTicker) {

        int safeYear = Math.max(2000, Math.min(2100, year));
        ArrayList<PortfolioValuePoint> timeline = buildPortfolioValueTimeline(store, ratesToNok, 120);

        PortfolioValuePoint firstInYear = null;
        PortfolioValuePoint lastInYear = null;
        for (PortfolioValuePoint point : timeline) {
            if (point.monthEnd.getYear() != safeYear) {
                continue;
            }
            if (firstInYear == null) {
                firstInYear = point;
            }
            lastInYear = point;
        }

        boolean hasPortfolioData = firstInYear != null && lastInYear != null;
        double startValueNok = hasPortfolioData ? firstInYear.value : 0.0;
        double endValueNok = hasPortfolioData ? lastInYear.value : 0.0;
        double portfolioReturnPct = 0.0;
        double portfolioReturnNok = 0.0;
        if (hasPortfolioData) {
            double startFactor = 1.0 + (firstInYear.twrCumulativeReturnPct / 100.0);
            double endFactor = 1.0 + (lastInYear.twrCumulativeReturnPct / 100.0);
            if (Math.abs(startFactor) > 1e-9 && Double.isFinite(startFactor) && Double.isFinite(endFactor)) {
                double yearFactor = endFactor / startFactor;
                portfolioReturnPct = (yearFactor - 1.0) * 100.0;
                portfolioReturnNok = startValueNok * (yearFactor - 1.0);
            }
        }

        double realizedGainNok = 0.0;
        double dividendsNok = 0.0;
        for (Security security : store.getSecurities()) {
            if (security == null) {
                continue;
            }

            String securityCurrency = normalizeCurrencyCode(security.getCurrencyCode());
            double rateToNok = ratesToNok == null ? 0.0 : ratesToNok.getOrDefault(securityCurrency, 0.0);
            if (rateToNok <= 0.0) {
                rateToNok = ratesToNok == null ? 1.0 : ratesToNok.getOrDefault(DEFAULT_CURRENCY_CODE, 1.0);
            }

            for (Security.SaleTrade saleTrade : security.getSaleTradesSortedByDate()) {
                if (saleTrade == null || saleTrade.getTradeDate() == null || saleTrade.getTradeDate().getYear() != safeYear) {
                    continue;
                }
                realizedGainNok += saleTrade.getGainLoss() * rateToNok;
            }

            for (Security.DividendEvent dividendEvent : security.getAllDividendEventsSortedByDate()) {
                if (dividendEvent == null || dividendEvent.getTradeDate() == null || dividendEvent.getTradeDate().getYear() != safeYear) {
                    continue;
                }
                dividendsNok += dividendEvent.getAmount() * rateToNok;
            }
        }
        double realizedTotalNok = realizedGainNok + dividendsNok;

        String safeBenchmarkTicker = (benchmarkTicker == null || benchmarkTicker.isBlank())
            ? "^OSEAX"
            : benchmarkTicker.trim();

        boolean hasBenchmarkData = false;
        double benchmarkReturnPct = 0.0;

        LocalDate from = LocalDate.of(safeYear, 1, 1);
        LocalDate to = LocalDate.of(safeYear, 12, 31);
        String resolvedBenchmarkTicker = safeBenchmarkTicker;
        for (String candidateTicker : buildBenchmarkTickerCandidates(safeBenchmarkTicker)) {
            NavigableMap<LocalDate, Double> benchmarkSeries = fetchHistoricalCloseSeries(
                candidateTicker,
                from.minusDays(BENCHMARK_FETCH_BUFFER_DAYS),
                to.plusDays(BENCHMARK_FETCH_BUFFER_DAYS)
            );

            double startValue = resolveBoundaryValue(
                benchmarkSeries,
                from,
                true,
                BENCHMARK_BOUNDARY_TOLERANCE_DAYS
            );
            double endValue = resolveBoundaryValue(
                benchmarkSeries,
                to,
                false,
                BENCHMARK_BOUNDARY_TOLERANCE_DAYS
            );

            if (startValue > 0.0 && endValue > 0.0) {
                benchmarkReturnPct = ((endValue - startValue) / startValue) * 100.0;
                hasBenchmarkData = true;
                resolvedBenchmarkTicker = candidateTicker;
                break;
            }
        }

        ArrayList<PortfolioValuePoint> yearPoints = buildAnnualYearPoints(store, ratesToNok, safeYear);
        AnalyticsSummary analyticsSummary = buildAnalyticsSummary(yearPoints, resolvedBenchmarkTicker, from, to);

        double simulationStartValueNok = endValueNok > 0.0
                ? endValueNok
                : (lastInYear != null && lastInYear.value > 0.0 ? lastInYear.value : 0.0);
        MonteCarloSummary monteCarloSummary = buildMonteCarloSummary(timeline, simulationStartValueNok, safeYear);

        return new AnnualPerformanceSummary(
                safeYear,
                resolvedBenchmarkTicker,
                hasPortfolioData,
                startValueNok,
                endValueNok,
                portfolioReturnNok,
                portfolioReturnPct,
                realizedGainNok,
                dividendsNok,
                realizedTotalNok,
                hasBenchmarkData,
                benchmarkReturnPct,
                analyticsSummary.hasAnalytics,
                analyticsSummary.annualizedVolatilityPct,
                analyticsSummary.sharpeRatio,
                analyticsSummary.hasBeta,
                analyticsSummary.beta,
                monteCarloSummary.hasMonteCarlo,
                monteCarloSummary.horizonMonths,
                monteCarloSummary.iterations,
                monteCarloSummary.startValueNok,
                monteCarloSummary.medianEndValueNok,
                monteCarloSummary.p10EndValueNok,
                monteCarloSummary.p90EndValueNok,
                monteCarloSummary.expectedEndValueNok
        );
    }

    public static StandardAnalyticsSummary buildStandardAnalyticsSummary(
            TransactionStore store,
            Map<String, Double> ratesToNok,
            String benchmarkTicker) {

        ArrayList<PortfolioValuePoint> timeline = buildPortfolioValueTimeline(store, ratesToNok, 120);
        if (timeline.size() < 3) {
            return new StandardAnalyticsSummary(
                    false, 0.0, 0.0, false, 0.0, benchmarkTicker == null ? "^OSEAX" : benchmarkTicker,
                    false, 0, 0, 0.0, 0.0, 0.0, 0.0, 0.0
            );
        }

        String safeBenchmarkTicker = (benchmarkTicker == null || benchmarkTicker.isBlank())
                ? "^OSEAX"
                : benchmarkTicker.trim();
        LocalDate from = timeline.get(Math.max(0, timeline.size() - 61)).monthEnd;
        LocalDate to = timeline.get(timeline.size() - 1).monthEnd;
        ArrayList<PortfolioValuePoint> recent = new ArrayList<>(timeline.subList(Math.max(0, timeline.size() - 61), timeline.size()));
        AnalyticsSummary analytics = buildAnalyticsSummary(recent, safeBenchmarkTicker, from, to);
        double startValue = recent.get(recent.size() - 1).value;
        MonteCarloSummary monteCarlo = buildMonteCarloSummary(timeline, startValue, to.getYear());

        return new StandardAnalyticsSummary(
                analytics.hasAnalytics,
                analytics.annualizedVolatilityPct,
                analytics.sharpeRatio,
                analytics.hasBeta,
                analytics.beta,
                safeBenchmarkTicker,
                monteCarlo.hasMonteCarlo,
                monteCarlo.horizonMonths,
                monteCarlo.iterations,
                monteCarlo.startValueNok,
                monteCarlo.medianEndValueNok,
                monteCarlo.p10EndValueNok,
                monteCarlo.p90EndValueNok,
                monteCarlo.expectedEndValueNok
        );
    }

    private static AnalyticsSummary buildAnalyticsSummary(
            ArrayList<PortfolioValuePoint> points,
            String benchmarkTicker,
            LocalDate from,
            LocalDate to) {

        ArrayList<Double> portfolioMonthlyReturns = buildMonthlyReturnSeries(points);
        if (portfolioMonthlyReturns.size() < 2) {
            return new AnalyticsSummary(false, 0.0, 0.0, false, 0.0);
        }

        double volatilityMonthly = computeStdDev(portfolioMonthlyReturns);
        double annualizedVolatilityPct = volatilityMonthly * Math.sqrt(12.0) * 100.0;

        double riskFreeAnnual = resolveRiskFreeRateAnnual();
        double riskFreeMonthly = Math.pow(1.0 + riskFreeAnnual, 1.0 / 12.0) - 1.0;
        double avgExcess = 0.0;
        for (double value : portfolioMonthlyReturns) {
            avgExcess += (value - riskFreeMonthly);
        }
        avgExcess /= portfolioMonthlyReturns.size();
        double sharpe = volatilityMonthly > 1e-12
                ? (avgExcess / volatilityMonthly) * Math.sqrt(12.0)
                : 0.0;
        if (!Double.isFinite(sharpe)) {
            sharpe = 0.0;
        }

        boolean hasBeta = false;
        double beta = 0.0;
        if (from != null && to != null) {
            ArrayList<Double> benchmarkMonthlyReturns = buildBenchmarkMonthlyReturnSeries(
                    benchmarkTicker, from, to, points
            );
            int alignedCount = Math.min(portfolioMonthlyReturns.size(), benchmarkMonthlyReturns.size());
            if (alignedCount >= 2) {
                ArrayList<Double> p = new ArrayList<>(portfolioMonthlyReturns.subList(portfolioMonthlyReturns.size() - alignedCount, portfolioMonthlyReturns.size()));
                ArrayList<Double> b = new ArrayList<>(benchmarkMonthlyReturns.subList(benchmarkMonthlyReturns.size() - alignedCount, benchmarkMonthlyReturns.size()));
                double variance = computeVariance(b);
                if (variance > 1e-12) {
                    beta = computeCovariance(p, b) / variance;
                    hasBeta = Double.isFinite(beta);
                    if (!hasBeta) {
                        beta = 0.0;
                    }
                }
            }
        }

        return new AnalyticsSummary(
                true,
                annualizedVolatilityPct,
                sharpe,
                hasBeta,
                beta
        );
    }

    private static MonteCarloSummary buildMonteCarloSummary(
            ArrayList<PortfolioValuePoint> timeline,
            double startValueNok,
            int seedBiasYear) {

        ArrayList<Double> historicalMonthlyReturns = buildMonthlyReturnSeries(timeline);
        int horizonMonths = resolveMonteCarloHorizonMonths();
        int iterations = resolveMonteCarloIterations();
        if (historicalMonthlyReturns.isEmpty() || startValueNok <= 0.0 || horizonMonths <= 0 || iterations <= 0) {
            return new MonteCarloSummary(false, 0, 0, 0.0, 0.0, 0.0, 0.0, 0.0);
        }

        Random random = new Random(1_000_003L + seedBiasYear);
        double[] simulatedEndValues = new double[iterations];
        double expected = 0.0;
        for (int i = 0; i < iterations; i++) {
            double value = startValueNok;
            for (int month = 0; month < horizonMonths; month++) {
                double sampledReturn = historicalMonthlyReturns.get(random.nextInt(historicalMonthlyReturns.size()));
                value *= Math.max(0.0, 1.0 + sampledReturn);
            }
            simulatedEndValues[i] = value;
            expected += value;
        }
        Arrays.sort(simulatedEndValues);
        expected /= iterations;
        double p10 = percentile(simulatedEndValues, 0.10);
        double median = percentile(simulatedEndValues, 0.50);
        double p90 = percentile(simulatedEndValues, 0.90);

        return new MonteCarloSummary(
                true,
                horizonMonths,
                iterations,
                startValueNok,
                median,
                p10,
                p90,
                expected
        );
    }

    private static ArrayList<Double> buildMonthlyReturnSeries(ArrayList<PortfolioValuePoint> points) {
        ArrayList<Double> returns = new ArrayList<>();
        if (points == null || points.size() < 2) {
            return returns;
        }

        for (int i = 1; i < points.size(); i++) {
            PortfolioValuePoint previous = points.get(i - 1);
            PortfolioValuePoint current = points.get(i);
            double prevFactor = 1.0 + (previous.twrCumulativeReturnPct / 100.0);
            double currFactor = 1.0 + (current.twrCumulativeReturnPct / 100.0);
            if (!Double.isFinite(prevFactor) || !Double.isFinite(currFactor) || Math.abs(prevFactor) < 1e-12) {
                continue;
            }
            double monthlyReturn = (currFactor / prevFactor) - 1.0;
            if (Double.isFinite(monthlyReturn)) {
                returns.add(monthlyReturn);
            }
        }
        return returns;
    }

    private static ArrayList<Double> buildBenchmarkMonthlyReturnSeries(
            String benchmarkTicker,
            LocalDate from,
            LocalDate to,
            ArrayList<PortfolioValuePoint> portfolioPoints) {

        ArrayList<Double> returns = new ArrayList<>();
        if (portfolioPoints == null || portfolioPoints.size() < 2 || from == null || to == null) {
            return returns;
        }

        NavigableMap<LocalDate, Double> benchmarkSeries = fetchHistoricalCloseSeries(
                benchmarkTicker,
                from.minusDays(BENCHMARK_FETCH_BUFFER_DAYS),
                to.plusDays(BENCHMARK_FETCH_BUFFER_DAYS)
        );
        if (benchmarkSeries == null || benchmarkSeries.isEmpty()) {
            return returns;
        }

        ArrayList<Double> monthEndValues = new ArrayList<>();
        for (PortfolioValuePoint point : portfolioPoints) {
            double value = resolveBoundaryValue(
                    benchmarkSeries,
                    point.monthEnd,
                    false,
                    BENCHMARK_BOUNDARY_TOLERANCE_DAYS
            );
            monthEndValues.add(value > 0.0 ? value : Double.NaN);
        }

        for (int i = 1; i < monthEndValues.size(); i++) {
            double prev = monthEndValues.get(i - 1);
            double curr = monthEndValues.get(i);
            if (!Double.isFinite(prev) || !Double.isFinite(curr) || prev <= 0.0) {
                continue;
            }
            double periodReturn = (curr / prev) - 1.0;
            if (Double.isFinite(periodReturn)) {
                returns.add(periodReturn);
            }
        }
        return returns;
    }

    private static double computeStdDev(ArrayList<Double> values) {
        if (values == null || values.size() < 2) {
            return 0.0;
        }
        double mean = 0.0;
        for (double value : values) {
            mean += value;
        }
        mean /= values.size();
        double variance = 0.0;
        for (double value : values) {
            double diff = value - mean;
            variance += diff * diff;
        }
        variance /= (values.size() - 1);
        if (!Double.isFinite(variance) || variance < 0.0) {
            return 0.0;
        }
        return Math.sqrt(variance);
    }

    private static double computeVariance(ArrayList<Double> values) {
        double stdDev = computeStdDev(values);
        return stdDev * stdDev;
    }

    private static double computeCovariance(ArrayList<Double> left, ArrayList<Double> right) {
        if (left == null || right == null || left.size() != right.size() || left.size() < 2) {
            return 0.0;
        }
        int n = left.size();
        double meanLeft = 0.0;
        double meanRight = 0.0;
        for (int i = 0; i < n; i++) {
            meanLeft += left.get(i);
            meanRight += right.get(i);
        }
        meanLeft /= n;
        meanRight /= n;

        double covariance = 0.0;
        for (int i = 0; i < n; i++) {
            covariance += (left.get(i) - meanLeft) * (right.get(i) - meanRight);
        }
        covariance /= (n - 1);
        return Double.isFinite(covariance) ? covariance : 0.0;
    }

    private static double percentile(double[] sorted, double quantile) {
        if (sorted == null || sorted.length == 0) {
            return 0.0;
        }
        if (quantile <= 0.0) {
            return sorted[0];
        }
        if (quantile >= 1.0) {
            return sorted[sorted.length - 1];
        }

        double position = quantile * (sorted.length - 1);
        int lower = (int) Math.floor(position);
        int upper = (int) Math.ceil(position);
        if (lower == upper) {
            return sorted[lower];
        }
        double weight = position - lower;
        return sorted[lower] * (1.0 - weight) + sorted[upper] * weight;
    }

    private static double resolveRiskFreeRateAnnual() {
        String raw = System.getProperty("portfolio.analytics.riskFreeRate", "");
        if (raw == null || raw.isBlank()) {
            return DEFAULT_ANALYTICS_RISK_FREE_RATE;
        }
        try {
            double parsed = Double.parseDouble(raw.trim());
            if (!Double.isFinite(parsed)) {
                return DEFAULT_ANALYTICS_RISK_FREE_RATE;
            }
            return Math.max(-0.5, Math.min(1.0, parsed));
        } catch (NumberFormatException ignored) {
            return DEFAULT_ANALYTICS_RISK_FREE_RATE;
        }
    }

    private static int resolveMonteCarloHorizonMonths() {
        String raw = System.getProperty("portfolio.analytics.montecarlo.horizonMonths", "");
        if (raw == null || raw.isBlank()) {
            return DEFAULT_MONTE_CARLO_HORIZON_MONTHS;
        }
        try {
            int parsed = Integer.parseInt(raw.trim());
            return Math.max(1, Math.min(120, parsed));
        } catch (NumberFormatException ignored) {
            return DEFAULT_MONTE_CARLO_HORIZON_MONTHS;
        }
    }

    private static int resolveMonteCarloIterations() {
        String raw = System.getProperty("portfolio.analytics.montecarlo.iterations", "");
        if (raw == null || raw.isBlank()) {
            return DEFAULT_MONTE_CARLO_ITERATIONS;
        }
        try {
            int parsed = Integer.parseInt(raw.trim());
            return Math.max(200, Math.min(20000, parsed));
        } catch (NumberFormatException ignored) {
            return DEFAULT_MONTE_CARLO_ITERATIONS;
        }
    }

    public static String buildAnnualPortfolioValueSparklineSvg(
            TransactionStore store,
            Map<String, Double> ratesToNok,
            int year) {

        ArrayList<PortfolioValuePoint> yearPoints = buildAnnualYearPoints(store, ratesToNok, year);
        if (yearPoints.isEmpty()) {
            return "";
        }

        return buildPortfolioValueSparkline(yearPoints, SparklineMetric.VALUE);
    }

    public static String buildStandardPortfolioValueSparklineSvg(
            TransactionStore store,
            Map<String, Double> ratesToNok) {

        ArrayList<PortfolioValuePoint> timelinePoints = buildPortfolioValueTimeline(store, ratesToNok, 60);
        if (timelinePoints.isEmpty()) {
            return "";
        }

        LinkedHashMap<String, ArrayList<PortfolioValuePoint>> byRange = buildStandardSparklineRanges(timelinePoints);
        if (byRange.isEmpty()) {
            return "";
        }

        String defaultRange = byRange.containsKey("1Y") ? "1Y" : byRange.keySet().iterator().next();
        StringBuilder html = new StringBuilder();
        html.append("<div class=\"sparkline-widget\">\n");
        html.append("<div class=\"sparkline-metric-controls\">\n");
        html.append("<button type=\"button\" class=\"sparkline-metric-btn is-active\" data-metric=\"value\">Value</button>\n");
        html.append("</div>\n");
        html.append("<div class=\"sparkline-controls\">\n");
        for (String range : byRange.keySet()) {
            String active = range.equals(defaultRange) ? " is-active" : "";
            html.append("<button type=\"button\" class=\"sparkline-range-btn")
                .append(active)
                .append("\" data-range=\"")
                .append(range)
                .append("\">")
                .append(range)
                .append("</button>\n");
        }
        html.append("</div>\n");

        for (Map.Entry<String, ArrayList<PortfolioValuePoint>> entry : byRange.entrySet()) {
            String range = entry.getKey();
            html.append("<div class=\"sparkline-panel")
                .append(range.equals(defaultRange) ? " is-active" : "")
                .append("\" data-range=\"")
                .append(range)
                .append("\" data-metric=\"value\">\n")
                .append(buildPortfolioValueSparkline(entry.getValue(), SparklineMetric.VALUE))
                .append("</div>\n");
        }

        html.append("</div>\n");
        return html.toString();
    }

    public static String buildStandardPortfolioReturnSparklineSvg(
            TransactionStore store,
            Map<String, Double> ratesToNok) {

        ArrayList<PortfolioValuePoint> timelinePoints = buildPortfolioValueTimeline(store, ratesToNok, 60);
        if (timelinePoints.isEmpty()) {
            return "";
        }

        LinkedHashMap<String, ArrayList<PortfolioValuePoint>> byRange = buildStandardSparklineRanges(timelinePoints);
        if (byRange.isEmpty()) {
            return "";
        }

        String defaultRange = byRange.containsKey("1Y") ? "1Y" : byRange.keySet().iterator().next();
        StringBuilder html = new StringBuilder();
        html.append("<div class=\"sparkline-widget\">\n");
        html.append("<div class=\"sparkline-metric-controls\">\n");
        html.append("<button type=\"button\" class=\"sparkline-metric-btn js-return-amount-label is-active\" data-metric=\"return-nok\">Return (NOK)</button>\n");
        html.append("<button type=\"button\" class=\"sparkline-metric-btn\" data-metric=\"return-pct\">Return (%)</button>\n");
        html.append("</div>\n");

        html.append("<div class=\"sparkline-controls\">\n");
        for (String range : byRange.keySet()) {
            String active = range.equals(defaultRange) ? " is-active" : "";
            html.append("<button type=\"button\" class=\"sparkline-range-btn")
                .append(active)
                .append("\" data-range=\"")
                .append(range)
                .append("\">")
                .append(range)
                .append("</button>\n");
        }
        html.append("</div>\n");

        for (Map.Entry<String, ArrayList<PortfolioValuePoint>> entry : byRange.entrySet()) {
            String range = entry.getKey();
            ArrayList<PortfolioValuePoint> points = entry.getValue();

            html.append("<div class=\"sparkline-panel")
                .append(range.equals(defaultRange) ? " is-active" : "")
                .append("\" data-range=\"")
                .append(range)
                .append("\" data-metric=\"return-nok\">\n")
                .append(buildPortfolioValueSparkline(points, SparklineMetric.RETURN_NOK))
                .append("</div>\n");

            html.append("<div class=\"sparkline-panel\" data-range=\"")
                .append(range)
                .append("\" data-metric=\"return-pct\">\n")
                .append(buildPortfolioValueSparkline(points, SparklineMetric.RETURN_PCT))
                .append("</div>\n");
        }

        html.append("</div>\n");
        return html.toString();
    }

    public static String buildAnnualPortfolioReturnSparklineSvg(
            TransactionStore store,
            Map<String, Double> ratesToNok,
            int year) {

        int safeYear = Math.max(2000, Math.min(2100, year));
        ArrayList<PortfolioValuePoint> yearPoints = buildAnnualYearPoints(store, ratesToNok, safeYear);
        if (yearPoints.isEmpty()) {
            return "";
        }

        String yearRangeKey = "YEAR";
        StringBuilder html = new StringBuilder();
        html.append("<div class=\"sparkline-widget\">\n");
        html.append("<div class=\"sparkline-metric-controls\">\n");
        html.append("<button type=\"button\" class=\"sparkline-metric-btn js-return-amount-label is-active\" data-metric=\"return-nok\">Return (NOK)</button>\n");
        html.append("<button type=\"button\" class=\"sparkline-metric-btn\" data-metric=\"return-pct\">Return (%)</button>\n");
        html.append("</div>\n");

        html.append("<div class=\"sparkline-panel is-active\" data-range=\"")
            .append(yearRangeKey)
            .append("\" data-metric=\"return-nok\">\n")
            .append(buildPortfolioValueSparkline(yearPoints, SparklineMetric.RETURN_NOK))
            .append("</div>\n");

        html.append("<div class=\"sparkline-panel\" data-range=\"")
            .append(yearRangeKey)
            .append("\" data-metric=\"return-pct\">\n")
            .append(buildPortfolioValueSparkline(yearPoints, SparklineMetric.RETURN_PCT))
            .append("</div>\n");

        html.append("</div>\n");
        return html.toString();
    }

    public static double resolvePriceAtDate(
            Security security,
            LocalDate date,
            Map<String, NavigableMap<LocalDate, Double>> priceSeriesCache) {

        return resolvePriceAtDateDetailed(security, date, priceSeriesCache).getPrice();
    }

    public static PriceResolution resolvePriceAtDateDetailed(
            Security security,
            LocalDate date,
            Map<String, NavigableMap<LocalDate, Double>> priceSeriesCache) {

        if (security == null || date == null) {
            return new PriceResolution(0.0, true);
        }

        Map<String, NavigableMap<LocalDate, Double>> safeCache =
                priceSeriesCache == null ? new HashMap<>() : priceSeriesCache;
        return resolveHistoricalPriceDetailed(security, date, safeCache);
    }

    private static ArrayList<PortfolioValuePoint> buildAnnualYearPoints(
            TransactionStore store,
            Map<String, Double> ratesToNok,
            int year) {

        int safeYear = Math.max(2000, Math.min(2100, year));
        YearMonth yearStart = YearMonth.of(safeYear, 1);
        long monthsFromStart = ChronoUnit.MONTHS.between(yearStart, YearMonth.now()) + 1L;
        int months = (int) Math.max(24L, monthsFromStart);

        ArrayList<PortfolioValuePoint> timeline = buildPortfolioValueTimeline(store, ratesToNok, months);
        ArrayList<PortfolioValuePoint> yearPoints = new ArrayList<>();
        for (PortfolioValuePoint point : timeline) {
            if (point != null && point.monthEnd != null && point.monthEnd.getYear() == safeYear) {
                yearPoints.add(point);
            }
        }

        if (yearPoints.isEmpty()) {
            return yearPoints;
        }

        return yearPoints;
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

    private static double resolveBoundaryValue(
            NavigableMap<LocalDate, Double> series,
            LocalDate targetDate,
            boolean preferForwardFromTarget,
            long maxDistanceDays) {
        if (series == null || series.isEmpty() || targetDate == null) {
            return 0.0;
        }
        long safeMaxDistance = Math.max(0L, maxDistanceDays);

        LocalDate bestDate = null;
        double bestValue = 0.0;

        for (Map.Entry<LocalDate, Double> entry : series.entrySet()) {
            if (entry == null || entry.getKey() == null || entry.getValue() == null || entry.getValue() <= 0.0) {
                continue;
            }

            long distanceDays = Math.abs(ChronoUnit.DAYS.between(targetDate, entry.getKey()));
            if (distanceDays > safeMaxDistance) {
                continue;
            }

            if (bestDate == null) {
                bestDate = entry.getKey();
                bestValue = entry.getValue();
                continue;
            }

            long currentDistance = Math.abs(ChronoUnit.DAYS.between(targetDate, bestDate));
            if (distanceDays < currentDistance) {
                bestDate = entry.getKey();
                bestValue = entry.getValue();
                continue;
            }

            if (distanceDays == currentDistance) {
                boolean candidateIsAfter = !entry.getKey().isBefore(targetDate);
                boolean bestIsAfter = !bestDate.isBefore(targetDate);
                if (preferForwardFromTarget) {
                    if (candidateIsAfter && !bestIsAfter) {
                        bestDate = entry.getKey();
                        bestValue = entry.getValue();
                    }
                } else {
                    if (!candidateIsAfter && bestIsAfter) {
                        bestDate = entry.getKey();
                        bestValue = entry.getValue();
                    }
                }
            }
        }

        return bestValue;
    }
}
