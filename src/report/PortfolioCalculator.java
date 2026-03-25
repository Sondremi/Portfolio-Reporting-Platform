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
import java.time.Instant;
import java.time.LocalDate;
import java.time.YearMonth;
import java.time.ZoneOffset;
import java.time.format.DateTimeFormatter;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.*;

public class PortfolioCalculator {

    private static final DateTimeFormatter DATE_FORMATTER = DateTimeFormatter.ofPattern("dd.MM.yyyy");
    private static final String DEFAULT_CURRENCY_CODE = "NOK";
    private static final Pattern YAHOO_TIMESTAMP_ARRAY = Pattern.compile("\\\"timestamp\\\"\\s*:\\s*\\[(.*?)\\]", Pattern.DOTALL);
    private static final Pattern YAHOO_CLOSE_ARRAY = Pattern.compile("\\\"close\\\"\\s*:\\s*\\[(.*?)\\]", Pattern.DOTALL);

    private static final class PortfolioValuePoint {
        private final LocalDate monthEnd;
        private final double value;

        private PortfolioValuePoint(LocalDate monthEnd, double value) {
            this.monthEnd = monthEnd;
            this.value = value;
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
        String totalCurrencyCode = null;

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

        for (OverviewRow row : overviewRows) {
            totalMarketValue += row.marketValue;
            totalReturn += row.totalReturn;
            totalHistoricalCostBasis += row.historicalCostBasis;
            totalCurrencyCode = mergeCurrencyCodes(totalCurrencyCode, row.currencyCode);

            if (best == null || row.totalReturn > best.totalReturn) {
                best = row;
            }
            if (worst == null || row.totalReturn < worst.totalReturn) {
                worst = row;
            }
        }

        if (best != null) {
            bestLabel = getOverviewRowLabel(best);
            bestCurrencyCode = normalizeCurrencyCode(best.currencyCode);
            bestReturn = best.totalReturn;
            bestReturnPct = best.totalReturnPct;
        }
        if (worst != null) {
            worstLabel = getOverviewRowLabel(worst);
            worstCurrencyCode = normalizeCurrencyCode(worst.currencyCode);
            worstReturn = worst.totalReturn;
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

    private static String buildPortfolioValueSparklineWidget(TransactionStore store, Map<String, Double> ratesToNok) {
        ArrayList<PortfolioValuePoint> allPoints = buildPortfolioValueTimeline(store, ratesToNok, 60);
        if (allPoints.isEmpty()) {
            return "";
        }

        LinkedHashMap<String, ArrayList<PortfolioValuePoint>> byRange = new LinkedHashMap<>();
        byRange.put("1M", takeLastPoints(allPoints, 2));
        byRange.put("3M", takeLastPoints(allPoints, 4));
        byRange.put("6M", takeLastPoints(allPoints, 7));
        byRange.put("1Y", takeLastPoints(allPoints, 12));
        byRange.put("YTD", takeYearToDatePoints(allPoints));
        byRange.put("3Y", takeLastPoints(allPoints, 36));
        byRange.put("5Y", takeLastPoints(allPoints, 60));

        String defaultRange = byRange.containsKey("1Y") ? "1Y" : byRange.keySet().iterator().next();
        StringBuilder html = new StringBuilder();
        html.append("<div class=\"sparkline-widget\">\n");
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
            String active = range.equals(defaultRange) ? " is-active" : "";
            html.append("<div class=\"sparkline-panel")
                .append(active)
                .append("\" data-range=\"")
                .append(range)
                .append("\">\n")
                .append(buildPortfolioValueSparkline(points))
                .append("</div>\n");
        }
        html.append("</div>\n");
        return html.toString();
    }

    private static String buildPortfolioValueSparkline(ArrayList<PortfolioValuePoint> points) {
        if (points == null || points.isEmpty()) {
            return "";
        }

        final double width = 500.0;
        final double height = 124.0;
        final double left = 46.0;
        final double right = 14.0;
        final double top = 8.0;
        final double bottom = 20.0;
        final double plotWidth = width - left - right;
        final double plotHeight = height - top - bottom;

        double min = Double.POSITIVE_INFINITY;
        double max = Double.NEGATIVE_INFINITY;
        int n = points.size();
        for (PortfolioValuePoint point : points) {
            min = Math.min(min, point.value);
            max = Math.max(max, point.value);
        }
        if (!Double.isFinite(min) || !Double.isFinite(max) || Math.abs(max - min) < 1e-9) {
            max += 1.0;
            min -= 1.0;
        }

        double[] xValues = new double[n];
        double[] yValues = new double[n];
        for (int i = 0; i < n; i++) {
            double x = left + (plotWidth * i / Math.max(1.0, n - 1.0));
            double y = mapValueToY(points.get(i).value, min, max, top, plotHeight);
            xValues[i] = x;
            yValues[i] = y;
        }

        StringBuilder svg = new StringBuilder();
        svg.append("<svg viewBox=\"0 0 ").append(String.format(Locale.US, "%.0f %.0f", width, height))
                .append("\" xmlns=\"http://www.w3.org/2000/svg\" role=\"img\">");

        double axisY = top + plotHeight;
        svg.append("<line x1=\"").append(svgNumber(left)).append("\" y1=\"").append(svgNumber(top))
                .append("\" x2=\"").append(svgNumber(left)).append("\" y2=\"").append(svgNumber(axisY))
                .append("\" stroke=\"#c7d2df\" stroke-width=\"1\"/>");
        svg.append("<line x1=\"").append(svgNumber(left)).append("\" y1=\"").append(svgNumber(axisY))
                .append("\" x2=\"").append(svgNumber(left + plotWidth)).append("\" y2=\"").append(svgNumber(axisY))
                .append("\" stroke=\"#c7d2df\" stroke-width=\"1\"/>");

        svg.append("<text x=\"").append(svgNumber(left - 4.0)).append("\" y=\"").append(svgNumber(top + 1.0))
            .append("\" class=\"js-chart-money\" data-value-nok=\"").append(svgNumber(max)).append("\" data-format=\"compact\" text-anchor=\"end\" dominant-baseline=\"hanging\" font-size=\"8\" fill=\"#eaf2ff\">")
                .append(formatCompactKroner(max))
                .append("</text>");
        svg.append("<text x=\"").append(svgNumber(left - 4.0)).append("\" y=\"").append(svgNumber(axisY))
            .append("\" class=\"js-chart-money\" data-value-nok=\"").append(svgNumber(min)).append("\" data-format=\"compact\" text-anchor=\"end\" dominant-baseline=\"middle\" font-size=\"8\" fill=\"#eaf2ff\">")
                .append(formatCompactKroner(min))
                .append("</text>");

        svg.append("<line x1=\"").append(svgNumber(left)).append("\" y1=\"").append(svgNumber(top))
                .append("\" x2=\"").append(svgNumber(left + plotWidth)).append("\" y2=\"").append(svgNumber(top))
                .append("\" stroke=\"#eaf2ff\" stroke-width=\"0.8\" opacity=\"0.45\" stroke-dasharray=\"3 3\"/>");

        StringBuilder path = new StringBuilder();
        for (int i = 0; i < n; i++) {
            path.append(i == 0 ? "M " : " L ")
                    .append(svgNumber(xValues[i])).append(" ").append(svgNumber(yValues[i]));
        }

        svg.append("<path d=\"").append(path).append("\" fill=\"none\" stroke=\"#f1f6ff\" stroke-width=\"2.2\" stroke-linecap=\"round\"/>");

        DateTimeFormatter axisMonthFormat = DateTimeFormatter.ofPattern("MMM", Locale.ENGLISH);
        for (int i = 0; i < n; i++) {
            String monthLabel = points.get(i).monthEnd.format(axisMonthFormat);
            svg.append("<circle class=\"chart-hover-target chart-hover-point\" cx=\"").append(svgNumber(xValues[i])).append("\" cy=\"").append(svgNumber(yValues[i]))
                .append("\" r=\"2.2\" fill=\"#f1f6ff\">")
                .append("<title class=\"js-chart-money\" data-value-nok=\"").append(svgNumber(points.get(i).value))
                .append("\" data-format=\"compact\" data-prefix=\"").append(monthLabel).append(": \">")
                .append(monthLabel).append(": ").append(formatCompactKroner(points.get(i).value))
                .append("</title></circle>");

            svg.append("<line x1=\"").append(svgNumber(xValues[i])).append("\" y1=\"").append(svgNumber(axisY))
                    .append("\" x2=\"").append(svgNumber(xValues[i])).append("\" y2=\"").append(svgNumber(axisY + 2.8))
                    .append("\" stroke=\"#eaf2ff\" stroke-width=\"0.8\"/>");

            String tickAnchor = "middle";
            double tickLabelX = xValues[i];
            if (i == 0) {
                tickAnchor = "start";
                tickLabelX = Math.max(tickLabelX, left + 1.0);
            } else if (i == n - 1) {
                tickAnchor = "end";
                tickLabelX = Math.min(tickLabelX, left + plotWidth - 1.0);
            }

                svg.append("<text x=\"").append(svgNumber(tickLabelX)).append("\" y=\"").append(svgNumber(axisY + 13.0))
                    .append("\" text-anchor=\"").append(tickAnchor).append("\" font-size=\"7\" fill=\"#eaf2ff\">")
                    .append(monthLabel)
                    .append("</text>");
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

        ArrayList<Events.CashEvent> sortedCashEvents = new ArrayList<>();
        if (!useAuthoritativeSnapshots) {
            sortedCashEvents.addAll(cashEvents);
            sortedCashEvents.sort(Comparator.comparing(Events.CashEvent::tradeDate));
        }

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
        int unitEventIndex = 0;
        int cashEventIndex = 0;
        int[] cashSnapshotIndexRef = new int[]{0};
        double runningCash = 0.0;
        LinkedHashMap<String, Double> latestBalanceByPortfolio = new LinkedHashMap<>();

        for (int monthOffset = 0; monthOffset < monthCount; monthOffset++) {
            YearMonth currentMonth = startMonth.plusMonths(monthOffset);
            LocalDate monthEnd = currentMonth.atEndOfMonth();

            while (unitEventIndex < sortedUnitEvents.size()
                    && !sortedUnitEvents.get(unitEventIndex).tradeDate().isAfter(monthEnd)) {
                Events.UnitEvent event = sortedUnitEvents.get(unitEventIndex);
                unitsBySecurity.merge(event.securityKey(), event.unitsDelta(), Double::sum);
                unitEventIndex++;
            }

            while (cashEventIndex < sortedCashEvents.size()
                    && !sortedCashEvents.get(cashEventIndex).tradeDate().isAfter(monthEnd)) {
                Events.CashEvent cashEvent = sortedCashEvents.get(cashEventIndex);
                runningCash += cashEvent.cashDelta();
                cashEventIndex++;
            }

            double balanceSnapshotCash = sumPortfolioBalancesOnOrBefore(
                    monthEnd,
                    sortedCashSnapshots,
                    cashSnapshotIndexRef,
                    latestBalanceByPortfolio
            );

            double totalValue = useAuthoritativeSnapshots
                    ? balanceSnapshotCash
                    : runningCash + balanceSnapshotCash;

            for (Map.Entry<String, Double> entry : unitsBySecurity.entrySet()) {
                double units = entry.getValue();
                if (units <= 0.0000001) {
                    continue;
                }

                Security security = securitiesByTrackingKey.get(entry.getKey());
                if (security == null) {
                    continue;
                }

                double price = resolveHistoricalPrice(security, monthEnd, priceSeriesCache);
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

            timeline.add(new PortfolioValuePoint(monthEnd, totalValue));
        }

        return timeline;
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
        if (security == null) {
            return 0.0;
        }

        String ticker = security.getTicker();
        if (ticker == null || ticker.isBlank() || "-".equals(ticker)) {
            return Math.max(0.0, security.getLatestPrice());
        }

        NavigableMap<LocalDate, Double> series = priceSeriesCache.computeIfAbsent(
                ticker,
                t -> fetchHistoricalCloseSeries(t, LocalDate.now().minusMonths(18), LocalDate.now())
        );

        if (series.isEmpty()) {
            return Math.max(0.0, security.getLatestPrice());
        }

        Map.Entry<LocalDate, Double> floor = series.floorEntry(monthEnd);
        if (floor != null && floor.getValue() != null && floor.getValue() > 0.0) {
            return floor.getValue();
        }

        Map.Entry<LocalDate, Double> first = series.firstEntry();
        if (first != null && first.getValue() != null && first.getValue() > 0.0) {
            return first.getValue();
        }

        return Math.max(0.0, security.getLatestPrice());
    }

    private static NavigableMap<LocalDate, Double> fetchHistoricalCloseSeries(String ticker, LocalDate fromDate, LocalDate toDate) {
        NavigableMap<LocalDate, Double> series = new TreeMap<>();
        if (ticker == null || ticker.isBlank()) {
            return series;
        }

        HttpURLConnection connection = null;
        try {
            long period1 = fromDate.atStartOfDay(ZoneOffset.UTC).toEpochSecond();
            long period2 = toDate.plusDays(1).atStartOfDay(ZoneOffset.UTC).toEpochSecond();
            String encodedTicker = URLEncoder.encode(ticker, StandardCharsets.UTF_8);
            String urlText = "https://query1.finance.yahoo.com/v8/finance/chart/" + encodedTicker
                    + "?period1=" + period1
                    + "&period2=" + period2
                    + "&interval=1d&events=history";

            URL url = URI.create(urlText).toURL();
            connection = (HttpURLConnection) url.openConnection();
            connection.setRequestMethod("GET");
            connection.setConnectTimeout(5000);
            connection.setReadTimeout(7000);

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
            Matcher closeMatcher = YAHOO_CLOSE_ARRAY.matcher(body);
            if (!timestampMatcher.find() || !closeMatcher.find()) {
                return series;
            }

            String[] timestamps = timestampMatcher.group(1).split(",");
            String[] closes = closeMatcher.group(1).split(",");
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

    private static String mergeCurrencyCodes(String currentCurrencyCode, String nextCurrencyCode) {
        String next = (nextCurrencyCode == null || nextCurrencyCode.isBlank())
                ? DEFAULT_CURRENCY_CODE
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
}