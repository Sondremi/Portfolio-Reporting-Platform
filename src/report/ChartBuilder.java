package report;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

public class ChartBuilder {

    private static final class AllocationBucket {
        private final String label;
        private final double value;

        private AllocationBucket(String label, double value) {
            this.label = label;
            this.value = value;
        }
    }

    private static final class PieSliceLabel {
        private final boolean rightSide;
        private final String color;
        private final String text;
        private final double anchorX;
        private final double anchorY;
        private final double bendX;
        private double labelY;

        private PieSliceLabel(boolean rightSide, String color, String text,
                              double anchorX, double anchorY, double bendX, double labelY) {
            this.rightSide = rightSide;
            this.color = color;
            this.text = text;
            this.anchorX = anchorX;
            this.anchorY = anchorY;
            this.bendX = bendX;
            this.labelY = labelY;
        }
    }

    public static String buildOverviewBarChartSvg(List<OverviewRow> rows, boolean percentChart, Map<String, Double> ratesToNok) {
        final double width = percentChart ? 1100.0 : 1040.0;
        final double height = 450.0;
        final double left = percentChart ? 68.0 : 92.0;
        final double right = percentChart ? 22.0 : 38.0;
        final double top = 14.0;
        final double bottom = 118.0;
        final double plotWidth = width - left - right;
        final double plotHeight = height - top - bottom;

        double minValue = 0.0;
        double maxValue = 0.0;
        for (OverviewRow row : rows) {
            double value = percentChart
                    ? row.totalReturnPct
                    : convertToNok(row.totalReturn, row.currencyCode, ratesToNok);
            minValue = Math.min(minValue, value);
            maxValue = Math.max(maxValue, value);
        }

        if (Math.abs(maxValue - minValue) < 1e-9) {
            maxValue += 1.0;
            minValue -= 1.0;
        }

        double valueRange = maxValue - minValue;
        double zeroY = mapValueToY(0.0, minValue, maxValue, top, plotHeight);
        double chartZeroY = Math.max(top, Math.min(top + plotHeight, zeroY));

        StringBuilder svg = new StringBuilder();
        svg.append("<svg class=\"chart-svg\" viewBox=\"0 0 ")
                .append(svgNumber(width))
                .append(" ")
                .append(svgNumber(height))
                .append("\" xmlns=\"http://www.w3.org/2000/svg\" role=\"img\">\n");

        int tickCount = percentChart ? 7 : 5;
        for (int i = 0; i <= tickCount; i++) {
            double tickValue = maxValue - ((valueRange / tickCount) * i);
            double y = mapValueToY(tickValue, minValue, maxValue, top, plotHeight);

            svg.append("<line x1=\"").append(svgNumber(left)).append("\" y1=\"").append(svgNumber(y))
                    .append("\" x2=\"").append(svgNumber(left + plotWidth)).append("\" y2=\"").append(svgNumber(y))
                    .append("\" stroke=\"#d2d8df\" stroke-width=\"1.3\"/>\n");

                if (percentChart) {
                svg.append("<text x=\"").append(svgNumber(left - 8.0)).append("\" y=\"").append(svgNumber(y + 4.0))
                    .append("\" text-anchor=\"end\" font-size=\"12\" font-weight=\"600\" fill=\"#4a5563\">")
                    .append(escapeHtml(formatChartValue(tickValue, true, true)))
                    .append("</text>\n");
                } else {
                svg.append("<text class=\"js-chart-money\" data-value-nok=\"").append(svgNumber(tickValue)).append("\" data-decimals=\"0\"")
                    .append(" x=\"").append(svgNumber(left - 8.0)).append("\" y=\"").append(svgNumber(y + 4.0))
                    .append("\" text-anchor=\"end\" font-size=\"12\" font-weight=\"600\" fill=\"#4a5563\">")
                    .append(escapeHtml(formatNumber(tickValue, 0) + " NOK"))
                    .append("</text>\n");
                }
        }

        svg.append("<rect x=\"").append(svgNumber(left)).append("\" y=\"").append(svgNumber(top))
                .append("\" width=\"").append(svgNumber(plotWidth)).append("\" height=\"").append(svgNumber(plotHeight))
                .append("\" fill=\"none\" stroke=\"#b8c0ca\" stroke-width=\"1.3\"/>\n");

        double slotWidth = rows.isEmpty() ? plotWidth : plotWidth / rows.size();
        double barWidth = Math.max(8.0, slotWidth * 0.62);
        double labelRotation = rows.size() > 14 ? -42.0 : -34.0;
        double labelFontSize = rows.size() > 14 ? 11.0 : 12.0;

        for (int i = 0; i < rows.size(); i++) {
            OverviewRow row = rows.get(i);
            double value = percentChart
                    ? row.totalReturnPct
                    : convertToNok(row.totalReturn, row.currencyCode, ratesToNok);
            double x = left + (i * slotWidth) + ((slotWidth - barWidth) / 2.0);
            double yValue = mapValueToY(value, minValue, maxValue, top, plotHeight);
            double barY = Math.min(yValue, chartZeroY);
            double barHeight = Math.abs(chartZeroY - yValue);
            if (barHeight < 1.0) {
                barHeight = 1.0;
            }

            String barColor = value >= 0.0 ? "#2b8a3e" : "#b23a31";
            String barBorderColor = value >= 0.0 ? "#1f6f31" : "#8f2b24";
            String label = getOverviewRowLabel(row);
            String compactLabel = getCompactOverviewLabel(row);

                svg.append("<rect class=\"chart-hover-target chart-hover-bar\" x=\"").append(svgNumber(x)).append("\" y=\"").append(svgNumber(barY))
                    .append("\" width=\"").append(svgNumber(barWidth)).append("\" height=\"").append(svgNumber(barHeight))
                    .append("\" fill=\"").append(barColor).append("\" stroke=\"").append(barBorderColor)
                    .append("\" stroke-width=\"1.1\" rx=\"1\">\n");
                if (percentChart) {
                svg.append("<title>")
                    .append(escapeHtml(label + ": " + formatChartValue(value, true, false)))
                    .append("</title>");
                } else {
                svg.append("<title class=\"js-chart-money\" data-value-nok=\"").append(svgNumber(value))
                    .append("\" data-decimals=\"2\" data-prefix=\"").append(escapeHtml(label + ": "))
                    .append("\">")
                    .append(escapeHtml(label + ": " + formatNumber(value, 2) + " NOK"))
                    .append("</title>");
                }
                svg.append("</rect>\n");

            double labelAnchorX = x + (barWidth / 2.0);
                double labelAnchorY = height - bottom + 16.0;
            svg.append("<text x=\"").append(svgNumber(labelAnchorX)).append("\" y=\"").append(svgNumber(labelAnchorY))
                    .append("\" transform=\"rotate(").append(svgNumber(labelRotation)).append(" ").append(svgNumber(labelAnchorX)).append(" ").append(svgNumber(labelAnchorY))
                    .append(")\" text-anchor=\"end\" font-size=\"").append(svgNumber(labelFontSize)).append("\" font-weight=\"700\" fill=\"#1f2933\" paint-order=\"stroke\" stroke=\"#ffffff\" stroke-width=\"2.2\" stroke-linejoin=\"round\">")
                    .append(escapeHtml(compactLabel))
                    .append("</text>\n");
        }

        svg.append("<line x1=\"").append(svgNumber(left)).append("\" y1=\"").append(svgNumber(chartZeroY))
                .append("\" x2=\"").append(svgNumber(left + plotWidth)).append("\" y2=\"").append(svgNumber(chartZeroY))
                .append("\" stroke=\"#4f5d6c\" stroke-width=\"1.8\"/>\n");
        svg.append("<line x1=\"").append(svgNumber(left)).append("\" y1=\"").append(svgNumber(top))
                .append("\" x2=\"").append(svgNumber(left)).append("\" y2=\"").append(svgNumber(top + plotHeight))
                .append("\" stroke=\"#4f5d6c\" stroke-width=\"1.8\"/>\n");

        svg.append("</svg>\n");
        return svg.toString();
    }

    public static String buildMarketValueBarChartSvg(List<OverviewRow> rows, Map<String, Double> ratesToNok) {
        final double width = 760.0;
        final double height = 390.0;
        final double left = 126.0;
        final double right = 20.0;
        final double top = 16.0;
        final double bottom = 108.0;
        final double plotWidth = width - left - right;
        final double plotHeight = height - top - bottom;

        List<OverviewRow> rowsWithValue = new ArrayList<>();
        List<Double> marketValuesNok = new ArrayList<>();
        double totalMarketValue = 0.0;
        double maxValue = 0.0;
        for (OverviewRow row : rows) {
            double marketValueNok = convertToNok(row.marketValue, row.currencyCode, ratesToNok);
            if (marketValueNok > 0.0) {
                rowsWithValue.add(row);
                marketValuesNok.add(marketValueNok);
                totalMarketValue += marketValueNok;
                maxValue = Math.max(maxValue, marketValueNok);
            }
        }

        StringBuilder svg = new StringBuilder();
        svg.append("<svg class=\"chart-svg market-value-bar-chart\" viewBox=\"0 0 ")
           .append(svgNumber(width)).append(" ").append(svgNumber(height))
           .append("\" xmlns=\"http://www.w3.org/2000/svg\" role=\"img\">\n");

        if (rowsWithValue.isEmpty() || maxValue <= 0.0) {
            svg.append("<text x=\"").append(svgNumber(width / 2.0))
               .append("\" y=\"").append(svgNumber(height / 2.0))
               .append("\" text-anchor=\"middle\" font-size=\"13\" fill=\"#666\">No market value data</text>\n");
            svg.append("</svg>\n");
            return svg.toString();
        }

        int tickCount = 5;
        for (int i = 0; i <= tickCount; i++) {
            double tickValue = maxValue - ((maxValue / tickCount) * i);
            double y = mapValueToY(tickValue, 0.0, maxValue, top, plotHeight);

            svg.append("<line x1=\"").append(svgNumber(left)).append("\" y1=\"").append(svgNumber(y))
               .append("\" x2=\"").append(svgNumber(left + plotWidth)).append("\" y2=\"").append(svgNumber(y))
               .append("\" stroke=\"#ececec\" stroke-width=\"1\"/>\n");

                svg.append("<text class=\"js-chart-money\" data-value-nok=\"").append(svgNumber(tickValue)).append("\" data-decimals=\"0\"")
                    .append(" x=\"").append(svgNumber(left - 8.0)).append("\" y=\"").append(svgNumber(y + 4.0))
                    .append("\" text-anchor=\"end\" font-size=\"10\" fill=\"#666\">")
                    .append(escapeHtml(formatNumber(tickValue, 0) + " NOK"))
                    .append("</text>\n");
        }

        double averageValue = totalMarketValue / rowsWithValue.size();
        double averageY = mapValueToY(averageValue, 0.0, maxValue, top, plotHeight);
          svg.append("<line x1=\"").append(svgNumber(left)).append("\" y1=\"").append(svgNumber(averageY))
              .append("\" x2=\"").append(svgNumber(left + plotWidth)).append("\" y2=\"").append(svgNumber(averageY))
              .append("\" stroke=\"#495057\" stroke-width=\"1.2\" stroke-dasharray=\"5 4\"/>\n");

          svg.append("<line class=\"chart-hover-target chart-hover-avg-hit\" x1=\"").append(svgNumber(left)).append("\" y1=\"").append(svgNumber(averageY))
              .append("\" x2=\"").append(svgNumber(left + plotWidth)).append("\" y2=\"").append(svgNumber(averageY))
              .append("\" stroke=\"transparent\" stroke-width=\"12\">\n")
              .append("<title class=\"js-chart-money\" data-value-nok=\"").append(svgNumber(averageValue)).append("\" data-decimals=\"0\" data-prefix=\"Avg: \">");
          svg.append(escapeHtml("Avg: " + formatNumber(averageValue, 0) + " NOK"));
          svg.append("</title></line>\n");

             svg.append("<text class=\"js-chart-money\" data-value-nok=\"").append(svgNumber(averageValue)).append("\" data-decimals=\"0\" data-prefix=\"Avg: \"")
                  .append(" x=\"").append(svgNumber(left - 8.0)).append("\" y=\"").append(svgNumber(averageY + 3.0))
                  .append("\" text-anchor=\"end\" font-size=\"10\" fill=\"#495057\">Avg: ")
              .append(escapeHtml(formatNumber(averageValue, 0) + " NOK"))
              .append("</text>\n");

        double slotWidth = plotWidth / rowsWithValue.size();
        double barWidth = Math.max(10.0, slotWidth * 0.92);
        String[] colors = {"#0b7285", "#2f9e44", "#f08c00", "#7048e8", "#c92a2a", "#1c7ed6"};

        for (int i = 0; i < rowsWithValue.size(); i++) {
            OverviewRow row = rowsWithValue.get(i);
            double marketValueNok = marketValuesNok.get(i);
            double x = left + (i * slotWidth) + ((slotWidth - barWidth) / 2.0);
            double y = mapValueToY(marketValueNok, 0.0, maxValue, top, plotHeight);
            double barHeight = (top + plotHeight) - y;

            String color = colors[i % colors.length];
            String label = getCompactBarLabel(row);

                svg.append("<rect class=\"chart-hover-target chart-hover-bar\" x=\"").append(svgNumber(x)).append("\" y=\"").append(svgNumber(y))
               .append("\" width=\"").append(svgNumber(barWidth)).append("\" height=\"").append(svgNumber(barHeight))
               .append("\" fill=\"").append(color).append("\" rx=\"2\">\n")
                    .append("<title class=\"js-chart-money\" data-value-nok=\"").append(svgNumber(marketValueNok))
                    .append("\" data-decimals=\"0\" data-prefix=\"").append(escapeHtml(row.securityDisplayName + ": "))
                    .append("\">")
                    .append(escapeHtml(row.securityDisplayName + ": " + formatNumber(marketValueNok, 0) + " NOK"))
                    .append("</title></rect>\n");

            double labelAnchorX = x + (barWidth / 2.0);
            double labelAnchorY = height - bottom + 16.0;
            svg.append("<text x=\"").append(svgNumber(labelAnchorX)).append("\" y=\"").append(svgNumber(labelAnchorY))
               .append("\" transform=\"rotate(-30 ").append(svgNumber(labelAnchorX)).append(" ").append(svgNumber(labelAnchorY)).append(")\" ")
               .append("text-anchor=\"end\" font-size=\"10\" font-weight=\"600\" fill=\"#1f2933\" paint-order=\"stroke\" stroke=\"#ffffff\" stroke-width=\"2\" stroke-linejoin=\"round\">")
               .append(escapeHtml(label)).append("</text>\n");
        }

        svg.append("<line x1=\"").append(svgNumber(left)).append("\" y1=\"").append(svgNumber(top + plotHeight))
           .append("\" x2=\"").append(svgNumber(left + plotWidth)).append("\" y2=\"").append(svgNumber(top + plotHeight))
           .append("\" stroke=\"#7a7a7a\" stroke-width=\"1.1\"/>\n");
        svg.append("<line x1=\"").append(svgNumber(left)).append("\" y1=\"").append(svgNumber(top))
           .append("\" x2=\"").append(svgNumber(left)).append("\" y2=\"").append(svgNumber(top + plotHeight))
           .append("\" stroke=\"#7a7a7a\" stroke-width=\"1.1\"/>\n");

        svg.append("</svg>\n");
        return svg.toString();
    }

    public static String buildMarketValueAllocationSvg(List<OverviewRow> rows, Map<String, Double> ratesToNok) {
        final double width = 760.0;
        final double height = 390.0;
        final double centerX = width / 2.0;
        final double centerY = 162.0;
        final double radius = 112.0;

        List<OverviewRow> rowsWithValue = new ArrayList<>();
        List<Double> marketValuesNok = new ArrayList<>();
        double totalMarketValue = 0.0;
        for (OverviewRow row : rows) {
            double marketValueNok = convertToNok(row.marketValue, row.currencyCode, ratesToNok);
            if (marketValueNok > 0.0) {
                rowsWithValue.add(row);
                marketValuesNok.add(marketValueNok);
                totalMarketValue += marketValueNok;
            }
        }

        StringBuilder svg = new StringBuilder();
        svg.append("<svg class=\"chart-svg\" viewBox=\"0 0 ")
                .append(svgNumber(width))
                .append(" ")
                .append(svgNumber(height))
                .append("\" xmlns=\"http://www.w3.org/2000/svg\" role=\"img\">\n");

        if (rowsWithValue.isEmpty() || totalMarketValue <= 0.0) {
            svg.append("<text x=\"").append(svgNumber(centerX)).append("\" y=\"").append(svgNumber(centerY))
                    .append("\" text-anchor=\"middle\" font-size=\"13\" fill=\"#666\">No market value data</text>\n");
            svg.append("</svg>\n");
            return svg.toString();
        }

        double currentAngle = -Math.PI / 2.0;
        List<PieSliceLabel> pieLabels = new ArrayList<>();
        for (int i = 0; i < rowsWithValue.size(); i++) {
            OverviewRow row = rowsWithValue.get(i);
            double marketValueNok = marketValuesNok.get(i);
            double fraction = marketValueNok / totalMarketValue;
            double sliceAngle = fraction * Math.PI * 2.0;
            double endAngle = currentAngle + sliceAngle;
            String color = getAllocationColor(i);

            double x1 = centerX + radius * Math.cos(currentAngle);
            double y1 = centerY + radius * Math.sin(currentAngle);
            double x2 = centerX + radius * Math.cos(endAngle);
            double y2 = centerY + radius * Math.sin(endAngle);
            int largeArcFlag = sliceAngle > Math.PI ? 1 : 0;

                svg.append("<path class=\"chart-hover-target chart-hover-slice\" d=\"M ").append(svgNumber(centerX)).append(" ").append(svgNumber(centerY))
                    .append(" L ").append(svgNumber(x1)).append(" ").append(svgNumber(y1))
                    .append(" A ").append(svgNumber(radius)).append(" ").append(svgNumber(radius)).append(" 0 ")
                    .append(largeArcFlag).append(" 1 ").append(svgNumber(x2)).append(" ").append(svgNumber(y2))
                    .append(" Z\" fill=\"").append(color).append("\">\n")
                    .append("<title class=\"js-chart-money\" data-value-nok=\"").append(svgNumber(marketValueNok)).append("\" data-decimals=\"2\" data-prefix=\"")
                    .append(escapeHtml(getOverviewRowLabel(row) + ": "))
                    .append("\" data-suffix=\" (").append(escapeHtml(formatNumber(fraction * 100.0, 2))).append("%)\">")
                    .append(escapeHtml(getOverviewRowLabel(row)
                        + ": " + formatNumber(marketValueNok, 2) + " NOK (" + formatNumber(fraction * 100.0, 2) + "%)"))
                    .append("</title></path>\n");

            double midAngle = currentAngle + (sliceAngle / 2.0);
            double anchorX = centerX + radius * Math.cos(midAngle);
            double anchorY = centerY + radius * Math.sin(midAngle);
            double bendX = centerX + (radius + 14.0) * Math.cos(midAngle);
            double bendY = centerY + (radius + 8.0) * Math.sin(midAngle);
            boolean rightSide = Math.cos(midAngle) >= 0.0;

            String labelText = getCompactPieLabel(row) + " " + formatNumber(fraction * 100.0, 1) + "%";
            pieLabels.add(new PieSliceLabel(rightSide, color, labelText, anchorX, anchorY, bendX, bendY));

            currentAngle = endAngle;
        }

        adjustPieLabelPositions(pieLabels, false, 22.0, 318.0, 15.0);
        adjustPieLabelPositions(pieLabels, true, 22.0, 318.0, 15.0);

        for (PieSliceLabel label : pieLabels) {
            double textX = label.rightSide ? (width - 224.0) : 224.0;
            double lineEndX = label.rightSide ? (width - 236.0) : 236.0;
            String textAnchor = label.rightSide ? "start" : "end";

            svg.append("<polyline points=\"")
                    .append(svgNumber(label.anchorX)).append(",").append(svgNumber(label.anchorY)).append(" ")
                    .append(svgNumber(label.bendX)).append(",").append(svgNumber(label.labelY)).append(" ")
                    .append(svgNumber(lineEndX)).append(",").append(svgNumber(label.labelY)).append("\"")
                    .append(" fill=\"none\" stroke=\"").append(label.color)
                    .append("\" stroke-width=\"0.9\" opacity=\"0.85\"/>\n");

            svg.append("<text x=\"").append(svgNumber(textX)).append("\" y=\"").append(svgNumber(label.labelY))
                    .append("\" text-anchor=\"").append(textAnchor)
                        .append("\" dominant-baseline=\"middle\" font-size=\"10\" fill=\"#2f2f2f\" paint-order=\"stroke\" stroke=\"#ffffff\" stroke-width=\"1.8\" stroke-linejoin=\"round\">")
                    .append(escapeHtml(label.text))
                    .append("</text>\n");
        }

            double summaryY = height - 28.0;
        svg.append("<text x=\"").append(svgNumber(centerX)).append("\" y=\"").append(svgNumber(summaryY))
                .append("\" text-anchor=\"middle\" font-size=\"12\" fill=\"#666\">Market Value Total</text>\n");
        svg.append("<text x=\"").append(svgNumber(centerX)).append("\" y=\"").append(svgNumber(summaryY + 16.0))
            .append("\" class=\"js-chart-money\" data-value-nok=\"").append(svgNumber(totalMarketValue)).append("\" data-decimals=\"0\" text-anchor=\"middle\" font-size=\"14\" fill=\"#222\" font-weight=\"600\">")
            .append(escapeHtml(formatNumber(totalMarketValue, 0) + " NOK"))
                .append("</text>\n");

        svg.append("</svg>\n");
        return svg.toString();
    }

    public static String buildAssetTypeAllocationSvg(List<OverviewRow> rows, double rawCashValue, Map<String, Double> ratesToNok) {
        double stockValue = 0.0;
        double fundValue = 0.0;
        double cashValue = Math.max(0.0, rawCashValue);
        double cashDebtValue = rawCashValue < 0.0 ? -rawCashValue : 0.0;
        double otherValue = 0.0;
        int stockCount = 0;
        int fundCount = 0;
        int cashCount = cashValue > 0.0 ? 1 : 0;
        int cashDebtCount = cashDebtValue > 0.0 ? 1 : 0;
        int otherCount = 0;

        for (OverviewRow row : rows) {
            double marketValueNok = convertToNok(row.marketValue, row.currencyCode, ratesToNok);
            if (marketValueNok <= 0.0) continue;
            switch (row.assetType) {
                case "STOCK" -> {
                    stockValue += marketValueNok;
                    stockCount++;
                }
                case "FUND" -> {
                    fundValue += marketValueNok;
                    fundCount++;
                }
                default -> {
                    otherValue += marketValueNok;
                    otherCount++;
                }
            }
        }

        final double width = 440.0;
        final double height = 330.0;
        final double centerX = width / 2.0;
        final double centerY = 118.0;
        final double radius = 98.0;

        double totalValue = stockValue + fundValue + cashValue + cashDebtValue + otherValue;
        if (totalValue <= 0.0) {
            return "<svg class=\"chart-svg\" viewBox=\"0 0 440 330\" xmlns=\"http://www.w3.org/2000/svg\">"
                    + "<text x=\"220\" y=\"118\" text-anchor=\"middle\" font-size=\"13\" fill=\"#666\">No asset type data</text></svg>";
        }

        List<String> labels = new ArrayList<>();
        List<Double> values = new ArrayList<>();
        List<String> colors = new ArrayList<>();

        if (stockValue > 0.0) {
            labels.add("Stocks");
            values.add(stockValue);
            colors.add("#1c7ed6");
        }
        if (fundValue > 0.0) {
            labels.add("Funds");
            values.add(fundValue);
            colors.add("#2f9e44");
        }
        if (cashValue > 0.0) {
            labels.add("Cash");
            values.add(cashValue);
            colors.add("#f59f00");
        }
        if (cashDebtValue > 0.0) {
            labels.add("Cash (Debt)");
            values.add(cashDebtValue);
            colors.add("#e03131");
        }
        if (otherValue > 0.0) {
            labels.add("Other");
            values.add(otherValue);
            colors.add("#868e96");
        }

        StringBuilder svg = new StringBuilder();
        svg.append("<svg class=\"chart-svg\" viewBox=\"0 0 ")
                .append(svgNumber(width)).append(" ").append(svgNumber(height))
                .append("\" xmlns=\"http://www.w3.org/2000/svg\" role=\"img\">\n");

        if (values.size() == 1) {
                svg.append("<circle class=\"chart-hover-target chart-hover-slice\" cx=\"").append(svgNumber(centerX)).append("\" cy=\"").append(svgNumber(centerY))
                    .append("\" r=\"").append(svgNumber(radius)).append("\" fill=\"").append(colors.get(0)).append("\">\n")
                    .append("<title class=\"js-chart-money\" data-value-nok=\"").append(svgNumber(values.get(0))).append("\" data-decimals=\"2\" data-prefix=\"")
                    .append(escapeHtml(labels.get(0) + ": "))
                    .append("\" data-suffix=\" (100.0%)\">")
                    .append(escapeHtml(labels.get(0) + ": " + formatNumber(values.get(0), 2) + " NOK (100.0%)"))
                    .append("</title></circle>\n");
        } else {
            double currentAngle = -Math.PI / 2.0;
            for (int i = 0; i < values.size(); i++) {
                double value = values.get(i);
                double fraction = value / totalValue;
                double sliceAngle = fraction * Math.PI * 2.0;
                double endAngle = currentAngle + sliceAngle;
                String color = colors.get(i);

                double x1 = centerX + radius * Math.cos(currentAngle);
                double y1 = centerY + radius * Math.sin(currentAngle);
                double x2 = centerX + radius * Math.cos(endAngle);
                double y2 = centerY + radius * Math.sin(endAngle);
                int largeArcFlag = sliceAngle > Math.PI ? 1 : 0;

                svg.append("<path class=\"chart-hover-target chart-hover-slice\" d=\"M ").append(svgNumber(centerX)).append(" ").append(svgNumber(centerY))
                        .append(" L ").append(svgNumber(x1)).append(" ").append(svgNumber(y1))
                        .append(" A ").append(svgNumber(radius)).append(" ").append(svgNumber(radius)).append(" 0 ")
                        .append(largeArcFlag).append(" 1 ").append(svgNumber(x2)).append(" ").append(svgNumber(y2))
                        .append(" Z\" fill=\"").append(color).append("\">\n")
                    .append("<title class=\"js-chart-money\" data-value-nok=\"").append(svgNumber(value)).append("\" data-decimals=\"2\" data-prefix=\"")
                    .append(escapeHtml(labels.get(i) + ": "))
                    .append("\" data-suffix=\" (").append(escapeHtml(formatNumber(fraction * 100.0, 1))).append("%)\">")
                    .append(escapeHtml(labels.get(i) + ": " + formatNumber(value, 2) + " NOK (" + formatNumber(fraction * 100.0, 1) + "%)"))
                    .append("</title></path>\n");

                currentAngle = endAngle;
            }
        }

        double stockPct = totalValue > 0.0 ? (stockValue / totalValue) * 100.0 : 0.0;
        double fundPct = totalValue > 0.0 ? (fundValue / totalValue) * 100.0 : 0.0;
        double cashPct = totalValue > 0.0 ? (cashValue / totalValue) * 100.0 : 0.0;
        double cashDebtPct = totalValue > 0.0 ? (cashDebtValue / totalValue) * 100.0 : 0.0;
        double otherPct = totalValue > 0.0 ? (otherValue / totalValue) * 100.0 : 0.0;

        double summaryY = centerY + radius + 14.0;
        svg.append("<text x=\"").append(svgNumber(centerX)).append("\" y=\"").append(svgNumber(summaryY))
                .append("\" text-anchor=\"middle\" font-size=\"12\" fill=\"#666\">Asset Mix</text>\n");

        List<String> legendLabels = new ArrayList<>();
        List<Double> legendPcts = new ArrayList<>();
        List<String> legendColors = new ArrayList<>();

        if (stockCount > 0) {
            legendLabels.add("Stocks: " + stockCount);
            legendPcts.add(stockPct);
            legendColors.add("#1c7ed6");
        }
        if (fundCount > 0) {
            legendLabels.add("Funds: " + fundCount);
            legendPcts.add(fundPct);
            legendColors.add("#2f9e44");
        }
        if (cashCount > 0) {
            legendLabels.add("Cash");
            legendPcts.add(cashPct);
            legendColors.add("#f59f00");
        }
        if (cashDebtCount > 0) {
            legendLabels.add("Cash (Debt)");
            legendPcts.add(cashDebtPct);
            legendColors.add("#e03131");
        }
        if (otherCount > 0) {
            legendLabels.add("Other: " + otherCount);
            legendPcts.add(otherPct);
            legendColors.add("#868e96");
        }

        boolean useTwoColumns = legendLabels.size() > 4;
        int rowsPerColumn = useTwoColumns ? (legendLabels.size() + 1) / 2 : legendLabels.size();
        double legendYStart = summaryY + 18.0;
        double legendBottomY = height - 14.0;
        double legendRowGap = computeLegendRowGap(rowsPerColumn, legendYStart, legendBottomY, 15.0);
        for (int i = 0; i < legendLabels.size(); i++) {
            int columnIndex = (useTwoColumns && i >= rowsPerColumn) ? 1 : 0;
            int rowIndex = useTwoColumns ? (i % rowsPerColumn) : i;
            double y = legendYStart + (rowIndex * legendRowGap);
            double dotX = columnIndex == 0 ? 22.0 : 228.0;
            double labelX = dotX + 9.0;
            double pctX = useTwoColumns
                    ? (columnIndex == 0 ? 214.0 : (width - 16.0))
                    : (width - 16.0);
            String displayLabel = abbreviateLegendLabel(legendLabels.get(i), useTwoColumns ? 14 : 24);

            svg.append("<circle cx=\"").append(svgNumber(dotX)).append("\" cy=\"").append(svgNumber(y - 3.0)).append("\" r=\"3.7\" fill=\"")
                    .append(legendColors.get(i)).append("\"/>\n");
            svg.append("<text x=\"").append(svgNumber(labelX)).append("\" y=\"").append(svgNumber(y)).append("\" text-anchor=\"start\" font-size=\"12\" fill=\"#2f2f2f\">")
                    .append(escapeHtml(displayLabel))
                    .append("</text>\n");
            svg.append("<text x=\"").append(svgNumber(pctX)).append("\" y=\"").append(svgNumber(y))
                    .append("\" text-anchor=\"end\" font-size=\"12\" fill=\"#4a4a4a\">")
                    .append(escapeHtml(formatNumber(legendPcts.get(i), 1) + "%"))
                    .append("</text>\n");
        }

        svg.append("</svg>\n");
        return svg.toString();
    }

    public static String buildSectorAllocationSvg(List<OverviewRow> rows, Map<String, Double> ratesToNok) {
        return buildCategoricalAllocationSvg(rows, true, "Sector Mix", "No sector data", ratesToNok);
    }

    public static String buildRegionAllocationSvg(List<OverviewRow> rows, Map<String, Double> ratesToNok) {
        return buildCategoricalAllocationSvg(rows, false, "Region Mix", "No region data", ratesToNok);
    }

    private static String buildCategoricalAllocationSvg(List<OverviewRow> rows, boolean sectorChart,
                                                        String centerTitle, String emptyMessage,
                                                        Map<String, Double> ratesToNok) {
        final double width = 440.0;
        final double height = 330.0;
        final double centerX = width / 2.0;
        final double centerY = 118.0;
        final double radius = 98.0;

        LinkedHashMap<String, Double> valueByCategory = new LinkedHashMap<>();
        double totalMarketValue = 0.0;
        for (OverviewRow row : rows) {
            double marketValueNok = convertToNok(row.marketValue, row.currencyCode, ratesToNok);
            if (marketValueNok <= 0.0) {
                continue;
            }

            if (sectorChart && row.sectorWeights != null && !row.sectorWeights.isEmpty()) {
                double totalWeight = 0.0;
                for (double weight : row.sectorWeights.values()) {
                    if (Double.isFinite(weight) && weight > 0.0) {
                        totalWeight += weight;
                    }
                }

                if (totalWeight > 0.0) {
                    for (Map.Entry<String, Double> entry : row.sectorWeights.entrySet()) {
                        double weight = entry.getValue() == null ? 0.0 : entry.getValue();
                        if (!Double.isFinite(weight) || weight <= 0.0) {
                            continue;
                        }

                        String sectorLabel = entry.getKey() == null || entry.getKey().isBlank()
                                ? "Other"
                                : entry.getKey();
                        double weightedValue = marketValueNok * (weight / totalWeight);
                        valueByCategory.merge(sectorLabel, weightedValue, Double::sum);
                    }

                    totalMarketValue += marketValueNok;
                    continue;
                }
            }

            if (!sectorChart && row.regionWeights != null && !row.regionWeights.isEmpty()) {
                double totalWeight = 0.0;
                for (double weight : row.regionWeights.values()) {
                    if (Double.isFinite(weight) && weight > 0.0) {
                        totalWeight += weight;
                    }
                }

                if (totalWeight > 0.0) {
                    for (Map.Entry<String, Double> entry : row.regionWeights.entrySet()) {
                        double weight = entry.getValue() == null ? 0.0 : entry.getValue();
                        if (!Double.isFinite(weight) || weight <= 0.0) {
                            continue;
                        }

                        String regionLabel = entry.getKey() == null || entry.getKey().isBlank()
                                ? "Global"
                                : entry.getKey();
                        double weightedValue = marketValueNok * (weight / totalWeight);
                        valueByCategory.merge(regionLabel, weightedValue, Double::sum);
                    }

                    totalMarketValue += marketValueNok;
                    continue;
                }
            }

            String category = sectorChart ? classifySector(row) : classifyRegion(row);
            valueByCategory.merge(category, marketValueNok, Double::sum);
            totalMarketValue += marketValueNok;
        }

        int maxBuckets = sectorChart ? 12 : 10;
        List<AllocationBucket> buckets = compactAllocationBuckets(valueByCategory, maxBuckets);

        StringBuilder svg = new StringBuilder();
        svg.append("<svg class=\"chart-svg\" viewBox=\"0 0 ")
                .append(svgNumber(width)).append(" ").append(svgNumber(height))
                .append("\" xmlns=\"http://www.w3.org/2000/svg\" role=\"img\">\n");

        if (buckets.isEmpty() || totalMarketValue <= 0.0) {
            svg.append("<text x=\"").append(svgNumber(centerX)).append("\" y=\"").append(svgNumber(centerY))
                    .append("\" text-anchor=\"middle\" font-size=\"13\" fill=\"#666\">")
                    .append(escapeHtml(emptyMessage))
                    .append("</text>\n");
            svg.append("</svg>\n");
            return svg.toString();
        }

        double currentAngle = -Math.PI / 2.0;
        for (int i = 0; i < buckets.size(); i++) {
            AllocationBucket bucket = buckets.get(i);
            double fraction = bucket.value / totalMarketValue;
            double sliceAngle = fraction * Math.PI * 2.0;
            double endAngle = currentAngle + sliceAngle;
            String color = getAllocationColor(i);

            double x1 = centerX + radius * Math.cos(currentAngle);
            double y1 = centerY + radius * Math.sin(currentAngle);
            double x2 = centerX + radius * Math.cos(endAngle);
            double y2 = centerY + radius * Math.sin(endAngle);
            int largeArcFlag = sliceAngle > Math.PI ? 1 : 0;

                svg.append("<path class=\"chart-hover-target chart-hover-slice\" d=\"M ").append(svgNumber(centerX)).append(" ").append(svgNumber(centerY))
                    .append(" L ").append(svgNumber(x1)).append(" ").append(svgNumber(y1))
                    .append(" A ").append(svgNumber(radius)).append(" ").append(svgNumber(radius)).append(" 0 ")
                    .append(largeArcFlag).append(" 1 ").append(svgNumber(x2)).append(" ").append(svgNumber(y2))
                    .append(" Z\" fill=\"").append(color).append("\">\n")
                    .append("<title class=\"js-chart-money\" data-value-nok=\"").append(svgNumber(bucket.value)).append("\" data-decimals=\"2\" data-prefix=\"")
                    .append(escapeHtml(bucket.label + ": "))
                    .append("\" data-suffix=\" (").append(escapeHtml(formatNumber(fraction * 100.0, 1))).append("%)\">")
                    .append(escapeHtml(bucket.label + ": " + formatNumber(bucket.value, 2) + " NOK (" + formatNumber(fraction * 100.0, 1) + "%)"))
                    .append("</title></path>\n");

            currentAngle = endAngle;
        }

        double summaryY = centerY + radius + 14.0;
        svg.append("<text x=\"").append(svgNumber(centerX)).append("\" y=\"").append(svgNumber(summaryY))
                .append("\" text-anchor=\"middle\" font-size=\"12\" fill=\"#666\">")
                .append(escapeHtml(centerTitle))
                .append("</text>\n");

        boolean useTwoColumns = buckets.size() > 6;
        int rowsPerColumn = useTwoColumns ? (buckets.size() + 1) / 2 : buckets.size();
        double legendYStart = summaryY + 18.0;
        double legendBottomY = height - 14.0;
        double legendRowGap = computeLegendRowGap(rowsPerColumn, legendYStart, legendBottomY, 15.0);
        for (int i = 0; i < buckets.size(); i++) {
            AllocationBucket bucket = buckets.get(i);
            double pct = (bucket.value / totalMarketValue) * 100.0;
            int columnIndex = (useTwoColumns && i >= rowsPerColumn) ? 1 : 0;
            int rowIndex = useTwoColumns ? (i % rowsPerColumn) : i;
            double y = legendYStart + (rowIndex * legendRowGap);
            double dotX = columnIndex == 0 ? 22.0 : 228.0;
            double labelX = dotX + 9.0;
            double pctX = useTwoColumns
                    ? (columnIndex == 0 ? 214.0 : (width - 16.0))
                    : (width - 16.0);
            String color = getAllocationColor(i);
            String displayLabel = abbreviateLegendLabel(bucket.label, useTwoColumns ? 22 : 36);

            svg.append("<circle cx=\"").append(svgNumber(dotX)).append("\" cy=\"").append(svgNumber(y - 3.0)).append("\" r=\"3.7\" fill=\"")
                    .append(color).append("\"/>\n");
            svg.append("<text x=\"").append(svgNumber(labelX)).append("\" y=\"").append(svgNumber(y)).append("\" text-anchor=\"start\" font-size=\"12\" fill=\"#2f2f2f\">")
                    .append(escapeHtml(displayLabel))
                    .append("</text>\n");
            svg.append("<text x=\"").append(svgNumber(pctX)).append("\" y=\"").append(svgNumber(y))
                    .append("\" text-anchor=\"end\" font-size=\"12\" fill=\"#4a4a4a\">")
                    .append(escapeHtml(formatNumber(pct, 1) + "%"))
                    .append("</text>\n");
        }

        svg.append("</svg>\n");
        return svg.toString();
    }

    private static List<AllocationBucket> compactAllocationBuckets(Map<String, Double> rawBuckets, int maxBuckets) {
        List<AllocationBucket> sorted = new ArrayList<>();
        for (Map.Entry<String, Double> entry : rawBuckets.entrySet()) {
            if (entry.getValue() <= 0.0) {
                continue;
            }
            sorted.add(new AllocationBucket(entry.getKey(), entry.getValue()));
        }
        sorted.sort((a, b) -> Double.compare(b.value, a.value));

        if (sorted.size() <= maxBuckets) {
            return sorted;
        }

        int directBuckets = Math.max(1, maxBuckets - 1);
        List<AllocationBucket> compacted = new ArrayList<>();
        double otherValue = 0.0;
        for (int i = 0; i < sorted.size(); i++) {
            AllocationBucket bucket = sorted.get(i);
            if (i < directBuckets) {
                compacted.add(bucket);
            } else {
                otherValue += bucket.value;
            }
        }

        if (otherValue > 0.0) {
            compacted.add(new AllocationBucket("Other", otherValue));
        }
        return compacted;
    }

    private static double mapValueToY(double value, double minValue, double maxValue, double chartTop, double chartHeight) {
        if (Math.abs(maxValue - minValue) < 1e-12) return chartTop + chartHeight / 2.0;
        return chartTop + ((maxValue - value) / (maxValue - minValue)) * chartHeight;
    }

    private static String svgNumber(double value) {
        return String.format(Locale.US, "%.2f", value);
    }

    private static String formatChartValue(double value, boolean percentChart, boolean compact) {
        if (percentChart) {
            return formatNumber(value, 2) + "%";
        }
        int decimals = compact ? 0 : 2;
        return formatNumber(value, decimals) + " NOK";
    }

    private static double convertToNok(double amount, String currencyCode, Map<String, Double> ratesToNok) {
        if (!Double.isFinite(amount)) {
            return 0.0;
        }
        if (ratesToNok == null || ratesToNok.isEmpty()) {
            return amount;
        }
        String normalized = normalizeCurrencyCode(currencyCode);
        double rate = ratesToNok.getOrDefault(normalized, 0.0);
        if (rate <= 0.0) {
            rate = ratesToNok.getOrDefault("NOK", 1.0);
        }
        return amount * rate;
    }

    private static String normalizeCurrencyCode(String currencyCode) {
        if (currencyCode == null || currencyCode.isBlank()) {
            return "NOK";
        }
        String normalized = currencyCode.trim().toUpperCase(Locale.ROOT);
        if (!normalized.matches("[A-Z]{3}")) {
            return "NOK";
        }
        return normalized;
    }

    private static String formatNumber(double value, int decimals) {
        return String.format(Locale.US, "%,." + decimals + "f", value);
    }

    private static String escapeHtml(String text) {
        if (text == null) return "";
        return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\"", "&quot;");
    }

    private static String getOverviewRowLabel(OverviewRow row) {
        return (row.securityDisplayName == null || row.securityDisplayName.isBlank()) ? row.tickerText : row.securityDisplayName;
    }

    private static String getCompactPieLabel(OverviewRow row) {
        String label = getOverviewRowLabel(row);
        return (label == null || label.isBlank()) ? "-" : label;
    }

    private static String getCompactBarLabel(OverviewRow row) {
        String fullLabel = getOverviewRowLabel(row);
        if (fullLabel == null || fullLabel.isBlank()) {
            return "-";
        }

        if (fullLabel.length() <= 24) {
            return fullLabel;
        }

        boolean hasTicker = row.tickerText != null && !row.tickerText.isBlank() && !"-".equals(row.tickerText);
        if ("STOCK".equals(row.assetType) && hasTicker) {
            return row.tickerText;
        }

        return fullLabel.substring(0, 21) + "...";
    }

    private static String getCompactOverviewLabel(OverviewRow row) {
        String fullLabel = getOverviewRowLabel(row);
        if (fullLabel == null || fullLabel.isBlank()) {
            return "-";
        }

        if (fullLabel.length() <= 26) {
            return fullLabel;
        }

        boolean hasTicker = row.tickerText != null && !row.tickerText.isBlank() && !"-".equals(row.tickerText);
        if ("STOCK".equals(row.assetType) && hasTicker) {
            return row.tickerText;
        }

        return fullLabel.substring(0, 23) + "...";
    }

    private static String getAllocationColor(int index) {
        String[] palette = new String[] {
                "#0b7285", "#2f9e44", "#f08c00", "#7048e8", "#c92a2a", "#1c7ed6", "#5f3dc4", "#2b8a3e", "#e67700", "#0ca678"
        };
        return palette[index % palette.length];
    }

    private static void adjustPieLabelPositions(List<PieSliceLabel> labels, boolean rightSide,
                                                double minY, double maxY, double minGap) {
        List<PieSliceLabel> sideLabels = new ArrayList<>();
        for (PieSliceLabel label : labels) {
            if (label.rightSide == rightSide) {
                sideLabels.add(label);
            }
        }

        sideLabels.sort(Comparator.comparingDouble(label -> label.labelY));
        if (sideLabels.isEmpty()) {
            return;
        }

        sideLabels.get(0).labelY = Math.max(minY, sideLabels.get(0).labelY);
        for (int i = 1; i < sideLabels.size(); i++) {
            PieSliceLabel current = sideLabels.get(i);
            PieSliceLabel previous = sideLabels.get(i - 1);
            current.labelY = Math.max(current.labelY, previous.labelY + minGap);
        }

        double overflow = sideLabels.get(sideLabels.size() - 1).labelY - maxY;
        if (overflow > 0.0) {
            for (PieSliceLabel label : sideLabels) {
                label.labelY -= overflow;
            }
        }

        double underflow = minY - sideLabels.get(0).labelY;
        if (underflow > 0.0) {
            for (PieSliceLabel label : sideLabels) {
                label.labelY += underflow;
            }
        }
    }

    private static double computeLegendRowGap(int rowsPerColumn, double startY, double maxY, double preferredGap) {
        if (rowsPerColumn <= 1) {
            return 0.0;
        }

        double availableHeight = Math.max(0.0, maxY - startY);
        return Math.min(preferredGap, availableHeight / (rowsPerColumn - 1));
    }

    private static String abbreviateLegendLabel(String label, int maxChars) {
        String text = label == null ? "" : label.trim();
        if (text.isEmpty()) {
            return "-";
        }

        int safeMax = Math.max(6, maxChars);
        if (text.length() <= safeMax) {
            return text;
        }

        return text.substring(0, safeMax - 3) + "...";
    }

    private static String classifySector(OverviewRow row) {
        if (row == null || row.sectorLabel == null || row.sectorLabel.isBlank()) {
            return "Other";
        }
        return row.sectorLabel;
    }

    private static String classifyRegion(OverviewRow row) {
        if (row == null || row.regionLabel == null || row.regionLabel.isBlank()) {
            return "Global";
        }
        return row.regionLabel;
    }
}