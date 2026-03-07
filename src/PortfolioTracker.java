import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.Locale;
import java.util.Map;

public class PortfolioTracker {
    private static final String INPUT_DIRECTORY = "transaction_files";
    private static final String OUTPUT_FILE = "portfolio-report.html";
        private static final Map<String, String> RENAMED_SECURITY_ISIN = Map.of(
            "NO0010782519", "NO0012948878"
        );

    private static final ArrayList<Security> securities = new ArrayList<>();
    private static final Map<String, Security> securitiesByKey = new LinkedHashMap<>();

    public static void main(String[] args) throws IOException {
        ensureInputDirectoryExists();
        int filesProcessed = readAllTransactionFiles();

        if (filesProcessed == 0) {
            System.out.println("No CSV files found in '" + INPUT_DIRECTORY + "'.");
            System.out.println("Place one or more transaction exports in that folder and run again.");
        } else {
            System.out.println("Loaded " + filesProcessed + " CSV file(s) from '" + INPUT_DIRECTORY + "'.");
        }

        writeHtmlReport();
    }

    private static void ensureInputDirectoryExists() {
        File directory = new File(INPUT_DIRECTORY);
        if (!directory.exists() && directory.mkdirs()) {
            System.out.println("Created input folder: " + INPUT_DIRECTORY);
        }
    }

    private static int readAllTransactionFiles() throws IOException {
        File directory = new File(INPUT_DIRECTORY);
        if (!directory.exists() || !directory.isDirectory()) {
            return 0;
        }

        File[] csvFiles = directory.listFiles((dir, name) ->
            name.toLowerCase(Locale.ROOT).endsWith(".csv")
                && !name.equalsIgnoreCase(OUTPUT_FILE)
                && !name.toLowerCase(Locale.ROOT).contains("example"));

        if (csvFiles == null || csvFiles.length == 0) {
            return 0;
        }

        Arrays.sort(csvFiles, Comparator.comparing(File::getName, String.CASE_INSENSITIVE_ORDER));

        int processedCount = 0;
        for (File csvFile : csvFiles) {
            if (readFile(csvFile)) {
                processedCount++;
            }
        }
        return processedCount;
    }

    private static double parseDoubleOrZero(String value) {
        if (value == null || value.trim().isEmpty()) {
            return 0.0;
        }

        String normalized = value.trim()
                .replace("\u00A0", "")
                .replace("−", "-")
                .replace(" ", "");

        if (normalized.contains(",") && normalized.contains(".")) {
            normalized = normalized.replace(".", "").replace(",", ".");
        } else {
            normalized = normalized.replace(",", ".");
        }

        try {
            return Double.parseDouble(normalized);
        } catch (NumberFormatException ex) {
            return 0.0;
        }
    }

    private static ArrayList<String> parseDelimitedLine(String line, char delimiter) {
        ArrayList<String> fields = new ArrayList<>();
        StringBuilder current = new StringBuilder();
        boolean inQuotes = false;

        for (int i = 0; i < line.length(); i++) {
            char currentChar = line.charAt(i);

            if (currentChar == '"') {
                boolean escapedQuote = inQuotes && i + 1 < line.length() && line.charAt(i + 1) == '"';
                if (escapedQuote) {
                    current.append('"');
                    i++;
                } else {
                    inQuotes = !inQuotes;
                }
            } else if (currentChar == delimiter && !inQuotes) {
                fields.add(current.toString());
                current.setLength(0);
            } else {
                current.append(currentChar);
            }
        }

        fields.add(current.toString());
        return fields;
    }

    private static char detectDelimiter(String line) {
        int tabCount = countOccurrences(line, '\t');
        int semicolonCount = countOccurrences(line, ';');
        int commaCount = countOccurrences(line, ',');

        if (tabCount >= semicolonCount && tabCount >= commaCount) {
            return '\t';
        }
        if (semicolonCount >= tabCount && semicolonCount >= commaCount) {
            return ';';
        }
        return ',';
    }

    private static int countOccurrences(String text, char target) {
        int count = 0;
        for (int i = 0; i < text.length(); i++) {
            if (text.charAt(i) == target) {
                count++;
            }
        }
        return count;
    }

    private static String getCell(ArrayList<String> row, int index) {
        if (index < 0 || index >= row.size()) {
            return "";
        }
        return row.get(index).trim().replace("\uFEFF", "");
    }

    private static String normalizeHeader(String value) {
        return value
                .replace("\uFEFF", "")
                .trim()
                .toLowerCase(Locale.ROOT);
    }

    private static Security getOrCreateSecurity(String name, String isin) {
        String key = isin == null || isin.isBlank() ? name : isin;
        Security security = securitiesByKey.get(key);
        if (security != null) {
            return security;
        }

        Security newSecurity = new Security(name, isin);
        securitiesByKey.put(key, newSecurity);
        securities.add(newSecurity);
        return newSecurity;
    }

    private static class HeaderIndexes {
        int transactionId;
        int securityName;
        int securityType;
        int isin;
        int transactionType;
        int amount;
        int quantity;
        int price;
        int result;
        int totalFees;
        int tradeDate;

        HeaderIndexes() {
            transactionId = securityName = securityType = isin = transactionType = amount = quantity = price = result = totalFees = tradeDate = -1;
        }

        boolean hasRequiredColumns() {
            return securityName >= 0 && transactionType >= 0;
        }
    }

    private static final class OverviewRow {
        private final String tickerText;
        private final String securityDisplayName;
        private final String assetType;
        private final String latestPriceText;
        private final String realizedReturnPctText;
        private final String realizedGainText;
        private final String dividendsText;
        private final double units;
        private final double averageCost;
        private final double latestPrice;
        private final double positionCostBasis;
        private final double marketValue;
        private final double unrealized;
        private final double unrealizedPct;
        private final double realized;
        private final double dividends;
        private final double totalReturn;
        private final double totalReturnPct;
        private final boolean hasPrice;

        private OverviewRow(
                String tickerText,
                String securityDisplayName,
                String assetType,
                String latestPriceText,
                String realizedReturnPctText,
                String realizedGainText,
                String dividendsText,
                double units,
                double averageCost,
                double latestPrice,
                double positionCostBasis,
                double marketValue,
                double unrealized,
                double unrealizedPct,
                double realized,
                double dividends,
                double totalReturn,
                double totalReturnPct,
                boolean hasPrice) {
            this.tickerText = tickerText;
            this.securityDisplayName = securityDisplayName;
            this.assetType = assetType;
            this.latestPriceText = latestPriceText;
            this.realizedReturnPctText = realizedReturnPctText;
            this.realizedGainText = realizedGainText;
            this.dividendsText = dividendsText;
            this.units = units;
            this.averageCost = averageCost;
            this.latestPrice = latestPrice;
            this.positionCostBasis = positionCostBasis;
            this.marketValue = marketValue;
            this.unrealized = unrealized;
            this.unrealizedPct = unrealizedPct;
            this.realized = realized;
            this.dividends = dividends;
            this.totalReturn = totalReturn;
            this.totalReturnPct = totalReturnPct;
            this.hasPrice = hasPrice;
        }
    }

    private static HeaderIndexes findHeaderIndexes(ArrayList<String> headerColumns) {
        HeaderIndexes indexes = new HeaderIndexes();

        for (int i = 0; i < headerColumns.size(); i++) {
            String column = normalizeHeader(headerColumns.get(i));
            switch (column) {
                case "id", "transaksjon" -> indexes.transactionId = i;
                case "verdipapir" -> indexes.securityName = i;
                case "verdipapirtype" -> indexes.securityType = i;
                case "isin" -> indexes.isin = i;
                case "transaksjonstype" -> indexes.transactionType = i;
                case "beløp", "belop", "handelsbeløp", "handelsbelop" -> indexes.amount = i;
                case "antall" -> indexes.quantity = i;
                case "kurs", "kurs per andel" -> indexes.price = i;
                case "resultat" -> indexes.result = i;
                case "totale avgifter", "omkostninger" -> indexes.totalFees = i;
                case "handelsdag", "handelsdato" -> indexes.tradeDate = i;
                default -> {
                    // Ignored.
                }
            }
        }
        return indexes;
    }

    private static void processRow(ArrayList<String> row, HeaderIndexes indexes) {
        if (row == null || row.isEmpty()) {
            return;
        }

        String securityName = getCell(row, indexes.securityName);
        if (securityName.isEmpty()) {
            return;
        }

        String isin = getCell(row, indexes.isin);
        String securityType = getCell(row, indexes.securityType);
        Security security = getOrCreateSecurity(securityName, isin);
        if (security == null) {
            return;
        }

        security.updateAssetTypeFromHint(securityType, securityName);

        processTransaction(security, row, indexes);
    }

    private static LocalDate parseTradeDateForSort(String tradeDateText) {
        if (tradeDateText == null || tradeDateText.isBlank()) {
            return LocalDate.MIN;
        }

        String value = tradeDateText.trim();
        DateTimeFormatter[] formats = new DateTimeFormatter[] {
            DateTimeFormatter.ISO_LOCAL_DATE,
            DateTimeFormatter.ofPattern("dd.MM.yyyy")
        };

        for (DateTimeFormatter formatter : formats) {
            try {
                return LocalDate.parse(value, formatter);
            } catch (DateTimeParseException ignored) {
                // Try next format.
            }
        }

        return LocalDate.MIN;
    }

    private static long parseSortIdOrDefault(String value) {
        if (value == null || value.isBlank()) {
            return Long.MAX_VALUE;
        }

        try {
            return Long.parseLong(value.trim());
        } catch (NumberFormatException ignored) {
            return Long.MAX_VALUE;
        }
    }

    private static int getAssetPriority(Security security) {
        return switch (security.getAssetType()) {
            case STOCK -> 0;
            case FUND -> 1;
            case UNKNOWN -> 0;
        };
    }

    private static ArrayList<Security> getSortedSecuritiesForOverview() {
        ArrayList<Security> sorted = new ArrayList<>();
        for (Security security : securities) {
            if (security.getUnitsOwned() > 0.0000001
                    && !isReplacedSecurity(security)
                    && !isTemporaryRightsSecurity(security)) {
                sorted.add(security);
            }
        }

        sorted.sort(
                Comparator.comparingInt(PortfolioTracker::getAssetPriority)
                        .thenComparing(Security::getName, String.CASE_INSENSITIVE_ORDER)
        );
        return sorted;
    }

    private static boolean isReplacedSecurity(Security security) {
        String replacementIsin = RENAMED_SECURITY_ISIN.get(security.getIsin());
        if (replacementIsin == null || replacementIsin.isBlank()) {
            return false;
        }

        Security replacement = securitiesByKey.get(replacementIsin);
        return replacement != null && replacement.getUnitsOwned() > 0.0000001;
    }

    private static boolean isTemporaryRightsSecurity(Security security) {
        String name = security.getName();
        if (name == null || name.isBlank()) {
            return false;
        }

        String normalized = name.toUpperCase(Locale.ROOT);
        return normalized.contains("T-RETT") || normalized.contains("TEGNINGSRETT");
    }

    private static ArrayList<Security> getSortedSoldSecurities() {
        ArrayList<Security> soldSecurities = new ArrayList<>();
        for (Security security : securities) {
            if (security.hasSales()) {
                soldSecurities.add(security);
            }
        }

        soldSecurities.sort(
                Comparator.comparingInt(PortfolioTracker::getAssetPriority)
                        .thenComparing(Security::getRealizedSalesValue, Comparator.reverseOrder())
                        .thenComparing(Security::getName, String.CASE_INSENSITIVE_ORDER)
        );

        return soldSecurities;
    }

    private static void processTransaction(Security security, ArrayList<String> row, HeaderIndexes indexes) {
        String transactionType = getCell(row, indexes.transactionType)
                .toUpperCase(Locale.ROOT)
                .replaceAll("\\s+", " ");

        String tradeDate = getCell(row, indexes.tradeDate);

        double amount = parseDoubleOrZero(getCell(row, indexes.amount));
        double quantity = parseDoubleOrZero(getCell(row, indexes.quantity));
        double price = parseDoubleOrZero(getCell(row, indexes.price));
        double result = parseDoubleOrZero(getCell(row, indexes.result));
        double totalFees = parseDoubleOrZero(getCell(row, indexes.totalFees));

        switch (transactionType) {
            case "SALG", "SELL", "KJØPT", "KJOPT", "KJØP", "KJOP", "BUY", "REINVESTERT UTBYTTE", "REINVESTERTUTBYTTE" ->
                    security.addTransaction(tradeDate, transactionType, amount, quantity, price, result, totalFees);
            case "UTBYTTE INNLEGG VP", "BYTTE INNLEGG VP", "MAK BYTTE INNLEGG VP", "TILDELING INNLEGG RE" -> {
                if (!isTemporaryRightsSecurity(security)) {
                    security.addZeroCostUnits(quantity);
                }
            }
            case "UTBYTTE", "DIVIDEND" -> security.addDividend(amount);
            case "INNSKUDD", "UTTAK INTERNET", "PLATTFORMAVGIFT", "TILBAKEBET. FOND AVG",
                    "OVERBELÅNINGSRENTE", "TILBAKEBETALING", "OVERFØRING VIA TRUSTLY", "OVERFORING VIA TRUSTLY" -> {
                // Ignored on purpose; these are cash-account events.
            }
            default -> {
                // Intentionally ignored unknown transaction types.
            }
        }
    }

    private static boolean readFile(File file) throws IOException {
        try {
            byte[] bytes = Files.readAllBytes(file.toPath());
            if (bytes.length == 0) {
                return false;
            }

            Charset charset = detectCharset(bytes);
            String content = new String(bytes, charset).replace("\uFEFF", "");
            String[] lines = content.split("\\R");

            if (lines.length == 0 || lines[0].isBlank()) {
                return false;
            }

            String header = lines[0];
            char delimiter = detectDelimiter(header);
            HeaderIndexes indexes = findHeaderIndexes(parseDelimitedLine(header, delimiter));

            if (!indexes.hasRequiredColumns()) {
                System.out.println("Skipping unsupported CSV format: " + file.getName());
                return false;
            }

            ArrayList<ArrayList<String>> rows = new ArrayList<>();
            for (int i = 1; i < lines.length; i++) {
                String line = lines[i];
                if (line == null || line.isBlank()) {
                    continue;
                }
                rows.add(parseDelimitedLine(line.replace("−", "-"), delimiter));
            }

            if (indexes.tradeDate >= 0) {
                rows.sort(
                    Comparator
                        .comparing((ArrayList<String> row) -> parseTradeDateForSort(getCell(row, indexes.tradeDate)))
                        .thenComparingLong(row -> parseSortIdOrDefault(getCell(row, indexes.transactionId)))
                );
            }

            for (ArrayList<String> row : rows) {
                processRow(row, indexes);
            }
            return true;
        } catch (FileNotFoundException e) {
            System.out.println(e.getMessage());
            return false;
        }
    }

    private static Charset detectCharset(byte[] bytes) {
        if (bytes.length >= 3
                && bytes[0] == (byte) 0xEF
                && bytes[1] == (byte) 0xBB
                && bytes[2] == (byte) 0xBF) {
            return StandardCharsets.UTF_8;
        }

        if (bytes.length >= 2
                && bytes[0] == (byte) 0xFF
                && bytes[1] == (byte) 0xFE) {
            return StandardCharsets.UTF_16LE;
        }

        if (bytes.length >= 2
                && bytes[0] == (byte) 0xFE
                && bytes[1] == (byte) 0xFF) {
            return StandardCharsets.UTF_16BE;
        }

        int sampleSize = Math.min(bytes.length, 4000);
        int zeroCount = 0;
        for (int i = 0; i < sampleSize; i++) {
            if (bytes[i] == 0) {
                zeroCount++;
            }
        }

        if (zeroCount > sampleSize / 10) {
            return StandardCharsets.UTF_16LE;
        }

        return StandardCharsets.UTF_8;
    }

    private static void writeHtmlReport() {
        File file = new File(OUTPUT_FILE);
        try (FileWriter writer = new FileWriter(file)) {
            writeHtmlHeader(writer);
            writeOverviewTableHtml(writer);
            writeRealizedSummaryTableHtml(writer);
            writeSaleTradesTablesHtml(writer);
            writeHtmlFooter(writer);
            new ProcessBuilder("open", OUTPUT_FILE).start();
        } catch (IOException e) {
            System.out.println("Failed to write report: " + e.getMessage());
        }
    }

    private static void writeHtmlHeader(FileWriter writer) throws IOException {
        writer.write("<!DOCTYPE html>\n");
        writer.write("<html lang=\"en\">\n");
        writer.write("<head>\n");
        writer.write("  <meta charset=\"utf-8\">\n");
        writer.write("  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n");
        writer.write("  <title>Portfolio Report</title>\n");
        writer.write("  <style>\n");
        writer.write("    body { font-family: -apple-system, BlinkMacSystemFont, Segoe UI, sans-serif; margin: 24px; color: #111; }\n");
        writer.write("    h1 { margin: 0 0 20px 0; font-size: 26px; }\n");
        writer.write("    h2 { margin: 28px 0 8px 0; font-size: 18px; }\n");
        writer.write("    table { border-collapse: collapse; width: 100%; margin: 8px 0 18px 0; table-layout: auto; }\n");
        writer.write("    th, td { border: 1px solid #d0d0d0; padding: 6px 8px; font-size: 13px; text-align: left; white-space: nowrap; }\n");
        writer.write("    th { background: #f3f3f3; font-weight: 600; }\n");
        writer.write("    td.num { text-align: right; font-variant-numeric: tabular-nums; }\n");
        writer.write("    td.text { text-align: left; }\n");
        writer.write("    tr.total-row td { font-weight: 700; background: #fafafa; }\n");
        writer.write("    .muted { color: #666; font-size: 12px; margin-top: -8px; }\n");
        writer.write("    .overview-charts { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 14px; margin: 8px 0 14px 0; align-items: stretch; }\n");
        writer.write("    .overview-chart { border: 1px solid #d0d0d0; border-radius: 6px; background: #fff; padding: 10px; }\n");
        writer.write("    .overview-chart h3 { margin: 0 0 8px 0; font-size: 14px; font-weight: 600; }\n");
        writer.write("    .chart-svg { width: 100%; height: 290px; display: block; }\n");
        writer.write("    .overview-chart.allocation-card { grid-column: 1 / -1; }\n");
        writer.write("    .overview-chart.allocation-card .chart-svg { height: 240px; }\n");
        writer.write("    .allocation-visuals { display: grid; grid-template-columns: 0.75fr 2.2fr 0.75fr; gap: 12px; }\n");
        writer.write("    .allocation-panel { border: 1px solid #ececec; border-radius: 6px; padding: 6px; }\n");
        writer.write("    .allocation-legend { margin: 8px 0 0 0; padding: 0; list-style: none; display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); column-gap: 10px; row-gap: 4px; }\n");
        writer.write("    .allocation-legend li { display: flex; align-items: center; justify-content: space-between; gap: 8px; font-size: 12px; color: #333; }\n");
        writer.write("    .allocation-legend .name { display: inline-flex; align-items: center; gap: 6px; min-width: 0; }\n");
        writer.write("    .allocation-legend .name-text { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; max-width: 150px; }\n");
        writer.write("    .allocation-legend .dot { width: 9px; height: 9px; border-radius: 50%; flex: 0 0 auto; }\n");
        writer.write("    .allocation-legend .value { font-variant-numeric: tabular-nums; color: #555; }\n");
        writer.write("    @media (max-width: 980px) { .allocation-visuals { grid-template-columns: 1fr; } }\n");
        writer.write("    @media (max-width: 980px) { .overview-charts { grid-template-columns: 1fr; } }\n");
        writer.write("  </style>\n");
        writer.write("</head>\n");
        writer.write("<body>\n");
        writer.write("  <h1>Portfolio Report</h1>\n");
        writer.write("  <div class=\"muted\">Generated from all CSV files in transaction_files/</div>\n");
    }

    private static void writeHtmlFooter(FileWriter writer) throws IOException {
        writer.write("</body>\n");
        writer.write("</html>\n");
    }

    private static String escapeHtml(String value) {
        String text = value == null ? "" : value;
        return text
                .replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace("\"", "&quot;")
                .replace("'", "&#39;");
    }

    private static boolean isNumericCellValue(String value) {
        if (value == null) {
            return false;
        }

        String text = value.trim();
        if (text.isEmpty() || "-".equals(text)) {
            return false;
        }

        return text.matches("-?[0-9][0-9 ]*(\\.[0-9]+)?");
    }

    private static String toDataCell(String value) {
        String cellClass = isNumericCellValue(value) ? "num" : "text";
        return "<td class=\"" + cellClass + "\">" + escapeHtml(value) + "</td>";
    }

    private static void writeHtmlRow(FileWriter writer, boolean header, String... fields) throws IOException {
        writer.write("<tr>");
        for (int i = 0; i < fields.length; i++) {
            if (header) {
                writer.write("<th>" + escapeHtml(fields[i]) + "</th>");
            } else {
                writer.write(toDataCell(fields[i]));
            }
        }
        writer.write("</tr>\n");
    }

    private static OverviewRow buildOverviewRow(Security security) {
        String ticker = security.getTicker();
        String tickerText = (ticker == null || ticker.isBlank()) ? "-" : ticker;
        double units = security.getUnitsOwned();
        double averageCost = security.getAverageCost();
        double latestPrice = security.getLatestPrice();
        double positionCostBasis = units * averageCost;
        boolean hasPrice = latestPrice > 0.0;
        double marketValue = hasPrice ? units * latestPrice : 0.0;
        double unrealized = hasPrice ? (marketValue - positionCostBasis) : 0.0;
        double unrealizedPct = hasPrice && positionCostBasis > 0 ? (unrealized / positionCostBasis) * 100.0 : 0.0;
        double realized = parseDoubleOrZero(security.getRealizedGainAsText());
        double dividends = parseDoubleOrZero(security.getDividendsAsText());
        double totalReturn = unrealized + realized + dividends;
        double totalReturnPct = positionCostBasis > 0 ? (totalReturn / positionCostBasis) * 100.0 : 0.0;

        return new OverviewRow(
                tickerText,
                security.getDisplayName(),
                security.getAssetType().name(),
                security.getLatestPriceAsText(),
                security.getRealizedReturnPctAsText(),
                security.getRealizedGainAsText(),
                security.getDividendsAsText(),
                units,
                averageCost,
                latestPrice,
                positionCostBasis,
                marketValue,
                unrealized,
                unrealizedPct,
                realized,
                dividends,
                totalReturn,
                totalReturnPct,
                hasPrice
        );
    }

    private static void writeOverviewChartsHtml(FileWriter writer, ArrayList<OverviewRow> rows) throws IOException {
        if (rows.isEmpty()) {
            return;
        }

        writer.write("<div class=\"overview-charts\">\n");
        writeOverviewChartCard(writer, "Total Return (NOK)", rows, false);
        writeOverviewChartCard(writer, "Total Return (%)", rows, true);
        writeMarketValueAllocationCard(writer, rows);
        writer.write("</div>\n");
    }

    private static void writeOverviewChartCard(FileWriter writer, String title, ArrayList<OverviewRow> rows, boolean percentChart) throws IOException {
        writer.write("<section class=\"overview-chart\">\n");
        writer.write("<h3>" + escapeHtml(title) + "</h3>\n");
        writer.write(buildOverviewBarChartSvg(rows, percentChart));
        writer.write("</section>\n");
    }

    private static void writeMarketValueAllocationCard(FileWriter writer, ArrayList<OverviewRow> rows) throws IOException {
        writer.write("<section class=\"overview-chart allocation-card\">\n");
        writer.write("<h3>Market Value Allocation</h3>\n");
        writer.write("<div class=\"allocation-visuals\">\n");
        writer.write("<div class=\"allocation-panel\">\n");
        writer.write(buildMarketValueAllocationSvg(rows));
        writer.write("</div>\n");
        writer.write("<div class=\"allocation-panel\">\n");
        writer.write(buildMarketValueBarChartSvg(rows));
        writer.write("</div>\n");
        writer.write("<div class=\"allocation-panel\">\n");
        writer.write(buildAssetTypeAllocationSvg(rows));
        writer.write("</div>\n");
        writer.write("</div>\n");

        ArrayList<OverviewRow> rowsWithValue = new ArrayList<>();
        double totalMarketValue = 0.0;
        for (OverviewRow row : rows) {
            if (row.marketValue > 0.0) {
                rowsWithValue.add(row);
                totalMarketValue += row.marketValue;
            }
        }

        if (!rowsWithValue.isEmpty() && totalMarketValue > 0.0) {
            writer.write("<ul class=\"allocation-legend\">\n");
            for (int i = 0; i < rowsWithValue.size(); i++) {
                OverviewRow row = rowsWithValue.get(i);
                String label = (row.securityDisplayName == null || row.securityDisplayName.isBlank()) ? row.tickerText : row.securityDisplayName;
                double pct = (row.marketValue / totalMarketValue) * 100.0;
                String color = getAllocationColor(i);
                writer.write("<li><span class=\"name\"><span class=\"dot\" style=\"background:" + color + "\"></span><span class=\"name-text\">"
                        + escapeHtml(label)
                        + "</span></span><span class=\"value\">"
                        + escapeHtml(formatNumber(pct, 1))
                        + "%</span></li>\n");
            }
            writer.write("</ul>\n");
        }

        writer.write("</section>\n");
    }

    private static String getOverviewRowLabel(OverviewRow row) {
        return (row.securityDisplayName == null || row.securityDisplayName.isBlank()) ? row.tickerText : row.securityDisplayName;
    }

    private static String buildOverviewBarChartSvg(ArrayList<OverviewRow> rows, boolean percentChart) {
        final double width = 1100.0;
        final double height = 360.0;
        final double left = 68.0;
        final double right = 22.0;
        final double top = 18.0;
        final double bottom = 122.0;
        final double plotWidth = width - left - right;
        final double plotHeight = height - top - bottom;

        double minValue = 0.0;
        double maxValue = 0.0;
        for (OverviewRow row : rows) {
            double value = percentChart ? row.totalReturnPct : row.totalReturn;
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

        int tickCount = 5;
        for (int i = 0; i <= tickCount; i++) {
            double tickValue = maxValue - ((valueRange / tickCount) * i);
            double y = mapValueToY(tickValue, minValue, maxValue, top, plotHeight);

            svg.append("<line x1=\"").append(svgNumber(left)).append("\" y1=\"").append(svgNumber(y))
                    .append("\" x2=\"").append(svgNumber(left + plotWidth)).append("\" y2=\"").append(svgNumber(y))
                    .append("\" stroke=\"#ececec\" stroke-width=\"1\"/>\n");

            svg.append("<text x=\"").append(svgNumber(left - 8.0)).append("\" y=\"").append(svgNumber(y + 4.0))
                    .append("\" text-anchor=\"end\" font-size=\"10\" fill=\"#666\">")
                    .append(escapeHtml(formatChartValue(tickValue, percentChart, true)))
                    .append("</text>\n");
        }

        double slotWidth = rows.isEmpty() ? plotWidth : plotWidth / rows.size();
        double barWidth = Math.max(6.0, slotWidth * 0.62);

        for (int i = 0; i < rows.size(); i++) {
            OverviewRow row = rows.get(i);
            double value = percentChart ? row.totalReturnPct : row.totalReturn;
            double x = left + (i * slotWidth) + ((slotWidth - barWidth) / 2.0);
            double yValue = mapValueToY(value, minValue, maxValue, top, plotHeight);
            double barY = Math.min(yValue, chartZeroY);
            double barHeight = Math.abs(chartZeroY - yValue);
            if (barHeight < 1.0) {
                barHeight = 1.0;
            }

            String barColor = value >= 0.0 ? "#2f9e44" : "#d94841";
            String label = (row.securityDisplayName == null || row.securityDisplayName.isBlank()) ? row.tickerText : row.securityDisplayName;

            svg.append("<rect x=\"").append(svgNumber(x)).append("\" y=\"").append(svgNumber(barY))
                    .append("\" width=\"").append(svgNumber(barWidth)).append("\" height=\"").append(svgNumber(barHeight))
                    .append("\" fill=\"").append(barColor).append("\" rx=\"2\">\n")
                    .append("<title>")
                    .append(escapeHtml(label + ": " + formatChartValue(value, percentChart, false)))
                    .append("</title></rect>\n");

            double labelAnchorX = x + (barWidth / 2.0);
            double labelAnchorY = height - bottom + 64.0;
            svg.append("<text x=\"").append(svgNumber(labelAnchorX)).append("\" y=\"").append(svgNumber(labelAnchorY))
                    .append("\" transform=\"rotate(-45 ").append(svgNumber(labelAnchorX)).append(" ").append(svgNumber(labelAnchorY))
                    .append(")\" text-anchor=\"end\" font-size=\"9\" fill=\"#444\">")
                    .append(escapeHtml(label))
                    .append("</text>\n");
        }

        svg.append("<line x1=\"").append(svgNumber(left)).append("\" y1=\"").append(svgNumber(chartZeroY))
                .append("\" x2=\"").append(svgNumber(left + plotWidth)).append("\" y2=\"").append(svgNumber(chartZeroY))
                .append("\" stroke=\"#7a7a7a\" stroke-width=\"1.1\"/>\n");
        svg.append("<line x1=\"").append(svgNumber(left)).append("\" y1=\"").append(svgNumber(top))
                .append("\" x2=\"").append(svgNumber(left)).append("\" y2=\"").append(svgNumber(top + plotHeight))
                .append("\" stroke=\"#7a7a7a\" stroke-width=\"1.1\"/>\n");

        svg.append("</svg>\n");
        return svg.toString();
    }

        private static String buildMarketValueAllocationSvg(ArrayList<OverviewRow> rows) {
        final double width = 440.0;
        final double height = 330.0;
        final double centerX = width / 2.0;
        final double centerY = 142.0;
        final double radius = 92.0;
        final double innerRadius = 50.0;

        ArrayList<OverviewRow> rowsWithValue = new ArrayList<>();
        double totalMarketValue = 0.0;
        for (OverviewRow row : rows) {
            if (row.marketValue > 0.0) {
            rowsWithValue.add(row);
            totalMarketValue += row.marketValue;
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
        for (int i = 0; i < rowsWithValue.size(); i++) {
            OverviewRow row = rowsWithValue.get(i);
            double fraction = row.marketValue / totalMarketValue;
            double sliceAngle = fraction * Math.PI * 2.0;
            double endAngle = currentAngle + sliceAngle;
            String color = getAllocationColor(i);

            double x1 = centerX + radius * Math.cos(currentAngle);
            double y1 = centerY + radius * Math.sin(currentAngle);
            double x2 = centerX + radius * Math.cos(endAngle);
            double y2 = centerY + radius * Math.sin(endAngle);
            int largeArcFlag = sliceAngle > Math.PI ? 1 : 0;

            svg.append("<path d=\"M ").append(svgNumber(centerX)).append(" ").append(svgNumber(centerY))
                .append(" L ").append(svgNumber(x1)).append(" ").append(svgNumber(y1))
                .append(" A ").append(svgNumber(radius)).append(" ").append(svgNumber(radius)).append(" 0 ")
                .append(largeArcFlag).append(" 1 ").append(svgNumber(x2)).append(" ").append(svgNumber(y2))
                .append(" Z\" fill=\"").append(color).append("\">\n")
                .append("<title>")
                    .append(escapeHtml(getOverviewRowLabel(row)
                    + ": " + formatNumber(row.marketValue, 2) + " kr (" + formatNumber(fraction * 100.0, 2) + "%)"))
                .append("</title></path>\n");

            currentAngle = endAngle;
        }

        svg.append("<circle cx=\"").append(svgNumber(centerX)).append("\" cy=\"").append(svgNumber(centerY))
            .append("\" r=\"").append(svgNumber(innerRadius)).append("\" fill=\"#fff\"/>\n");
        svg.append("<text x=\"").append(svgNumber(centerX)).append("\" y=\"").append(svgNumber(centerY - 4.0))
            .append("\" text-anchor=\"middle\" font-size=\"11\" fill=\"#666\">Market Value</text>\n");
        svg.append("<text x=\"").append(svgNumber(centerX)).append("\" y=\"").append(svgNumber(centerY + 14.0))
            .append("\" text-anchor=\"middle\" font-size=\"12\" fill=\"#222\" font-weight=\"600\">")
            .append(escapeHtml(formatNumber(totalMarketValue, 0) + " kr"))
            .append("</text>\n");

        svg.append("</svg>\n");
        return svg.toString();
        }

        private static String buildMarketValueBarChartSvg(ArrayList<OverviewRow> rows) {
        final double width = 860.0;
        final double height = 330.0;
        final double left = 74.0;
        final double right = 86.0;
        final double top = 18.0;
        final double bottom = 118.0;
        final double plotWidth = width - left - right;
        final double plotHeight = height - top - bottom;

        ArrayList<OverviewRow> rowsWithValue = new ArrayList<>();
        double totalMarketValue = 0.0;
        double maxValue = 0.0;
        for (OverviewRow row : rows) {
            if (row.marketValue > 0.0) {
            rowsWithValue.add(row);
            totalMarketValue += row.marketValue;
            maxValue = Math.max(maxValue, row.marketValue);
            }
        }

        StringBuilder svg = new StringBuilder();
        svg.append("<svg class=\"chart-svg\" viewBox=\"0 0 ")
            .append(svgNumber(width))
            .append(" ")
            .append(svgNumber(height))
            .append("\" xmlns=\"http://www.w3.org/2000/svg\" role=\"img\">\n");

        if (rowsWithValue.isEmpty() || maxValue <= 0.0) {
            svg.append("<text x=\"").append(svgNumber(width / 2.0)).append("\" y=\"").append(svgNumber(height / 2.0))
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

            svg.append("<text x=\"").append(svgNumber(left - 8.0)).append("\" y=\"").append(svgNumber(y + 4.0))
                .append("\" text-anchor=\"end\" font-size=\"10\" fill=\"#666\">")
                .append(escapeHtml(formatNumber(tickValue, 0) + " kr"))
                .append("</text>\n");
        }

        double averageValue = totalMarketValue / rowsWithValue.size();
        double averageY = mapValueToY(averageValue, 0.0, maxValue, top, plotHeight);
        svg.append("<line x1=\"").append(svgNumber(left)).append("\" y1=\"").append(svgNumber(averageY))
            .append("\" x2=\"").append(svgNumber(left + plotWidth)).append("\" y2=\"").append(svgNumber(averageY))
            .append("\" stroke=\"#495057\" stroke-width=\"1.2\" stroke-dasharray=\"5 4\"/>\n");
        svg.append("<text x=\"").append(svgNumber(left + plotWidth + 8.0)).append("\" y=\"").append(svgNumber(averageY + 3.0))
            .append("\" text-anchor=\"start\" font-size=\"10\" fill=\"#495057\">")
            .append(escapeHtml("Avg: " + formatNumber(averageValue, 0) + " kr"))
            .append("</text>\n");

        double slotWidth = plotWidth / rowsWithValue.size();
        double barWidth = Math.max(10.0, slotWidth * 0.92);
        for (int i = 0; i < rowsWithValue.size(); i++) {
            OverviewRow row = rowsWithValue.get(i);
            double x = left + (i * slotWidth) + ((slotWidth - barWidth) / 2.0);
            double y = mapValueToY(row.marketValue, 0.0, maxValue, top, plotHeight);
            double barHeight = (top + plotHeight) - y;

            String label = getOverviewRowLabel(row);
            String color = getAllocationColor(i);

            svg.append("<rect x=\"").append(svgNumber(x)).append("\" y=\"").append(svgNumber(y))
                .append("\" width=\"").append(svgNumber(barWidth)).append("\" height=\"").append(svgNumber(barHeight))
                .append("\" fill=\"").append(color).append("\" rx=\"2\">\n")
                .append("<title>")
                .append(escapeHtml(label + ": " + formatNumber(row.marketValue, 2) + " kr"))
                .append("</title></rect>\n");

            double labelAnchorX = x + (barWidth / 2.0);
            double labelAnchorY = height - bottom + 28.0;
            svg.append("<text x=\"").append(svgNumber(labelAnchorX)).append("\" y=\"").append(svgNumber(labelAnchorY))
                .append("\" transform=\"rotate(-35 ").append(svgNumber(labelAnchorX)).append(" ").append(svgNumber(labelAnchorY))
                .append(")\" text-anchor=\"end\" font-size=\"9\" fill=\"#444\">")
                .append(escapeHtml(label))
                .append("</text>\n");
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

        private static String buildAssetTypeAllocationSvg(ArrayList<OverviewRow> rows) {
        final double width = 360.0;
        final double height = 330.0;
        final double centerX = width / 2.0;
        final double centerY = 142.0;
        final double radius = 86.0;
        final double innerRadius = 44.0;

        double stockValue = 0.0;
        double fundValue = 0.0;
        double otherValue = 0.0;

        for (OverviewRow row : rows) {
            if (row.marketValue <= 0.0) {
                continue;
            }

            switch (row.assetType) {
                case "STOCK" -> stockValue += row.marketValue;
                case "FUND" -> fundValue += row.marketValue;
                default -> otherValue += row.marketValue;
            }
        }

        double totalValue = stockValue + fundValue + otherValue;

        StringBuilder svg = new StringBuilder();
        svg.append("<svg class=\"chart-svg\" viewBox=\"0 0 ")
            .append(svgNumber(width))
            .append(" ")
            .append(svgNumber(height))
            .append("\" xmlns=\"http://www.w3.org/2000/svg\" role=\"img\">\n");

        if (totalValue <= 0.0) {
            svg.append("<text x=\"").append(svgNumber(centerX)).append("\" y=\"").append(svgNumber(centerY))
                .append("\" text-anchor=\"middle\" font-size=\"13\" fill=\"#666\">No asset type data</text>\n");
            svg.append("</svg>\n");
            return svg.toString();
        }

        ArrayList<String> labels = new ArrayList<>();
        ArrayList<Double> values = new ArrayList<>();
        ArrayList<String> colors = new ArrayList<>();

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
        if (otherValue > 0.0) {
            labels.add("Other");
            values.add(otherValue);
            colors.add("#868e96");
        }

        if (values.size() == 1) {
            svg.append("<circle cx=\"").append(svgNumber(centerX)).append("\" cy=\"").append(svgNumber(centerY))
                .append("\" r=\"").append(svgNumber(radius)).append("\" fill=\"").append(colors.get(0)).append("\">\n")
                .append("<title>")
                .append(escapeHtml(labels.get(0) + ": " + formatNumber(values.get(0), 2) + " kr (100.0%)"))
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

                svg.append("<path d=\"M ").append(svgNumber(centerX)).append(" ").append(svgNumber(centerY))
                    .append(" L ").append(svgNumber(x1)).append(" ").append(svgNumber(y1))
                    .append(" A ").append(svgNumber(radius)).append(" ").append(svgNumber(radius)).append(" 0 ")
                    .append(largeArcFlag).append(" 1 ").append(svgNumber(x2)).append(" ").append(svgNumber(y2))
                    .append(" Z\" fill=\"").append(color).append("\">\n")
                    .append("<title>")
                    .append(escapeHtml(labels.get(i)
                        + ": " + formatNumber(value, 2)
                        + " kr (" + formatNumber(fraction * 100.0, 1) + "%)"))
                    .append("</title></path>\n");

                currentAngle = endAngle;
            }
        }

        svg.append("<circle cx=\"").append(svgNumber(centerX)).append("\" cy=\"").append(svgNumber(centerY))
            .append("\" r=\"").append(svgNumber(innerRadius)).append("\" fill=\"#fff\"/>\n");

        svg.append("<text x=\"").append(svgNumber(centerX)).append("\" y=\"").append(svgNumber(centerY - 4.0))
            .append("\" text-anchor=\"middle\" font-size=\"11\" fill=\"#666\">Asset Mix</text>\n");
        svg.append("<text x=\"").append(svgNumber(centerX)).append("\" y=\"").append(svgNumber(centerY + 14.0))
            .append("\" text-anchor=\"middle\" font-size=\"12\" fill=\"#222\" font-weight=\"600\">")
            .append(escapeHtml(formatNumber(totalValue, 0) + " kr"))
            .append("</text>\n");

        double stockPct = totalValue > 0.0 ? (stockValue / totalValue) * 100.0 : 0.0;
        double fundPct = totalValue > 0.0 ? (fundValue / totalValue) * 100.0 : 0.0;

        svg.append("<circle cx=\"").append(svgNumber(centerX - 60.0)).append("\" cy=\"258.00\" r=\"4.00\" fill=\"#1c7ed6\"/>\n");
        svg.append("<text x=\"").append(svgNumber(centerX - 50.0)).append("\" y=\"261.00\" font-size=\"10\" fill=\"#444\">")
            .append(escapeHtml("Stocks: " + formatNumber(stockPct, 1) + "%"))
            .append("</text>\n");

        svg.append("<circle cx=\"").append(svgNumber(centerX - 60.0)).append("\" cy=\"276.00\" r=\"4.00\" fill=\"#2f9e44\"/>\n");
        svg.append("<text x=\"").append(svgNumber(centerX - 50.0)).append("\" y=\"279.00\" font-size=\"10\" fill=\"#444\">")
            .append(escapeHtml("Funds: " + formatNumber(fundPct, 1) + "%"))
            .append("</text>\n");

        svg.append("</svg>\n");
        return svg.toString();
        }

        private static String getAllocationColor(int index) {
        String[] palette = new String[] {
            "#0b7285", "#2f9e44", "#f08c00", "#7048e8", "#c92a2a", "#1c7ed6", "#5f3dc4", "#2b8a3e", "#e67700", "#0ca678"
        };
        return palette[index % palette.length];
        }

    private static double mapValueToY(double value, double minValue, double maxValue, double chartTop, double chartHeight) {
        if (Math.abs(maxValue - minValue) < 1e-12) {
            return chartTop + (chartHeight / 2.0);
        }
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
        return formatNumber(value, decimals) + " kr";
    }

    private static void writeOverviewTableHtml(FileWriter writer) throws IOException {
        writer.write("<h2>PORTFOLIO OVERVIEW - CURRENT HOLDINGS</h2>\n");
        ArrayList<OverviewRow> overviewRows = new ArrayList<>();
        for (Security security : getSortedSecuritiesForOverview()) {
            overviewRows.add(buildOverviewRow(security));
        }

        writeOverviewChartsHtml(writer, overviewRows);

        writer.write("<table>\n");
        writeHtmlRow(writer, true,
                "Ticker",
                "Security",
                "Type",
                "Units",
                "Avg Cost",
                "Last Price",
                "Market Value",
                "Cost Basis",
                "Unrealized Return",
                "Unrealized Return (%)",
                "Realized Return (%)",
                "Realized Return",
                "Dividends",
                "Total Return",
                "Total Return (%)"
        );

        double totalCostBasis = 0.0;
        double totalMarketValue = 0.0;
        double totalUnrealized = 0.0;
        double totalCostBasisWithPrice = 0.0;
        double totalRealized = 0.0;
        double totalDividends = 0.0;
        for (OverviewRow row : overviewRows) {
            totalCostBasis += row.positionCostBasis;
            totalMarketValue += row.marketValue;
            if (row.hasPrice) {
                totalUnrealized += row.unrealized;
                totalCostBasisWithPrice += row.positionCostBasis;
            }
            totalRealized += row.realized;
            totalDividends += row.dividends;

            writeHtmlRow(
                writer,
                false,
                row.tickerText,
                row.securityDisplayName,
                row.assetType,
                formatUnits(row.units),
                formatNumber(row.averageCost, 2),
                row.latestPriceText,
                row.latestPrice > 0.0 ? formatNumber(row.marketValue, 2) : "-",
                formatNumber(row.positionCostBasis, 2),
                row.hasPrice ? formatNumber(row.unrealized, 2) : "-",
                row.hasPrice ? formatNumber(row.unrealizedPct, 2) : "-",
                row.realizedReturnPctText,
                row.realizedGainText,
                row.dividendsText,
                formatNumber(row.totalReturn, 2),
                formatNumber(row.totalReturnPct, 2)
            );
        }

        double totalReturn = totalUnrealized + totalRealized + totalDividends;
        double totalUnrealizedPct = totalCostBasisWithPrice > 0 ? (totalUnrealized / totalCostBasisWithPrice) * 100.0 : 0.0;
        double totalRealizedPct = totalCostBasis > 0 ? (totalRealized / totalCostBasis) * 100.0 : 0.0;
        double totalReturnPct = totalCostBasis > 0 ? (totalReturn / totalCostBasis) * 100.0 : 0.0;
        writer.write("<tr class=\"total-row\">"
            + toDataCell("")
            + toDataCell("TOTAL")
            + toDataCell("")
            + toDataCell("")
            + toDataCell("")
            + toDataCell("")
            + toDataCell(formatNumber(totalMarketValue, 2))
            + toDataCell(formatNumber(totalCostBasis, 2))
            + toDataCell(formatNumber(totalUnrealized, 2))
            + toDataCell(formatNumber(totalUnrealizedPct, 2))
            + toDataCell(formatNumber(totalRealizedPct, 2))
            + toDataCell(formatNumber(totalRealized, 2))
            + toDataCell(formatNumber(totalDividends, 2))
            + toDataCell(formatNumber(totalReturn, 2))
            + toDataCell(formatNumber(totalReturnPct, 2))
            + "</tr>\n"
        );

        writer.write("</table>\n\n");
    }

    private static void writeRealizedSummaryTableHtml(FileWriter writer) throws IOException {
        writer.write("<h2>REALIZED OVERVIEW - ALL SALES</h2>\n");
        writer.write("<table>\n");
        writeHtmlRow(writer, true, "Security", "Sales Value", "Cost Basis", "Realized Gain/Loss", "Dividends", "Return (%)");

        ArrayList<Security> soldSecurities = getSortedSoldSecurities();

        double totalSalesValue = 0.0;
        double totalCostBasis = 0.0;
        double totalRealizedGain = 0.0;
        double totalRealizedDividends = 0.0;

        for (Security security : soldSecurities) {
            double salesValue = security.getRealizedSalesValue();
            double costBasis = security.getRealizedCostBasis();
            double gain = security.getRealizedGain();
            double realizedDividends = security.isFullyRealized() ? security.getDividends() : 0.0;
            double returnPct = costBasis > 0 ? (gain / costBasis) * 100.0 : (gain > 0 ? 100.0 : 0.0);

            totalSalesValue += salesValue;
            totalCostBasis += costBasis;
            totalRealizedGain += gain;
            totalRealizedDividends += realizedDividends;

            writeHtmlRow(
                writer,
                false,
                security.getName(),
                formatNumber(salesValue, 2),
                formatNumber(costBasis, 2),
                formatNumber(gain, 2),
                formatNumber(realizedDividends, 2),
                formatNumber(returnPct, 2)
            );
        }

        double totalReturnPct = totalCostBasis > 0 ? (totalRealizedGain / totalCostBasis) * 100.0 : (totalRealizedGain > 0 ? 100.0 : 0.0);
        writer.write("<tr class=\"total-row\">"
            + toDataCell("TOTAL")
            + toDataCell(formatNumber(totalSalesValue, 2))
            + toDataCell(formatNumber(totalCostBasis, 2))
            + toDataCell(formatNumber(totalRealizedGain, 2))
            + toDataCell(formatNumber(totalRealizedDividends, 2))
            + toDataCell(formatNumber(totalReturnPct, 2))
            + "</tr>\n"
        );

        writer.write("</table>\n\n");
    }

    private static void writeSaleTradesTablesHtml(FileWriter writer) throws IOException {
        ArrayList<Security> soldSecurities = getSortedSoldSecurities();

        for (Security security : soldSecurities) {
            writer.write("<h2>SALE TRADES - " + escapeHtml(security.getName()) + "</h2>\n");
            writer.write("<table>\n");
            writeHtmlRow(writer, true, "Sale Date", "Units", "Price/Unit", "Sale Value", "Cost Basis", "Gain/Loss", "Return (%)");

            double totalUnits = 0.0;
            double totalSalesValue = 0.0;
            double totalCostBasis = 0.0;
            double totalGain = 0.0;

            for (Security.SaleTrade trade : security.getSaleTradesSortedByDate()) {
                totalUnits += trade.getUnits();
                totalSalesValue += trade.getSaleValue();
                totalCostBasis += trade.getCostBasis();
                totalGain += trade.getGainLoss();

                writeHtmlRow(
                    writer,
                    false,
                    trade.getTradeDateAsCsv(),
                    formatUnits(trade.getUnits()),
                    formatNumber(trade.getUnitPrice(), 2),
                    formatNumber(trade.getSaleValue(), 2),
                    formatNumber(trade.getCostBasis(), 2),
                    formatNumber(trade.getGainLoss(), 2),
                    formatNumber(trade.getReturnPct(), 2)
                );
            }

            double totalReturnPct = totalCostBasis > 0 ? (totalGain / totalCostBasis) * 100.0 : (totalGain > 0 ? 100.0 : 0.0);
                writer.write("<tr class=\"total-row\">"
                    + toDataCell("TOTAL")
                    + toDataCell(formatUnits(totalUnits))
                    + toDataCell("")
                    + toDataCell(formatNumber(totalSalesValue, 2))
                    + toDataCell(formatNumber(totalCostBasis, 2))
                    + toDataCell(formatNumber(totalGain, 2))
                    + toDataCell(formatNumber(totalReturnPct, 2))
                    + "</tr>\n"
                );

            writer.write("</table>\n\n");
        }
    }

    private static String formatNumber(double value, int decimals) {
        String pattern = "%,." + decimals + "f";
        return String.format(Locale.US, pattern, value).replace(',', ' ');
    }

    private static String formatUnits(double value) {
        String text = formatNumber(value, 4);
        int decimalPoint = text.indexOf('.');
        if (decimalPoint < 0) {
            return text;
        }

        int end = text.length();
        while (end > decimalPoint + 1 && text.charAt(end - 1) == '0') {
            end--;
        }
        if (end == decimalPoint + 1) {
            end = decimalPoint;
        }

        return text.substring(0, end);
    }
}
