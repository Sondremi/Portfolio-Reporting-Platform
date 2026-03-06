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

    private static void writeOverviewTableHtml(FileWriter writer) throws IOException {
        writer.write("<h2>PORTFOLIO OVERVIEW - CURRENT HOLDINGS</h2>\n");
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
        for (Security security : getSortedSecuritiesForOverview()) {
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

            totalCostBasis += positionCostBasis;
            totalMarketValue += marketValue;
            if (hasPrice) {
                totalUnrealized += unrealized;
                totalCostBasisWithPrice += positionCostBasis;
            }
            totalRealized += realized;
            totalDividends += dividends;

            writeHtmlRow(
                writer,
                false,
                tickerText,
                security.getDisplayName(),
                security.getAssetType().name(),
                formatUnits(units),
                formatNumber(averageCost, 2),
                security.getLatestPriceAsText(),
                latestPrice > 0.0 ? formatNumber(marketValue, 2) : "-",
                formatNumber(positionCostBasis, 2),
                hasPrice ? formatNumber(unrealized, 2) : "-",
                hasPrice ? formatNumber(unrealizedPct, 2) : "-",
                security.getRealizedReturnPctAsText(),
                security.getRealizedGainAsText(),
                security.getDividendsAsText(),
                formatNumber(totalReturn, 2),
                formatNumber(totalReturnPct, 2)
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
