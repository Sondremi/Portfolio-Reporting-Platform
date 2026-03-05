import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.Locale;
import java.util.Map;

public class PortfolioTracker {
    private static final String CSV_SEPARATOR = ";";
    private static final String INPUT_DIRECTORY = "transaction_files";
    private static final String OUTPUT_FILE = "portfolio.csv";

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

        writeToCsv();
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
                        && !name.equalsIgnoreCase(OUTPUT_FILE));

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
            securityName = securityType = isin = transactionType = amount = quantity = price = result = totalFees = tradeDate = -1;
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

    private static void processLine(String line, char delimiter, HeaderIndexes indexes) {
        if (line == null || line.isBlank()) {
            return;
        }

        ArrayList<String> row = parseDelimitedLine(line.replace("−", "-"), delimiter);

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

    private static int getAssetPriority(Security security) {
        return switch (security.getAssetType()) {
            case STOCK -> 0;
            case FUND -> 1;
            case UNKNOWN -> 0;
        };
    }

    private static ArrayList<Security> getSortedSecuritiesForOverview() {
        ArrayList<Security> sorted = new ArrayList<>(securities);
        sorted.sort(
                Comparator.comparingInt(PortfolioTracker::getAssetPriority)
                        .thenComparing(Security::getName, String.CASE_INSENSITIVE_ORDER)
        );
        return sorted;
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

            for (int i = 1; i < lines.length; i++) {
                processLine(lines[i], delimiter, indexes);
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

    private static void writeToCsv() {
        File file = new File(OUTPUT_FILE);
        try (FileWriter writer = new FileWriter(file)) {
            writeOverviewAsCsv(writer);
            writer.write("\n\n");
            writeRealizedSummaryAsCsv(writer);
            writer.write("\n\n");
            writeSaleTradesPerSecurityAsCsv(writer);
            new ProcessBuilder("open", OUTPUT_FILE).start();
        } catch (IOException e) {
            System.out.println("Failed to write CSV: " + e.getMessage());
        }
    }

    private static void writeSectionTitle(FileWriter writer, String title) throws IOException {
        writer.write(title);
        writer.write("\n");
    }

    private static String escapeCsvCell(String value) {
        String text = value == null ? "" : value;
        return "\"" + text.replace("\"", "\"\"") + "\"";
    }

    private static void writeCsvRow(FileWriter writer, String... fields) throws IOException {
        for (int i = 0; i < fields.length; i++) {
            if (i > 0) {
                writer.write(CSV_SEPARATOR);
            }
            writer.write(escapeCsvCell(fields[i]));
        }
        writer.write("\n");
    }

    private static void writeOverviewAsCsv(FileWriter writer) throws IOException {
        writeCsvRow(
            writer,
            "Ticker",
            "Verdipapir",
            "Antall",
            "GAV",
            "Kurs",
            "Kostpris",
            "Markedsverdi",
            "Urealisert Avkastning (%)",
            "Urealisert Avkastning",
            "Realisert Avkastning (%)",
            "Realisert Avkastning",
            "Utbytte",
            "Avkastning (%)",
            "Avkastning"
        );

        int row = 2;
        int startRow = row;

        for (Security security : getSortedSecuritiesForOverview()) {
            String ticker = security.getTicker();
            String name = security.getName();

            String tickerCell = ticker.isEmpty() ? name : "=HVISFEIL(AKSJE(\"" + ticker + "\";25);\"-\")";
            String nameCell = ticker.isEmpty() ? name : "=HVISFEIL(AKSJE(A" + row + ";1);\"" + name + "\")";

            String marketPriceFormula = "=HVISFEIL(AKSJE(A" + row + ";0);0)";

            String costBasisFormula = "VERDI(C" + row + ")*VERDI(D" + row + ")";
            String marketValueFormula = "VERDI(C" + row + ")*VERDI(E" + row + ")";

            String unrealizedGainFormula = "=" + marketValueFormula + "-" + costBasisFormula;
            String unrealizedGainPctFormula = "=AVRUND(HVISFEIL(((" + marketValueFormula + ")-(" + costBasisFormula + "))/(" + costBasisFormula + "); 0); 2)";

            String totalGainFormula = "I" + row + "+K" + row + "+L" + row;
            String totalGainPctFormula = "=AVRUND(HVISFEIL((" + totalGainFormula + ")/F" + row + "; 0); 2)";

            writeCsvRow(
                writer,
                tickerCell,
                nameCell,
                security.getUnitsOwnedAsText(),
                security.getAverageCostAsText(),
                marketPriceFormula,
                "=" + costBasisFormula,
                "=" + marketValueFormula,
                unrealizedGainPctFormula,
                unrealizedGainFormula,
                security.getRealizedReturnPctAsText(),
                security.getRealizedGainAsText(),
                security.getDividendsAsText(),
                totalGainPctFormula,
                "=" + totalGainFormula
            );

            row++;
        }

        writeCsvRow(
            writer,
            "",
            "",
            "",
            "",
            "",
            "=SUMMER(F" + startRow + ":F" + (row - 1) + ")",
            "=SUMMER(G" + startRow + ":G" + (row - 1) + ")",
            "=SUMMER(H" + startRow + ":H" + (row - 1) + ")",
            "=SUMMER(I" + startRow + ":I" + (row - 1) + ")",
            "=SUMMER(J" + startRow + ":J" + (row - 1) + ")",
            "=SUMMER(K" + startRow + ":K" + (row - 1) + ")",
            "=SUMMER(L" + startRow + ":L" + (row - 1) + ")",
            "=SUMMER(M" + startRow + ":M" + (row - 1) + ")",
            "=SUMMER(N" + startRow + ":N" + (row - 1) + ")"
        );
    }

    private static void writeRealizedSummaryAsCsv(FileWriter writer) throws IOException {
        writeSectionTitle(writer, "TOTALOVERSIKT - ALLE SALG");
        writeCsvRow(writer, "Verdipapir", "Salgssum", "Kostnad", "Realisert gevinst/tap", "Avkastning (%)");

        ArrayList<Security> soldSecurities = getSortedSoldSecurities();

        double totalSalesValue = 0.0;
        double totalCostBasis = 0.0;
        double totalRealizedGain = 0.0;

        for (Security security : soldSecurities) {
            double salesValue = security.getRealizedSalesValue();
            double costBasis = security.getRealizedCostBasis();
            double gain = security.getRealizedGain();
            double returnPct = costBasis > 0 ? (gain / costBasis) * 100.0 : 0.0;

            totalSalesValue += salesValue;
            totalCostBasis += costBasis;
            totalRealizedGain += gain;

            writeCsvRow(
                writer,
                security.getName(),
                formatNumber(salesValue, 2),
                formatNumber(costBasis, 2),
                formatNumber(gain, 2),
                formatNumber(returnPct, 2)
            );
        }

        double totalReturnPct = totalCostBasis > 0 ? (totalRealizedGain / totalCostBasis) * 100.0 : 0.0;
        writeCsvRow(
            writer,
            "TOTALT",
            formatNumber(totalSalesValue, 2),
            formatNumber(totalCostBasis, 2),
            formatNumber(totalRealizedGain, 2),
            formatNumber(totalReturnPct, 2)
        );
    }

    private static void writeSaleTradesPerSecurityAsCsv(FileWriter writer) throws IOException {
        ArrayList<Security> soldSecurities = getSortedSoldSecurities();

        for (Security security : soldSecurities) {
            writeSectionTitle(writer, "SALGSTRADES - " + security.getName());
            writeCsvRow(writer, "Salgsdato", "Andeler", "Pris/andel", "Salgssum", "Kostnad", "Gevinst/Tap", "Avkastning (%)");

            double totalUnits = 0.0;
            double totalSalesValue = 0.0;
            double totalCostBasis = 0.0;
            double totalGain = 0.0;

            for (Security.SaleTrade trade : security.getSaleTradesSortedByDate()) {
                totalUnits += trade.getUnits();
                totalSalesValue += trade.getSaleValue();
                totalCostBasis += trade.getCostBasis();
                totalGain += trade.getGainLoss();

                writeCsvRow(
                    writer,
                    trade.getTradeDateAsCsv(),
                    formatNumber(trade.getUnits(), 4),
                    formatNumber(trade.getUnitPrice(), 2),
                    formatNumber(trade.getSaleValue(), 2),
                    formatNumber(trade.getCostBasis(), 2),
                    formatNumber(trade.getGainLoss(), 2),
                    formatNumber(trade.getReturnPct(), 2)
                );
            }

            double totalReturnPct = totalCostBasis > 0 ? (totalGain / totalCostBasis) * 100.0 : 0.0;
            writeCsvRow(
                writer,
                "TOTALT",
                formatNumber(totalUnits, 4),
                "",
                formatNumber(totalSalesValue, 2),
                formatNumber(totalCostBasis, 2),
                formatNumber(totalGain, 2),
                formatNumber(totalReturnPct, 2)
            );

            writer.write("\n\n");
        }
    }

    private static String formatNumber(double value, int decimals) {
        String pattern = "%." + decimals + "f";
        return String.format(Locale.US, pattern, value);
    }
}
