import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.BufferedReader;
import java.net.HttpURLConnection;
import java.net.URI;
import java.net.URL;
import java.net.URLEncoder;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.time.Instant;
import java.time.LocalDate;
import java.time.YearMonth;
import java.time.ZoneOffset;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Locale;
import java.util.Map;
import java.util.NavigableMap;
import java.util.TreeMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class PortfolioReportGenerator {
    private static final String INPUT_DIRECTORY = "transaction_files";
    private static final String OUTPUT_FILE = "portfolio-report.html";

    // User-editable mapping for renamed/merged instruments: old ISIN -> new ISIN.
    // Add entries here when the broker reports a replacement as separate ISINs,
    // so history and current holdings are treated as one continuous position.
    private static final Map<String, String> RENAMED_SECURITY_ISIN = Map.of(
            "NO0010782519", "NO0012948878"
    );

    private static final ArrayList<Security> securities = new ArrayList<>();
    private static final Map<String, Security> securitiesByKey = new LinkedHashMap<>();
    private static final Map<String, String> canonicalSecurityNameByIsin = new LinkedHashMap<>();
    private static final ArrayList<UnitEvent> unitEvents = new ArrayList<>();
    private static final ArrayList<CashEvent> cashEvents = new ArrayList<>();
    private static final ArrayList<PortfolioCashSnapshot> portfolioCashSnapshots = new ArrayList<>();
    private static int loadedCsvFileCount = 0;
    private static int loadedTransactionRowCount = 0;

    private static final Pattern YAHOO_TIMESTAMP_ARRAY = Pattern.compile("\\\"timestamp\\\"\\s*:\\s*\\[(.*?)\\]", Pattern.DOTALL);
    private static final Pattern YAHOO_CLOSE_ARRAY = Pattern.compile("\\\"close\\\"\\s*:\\s*\\[(.*?)\\]", Pattern.DOTALL);
    public static void main(String[] args) throws IOException {
        ensureInputDirectoryExists();
        loadedCsvFileCount = 0;
        loadedTransactionRowCount = 0;
        unitEvents.clear();
        cashEvents.clear();
        portfolioCashSnapshots.clear();
        int filesProcessed = readAllTransactionFiles();
        loadedCsvFileCount = filesProcessed;

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

    private static String canonicalizeIsin(String isin) {
        if (isin == null || isin.isBlank()) {
            return isin;
        }

        String normalized = isin.trim().toUpperCase(Locale.ROOT);
        return RENAMED_SECURITY_ISIN.getOrDefault(normalized, normalized);
    }

    private static boolean isRenamedSecurityIsin(String isin) {
        if (isin == null || isin.isBlank()) {
            return false;
        }

        String normalized = isin.trim().toUpperCase(Locale.ROOT);
        return RENAMED_SECURITY_ISIN.containsKey(normalized)
                || RENAMED_SECURITY_ISIN.containsValue(normalized);
    }

    private static boolean isRenameBookkeepingTransaction(String transactionType, String securityIsin) {
        if (!isRenamedSecurityIsin(securityIsin) || transactionType == null || transactionType.isBlank()) {
            return false;
        }

        String normalized = transactionType.toUpperCase(Locale.ROOT).replaceAll("\\s+", " ").trim();
        return "BYTTE INNLEGG VP".equals(normalized)
                || "BYTTE UTTAK VP".equals(normalized)
                || "MAK BYTTE INNLEGG VP".equals(normalized)
                || "MAK BYTTE UTTAK VP".equals(normalized);
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
        int cancellationDate;
        int portfolioId;
        int cashBalance;

        HeaderIndexes() {
            transactionId = securityName = securityType = isin = transactionType = amount = quantity = price = result = totalFees = tradeDate = cancellationDate = portfolioId = cashBalance = -1;
        }

        boolean hasRequiredColumns() {
            return securityName >= 0 && transactionType >= 0;
        }
    }

    private static final class UnitEvent {
        private final LocalDate tradeDate;
        private final String securityKey;
        private final double unitsDelta;

        private UnitEvent(LocalDate tradeDate, String securityKey, double unitsDelta) {
            this.tradeDate = tradeDate;
            this.securityKey = securityKey;
            this.unitsDelta = unitsDelta;
        }
    }

    private static final class CashEvent {
        private final LocalDate tradeDate;
        private final double cashDelta;

        private CashEvent(LocalDate tradeDate, double cashDelta) {
            this.tradeDate = tradeDate;
            this.cashDelta = cashDelta;
        }
    }

    private static final class PortfolioCashSnapshot {
        private final LocalDate tradeDate;
        private final long sortId;
        private final String portfolioId;
        private final double balance;

        private PortfolioCashSnapshot(LocalDate tradeDate, long sortId, String portfolioId, double balance) {
            this.tradeDate = tradeDate;
            this.sortId = sortId;
            this.portfolioId = portfolioId;
            this.balance = balance;
        }
    }

    private static final class PortfolioValuePoint {
        private final LocalDate monthEnd;
        private final double value;

        private PortfolioValuePoint(LocalDate monthEnd, double value) {
            this.monthEnd = monthEnd;
            this.value = value;
        }
    }

    private static final class OverviewRow {
        private final String tickerText;
        private final String securityDisplayName;
        private final String assetType;
        private final String sectorLabel;
        private final Map<String, Double> sectorWeights;
        private final String regionLabel;
        private final Map<String, Double> regionWeights;
        private final String currencyCode;
        private final String realizedReturnPctText;
        private final double units;
        private final double averageCost;
        private final double latestPrice;
        private final double positionCostBasis;
        private final double realizedCostBasis;
        private final double historicalCostBasis;
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
                String sectorLabel,
                Map<String, Double> sectorWeights,
                String regionLabel,
                Map<String, Double> regionWeights,
                String currencyCode,
                String realizedReturnPctText,
                double units,
                double averageCost,
                double latestPrice,
                double positionCostBasis,
                double realizedCostBasis,
                double historicalCostBasis,
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
            this.sectorLabel = sectorLabel;
            this.sectorWeights = sectorWeights == null ? Map.of() : new LinkedHashMap<>(sectorWeights);
            this.regionLabel = regionLabel;
            this.regionWeights = regionWeights == null ? Map.of() : new LinkedHashMap<>(regionWeights);
            this.currencyCode = currencyCode;
            this.realizedReturnPctText = realizedReturnPctText;
            this.units = units;
            this.averageCost = averageCost;
            this.latestPrice = latestPrice;
            this.positionCostBasis = positionCostBasis;
            this.realizedCostBasis = realizedCostBasis;
            this.historicalCostBasis = historicalCostBasis;
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

    private static final class HeaderSummary {
        private final String generatedDate;
        private final int fileCount;
        private final int transactionCount;
        private final int holdingsCount;
        private final String totalCurrencyCode;
        private final double totalMarketValue;
        private final double cashHoldings;
        private final double totalReturn;
        private final double totalReturnPct;
        private final String bestLabel;
        private final double bestReturn;
        private final double bestReturnPct;
        private final String worstLabel;
        private final double worstReturn;
        private final double worstReturnPct;
        private final String sparklineSvg;

        private HeaderSummary(
                String generatedDate,
                int fileCount,
                int transactionCount,
                int holdingsCount,
                String totalCurrencyCode,
                double totalMarketValue,
                double cashHoldings,
                double totalReturn,
                double totalReturnPct,
                String bestLabel,
                double bestReturn,
                double bestReturnPct,
                String worstLabel,
                double worstReturn,
                double worstReturnPct,
                String sparklineSvg) {
            this.generatedDate = generatedDate;
            this.fileCount = fileCount;
            this.transactionCount = transactionCount;
            this.holdingsCount = holdingsCount;
            this.totalCurrencyCode = totalCurrencyCode;
            this.totalMarketValue = totalMarketValue;
            this.cashHoldings = cashHoldings;
            this.totalReturn = totalReturn;
            this.totalReturnPct = totalReturnPct;
            this.bestLabel = bestLabel;
            this.bestReturn = bestReturn;
            this.bestReturnPct = bestReturnPct;
            this.worstLabel = worstLabel;
            this.worstReturn = worstReturn;
            this.worstReturnPct = worstReturnPct;
            this.sparklineSvg = sparklineSvg;
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
                case "makuleringsdato" -> indexes.cancellationDate = i;
                case "portefølje", "portefolje", "konto" -> indexes.portfolioId = i;
                case "saldo" -> indexes.cashBalance = i;
                default -> {
                    // Ignored.
                }
            }

            if (indexes.portfolioId < 0 && (column.contains("portef") || "konto".equals(column))) {
                indexes.portfolioId = i;
            }

            if (indexes.cashBalance < 0 && column.contains("saldo")) {
                indexes.cashBalance = i;
            }
        }
        return indexes;
    }

    private static void processRow(ArrayList<String> row, HeaderIndexes indexes) {
        if (row == null || row.isEmpty()) {
            return;
        }

        String transactionType = getCell(row, indexes.transactionType)
                .toUpperCase(Locale.ROOT)
                .replaceAll("\\s+", " ")
                .trim();

        LocalDate tradeDateForTracking = indexes.tradeDate >= 0
            ? parseTradeDateForSort(getCell(row, indexes.tradeDate))
            : LocalDate.MIN;
        long sortId = parseSortIdOrDefault(getCell(row, indexes.transactionId));
        boolean balanceBackedRow = recordPortfolioCashSnapshot(row, indexes, tradeDateForTracking, sortId);

        String securityName = getCell(row, indexes.securityName);
        if (securityName.isEmpty()) {
            processStandaloneCashTransaction(row, indexes, transactionType, tradeDateForTracking, balanceBackedRow);
            return;
        }

        String isin = getCell(row, indexes.isin);
        String canonicalIsin = canonicalizeIsin(isin);
        rememberCanonicalSecurityName(isin, canonicalIsin, securityName);
        String securityType = getCell(row, indexes.securityType);
        Security security = getOrCreateSecurity(securityName, canonicalIsin);
        if (security == null) {
            return;
        }

        security.updateAssetTypeFromHint(securityType, securityName);

        processTransaction(security, row, indexes, isin, transactionType, tradeDateForTracking, balanceBackedRow);
    }

    private static boolean recordPortfolioCashSnapshot(ArrayList<String> row, HeaderIndexes indexes,
                                                       LocalDate tradeDate, long sortId) {
        if (indexes.portfolioId < 0 || indexes.cashBalance < 0) {
            return false;
        }

        String portfolioId = getCell(row, indexes.portfolioId);
        String balanceText = getCell(row, indexes.cashBalance);
        if (portfolioId.isBlank() || balanceText.isBlank()) {
            return false;
        }

        double balance = parseDoubleOrZero(balanceText);
        portfolioCashSnapshots.add(new PortfolioCashSnapshot(
                tradeDate == null ? LocalDate.MIN : tradeDate,
                sortId,
                portfolioId,
                balance
        ));
        return true;
    }

    private static void processStandaloneCashTransaction(ArrayList<String> row, HeaderIndexes indexes,
                                                         String transactionType, LocalDate tradeDateForTracking,
                                                         boolean balanceBackedRow) {
        if (balanceBackedRow) {
            return;
        }

        if (!isCashAccountEventType(transactionType)) {
            return;
        }

        double amount = parseDoubleOrZero(getCell(row, indexes.amount));
        double totalFees = parseDoubleOrZero(getCell(row, indexes.totalFees));

        double cashDelta = resolveCashAccountEventDelta(transactionType, amount, totalFees);
        if (Math.abs(cashDelta) > 1e-9) {
            recordCashEvent(tradeDateForTracking, cashDelta);
        }
    }

    private static void rememberCanonicalSecurityName(String originalIsin, String canonicalIsin, String securityName) {
        if (securityName == null || securityName.isBlank() || canonicalIsin == null || canonicalIsin.isBlank()) {
            return;
        }

        String normalizedCanonicalIsin = canonicalIsin.trim().toUpperCase(Locale.ROOT);
        String existingName = canonicalSecurityNameByIsin.get(normalizedCanonicalIsin);
        boolean fromCanonicalIsin = originalIsin != null
                && !originalIsin.isBlank()
                && normalizedCanonicalIsin.equals(originalIsin.trim().toUpperCase(Locale.ROOT));

        if (fromCanonicalIsin || existingName == null || existingName.isBlank()) {
            canonicalSecurityNameByIsin.put(normalizedCanonicalIsin, securityName);
        }
    }

    private static String getPreferredSecurityName(Security security) {
        if (security == null) {
            return "-";
        }

        // For stocks, prefer resolved provider/company name when available.
        if (security.getAssetType() == Security.AssetType.STOCK) {
            String resolvedStockName = security.getDisplayName();
            if (resolvedStockName != null && !resolvedStockName.isBlank()) {
                return resolvedStockName;
            }

            String csvName = security.getName();
            return (csvName == null || csvName.isBlank()) ? "-" : csvName;
        }

        String isin = security.getIsin();
        if (isin != null && !isin.isBlank()) {
            String preferred = canonicalSecurityNameByIsin.get(isin.trim().toUpperCase(Locale.ROOT));
            if (preferred != null && !preferred.isBlank()) {
                return preferred;
            }
        }

        String displayName = security.getDisplayName();
        return (displayName == null || displayName.isBlank()) ? "-" : displayName;
    }

    private static String getTickerText(Security security) {
        if (security == null) {
            return "-";
        }

        String ticker = security.getTicker();
        return (ticker == null || ticker.isBlank()) ? "-" : ticker;
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
                Comparator.comparingInt(PortfolioReportGenerator::getAssetPriority)
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
                Comparator.comparingInt(PortfolioReportGenerator::getAssetPriority)
                        .thenComparing(Security::getRealizedSalesValue, Comparator.reverseOrder())
                        .thenComparing(Security::getName, String.CASE_INSENSITIVE_ORDER)
        );

        return soldSecurities;
    }

    private static void processTransaction(Security security, ArrayList<String> row, HeaderIndexes indexes,
                                           String originalIsin, String transactionType,
                                           LocalDate tradeDateForTracking, boolean balanceBackedRow) {

        double quantity = parseDoubleOrZero(getCell(row, indexes.quantity));
        double amount = parseDoubleOrZero(getCell(row, indexes.amount));
        double price = parseDoubleOrZero(getCell(row, indexes.price));
        double result = parseDoubleOrZero(getCell(row, indexes.result));
        double totalFees = parseDoubleOrZero(getCell(row, indexes.totalFees));
        if (isRenameBookkeepingTransaction(transactionType, originalIsin)) {
            boolean isCancelled = indexes.cancellationDate >= 0 && !getCell(row, indexes.cancellationDate).isBlank();
            if (!isCancelled && "BYTTE INNLEGG VP".equals(transactionType) && quantity > 0.0) {
                security.reconcileUnitsFromCorporateAction(quantity);
            }
            return;
        }

        String tradeDate = getCell(row, indexes.tradeDate);

        if (isTradeLikeTransaction(transactionType)) {
            double tradeCashDelta = resolveTradeCashDelta(transactionType, amount, quantity, price, totalFees);
            if (!balanceBackedRow && Math.abs(tradeCashDelta) > 1e-9) {
                recordCashEvent(tradeDateForTracking, tradeCashDelta);
            }
            if (Math.abs(quantity) > 0.0) {
                double unitsDelta = isBuyLikeTransaction(transactionType, amount)
                        ? Math.abs(quantity)
                        : -Math.abs(quantity);
                recordUnitEvent(security, tradeDateForTracking, unitsDelta);
            }
            security.addTransaction(tradeDate, transactionType, amount, quantity, price, result, totalFees);
            return;
        }

        if (isUnitsCorporateActionTransaction(transactionType)) {
            if (!isTemporaryRightsSecurity(security)) {
                if (Math.abs(quantity) > 0.0) {
                    recordUnitEvent(security, tradeDateForTracking, Math.abs(quantity));
                }
                security.addZeroCostUnits(quantity, tradeDate);
            }
            return;
        }

        if (isDividendCashTransaction(transactionType)) {
            if (!balanceBackedRow && Math.abs(amount) > 1e-9) {
                recordCashEvent(tradeDateForTracking, Math.abs(amount));
            }
            security.addDividend(Math.abs(amount));
            return;
        }

        if (isFundCostRefundTransaction(transactionType)) {
            double refundAmount = totalFees > 0.0 ? totalFees : Math.abs(amount);
            security.applyCostRefund(refundAmount);
            if (!balanceBackedRow && Math.abs(refundAmount) > 0.0) {
                recordCashEvent(tradeDateForTracking, refundAmount);
            }
            return;
        }

        if (isCashAccountEventType(transactionType)) {
            double cashDelta = resolveCashAccountEventDelta(transactionType, amount, totalFees);
            if (!balanceBackedRow && Math.abs(cashDelta) > 1e-9) {
                recordCashEvent(tradeDateForTracking, cashDelta);
            }
        }
    }

    private static boolean isBuyLikeTransaction(String transactionType, double amount) {
        String normalized = transactionType == null ? "" : transactionType.toUpperCase(Locale.ROOT);
        if (normalized.contains("KJØP") || normalized.contains("KJOP") || normalized.contains("BUY")
                || normalized.contains("REINVEST")) {
            return true;
        }
        if (normalized.contains("SALG") || normalized.contains("SELL")) {
            return false;
        }
        return amount < 0.0;
    }

    private static boolean isSellLikeTransaction(String transactionType, double amount) {
        return !isBuyLikeTransaction(transactionType, amount);
    }

    private static boolean containsAnyKeyword(String value, String... keywords) {
        if (value == null || value.isBlank() || keywords == null || keywords.length == 0) {
            return false;
        }

        for (String keyword : keywords) {
            if (keyword != null && !keyword.isBlank() && value.contains(keyword)) {
                return true;
            }
        }
        return false;
    }

    private static boolean isTradeLikeTransaction(String transactionType) {
        String normalized = transactionType == null ? "" : transactionType.toUpperCase(Locale.ROOT);
        return containsAnyKeyword(normalized,
                "KJØP", "KJOP", "BUY", "SALG", "SELL", "REINVEST");
    }

    private static boolean isUnitsCorporateActionTransaction(String transactionType) {
        String normalized = transactionType == null ? "" : transactionType.toUpperCase(Locale.ROOT);
        return containsAnyKeyword(normalized,
                "UTBYTTE INNLEGG VP", "BYTTE INNLEGG VP", "MAK BYTTE INNLEGG VP", "TILDELING INNLEGG RE");
    }

    private static boolean isDividendCashTransaction(String transactionType) {
        String normalized = transactionType == null ? "" : transactionType.toUpperCase(Locale.ROOT);
        return !normalized.contains("REINVEST") && containsAnyKeyword(normalized, "UTBYTTE", "DIVIDEND");
    }

    private static boolean isFundCostRefundTransaction(String transactionType) {
        String normalized = transactionType == null ? "" : transactionType.toUpperCase(Locale.ROOT);
        return normalized.contains("TILBAKEBET") && normalized.contains("FOND");
    }

    private static double resolveTradeCashDelta(String transactionType, double amount,
                                                double quantity, double price, double totalFees) {
        if (isBuyLikeTransaction(transactionType, amount)) {
            double buyCashOut = Math.abs(amount);
            if (buyCashOut <= 1e-9 && Math.abs(quantity) > 1e-9 && price > 0.0) {
                buyCashOut = Math.abs(quantity) * price + Math.max(totalFees, 0.0);
            }
            return -buyCashOut;
        }

        if (isSellLikeTransaction(transactionType, amount)) {
            double sellCashIn = Math.abs(amount);
            if (sellCashIn <= 1e-9 && Math.abs(quantity) > 1e-9 && price > 0.0) {
                sellCashIn = Math.max(0.0, (Math.abs(quantity) * price) - Math.max(totalFees, 0.0));
            }
            return sellCashIn;
        }

        return 0.0;
    }

    private static boolean isCashAccountEventType(String transactionType) {
        if (transactionType == null || transactionType.isBlank()) {
            return false;
        }

        String normalized = transactionType.toUpperCase(Locale.ROOT);
        return containsAnyKeyword(normalized,
                "INNSKUDD", "INNBETAL", "INNSATT", "DEPOSIT", "TRANSFER IN",
                "UTTAK", "UTBETAL", "WITHDRAW", "TRANSFER OUT",
                "PLATTFORMAVG", "GEBYR", "KURTASJE", "FEE", "KOSTNAD",
                "DEBETRENTE", "RENTE", "OVERBEL",
                "OVERFØRING", "OVERFORING", "INTERNAL FROM", "INTERNAL TO",
                "TILBAKEBETALING", "REFUND");
    }

    private static double resolveCashAccountEventDelta(String transactionType, double amount, double totalFees) {
        String normalized = transactionType == null ? "" : transactionType.toUpperCase(Locale.ROOT);

        if (containsAnyKeyword(normalized,
                "INNSKUDD", "INNBETAL", "INNSATT", "DEPOSIT", "TRANSFER IN",
                "OVERFØRING", "OVERFORING", "INTERNAL FROM", "TILBAKEBETALING", "REFUND")) {
            double baseAmount = Math.abs(amount);
            return baseAmount > 1e-9 ? baseAmount : Math.max(0.0, totalFees);
        }

        if (containsAnyKeyword(normalized,
                "UTTAK", "UTBETAL", "WITHDRAW", "TRANSFER OUT", "INTERNAL TO")) {
            double baseAmount = Math.abs(amount);
            return baseAmount > 1e-9 ? -baseAmount : -Math.max(0.0, totalFees);
        }

        if (containsAnyKeyword(normalized,
                "PLATTFORMAVG", "GEBYR", "KURTASJE", "FEE", "KOSTNAD", "DEBETRENTE", "RENTE", "OVERBEL")) {
            if (Math.abs(amount) > 1e-9) {
                return amount;
            }
            return -Math.max(0.0, totalFees);
        }

        return amount;
    }

    private static void recordUnitEvent(Security security, LocalDate tradeDate, double unitsDelta) {
        if (security == null || Math.abs(unitsDelta) < 1e-9) {
            return;
        }

        String securityKey = getTrackingSecurityKey(security);
        if (securityKey.isBlank()) {
            return;
        }

        LocalDate eventDate = tradeDate == null ? LocalDate.MIN : tradeDate;
        unitEvents.add(new UnitEvent(eventDate, securityKey, unitsDelta));
    }

    private static void recordCashEvent(LocalDate tradeDate, double cashDelta) {
        if (Math.abs(cashDelta) < 1e-9) {
            return;
        }

        LocalDate eventDate = tradeDate == null ? LocalDate.MIN : tradeDate;
        cashEvents.add(new CashEvent(eventDate, cashDelta));
    }

    private static double getCurrentCashHoldings() {
        if (portfolioCashSnapshots.isEmpty()) {
            double cash = 0.0;
            for (CashEvent event : cashEvents) {
                cash += event.cashDelta;
            }
            return cash;
        }

        LinkedHashMap<String, PortfolioCashSnapshot> latestByPortfolio = new LinkedHashMap<>();
        for (PortfolioCashSnapshot snapshot : portfolioCashSnapshots) {
            if (snapshot == null || snapshot.portfolioId == null || snapshot.portfolioId.isBlank()) {
                continue;
            }

            PortfolioCashSnapshot existing = latestByPortfolio.get(snapshot.portfolioId);
            if (existing == null
                    || snapshot.tradeDate.isAfter(existing.tradeDate)
                    || (snapshot.tradeDate.equals(existing.tradeDate) && snapshot.sortId >= existing.sortId)) {
                latestByPortfolio.put(snapshot.portfolioId, snapshot);
            }
        }

        double authoritativePortfolioCash = 0.0;
        for (PortfolioCashSnapshot snapshot : latestByPortfolio.values()) {
            authoritativePortfolioCash += snapshot.balance;
        }

        return authoritativePortfolioCash;
    }

    private static double sumPortfolioBalancesOnOrBefore(
            LocalDate monthEnd,
            ArrayList<PortfolioCashSnapshot> sortedSnapshots,
            int[] snapshotIndexRef,
            LinkedHashMap<String, Double> latestBalanceByPortfolio) {
        int snapshotIndex = snapshotIndexRef[0];
        while (snapshotIndex < sortedSnapshots.size()
                && !sortedSnapshots.get(snapshotIndex).tradeDate.isAfter(monthEnd)) {
            PortfolioCashSnapshot snapshot = sortedSnapshots.get(snapshotIndex);
            if (snapshot.portfolioId != null && !snapshot.portfolioId.isBlank()) {
                latestBalanceByPortfolio.put(snapshot.portfolioId, snapshot.balance);
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
                loadedTransactionRowCount++;
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
            ArrayList<OverviewRow> overviewRows = buildOverviewRows();
            HeaderSummary headerSummary = buildHeaderSummary(overviewRows);

            writeHtmlHeader(writer, headerSummary);
            writeOverviewTableHtml(writer, overviewRows);
            writeRealizedSummaryTableHtml(writer);
            writeSaleTradesTablesHtml(writer);
            writeHtmlFooter(writer);
            new ProcessBuilder("open", OUTPUT_FILE).start();
        } catch (IOException e) {
            System.out.println("Failed to write report: " + e.getMessage());
        }
    }

    private static ArrayList<OverviewRow> buildOverviewRows() {
        ArrayList<OverviewRow> overviewRows = new ArrayList<>();
        for (Security security : getSortedSecuritiesForOverview()) {
            overviewRows.add(buildOverviewRow(security));
        }
        return overviewRows;
    }

    private static HeaderSummary buildHeaderSummary(ArrayList<OverviewRow> overviewRows) {
        String generatedDate = LocalDate.now().format(DateTimeFormatter.ofPattern("dd.MM.yyyy"));
        int holdingsCount = overviewRows == null ? 0 : overviewRows.size();

        double totalMarketValue = 0.0;
        double cashHoldings = getCurrentCashHoldings();
        double totalReturn = 0.0;
        double totalHistoricalCostBasis = 0.0;
        String totalCurrencyCode = null;

        OverviewRow best = null;
        OverviewRow worst = null;

        if (overviewRows != null) {
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
        }

        double totalReturnPct = totalHistoricalCostBasis > 0.0
                ? (totalReturn / totalHistoricalCostBasis) * 100.0
                : 0.0;

        String bestLabel = best == null ? "-" : getOverviewRowLabel(best);
        String worstLabel = worst == null ? "-" : getOverviewRowLabel(worst);
        double bestReturn = best == null ? 0.0 : best.totalReturn;
        double worstReturn = worst == null ? 0.0 : worst.totalReturn;
        double bestReturnPct = best == null ? 0.0 : best.totalReturnPct;
        double worstReturnPct = worst == null ? 0.0 : worst.totalReturnPct;

        String sparklineSvg = buildHeaderSparklineSvg(overviewRows);

        return new HeaderSummary(
                generatedDate,
                loadedCsvFileCount,
                loadedTransactionRowCount,
                holdingsCount,
                totalCurrencyCode,
                totalMarketValue,
                cashHoldings,
                totalReturn,
                totalReturnPct,
                bestLabel,
                bestReturn,
                bestReturnPct,
                worstLabel,
                worstReturn,
                worstReturnPct,
                sparklineSvg
        );
    }

    private static String buildHeaderSparklineSvg(ArrayList<OverviewRow> overviewRows) {
        final double width = 320.0;
        final double height = 72.0;
        final double left = 50.0;
        final double right = 8.0;
        final double top = 8.0;
        final double bottom = 18.0;
        final double plotWidth = width - left - right;
        final double plotHeight = height - top - bottom;

        ArrayList<PortfolioValuePoint> points = buildPortfolioValueTimelineLast12Months();

        StringBuilder svg = new StringBuilder();
        svg.append("<svg class=\"hero-sparkline\" viewBox=\"0 0 ")
                .append(svgNumber(width)).append(" ").append(svgNumber(height))
            .append("\" preserveAspectRatio=\"xMinYMid meet\" xmlns=\"http://www.w3.org/2000/svg\" role=\"img\">\n");

        if (points.isEmpty()) {
            svg.append("<text x=\"").append(svgNumber(width / 2.0)).append("\" y=\"").append(svgNumber(height / 2.0 + 3.0))
                .append("\" text-anchor=\"middle\" font-size=\"9\" fill=\"#eaf2ff\">No 12M portfolio history</text>\n");
            svg.append("</svg>\n");
            return svg.toString();
        }

        int count = points.size();
        double minValue = Double.POSITIVE_INFINITY;
        double maxValue = Double.NEGATIVE_INFINITY;
        int minIndex = -1;
        int maxIndex = -1;
        for (int i = 0; i < count; i++) {
            double value = points.get(i).value;
            if (value < minValue) {
                minValue = value;
                minIndex = i;
            }
            if (value > maxValue) {
                maxValue = value;
                maxIndex = i;
            }
        }

        if (!Double.isFinite(minValue) || !Double.isFinite(maxValue)) {
            minValue = 0.0;
            maxValue = 1.0;
        }
        if (Math.abs(maxValue - minValue) < 1e-9) {
            maxValue += Math.max(1.0, maxValue * 0.02);
            minValue = Math.max(0.0, minValue - Math.max(1.0, minValue * 0.02));
        }

        svg.append("<line x1=\"").append(svgNumber(left)).append("\" y1=\"").append(svgNumber(top))
            .append("\" x2=\"").append(svgNumber(left)).append("\" y2=\"").append(svgNumber(top + plotHeight))
            .append("\" stroke=\"#c7d2df\" stroke-width=\"1\"/>\n");
        svg.append("<line x1=\"").append(svgNumber(left)).append("\" y1=\"").append(svgNumber(top + plotHeight))
            .append("\" x2=\"").append(svgNumber(left + plotWidth)).append("\" y2=\"").append(svgNumber(top + plotHeight))
            .append("\" stroke=\"#c7d2df\" stroke-width=\"1\"/>\n");

        String maxText = formatCompactKroner(maxValue);
        String minText = formatCompactKroner(minValue);
        svg.append("<text x=\"").append(svgNumber(left - 5.0)).append("\" y=\"").append(svgNumber(top + 1.0))
            .append("\" text-anchor=\"end\" dominant-baseline=\"hanging\" font-size=\"6\" fill=\"#eaf2ff\">")
            .append(escapeHtml(maxText)).append("</text>\n");
        svg.append("<text x=\"").append(svgNumber(left - 5.0)).append("\" y=\"").append(svgNumber(top + plotHeight))
            .append("\" text-anchor=\"end\" dominant-baseline=\"middle\" font-size=\"6\" fill=\"#eaf2ff\">")
            .append(escapeHtml(minText)).append("</text>\n");

        double[] xValues = new double[count];
        double[] yValues = new double[count];
        for (int i = 0; i < count; i++) {
            double x = count == 1
                    ? left + (plotWidth / 2.0)
                    : left + ((plotWidth / (count - 1.0)) * i);
            double y = mapValueToY(points.get(i).value, minValue, maxValue, top, plotHeight);
            xValues[i] = x;
            yValues[i] = y;
        }

        double dashedGuideY = maxIndex >= 0 ? yValues[maxIndex] : top;
        svg.append("<line x1=\"").append(svgNumber(left)).append("\" y1=\"").append(svgNumber(dashedGuideY))
            .append("\" x2=\"").append(svgNumber(left + plotWidth)).append("\" y2=\"").append(svgNumber(dashedGuideY))
            .append("\" stroke=\"#eaf2ff\" stroke-width=\"0.8\" opacity=\"0.45\" stroke-dasharray=\"3 3\"/>\n");

        for (int i = 1; i < count; i++) {
            svg.append("<line x1=\"").append(svgNumber(xValues[i - 1])).append("\" y1=\"").append(svgNumber(yValues[i - 1]))
                .append("\" x2=\"").append(svgNumber(xValues[i])).append("\" y2=\"").append(svgNumber(yValues[i]))
                .append("\" stroke=\"#f1f6ff\" stroke-width=\"2\" stroke-linecap=\"round\"/>\n");
        }

        for (int i = 0; i < count; i++) {
            svg.append("<circle cx=\"").append(svgNumber(xValues[i])).append("\" cy=\"").append(svgNumber(yValues[i]))
                .append("\" r=\"2.2\" fill=\"#f1f6ff\">\n")
                .append("<title>")
                .append(escapeHtml(points.get(i).monthEnd.format(DateTimeFormatter.ofPattern("MMM yyyy", Locale.ENGLISH))
                        + ": " + formatNumber(points.get(i).value, 0) + " kr"))
                .append("</title></circle>\n");
        }

        if (minIndex >= 0) {
            svg.append("<circle cx=\"").append(svgNumber(xValues[minIndex])).append("\" cy=\"").append(svgNumber(yValues[minIndex]))
                .append("\" r=\"3.7\" fill=\"none\" stroke=\"#ffd0cd\" stroke-width=\"1.1\"/>\n");
        }
        if (maxIndex >= 0 && maxIndex != minIndex) {
            svg.append("<circle cx=\"").append(svgNumber(xValues[maxIndex])).append("\" cy=\"").append(svgNumber(yValues[maxIndex]))
                .append("\" r=\"3.7\" fill=\"none\" stroke=\"#c9f7cf\" stroke-width=\"1.1\"/>\n");
        }

        double axisY = top + plotHeight;
        DateTimeFormatter axisMonthFormat = DateTimeFormatter.ofPattern("MMM yy", Locale.ENGLISH);
        int[] tickIndices = count >= 4
                ? new int[] {0, count / 3, (count * 2) / 3, count - 1}
                : new int[] {0, count - 1};

        int previousIndex = -1;
        for (int tickIndex : tickIndices) {
            if (tickIndex < 0 || tickIndex >= count || tickIndex == previousIndex) {
                continue;
            }

            svg.append("<line x1=\"").append(svgNumber(xValues[tickIndex])).append("\" y1=\"").append(svgNumber(axisY))
                .append("\" x2=\"").append(svgNumber(xValues[tickIndex])).append("\" y2=\"").append(svgNumber(axisY + 2.8))
                .append("\" stroke=\"#eaf2ff\" stroke-width=\"0.8\"/>\n");
            svg.append("<text x=\"").append(svgNumber(xValues[tickIndex])).append("\" y=\"").append(svgNumber(axisY + 8.8))
                .append("\" text-anchor=\"middle\" font-size=\"5.7\" fill=\"#eaf2ff\">")
                .append(escapeHtml(points.get(tickIndex).monthEnd.format(axisMonthFormat)))
                .append("</text>\n");
            previousIndex = tickIndex;
        }

        svg.append("</svg>\n");
        return svg.toString();
    }

    private static ArrayList<PortfolioValuePoint> buildPortfolioValueTimelineLast12Months() {
        ArrayList<PortfolioValuePoint> timeline = new ArrayList<>();
        if (unitEvents.isEmpty() && cashEvents.isEmpty() && portfolioCashSnapshots.isEmpty()) {
            return timeline;
        }

        boolean useAuthoritativeSnapshots = !portfolioCashSnapshots.isEmpty();

        ArrayList<UnitEvent> sortedEvents = new ArrayList<>(unitEvents);
        sortedEvents.sort(Comparator.comparing((UnitEvent e) -> e.tradeDate)
                .thenComparing(e -> e.securityKey));

        ArrayList<CashEvent> sortedCashEvents = new ArrayList<>();
        if (!useAuthoritativeSnapshots) {
            sortedCashEvents.addAll(cashEvents);
            sortedCashEvents.sort(Comparator.comparing((CashEvent e) -> e.tradeDate));
        }

        ArrayList<PortfolioCashSnapshot> sortedCashSnapshots = new ArrayList<>(portfolioCashSnapshots);
        sortedCashSnapshots.sort(
            Comparator.comparing((PortfolioCashSnapshot s) -> s.tradeDate)
                .thenComparingLong(s -> s.sortId)
        );

        Map<String, Security> securitiesByTrackingKey = new HashMap<>();
        for (Security security : securities) {
            securitiesByTrackingKey.put(getTrackingSecurityKey(security), security);
        }

        Map<String, Double> unitsBySecurity = new HashMap<>();
        Map<String, NavigableMap<LocalDate, Double>> priceSeriesCache = new HashMap<>();

        YearMonth endMonth = YearMonth.now();
        YearMonth startMonth = endMonth.minusMonths(11);
        int eventIndex = 0;
        int cashEventIndex = 0;
        int[] cashSnapshotIndexRef = new int[] {0};
        double runningCash = 0.0;
        LinkedHashMap<String, Double> latestBalanceByPortfolio = new LinkedHashMap<>();

        for (int monthOffset = 0; monthOffset < 12; monthOffset++) {
            YearMonth currentMonth = startMonth.plusMonths(monthOffset);
            LocalDate monthEnd = currentMonth.atEndOfMonth();

            while (eventIndex < sortedEvents.size() && !sortedEvents.get(eventIndex).tradeDate.isAfter(monthEnd)) {
                UnitEvent event = sortedEvents.get(eventIndex);
                unitsBySecurity.merge(event.securityKey, event.unitsDelta, Double::sum);
                eventIndex++;
            }

            while (cashEventIndex < sortedCashEvents.size()
                    && !sortedCashEvents.get(cashEventIndex).tradeDate.isAfter(monthEnd)) {
                CashEvent cashEvent = sortedCashEvents.get(cashEventIndex);
                runningCash += cashEvent.cashDelta;
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

                totalValue += units * price;
            }

            timeline.add(new PortfolioValuePoint(monthEnd, totalValue));
        }

        return timeline;
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

    private static String formatCompactKroner(double value) {
        double absValue = Math.abs(value);
        String prefix = value < 0.0 ? "-" : "";
        if (absValue >= 1_000_000_000.0) {
            return prefix + formatNumber(absValue / 1_000_000_000.0, 1) + "B kr";
        }
        if (absValue >= 1_000_000.0) {
            return prefix + formatNumber(absValue / 1_000_000.0, 1) + "M kr";
        }
        if (absValue >= 1_000.0) {
            return prefix + formatNumber(absValue / 1_000.0, 0) + "k kr";
        }
        return prefix + formatNumber(absValue, 0) + " kr";
    }

    private static void writeHtmlHeader(FileWriter writer, HeaderSummary summary) throws IOException {
        String generatedDate = summary.generatedDate;
        String totalMarketValueText = formatTotalMoney(summary.totalMarketValue, summary.totalCurrencyCode, 0);
        String cashHoldingsText = formatTotalMoney(summary.cashHoldings, summary.totalCurrencyCode, 0);
        String totalReturnText = formatTotalMoney(summary.totalReturn, summary.totalCurrencyCode, 0);
        String totalReturnPctText = formatPercent(summary.totalReturnPct, 2);
        String bestReturnText = formatTotalMoney(summary.bestReturn, summary.totalCurrencyCode, 0) + " (" + formatPercent(summary.bestReturnPct, 2) + ")";
        String worstReturnText = formatTotalMoney(summary.worstReturn, summary.totalCurrencyCode, 0) + " (" + formatPercent(summary.worstReturnPct, 2) + ")";

        writer.write("<!DOCTYPE html>\n");
        writer.write("<html lang=\"en\">\n");
        writer.write("<head>\n");
        writer.write("  <meta charset=\"utf-8\">\n");
        writer.write("  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n");
        writer.write("  <title>Portfolio Report</title>\n");
        writer.write("  <style>\n");
        writer.write("    body { font-family: -apple-system, BlinkMacSystemFont, Segoe UI, sans-serif; margin: 24px; color: #111; background: #f7f8fb; }\n");
        writer.write("    h1 { margin: 0 0 20px 0; font-size: 26px; }\n");
        writer.write("    h2 { margin: 28px 0 8px 0; font-size: 18px; }\n");
        writer.write("    .report-hero { margin: 0 0 18px 0; padding: 16px 18px; border-radius: 10px; background: linear-gradient(135deg, #0b7285, #1c7ed6); color: #fff; box-shadow: 0 8px 22px rgba(15, 23, 42, 0.15); }\n");
        writer.write("    .report-hero h1 { margin: 0; font-size: 26px; letter-spacing: 0.2px; }\n");
        writer.write("    .report-hero .meta { margin-top: 8px; display: flex; flex-wrap: wrap; gap: 10px 16px; font-size: 12px; opacity: 0.95; }\n");
        writer.write("    .report-hero .meta span { background: rgba(255, 255, 255, 0.14); border: 1px solid rgba(255, 255, 255, 0.22); border-radius: 999px; padding: 3px 10px; }\n");
        writer.write("    .hero-grid { margin-top: 12px; display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 12px; align-items: stretch; }\n");
        writer.write("    .hero-kpi-grid { display: contents; }\n");
        writer.write("    .hero-kpi-col { display: grid; gap: 10px; align-content: stretch; height: 100%; }\n");
        writer.write("    .hero-card, .hero-spark-card { border: 1px solid rgba(255,255,255,0.28); border-radius: 8px; background: rgba(255,255,255,0.12); padding: 9px 10px; }\n");
        writer.write("    .hero-card .label, .hero-spark-card .label { font-size: 11px; opacity: 0.9; }\n");
        writer.write("    .hero-card .value { margin-top: 4px; font-size: 18px; font-weight: 700; letter-spacing: 0.2px; }\n");
        writer.write("    .hero-card .name { font-size: 12px; font-weight: 600; margin-bottom: 2px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }\n");
        writer.write("    .hero-card .subvalue { font-size: 12px; opacity: 0.95; }\n");
        writer.write("    .hero-spark-card { display: flex; flex-direction: column; justify-content: center; padding: 8px 9px; height: 100%; }\n");
        writer.write("    .hero-sparkline { width: 100%; height: auto; display: block; }\n");
        writer.write("    table { border-collapse: collapse; width: 100%; margin: 8px 0 18px 0; table-layout: auto; }\n");
        writer.write("    th, td { border: 1px solid #d0d0d0; padding: 6px 8px; font-size: 13px; text-align: left; white-space: nowrap; }\n");
        writer.write("    .sale-trades-table { table-layout: fixed; }\n");
        writer.write("    .sale-trades-table th, .sale-trades-table td { overflow: hidden; text-overflow: ellipsis; }\n");
        writer.write("    th { background: #f3f3f3; font-weight: 600; }\n");
        writer.write("    td.num { text-align: right; font-variant-numeric: tabular-nums; }\n");
        writer.write("    td.text { text-align: left; }\n");
        writer.write("    tr.total-row td { font-weight: 700; background: #fafafa; }\n");
        writer.write("    tr.asset-split td { border-top: 2px solid #9a9a9a; }\n");
        writer.write("    .muted { color: #666; font-size: 12px; margin-top: -8px; }\n");
        writer.write("    .overview-charts { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 14px; margin: 8px 0 14px 0; align-items: stretch; }\n");
        writer.write("    .overview-chart { border: 1px solid #d0d0d0; border-radius: 6px; background: #fff; padding: 10px; }\n");
        writer.write("    .overview-chart h3 { margin: 0 0 8px 0; font-size: 15px; font-weight: 600; }\n");
        writer.write("    .chart-svg { width: 100%; height: 350px; display: block; }\n");
        writer.write("    .overview-chart.total-return-chart .chart-svg { height: 430px; }\n");
        writer.write("    .overview-chart.allocation-card { margin-top: 12px; }\n");
        writer.write("    .overview-chart.allocation-card .chart-svg { height: 230px; }\n");
        writer.write("    .overview-chart.allocation-card .chart-svg.market-value-bar-chart { height: 290px; }\n");
        writer.write("    .allocation-visuals { display: grid; grid-template-columns: repeat(6, minmax(0, 1fr)); gap: 10px; }\n");
        writer.write("    .allocation-panel { border: 1px solid #ececec; border-radius: 6px; padding: 6px; }\n");
        writer.write("    .allocation-panel-title { margin: 0 0 6px 0; font-size: 12px; color: #2b2b2b; font-weight: 600; }\n");
        writer.write("    .allocation-panel.asset-type-panel, .allocation-panel.sector-panel, .allocation-panel.region-panel { grid-column: span 2; }\n");
        writer.write("    .allocation-panel.security-pie-panel { grid-column: span 2; }\n");
        writer.write("    .allocation-panel.security-bar-panel { grid-column: span 4; }\n");
        writer.write("    @media (max-width: 1200px) { .hero-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); } .hero-spark-card { grid-column: span 2; } }\n");
        writer.write("    @media (max-width: 980px) { .hero-grid { grid-template-columns: 1fr; } .hero-spark-card { grid-column: span 1; } }\n");
        writer.write("    @media (max-width: 700px) { .report-hero { padding: 14px; } .report-hero h1 { font-size: 22px; } }\n");
        writer.write("    @media (max-width: 1200px) { .allocation-visuals { grid-template-columns: repeat(2, minmax(0, 1fr)); } .allocation-panel.asset-type-panel, .allocation-panel.sector-panel, .allocation-panel.region-panel, .allocation-panel.security-pie-panel, .allocation-panel.security-bar-panel { grid-column: span 1; } .allocation-panel.security-bar-panel { grid-column: span 2; } }\n");
        writer.write("    @media (max-width: 980px) { .allocation-visuals { grid-template-columns: 1fr; } .allocation-panel.security-bar-panel { grid-column: span 1; } }\n");
        writer.write("    @media (max-width: 980px) { .overview-charts { grid-template-columns: 1fr; } }\n");
        writer.write("  </style>\n");
        writer.write("</head>\n");
        writer.write("<body>\n");
        writer.write("  <header class=\"report-hero\">\n");
        writer.write("    <h1>Portfolio Report</h1>\n");
        writer.write("    <div class=\"meta\"><span>Date: " + escapeHtml(generatedDate) + "</span><span>Files: " + summary.fileCount + "</span><span>Transactions: " + summary.transactionCount + "</span><span>Holdings: " + summary.holdingsCount + "</span></div>\n");
        writer.write("    <div class=\"hero-grid\">\n");
        writer.write("      <div class=\"hero-kpi-grid\">\n");
        writer.write("        <div class=\"hero-kpi-col\">\n");
        writer.write("          <div class=\"hero-card\"><div class=\"label\">Market Value</div><div class=\"value\">" + escapeHtml(totalMarketValueText) + "</div></div>\n");
        writer.write("          <div class=\"hero-card\"><div class=\"label\">Best / Worst Holding</div><div class=\"name\">Best: " + escapeHtml(summary.bestLabel) + "</div><div class=\"subvalue\">" + escapeHtml(bestReturnText) + "</div><div class=\"name\" style=\"margin-top:6px;\">Worst: " + escapeHtml(summary.worstLabel) + "</div><div class=\"subvalue\">" + escapeHtml(worstReturnText) + "</div></div>\n");
        writer.write("        </div>\n");
        writer.write("        <div class=\"hero-kpi-col\">\n");
        writer.write("          <div class=\"hero-card\"><div class=\"label\">Cash Holdings</div><div class=\"value\">" + escapeHtml(cashHoldingsText) + "</div></div>\n");
        writer.write("          <div class=\"hero-card\"><div class=\"label\">Total Return</div><div class=\"value\">" + escapeHtml(totalReturnText) + "</div></div>\n");
        writer.write("          <div class=\"hero-card\"><div class=\"label\">Total Return (%)</div><div class=\"value\">" + escapeHtml(totalReturnPctText) + "</div></div>\n");
        writer.write("        </div>\n");
        writer.write("      </div>\n");
        writer.write("      <div class=\"hero-spark-card\"><div class=\"label\">Portfolio Value (12M)</div>" + summary.sparklineSvg + "</div>\n");
        writer.write("    </div>\n");
        writer.write("  </header>\n");
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

        return text.matches("-?[0-9][0-9 ]*(\\.[0-9]+)?( [A-Za-z%]+)?");
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

    private static void writeHtmlRowWithClass(FileWriter writer, String rowClass, String... fields) throws IOException {
        String classAttribute = (rowClass == null || rowClass.isBlank()) ? "" : " class=\"" + escapeHtml(rowClass) + "\"";
        writer.write("<tr" + classAttribute + ">");
        for (int i = 0; i < fields.length; i++) {
            writer.write(toDataCell(fields[i]));
        }
        writer.write("</tr>\n");
    }

    private static String getCurrencySuffix(String currencyCode) {
        if (currencyCode == null || currencyCode.isBlank()) {
            return "kr";
        }

        String normalized = currencyCode.trim().toUpperCase(Locale.ROOT);
        return switch (normalized) {
            case "NOK" -> "kr";
            default -> normalized.toLowerCase(Locale.ROOT);
        };
    }

    private static String formatMoney(double value, String currencyCode, int decimals) {
        return formatNumber(value, decimals) + " " + getCurrencySuffix(currencyCode);
    }

    private static String formatPercent(double value, int decimals) {
        return formatNumber(value, decimals) + " %";
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
            return formatNumber(value, decimals) + " mixed";
        }
        return formatMoney(value, aggregateCurrencyCode, decimals);
    }

    private static boolean isStockFundBoundary(String previousAssetType, String currentAssetType) {
        if (previousAssetType == null || currentAssetType == null || previousAssetType.equals(currentAssetType)) {
            return false;
        }

        return ("STOCK".equals(previousAssetType) && "FUND".equals(currentAssetType))
                || ("FUND".equals(previousAssetType) && "STOCK".equals(currentAssetType));
    }

    private static OverviewRow buildOverviewRow(Security security) {
        String ticker = security.getTicker();
        String tickerText = (ticker == null || ticker.isBlank()) ? "-" : ticker;
        String currencyCode = security.getCurrencyCode();
        double units = security.getUnitsOwned();
        double averageCost = security.getAverageCost();
        double latestPrice = security.getLatestPrice();
        double positionCostBasis = units * averageCost;
        double realizedCostBasis = security.getRealizedCostBasis();
        double historicalCostBasis = positionCostBasis + realizedCostBasis;
        boolean hasPrice = latestPrice > 0.0;
        double marketValue = hasPrice ? units * latestPrice : 0.0;
        double unrealized = hasPrice ? (marketValue - positionCostBasis) : 0.0;
        double unrealizedPct = hasPrice && positionCostBasis > 0 ? (unrealized / positionCostBasis) * 100.0 : 0.0;
        double realized = parseDoubleOrZero(security.getRealizedGainAsText());
        double dividends = parseDoubleOrZero(security.getDividendsAsText());
        double totalReturn = unrealized + realized + dividends;
        double totalReturnPct = historicalCostBasis > 0 ? (totalReturn / historicalCostBasis) * 100.0 : 0.0;

        return new OverviewRow(
                tickerText,
            getPreferredSecurityName(security),
                security.getAssetType().name(),
                security.getResolvedSector(),
                security.getResolvedSectorWeights(),
                security.getResolvedRegion(),
                security.getResolvedRegionWeights(),
                currencyCode,
                security.getRealizedReturnPctAsText(),
                units,
                averageCost,
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
    }

    private static void writeOverviewChartsHtml(FileWriter writer, ArrayList<OverviewRow> rows) throws IOException {
        if (rows.isEmpty()) {
            return;
        }

        writer.write("<div class=\"overview-charts\">\n");
        writeOverviewChartCard(writer, "Total Return (NOK)", rows, false);
        writeOverviewChartCard(writer, "Total Return (%)", rows, true);
        writer.write("</div>\n");
    }

    private static void writeOverviewChartCard(FileWriter writer, String title, ArrayList<OverviewRow> rows, boolean percentChart) throws IOException {
        writer.write("<section class=\"overview-chart total-return-chart\">\n");
        writer.write("<h3>" + escapeHtml(title) + "</h3>\n");
        writer.write(buildOverviewBarChartSvg(rows, percentChart));
        writer.write("</section>\n");
    }

    private static void writeMarketValueAllocationCard(FileWriter writer, ArrayList<OverviewRow> rows) throws IOException {
        writer.write("<section class=\"overview-chart allocation-card\">\n");
        writer.write("<h3>Market Value Allocation</h3>\n");
        writer.write("<div class=\"allocation-visuals\">\n");
        writer.write("<div class=\"allocation-panel asset-type-panel\">\n");
        writer.write("<h4 class=\"allocation-panel-title\">By Asset Type</h4>\n");
        writer.write(buildAssetTypeAllocationSvg(rows));
        writer.write("</div>\n");
        writer.write("<div class=\"allocation-panel sector-panel\">\n");
        writer.write("<h4 class=\"allocation-panel-title\">By Sector</h4>\n");
        writer.write(buildSectorAllocationSvg(rows));
        writer.write("</div>\n");
        writer.write("<div class=\"allocation-panel region-panel\">\n");
        writer.write("<h4 class=\"allocation-panel-title\">By Region</h4>\n");
        writer.write(buildRegionAllocationSvg(rows));
        writer.write("</div>\n");
        writer.write("<div class=\"allocation-panel security-pie-panel\">\n");
        writer.write("<h4 class=\"allocation-panel-title\">By Security (Pie)</h4>\n");
        writer.write(buildMarketValueAllocationSvg(rows));
        writer.write("</div>\n");
        writer.write("<div class=\"allocation-panel security-bar-panel\">\n");
        writer.write("<h4 class=\"allocation-panel-title\">By Security (Bar)</h4>\n");
        writer.write(buildMarketValueBarChartSvg(rows));
        writer.write("</div>\n");
        writer.write("</div>\n");
        writer.write("</section>\n");
    }

    private static String getOverviewRowLabel(OverviewRow row) {
        return (row.securityDisplayName == null || row.securityDisplayName.isBlank()) ? row.tickerText : row.securityDisplayName;
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

    private static String buildOverviewBarChartSvg(ArrayList<OverviewRow> rows, boolean percentChart) {
        final double width = 1100.0;
        final double height = 430.0;
        final double left = 68.0;
        final double right = 22.0;
        final double top = 26.0;
        final double bottom = 92.0;
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

        int tickCount = percentChart ? 7 : 5;
        for (int i = 0; i <= tickCount; i++) {
            double tickValue = maxValue - ((valueRange / tickCount) * i);
            double y = mapValueToY(tickValue, minValue, maxValue, top, plotHeight);

            svg.append("<line x1=\"").append(svgNumber(left)).append("\" y1=\"").append(svgNumber(y))
                    .append("\" x2=\"").append(svgNumber(left + plotWidth)).append("\" y2=\"").append(svgNumber(y))
                    .append("\" stroke=\"#ececec\" stroke-width=\"1\"/>\n");

                svg.append("<text x=\"").append(svgNumber(left - 8.0)).append("\" y=\"").append(svgNumber(y + 4.0))
                    .append("\" text-anchor=\"end\" font-size=\"11\" fill=\"#666\">")
                    .append(escapeHtml(formatChartValue(tickValue, percentChart, true)))
                    .append("</text>\n");
        }

        double slotWidth = rows.isEmpty() ? plotWidth : plotWidth / rows.size();
        double barWidth = Math.max(6.0, slotWidth * 0.48);

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
            String label = getOverviewRowLabel(row);
            String compactLabel = getCompactBarLabel(row);

            svg.append("<rect x=\"").append(svgNumber(x)).append("\" y=\"").append(svgNumber(barY))
                    .append("\" width=\"").append(svgNumber(barWidth)).append("\" height=\"").append(svgNumber(barHeight))
                    .append("\" fill=\"").append(barColor).append("\" rx=\"2\">\n")
                    .append("<title>")
                    .append(escapeHtml(label + ": " + formatChartValue(value, percentChart, false)))
                    .append("</title></rect>\n");

            double labelAnchorX = x + (barWidth / 2.0);
            double labelAnchorY = height - bottom + 16.0;
            svg.append("<text x=\"").append(svgNumber(labelAnchorX)).append("\" y=\"").append(svgNumber(labelAnchorY))
                    .append("\" transform=\"rotate(-38 ").append(svgNumber(labelAnchorX)).append(" ").append(svgNumber(labelAnchorY))
                    .append(")\" text-anchor=\"end\" font-size=\"11\" font-weight=\"600\" fill=\"#1f2933\" paint-order=\"stroke\" stroke=\"#ffffff\" stroke-width=\"2\" stroke-linejoin=\"round\">")
                    .append(escapeHtml(compactLabel))
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
        ArrayList<PieSliceLabel> pieLabels = new ArrayList<>();
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

            double midAngle = currentAngle + (sliceAngle / 2.0);
            double anchorX = centerX + radius * Math.cos(midAngle);
            double anchorY = centerY + radius * Math.sin(midAngle);
            double bendX = centerX + (radius + 16.0) * Math.cos(midAngle);
            double bendY = centerY + (radius + 10.0) * Math.sin(midAngle);
            boolean rightSide = Math.cos(midAngle) >= 0.0;

            String labelText = getCompactPieLabel(row)
                    + " "
                    + formatNumber(fraction * 100.0, 1)
                    + "%";
            pieLabels.add(new PieSliceLabel(rightSide, color, labelText, anchorX, anchorY, bendX, bendY));

            currentAngle = endAngle;
        }

        adjustPieLabelPositions(pieLabels, false, 18.0, 250.0, 18.0);
        adjustPieLabelPositions(pieLabels, true, 18.0, 250.0, 18.0);

        for (PieSliceLabel label : pieLabels) {
            double textX = label.rightSide ? (336.0) : (104.0);
            double lineEndX = label.rightSide ? (330.0) : (110.0);
            String textAnchor = label.rightSide ? "start" : "end";

            svg.append("<polyline points=\"")
                .append(svgNumber(label.anchorX)).append(",").append(svgNumber(label.anchorY)).append(" ")
                .append(svgNumber(label.bendX)).append(",").append(svgNumber(label.labelY)).append(" ")
                .append(svgNumber(lineEndX)).append(",").append(svgNumber(label.labelY)).append("\"")
                .append(" fill=\"none\" stroke=\"").append(label.color)
                .append("\" stroke-width=\"0.9\" opacity=\"0.85\"/>\n");

            svg.append("<text x=\"").append(svgNumber(textX)).append("\" y=\"").append(svgNumber(label.labelY))
                .append("\" text-anchor=\"").append(textAnchor)
                .append("\" dominant-baseline=\"middle\" font-size=\"12\" fill=\"#2f2f2f\" paint-order=\"stroke\" stroke=\"#ffffff\" stroke-width=\"2\" stroke-linejoin=\"round\">")
                .append(escapeHtml(label.text))
                .append("</text>\n");
        }

        double summaryY = centerY + radius + 30.0;
        svg.append("<text x=\"").append(svgNumber(centerX)).append("\" y=\"").append(svgNumber(summaryY))
            .append("\" text-anchor=\"middle\" font-size=\"12\" fill=\"#666\">Market Value Total</text>\n");
        svg.append("<text x=\"").append(svgNumber(centerX)).append("\" y=\"").append(svgNumber(summaryY + 16.0))
            .append("\" text-anchor=\"middle\" font-size=\"14\" fill=\"#222\" font-weight=\"600\">")
            .append(escapeHtml(formatNumber(totalMarketValue, 0) + " kr"))
            .append("</text>\n");

        svg.append("</svg>\n");
        return svg.toString();
        }

    private static String getCompactPieLabel(OverviewRow row) {
        String label = getOverviewRowLabel(row);
        if (label == null || label.isBlank()) {
            return "-";
        }

        return label;
    }

    private static void adjustPieLabelPositions(ArrayList<PieSliceLabel> labels, boolean rightSide,
                                                double minY, double maxY, double minGap) {
        ArrayList<PieSliceLabel> sideLabels = new ArrayList<>();
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

        private static String buildMarketValueBarChartSvg(ArrayList<OverviewRow> rows) {
        final double width = 860.0;
        final double height = 330.0;
        final double left = 74.0;
        final double right = 86.0;
        final double top = 18.0;
            final double bottom = 92.0;
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
        svg.append("<svg class=\"chart-svg market-value-bar-chart\" viewBox=\"0 0 ")
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
            String compactLabel = getCompactBarLabel(row);
            String color = getAllocationColor(i);

            svg.append("<rect x=\"").append(svgNumber(x)).append("\" y=\"").append(svgNumber(y))
                .append("\" width=\"").append(svgNumber(barWidth)).append("\" height=\"").append(svgNumber(barHeight))
                .append("\" fill=\"").append(color).append("\" rx=\"2\">\n")
                .append("<title>")
                .append(escapeHtml(label + ": " + formatNumber(row.marketValue, 2) + " kr"))
                .append("</title></rect>\n");

            double labelAnchorX = x + (barWidth / 2.0);
            double labelAnchorY = height - bottom + 16.0;
            svg.append("<text x=\"").append(svgNumber(labelAnchorX)).append("\" y=\"").append(svgNumber(labelAnchorY))
                .append("\" transform=\"rotate(-30 ").append(svgNumber(labelAnchorX)).append(" ").append(svgNumber(labelAnchorY))
                .append(")\" text-anchor=\"end\" font-size=\"10\" font-weight=\"600\" fill=\"#1f2933\" paint-order=\"stroke\" stroke=\"#ffffff\" stroke-width=\"2\" stroke-linejoin=\"round\">")
                .append(escapeHtml(compactLabel))
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
        final double width = 440.0;
        final double height = 330.0;
        final double centerX = width / 2.0;
        final double centerY = 126.0;
        final double radius = 98.0;

        double stockValue = 0.0;
        double fundValue = 0.0;
        double rawCashValue = getCurrentCashHoldings();
        double cashValue = Math.max(0.0, rawCashValue);
        double cashDebtValue = rawCashValue < 0.0 ? -rawCashValue : 0.0;
        double otherValue = 0.0;
        int stockCount = 0;
        int fundCount = 0;
        int cashCount = cashValue > 0.0 ? 1 : 0;
        int cashDebtCount = cashDebtValue > 0.0 ? 1 : 0;
        int otherCount = 0;

        for (OverviewRow row : rows) {
            if (row.marketValue <= 0.0) {
                continue;
            }

            switch (row.assetType) {
                case "STOCK" -> {
                    stockValue += row.marketValue;
                    stockCount++;
                }
                case "FUND" -> {
                    fundValue += row.marketValue;
                    fundCount++;
                }
                default -> {
                    otherValue += row.marketValue;
                    otherCount++;
                }
            }
        }

        double totalValue = stockValue + fundValue + cashValue + cashDebtValue + otherValue;

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

        double stockPct = totalValue > 0.0 ? (stockValue / totalValue) * 100.0 : 0.0;
        double fundPct = totalValue > 0.0 ? (fundValue / totalValue) * 100.0 : 0.0;
        double cashPct = totalValue > 0.0 ? (cashValue / totalValue) * 100.0 : 0.0;
        double cashDebtPct = totalValue > 0.0 ? (cashDebtValue / totalValue) * 100.0 : 0.0;
        double otherPct = totalValue > 0.0 ? (otherValue / totalValue) * 100.0 : 0.0;

        double summaryY = centerY + radius + 12.0;
        svg.append("<text x=\"").append(svgNumber(centerX)).append("\" y=\"").append(svgNumber(summaryY))
            .append("\" text-anchor=\"middle\" font-size=\"10\" fill=\"#666\">Asset Mix</text>\n");
        ArrayList<String> legendLabels = new ArrayList<>();
        ArrayList<Double> legendPcts = new ArrayList<>();
        ArrayList<String> legendColors = new ArrayList<>();

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

        double legendYStart = 244.0;
        for (int i = 0; i < legendLabels.size(); i++) {
            double y = legendYStart + (i * 14.0);
            svg.append("<circle cx=\"22\" cy=\"").append(svgNumber(y - 3.0)).append("\" r=\"3.7\" fill=\"")
                .append(legendColors.get(i)).append("\"/>\n");
            svg.append("<text x=\"31\" y=\"").append(svgNumber(y)).append("\" text-anchor=\"start\" font-size=\"10\" fill=\"#2f2f2f\">")
                .append(escapeHtml(legendLabels.get(i)))
                .append("</text>\n");
            svg.append("<text x=\"").append(svgNumber(width - 16.0)).append("\" y=\"").append(svgNumber(y))
                .append("\" text-anchor=\"end\" font-size=\"10\" fill=\"#4a4a4a\">")
                .append(escapeHtml(formatNumber(legendPcts.get(i), 1) + "%"))
                .append("</text>\n");
        }

        svg.append("</svg>\n");
        return svg.toString();
        }

    private static String buildSectorAllocationSvg(ArrayList<OverviewRow> rows) {
        return buildCategoricalAllocationSvg(rows, true, "Sector Mix", "No sector data");
    }

    private static String buildRegionAllocationSvg(ArrayList<OverviewRow> rows) {
        return buildCategoricalAllocationSvg(rows, false, "Region Mix", "No region data");
    }

    private static String buildCategoricalAllocationSvg(ArrayList<OverviewRow> rows, boolean sectorChart,
                                                        String centerTitle, String emptyMessage) {
        final double width = 440.0;
        final double height = 330.0;
        final double centerX = width / 2.0;
        final double centerY = 126.0;
        final double radius = 98.0;

        LinkedHashMap<String, Double> valueByCategory = new LinkedHashMap<>();
        double totalMarketValue = 0.0;
        for (OverviewRow row : rows) {
            if (row.marketValue <= 0.0) {
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
                        double weightedValue = row.marketValue * (weight / totalWeight);
                        valueByCategory.merge(sectorLabel, weightedValue, Double::sum);
                    }

                    totalMarketValue += row.marketValue;
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
                        double weightedValue = row.marketValue * (weight / totalWeight);
                        valueByCategory.merge(regionLabel, weightedValue, Double::sum);
                    }

                    totalMarketValue += row.marketValue;
                    continue;
                }
            }

            String category = sectorChart ? classifySector(row) : classifyRegion(row);
            valueByCategory.merge(category, row.marketValue, Double::sum);
            totalMarketValue += row.marketValue;
        }

        int maxBuckets = sectorChart ? 12 : 10;
        ArrayList<AllocationBucket> buckets = compactAllocationBuckets(valueByCategory, maxBuckets);

        StringBuilder svg = new StringBuilder();
        svg.append("<svg class=\"chart-svg\" viewBox=\"0 0 ")
            .append(svgNumber(width))
            .append(" ")
            .append(svgNumber(height))
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

            svg.append("<path d=\"M ").append(svgNumber(centerX)).append(" ").append(svgNumber(centerY))
                .append(" L ").append(svgNumber(x1)).append(" ").append(svgNumber(y1))
                .append(" A ").append(svgNumber(radius)).append(" ").append(svgNumber(radius)).append(" 0 ")
                .append(largeArcFlag).append(" 1 ").append(svgNumber(x2)).append(" ").append(svgNumber(y2))
                .append(" Z\" fill=\"").append(color).append("\">\n")
                .append("<title>")
                .append(escapeHtml(bucket.label
                    + ": " + formatNumber(bucket.value, 2) + " kr (" + formatNumber(fraction * 100.0, 1) + "%)"))
                .append("</title></path>\n");

            currentAngle = endAngle;
        }

        double summaryY = centerY + radius + 12.0;
        svg.append("<text x=\"").append(svgNumber(centerX)).append("\" y=\"").append(svgNumber(summaryY))
            .append("\" text-anchor=\"middle\" font-size=\"10\" fill=\"#666\">")
            .append(escapeHtml(centerTitle))
            .append("</text>\n");

        double legendYStart = 244.0;
        for (int i = 0; i < buckets.size(); i++) {
            AllocationBucket bucket = buckets.get(i);
            double pct = (bucket.value / totalMarketValue) * 100.0;
            double y = legendYStart + (i * 14.0);
            String color = getAllocationColor(i);

            svg.append("<circle cx=\"22\" cy=\"").append(svgNumber(y - 3.0)).append("\" r=\"3.7\" fill=\"")
                .append(color).append("\"/>\n");
            svg.append("<text x=\"31\" y=\"").append(svgNumber(y)).append("\" text-anchor=\"start\" font-size=\"10\" fill=\"#2f2f2f\">")
                .append(escapeHtml(bucket.label))
                .append("</text>\n");
            svg.append("<text x=\"").append(svgNumber(width - 16.0)).append("\" y=\"").append(svgNumber(y))
                .append("\" text-anchor=\"end\" font-size=\"10\" fill=\"#4a4a4a\">")
                .append(escapeHtml(formatNumber(pct, 1) + "%"))
                .append("</text>\n");
        }

        svg.append("</svg>\n");
        return svg.toString();
    }

    private static ArrayList<AllocationBucket> compactAllocationBuckets(Map<String, Double> rawBuckets, int maxBuckets) {
        ArrayList<AllocationBucket> sorted = new ArrayList<>();
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
        ArrayList<AllocationBucket> compacted = new ArrayList<>();
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
            return formatPercent(value, 2);
        }
        int decimals = compact ? 0 : 2;
        return formatMoney(value, "NOK", decimals);
    }

    private static void writeOverviewTableHtml(FileWriter writer, ArrayList<OverviewRow> overviewRows) throws IOException {
        writer.write("<h2>PORTFOLIO OVERVIEW - CURRENT HOLDINGS</h2>\n");

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
        double totalRealizedCostBasis = 0.0;
        double totalHistoricalCostBasis = 0.0;
        double totalMarketValue = 0.0;
        double totalUnrealized = 0.0;
        double totalCostBasisWithPrice = 0.0;
        double totalRealized = 0.0;
        double totalDividends = 0.0;
        String totalCurrencyCode = null;
        String previousAssetType = null;
        for (OverviewRow row : overviewRows) {
            totalCostBasis += row.positionCostBasis;
            totalRealizedCostBasis += row.realizedCostBasis;
            totalHistoricalCostBasis += row.historicalCostBasis;
            totalMarketValue += row.marketValue;
            if (row.hasPrice) {
                totalUnrealized += row.unrealized;
                totalCostBasisWithPrice += row.positionCostBasis;
            }
            totalRealized += row.realized;
            totalDividends += row.dividends;
            totalCurrencyCode = mergeCurrencyCodes(totalCurrencyCode, row.currencyCode);

            double realizedPctValue = parseDoubleOrZero(row.realizedReturnPctText);

            String rowClass = isStockFundBoundary(previousAssetType, row.assetType) ? "asset-split" : null;

            writeHtmlRowWithClass(
                writer,
                rowClass,
                row.tickerText,
                row.securityDisplayName,
                row.assetType,
                formatUnits(row.units),
                formatMoney(row.averageCost, row.currencyCode, 2),
                row.latestPrice > 0.0 ? formatMoney(row.latestPrice, row.currencyCode, 2) : "-",
                row.latestPrice > 0.0 ? formatMoney(row.marketValue, row.currencyCode, 2) : "-",
                formatMoney(row.positionCostBasis, row.currencyCode, 2),
                row.hasPrice ? formatMoney(row.unrealized, row.currencyCode, 2) : "-",
                row.hasPrice ? formatPercent(row.unrealizedPct, 2) : "-",
                formatPercent(realizedPctValue, 2),
                formatMoney(row.realized, row.currencyCode, 2),
                formatMoney(row.dividends, row.currencyCode, 2),
                formatMoney(row.totalReturn, row.currencyCode, 2),
                formatPercent(row.totalReturnPct, 2)
            );

            previousAssetType = row.assetType;
        }

        double totalReturn = totalUnrealized + totalRealized + totalDividends;
        double totalUnrealizedPct = totalCostBasisWithPrice > 0 ? (totalUnrealized / totalCostBasisWithPrice) * 100.0 : 0.0;
        double totalRealizedPct = totalRealizedCostBasis > 0 ? (totalRealized / totalRealizedCostBasis) * 100.0 : 0.0;
        double totalReturnPct = totalHistoricalCostBasis > 0 ? (totalReturn / totalHistoricalCostBasis) * 100.0 : 0.0;
        writer.write("<tr class=\"total-row\">"
            + toDataCell("")
            + toDataCell("TOTAL")
            + toDataCell("")
            + toDataCell("")
            + toDataCell("")
            + toDataCell("")
            + toDataCell(formatTotalMoney(totalMarketValue, totalCurrencyCode, 2))
            + toDataCell(formatTotalMoney(totalCostBasis, totalCurrencyCode, 2))
            + toDataCell(formatTotalMoney(totalUnrealized, totalCurrencyCode, 2))
            + toDataCell(formatPercent(totalUnrealizedPct, 2))
            + toDataCell(formatPercent(totalRealizedPct, 2))
            + toDataCell(formatTotalMoney(totalRealized, totalCurrencyCode, 2))
            + toDataCell(formatTotalMoney(totalDividends, totalCurrencyCode, 2))
            + toDataCell(formatTotalMoney(totalReturn, totalCurrencyCode, 2))
            + toDataCell(formatPercent(totalReturnPct, 2))
            + "</tr>\n"
        );

        writer.write("</table>\n\n");

        writeMarketValueAllocationCard(writer, overviewRows);
    }

    private static void writeRealizedSummaryTableHtml(FileWriter writer) throws IOException {
        writer.write("<h2>REALIZED OVERVIEW - ALL SALES</h2>\n");
        writer.write("<table>\n");
        writeHtmlRow(writer, true, "Ticker", "Security", "Sales Value", "Cost Basis", "Realized Gain/Loss", "Dividends", "Return (%)");

        ArrayList<Security> soldSecurities = getSortedSoldSecurities();

        double totalSalesValue = 0.0;
        double totalCostBasis = 0.0;
        double totalRealizedGain = 0.0;
        double totalRealizedDividends = 0.0;
        String totalCurrencyCode = null;
        String previousAssetType = null;

        for (Security security : soldSecurities) {
            String currencyCode = security.getCurrencyCode();
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
            totalCurrencyCode = mergeCurrencyCodes(totalCurrencyCode, currencyCode);

            writeHtmlRowWithClass(
                writer,
                rowClass,
                getTickerText(security),
                getPreferredSecurityName(security),
                formatMoney(salesValue, currencyCode, 2),
                formatMoney(costBasis, currencyCode, 2),
                formatMoney(gain, currencyCode, 2),
                formatMoney(realizedDividends, currencyCode, 2),
                formatPercent(returnPct, 2)
            );

            previousAssetType = currentAssetType;
        }

        double totalReturnPct = totalCostBasis > 0 ? (totalRealizedGain / totalCostBasis) * 100.0 : (totalRealizedGain > 0 ? 100.0 : 0.0);
        writer.write("<tr class=\"total-row\">"
            + toDataCell("")
            + toDataCell("TOTAL")
            + toDataCell(formatTotalMoney(totalSalesValue, totalCurrencyCode, 2))
            + toDataCell(formatTotalMoney(totalCostBasis, totalCurrencyCode, 2))
            + toDataCell(formatTotalMoney(totalRealizedGain, totalCurrencyCode, 2))
            + toDataCell(formatTotalMoney(totalRealizedDividends, totalCurrencyCode, 2))
            + toDataCell(formatPercent(totalReturnPct, 2))
            + "</tr>\n"
        );

        writer.write("</table>\n\n");
    }

    private static void writeSaleTradesTablesHtml(FileWriter writer) throws IOException {
        ArrayList<Security> soldSecurities = getSortedSoldSecurities();

        for (Security security : soldSecurities) {
            String currencyCode = security.getCurrencyCode();
            writer.write("<h2>SALE TRADES - " + escapeHtml(getPreferredSecurityName(security)) + "</h2>\n");
            writer.write("<table class=\"sale-trades-table\">\n");
            writer.write("<colgroup>\n");
            writer.write("<col style=\"width:12%\">\n");
            writer.write("<col style=\"width:12%\">\n");
            writer.write("<col style=\"width:13%\">\n");
            writer.write("<col style=\"width:18%\">\n");
            writer.write("<col style=\"width:18%\">\n");
            writer.write("<col style=\"width:17%\">\n");
            writer.write("<col style=\"width:10%\">\n");
            writer.write("</colgroup>\n");
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
                    formatMoney(trade.getUnitPrice(), currencyCode, 2),
                    formatMoney(trade.getSaleValue(), currencyCode, 2),
                    formatMoney(trade.getCostBasis(), currencyCode, 2),
                    formatMoney(trade.getGainLoss(), currencyCode, 2),
                    formatPercent(trade.getReturnPct(), 2)
                );
            }

            double totalReturnPct = totalCostBasis > 0 ? (totalGain / totalCostBasis) * 100.0 : (totalGain > 0 ? 100.0 : 0.0);
                writer.write("<tr class=\"total-row\">"
                    + toDataCell("TOTAL")
                    + toDataCell(formatUnits(totalUnits))
                    + toDataCell("")
                    + toDataCell(formatMoney(totalSalesValue, currencyCode, 2))
                    + toDataCell(formatMoney(totalCostBasis, currencyCode, 2))
                    + toDataCell(formatMoney(totalGain, currencyCode, 2))
                    + toDataCell(formatPercent(totalReturnPct, 2))
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
