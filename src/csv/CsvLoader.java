package csv;

import model.Events;
import model.Security;
import util.DateParser;

import java.io.*;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.time.LocalDate;
import java.util.*;
import java.util.Locale;

public class CsvLoader {

    public static int readAllTransactionFiles(String inputDirectory, TransactionStore store) throws IOException {
        File directory = new File(inputDirectory);
        if (!directory.exists() || !directory.isDirectory()) {
            return 0;
        }

        File[] csvFiles = directory.listFiles((dir, name) ->
                name.toLowerCase(Locale.ROOT).endsWith(".csv") &&
                        !name.toLowerCase(Locale.ROOT).contains("example") &&
                        !name.equalsIgnoreCase("portfolio-report.html"));

        if (csvFiles == null || csvFiles.length == 0) {
            return 0;
        }

        Arrays.sort(csvFiles, Comparator.comparing(File::getName, String.CASE_INSENSITIVE_ORDER));

        int processedCount = 0;
        for (File csvFile : csvFiles) {
            if (readFile(csvFile, store)) {
                processedCount++;
            }
        }
        store.setLoadedCsvFileCount(processedCount);
        return processedCount;
    }

    private static boolean readFile(File csvFile, TransactionStore store) throws IOException {
        int rowsRead = 0;
        Charset charset = detectCharset(csvFile);
        try (BufferedReader reader = new BufferedReader(new InputStreamReader(new FileInputStream(csvFile), charset))) {
            String headerLine = null;
            while ((headerLine = reader.readLine()) != null) {
                headerLine = headerLine.trim();
                if (!headerLine.isEmpty()) {
                    break;
                }
            }

            if (headerLine == null || headerLine.isBlank()) {
                return false;
            }

            char delimiter = detectDelimiter(headerLine);
            HeaderIndexes indexes = findHeaderIndexes(parseDelimitedLine(headerLine, delimiter));
            if (!indexes.hasRequiredColumns()) {
                return false;
            }

            ArrayList<ArrayList<String>> rows = new ArrayList<>();
            String line;
            while ((line = reader.readLine()) != null) {
                line = line.trim();
                if (line.isEmpty()) {
                    continue;
                }
                rows.add(parseDelimitedLine(line.replace("−", "-"), delimiter));
            }

            if (indexes.tradeDate >= 0) {
                rows.sort(Comparator
                        .comparing((ArrayList<String> row) -> DateParser.parseTradeDate(getCell(row, indexes.tradeDate)))
                        .thenComparingLong(row -> parseSortIdOrDefault(getCell(row, indexes.transactionId))));
            }

            inferRenamedSecurityMappings(rows, indexes, store);

            for (ArrayList<String> row : rows) {
                if (processRow(row, indexes, store)) {
                    rowsRead++;
                }
            }
        }
        System.out.println("Loaded " + rowsRead + " transaction row(s) from '" + csvFile.getName() + "'.");
        return true;
    }

    private static Charset detectCharset(File csvFile) throws IOException {
        try (InputStream in = new FileInputStream(csvFile)) {
            byte[] sample = in.readNBytes(1024);
            if (sample.length >= 2) {
                int b0 = sample[0] & 0xFF;
                int b1 = sample[1] & 0xFF;
                if (b0 == 0xFF && b1 == 0xFE) {
                    return StandardCharsets.UTF_16LE;
                }
                if (b0 == 0xFE && b1 == 0xFF) {
                    return StandardCharsets.UTF_16BE;
                }
            }

            int zeroOnEven = 0;
            int zeroOnOdd = 0;
            for (int i = 0; i < sample.length; i++) {
                if (sample[i] == 0) {
                    if (i % 2 == 0) zeroOnEven++;
                    else zeroOnOdd++;
                }
            }

            if (zeroOnEven > 8 || zeroOnOdd > 8) {
                return zeroOnOdd > zeroOnEven ? StandardCharsets.UTF_16LE : StandardCharsets.UTF_16BE;
            }
        }

        return StandardCharsets.UTF_8;
    }

    private static boolean processRow(ArrayList<String> row, HeaderIndexes indexes, TransactionStore store) {
        if (row == null || row.isEmpty()) return false;

        String transactionType = getCell(row, indexes.transactionType)
                .toUpperCase(Locale.ROOT)
                .replaceAll("\\s+", " ")
                .trim();

        LocalDate tradeDateForTracking = indexes.tradeDate >= 0
                ? DateParser.parseTradeDate(getCell(row, indexes.tradeDate))
                : LocalDate.MIN;

        long sortId = parseSortIdOrDefault(getCell(row, indexes.transactionId));
        boolean balanceBackedRow = recordPortfolioCashSnapshot(row, indexes, tradeDateForTracking, sortId, store);
        store.incrementTransactionRowCount();

        String securityName = getCell(row, indexes.securityName);
        if (securityName.isEmpty()) {
            processStandaloneCashTransaction(row, indexes, transactionType, tradeDateForTracking, balanceBackedRow, store);
            return true;
        }

        String isin = getCell(row, indexes.isin);
        String canonicalIsin = canonicalizeIsin(isin, store);
        store.rememberCanonicalSecurityName(isin, canonicalIsin, securityName);

        String securityType = getCell(row, indexes.securityType);
        Security security = store.getOrCreateSecurity(securityName, canonicalIsin);
        if (security == null) return false;

        security.updateAssetTypeFromHint(securityType, securityName);

        processTransaction(security, row, indexes, isin, transactionType, tradeDateForTracking, balanceBackedRow, store);
        return true;
    }

    private static boolean recordPortfolioCashSnapshot(ArrayList<String> row, HeaderIndexes indexes,
                                                       LocalDate tradeDate, long sortId, TransactionStore store) {
        if (indexes.portfolioId < 0 || indexes.cashBalance < 0) return false;

        String portfolioId = getCell(row, indexes.portfolioId);
        String balanceText = getCell(row, indexes.cashBalance);
        if (portfolioId.isBlank() || balanceText.isBlank()) return false;

        Double parsedBalance = parseDoubleOrNull(balanceText);
        if (parsedBalance == null) {
            System.out.println("Warning: ignored malformed cash balance snapshot for portfolio '"
                    + portfolioId + "' on " + tradeDate + " (value='" + balanceText + "')");
            return false;
        }

        double balance = parsedBalance;
        store.addPortfolioCashSnapshot(new Events.PortfolioCashSnapshot(tradeDate, sortId, portfolioId, balance));
        return true;
    }

    private static void processStandaloneCashTransaction(ArrayList<String> row, HeaderIndexes indexes,
                                                         String transactionType, LocalDate tradeDate,
                                                         boolean balanceBackedRow, TransactionStore store) {
        if (!isCashAccountEventType(transactionType)) return;

        double amount = parseDoubleOrZero(getCell(row, indexes.amount));
        double totalFees = parseDoubleOrZero(getCell(row, indexes.totalFees));
        double cashDelta = resolveCashAccountEventDelta(transactionType, amount, totalFees);

        if (Math.abs(cashDelta) > 1e-9) {
            store.addCashEvent(new Events.CashEvent(tradeDate, cashDelta, isExternalCashFlowEventType(transactionType)));
        }
    }

    private static void processTransaction(Security security, ArrayList<String> row, HeaderIndexes indexes,
                                           String originalIsin, String transactionType,
                                           LocalDate tradeDate, boolean balanceBackedRow, TransactionStore store) {

        double quantity = parseDoubleOrZero(getCell(row, indexes.quantity));
        double amount = parseDoubleOrZero(getCell(row, indexes.amount));
        double price = parseDoubleOrZero(getCell(row, indexes.price));
        double result = parseDoubleOrZero(getCell(row, indexes.result));
        double totalFees = parseDoubleOrZero(getCell(row, indexes.totalFees));
        String transactionCurrency = resolveTransactionCurrency(row, indexes);
        security.setCurrencyCodeFromTransaction(transactionCurrency);

        if (isRenameBookkeepingTransaction(transactionType, originalIsin, store)) {
            boolean isCancelled = indexes.cancellationDate >= 0
                    && !getCell(row, indexes.cancellationDate).isBlank();
            if (!isCancelled && "BYTTE INNLEGG VP".equals(transactionType) && quantity > 0.0) {
                String tradeDateText = getCell(row, indexes.tradeDate);
                double corporateActionCostBasis = parseDoubleOrZero(getCell(row, indexes.purchaseValue));

                if (security.getUnitsOwned() <= 1e-9) {
                    if (corporateActionCostBasis > 1e-9) {
                        security.addCorporateActionUnits(quantity, tradeDateText, corporateActionCostBasis);
                    } else {
                        security.addZeroCostUnits(quantity, tradeDateText);
                    }
                } else {
                    security.reconcileUnitsFromCorporateAction(quantity);
                    if (corporateActionCostBasis > 1e-9) {
                        security.alignCurrentCostBasisToTotal(corporateActionCostBasis);
                    }
                }
            }
            return;
        }

        if (isTradeLikeTransaction(transactionType)) {
            double tradeCashDelta = resolveTradeCashDelta(transactionType, amount, quantity, price, totalFees);
            if (!balanceBackedRow && Math.abs(tradeCashDelta) > 1e-9) {
                boolean sourceTracksBrokerCash = indexes.cashBalance >= 0;
                if (sourceTracksBrokerCash) {
                    // Trade settles against tracked broker cash account.
                    store.addCashEvent(new Events.CashEvent(tradeDate, tradeCashDelta, false));
                } else {
                    // Source has no broker cash ledger (e.g. bank-funded trade exports).
                    // Treat principal as external contribution/withdrawal instead.
                    store.addCashEvent(new Events.CashEvent(tradeDate, -tradeCashDelta, true));
                }
            }
            if (Math.abs(quantity) > 0.0) {
                double unitsDelta = isBuyLikeTransaction(transactionType, amount) ? Math.abs(quantity) : -Math.abs(quantity);
                store.addUnitEvent(new Events.UnitEvent(tradeDate, security.getIsin(), unitsDelta));
            }
            security.addTransaction(getCell(row, indexes.tradeDate), transactionType, amount, quantity, price, result, totalFees);
            return;
        }

        if (isUnitsCorporateActionTransaction(transactionType)) {
            if (Math.abs(quantity) > 0.0) {
                store.addUnitEvent(new Events.UnitEvent(tradeDate, security.getIsin(), Math.abs(quantity)));
            }
            security.addZeroCostUnits(quantity, getCell(row, indexes.tradeDate));
            return;
        }

        if (isDividendCashTransaction(transactionType)) {
            if (!balanceBackedRow && Math.abs(amount) > 1e-9) {
                store.addCashEvent(new Events.CashEvent(tradeDate, Math.abs(amount), false));
            }
            security.addDividend(Math.abs(amount), getCell(row, indexes.tradeDate), Math.abs(quantity));
            return;
        }

        if (isFundCostRefundTransaction(transactionType)) {
            double refundAmount = totalFees > 0.0 ? totalFees : Math.abs(amount);
            security.applyCostRefund(refundAmount);
            if (!balanceBackedRow && Math.abs(refundAmount) > 0.0) {
                store.addCashEvent(new Events.CashEvent(tradeDate, refundAmount, false));
            }
            return;
        }

        if (isCashAccountEventType(transactionType)) {
            double cashDelta = resolveCashAccountEventDelta(transactionType, amount, totalFees);
            if (Math.abs(cashDelta) > 1e-9) {
                store.addCashEvent(new Events.CashEvent(tradeDate, cashDelta, isExternalCashFlowEventType(transactionType)));
            }
        }
    }

    // ====================== Parsing Helpers ======================

    private static double parseDoubleOrZero(String value) {
        Double parsed = parseDoubleOrNull(value);
        return parsed == null ? 0.0 : parsed;
    }

    private static Double parseDoubleOrNull(String value) {
        if (value == null || value.trim().isEmpty()) return null;
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
            return null;
        }
    }

    private static long parseSortIdOrDefault(String value) {
        if (value == null || value.isBlank()) return Long.MAX_VALUE;
        try {
            return Long.parseLong(value.trim());
        } catch (NumberFormatException ex) {
            return Long.MAX_VALUE;
        }
    }

    private static ArrayList<String> parseDelimitedLine(String line, char delimiter) {
        ArrayList<String> fields = new ArrayList<>();
        StringBuilder current = new StringBuilder();
        boolean inQuotes = false;

        for (int i = 0; i < line.length(); i++) {
            char c = line.charAt(i);
            if (c == '"') {
                boolean escapedQuote = inQuotes && i + 1 < line.length() && line.charAt(i + 1) == '"';
                if (escapedQuote) {
                    current.append('"');
                    i++;
                } else {
                    inQuotes = !inQuotes;
                }
            } else if (c == delimiter && !inQuotes) {
                fields.add(current.toString());
                current.setLength(0);
            } else {
                current.append(c);
            }
        }
        fields.add(current.toString());
        return fields;
    }

    private static char detectDelimiter(String line) {
        int tabCount = countOccurrences(line, '\t');
        int semicolonCount = countOccurrences(line, ';');
        int commaCount = countOccurrences(line, ',');
        if (tabCount >= semicolonCount && tabCount >= commaCount) return '\t';
        if (semicolonCount >= tabCount && semicolonCount >= commaCount) return ';';
        return ',';
    }

    private static int countOccurrences(String text, char target) {
        int count = 0;
        for (int i = 0; i < text.length(); i++) {
            if (text.charAt(i) == target) count++;
        }
        return count;
    }

    private static String getCell(ArrayList<String> row, int index) {
        if (index < 0 || index >= row.size()) return "";
        return row.get(index).trim().replace("\uFEFF", "");
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
                case "belop", "beløp", "handelsbelop", "handelsbeløp" -> indexes.amount = i;
                case "kjopsverdi", "kjøpsverdi", "anskaffelsesverdi" -> indexes.purchaseValue = i;
                case "antall" -> indexes.quantity = i;
                case "kurs", "kurs per andel" -> indexes.price = i;
                case "resultat" -> indexes.result = i;
                case "totale avgifter", "omkostninger" -> indexes.totalFees = i;
                case "transaksjonstekst", "transaction text" -> indexes.transactionText = i;
                case "bokforingsdag", "bokføringsdag" -> {
                    if (indexes.tradeDate < 0) indexes.tradeDate = i;
                }
                case "handelsdag", "handelsdato" -> {
                    if (indexes.tradeDate < 0) indexes.tradeDate = i;
                }
                case "makuleringsdato" -> indexes.cancellationDate = i;
                case "portefolje", "portefølje", "konto" -> indexes.portfolioId = i;
                case "saldo" -> indexes.cashBalance = i;
                case "valuta" -> mapCurrencyColumn(headerColumns, i, indexes);
            }
        }
        return indexes;
    }

    private static String normalizeHeader(String value) {
        return value
                .replace("\uFEFF", "")
                .trim()
                .toLowerCase(Locale.ROOT);
    }

    private static void mapCurrencyColumn(ArrayList<String> headerColumns, int index, HeaderIndexes indexes) {
        String previous = index > 0 ? normalizeHeader(headerColumns.get(index - 1)) : "";
        if ("belop".equals(previous) || "beløp".equals(previous)
                || "handelsbelop".equals(previous) || "handelsbeløp".equals(previous)) {
            indexes.amountCurrency = index;
            return;
        }

        if ("resultat".equals(previous)) {
            indexes.resultCurrency = index;
            return;
        }

        if ("totale avgifter".equals(previous) || "omkostninger".equals(previous)) {
            indexes.feeCurrency = index;
            return;
        }

        if (indexes.transactionCurrency < 0) {
            indexes.transactionCurrency = index;
        }
    }

    private static String resolveTransactionCurrency(ArrayList<String> row, HeaderIndexes indexes) {
        String amountCurrency = getCell(row, indexes.amountCurrency);
        if (!amountCurrency.isBlank()) {
            return amountCurrency;
        }

        String transactionCurrency = getCell(row, indexes.transactionCurrency);
        if (!transactionCurrency.isBlank()) {
            return transactionCurrency;
        }

        String resultCurrency = getCell(row, indexes.resultCurrency);
        if (!resultCurrency.isBlank()) {
            return resultCurrency;
        }

        return getCell(row, indexes.feeCurrency);
    }

    private static String canonicalizeIsin(String isin, TransactionStore store) {
        if (isin == null || isin.isBlank()) return isin;
        String normalized = isin.trim().toUpperCase(Locale.ROOT);
        String current = normalized;
        HashSet<String> seen = new HashSet<>();

        while (!current.isBlank() && !seen.contains(current)) {
            seen.add(current);
            String next = store.getRenamedSecurityIsin().get(current);
            if (next == null || next.isBlank() || next.equals(current)) {
                break;
            }
            current = next;
        }

        return current;
    }

    private static void inferRenamedSecurityMappings(ArrayList<ArrayList<String>> rows, HeaderIndexes indexes, TransactionStore store) {
        if (rows == null || rows.isEmpty()) {
            return;
        }
        if (indexes.transactionId < 0 || indexes.transactionType < 0 || indexes.isin < 0) {
            return;
        }

        Map<String, String> oldIsinBySwapKey = new LinkedHashMap<>();
        Map<String, String> newIsinBySwapKey = new LinkedHashMap<>();
        Map<String, LinkedHashSet<String>> oldIsinByBucket = new LinkedHashMap<>();
        Map<String, LinkedHashSet<String>> newIsinByBucket = new LinkedHashMap<>();

        for (ArrayList<String> row : rows) {
            String transactionType = normalizeTransactionType(getCell(row, indexes.transactionType));
            if (transactionType.startsWith("MAK ")) {
                continue;
            }
            if (!"BYTTE UTTAK VP".equals(transactionType) && !"BYTTE INNLEGG VP".equals(transactionType)) {
                continue;
            }

            String isin = getCell(row, indexes.isin);
            if (isin.isBlank()) {
                continue;
            }
            String normalizedIsin = isin.trim().toUpperCase(Locale.ROOT);
            String swapKey = buildSwapInferenceKey(row, indexes);
            String bucketKey = buildSwapInferenceBucketKey(row, indexes);

            if ("BYTTE UTTAK VP".equals(transactionType)) {
                oldIsinBySwapKey.putIfAbsent(swapKey, normalizedIsin);
                oldIsinByBucket.computeIfAbsent(bucketKey, ignored -> new LinkedHashSet<>()).add(normalizedIsin);
            } else if ("BYTTE INNLEGG VP".equals(transactionType)) {
                newIsinBySwapKey.putIfAbsent(swapKey, normalizedIsin);
                newIsinByBucket.computeIfAbsent(bucketKey, ignored -> new LinkedHashSet<>()).add(normalizedIsin);
            }
        }

        for (Map.Entry<String, String> entry : oldIsinBySwapKey.entrySet()) {
            String oldIsin = entry.getValue();
            String newIsin = newIsinBySwapKey.get(entry.getKey());
            if (newIsin == null || newIsin.isBlank()) {
                continue;
            }
            store.rememberRenamedSecurityIsin(oldIsin, newIsin);
        }

        LinkedHashSet<String> allBucketKeys = new LinkedHashSet<>();
        allBucketKeys.addAll(oldIsinByBucket.keySet());
        allBucketKeys.addAll(newIsinByBucket.keySet());

        for (String bucketKey : allBucketKeys) {
            LinkedHashSet<String> oldSet = oldIsinByBucket.getOrDefault(bucketKey, new LinkedHashSet<>());
            LinkedHashSet<String> newSet = newIsinByBucket.getOrDefault(bucketKey, new LinkedHashSet<>());
            if (oldSet.size() != 1 || newSet.size() != 1) {
                continue;
            }
            store.rememberRenamedSecurityIsin(oldSet.iterator().next(), newSet.iterator().next());
        }
    }

    private static String buildSwapInferenceKey(ArrayList<String> row, HeaderIndexes indexes) {
        String bucket = buildSwapInferenceBucketKey(row, indexes);
        String text = normalizeFreeText(getCell(row, indexes.transactionText));
        String transactionId = getCell(row, indexes.transactionId);

        if (!text.isBlank()) {
            return bucket + "|TXT:" + text;
        }
        if (!transactionId.isBlank()) {
            return bucket + "|ID:" + transactionId;
        }
        return bucket + "|*";
    }

    private static String buildSwapInferenceBucketKey(ArrayList<String> row, HeaderIndexes indexes) {
        String tradeDateKey = "";
        if (indexes.tradeDate >= 0) {
            LocalDate parsed = DateParser.parseTradeDate(getCell(row, indexes.tradeDate));
            if (parsed != null && !parsed.equals(LocalDate.MIN)) {
                tradeDateKey = parsed.toString();
            } else {
                tradeDateKey = getCell(row, indexes.tradeDate);
            }
        }

        String portfolio = getCell(row, indexes.portfolioId).toUpperCase(Locale.ROOT);
        return tradeDateKey + "|" + portfolio;
    }

    private static String normalizeFreeText(String value) {
        if (value == null || value.isBlank()) {
            return "";
        }
        return value.toUpperCase(Locale.ROOT).replaceAll("\\s+", " ").trim();
    }

    private static String normalizeTransactionType(String transactionType) {
        if (transactionType == null || transactionType.isBlank()) {
            return "";
        }
        return transactionType.toUpperCase(Locale.ROOT).replaceAll("\\s+", " ").trim();
    }

    // ====================== Transaction Type Helpers ======================

    private static boolean isRenameBookkeepingTransaction(String transactionType, String securityIsin, TransactionStore store) {
        if (securityIsin == null || securityIsin.isBlank()) return false;
        String normalizedIsin = securityIsin.trim().toUpperCase(Locale.ROOT);
        if (!store.getRenamedSecurityIsin().containsKey(normalizedIsin) &&
            !store.getRenamedSecurityIsin().containsValue(normalizedIsin)) {
            return false;
        }
        String normalized = normalizeTransactionType(transactionType);
        return "BYTTE INNLEGG VP".equals(normalized) ||
               "BYTTE UTTAK VP".equals(normalized) ||
               "MAK BYTTE INNLEGG VP".equals(normalized) ||
               "MAK BYTTE UTTAK VP".equals(normalized);
    }

    private static boolean isTradeLikeTransaction(String transactionType) {
        if (transactionType == null) return false;
        String n = transactionType.toUpperCase(Locale.ROOT);
        return n.contains("KJØP") || n.contains("KJOP") || n.contains("BUY") ||
               n.contains("SALG") || n.contains("SELL") || n.contains("REINVEST");
    }

    private static boolean isBuyLikeTransaction(String transactionType, double amount) {
        if (transactionType == null) return amount < 0;
        String n = transactionType.toUpperCase(Locale.ROOT);
        if (n.contains("KJØP") || n.contains("KJOP") || n.contains("BUY") || n.contains("REINVEST")) return true;
        if (n.contains("SALG") || n.contains("SELL")) return false;
        return amount < 0;
    }

    private static boolean isUnitsCorporateActionTransaction(String transactionType) {
        if (transactionType == null) return false;
        String n = transactionType.toUpperCase(Locale.ROOT);
        return n.contains("UTBYTTE INNLEGG VP") || n.contains("BYTTE INNLEGG VP") ||
               n.contains("MAK BYTTE INNLEGG VP") || n.contains("TILDELING INNLEGG RE");
    }

    private static boolean isDividendCashTransaction(String transactionType) {
        if (transactionType == null) return false;
        String n = transactionType.toUpperCase(Locale.ROOT);
        return !n.contains("REINVEST") && (n.contains("UTBYTTE") || n.contains("DIVIDEND"));
    }

    private static boolean isFundCostRefundTransaction(String transactionType) {
        if (transactionType == null) return false;
        String n = transactionType.toUpperCase(Locale.ROOT);
        return n.contains("TILBAKEBET") && n.contains("FOND");
    }

    private static boolean isCashAccountEventType(String transactionType) {
        if (transactionType == null || transactionType.isBlank()) return false;
        String n = transactionType.toUpperCase(Locale.ROOT);
        return n.contains("INNSKUDD") || n.contains("INNBETAL") || n.contains("DEPOSIT") ||
               n.contains("UTTAK") || n.contains("UTBETAL") || n.contains("WITHDRAW") ||
               n.contains("OVERF") || n.contains("TRANSFER") || n.contains("TRUSTLY") ||
               n.contains("GEBYR") || n.contains("FEE") || n.contains("KOSTNAD") ||
               n.contains("RENTE") || n.contains("OVERBEL") || n.contains("PLATTFORMAVG");
    }

    private static boolean isExternalCashFlowEventType(String transactionType) {
        if (transactionType == null || transactionType.isBlank()) return false;
        String n = transactionType.toUpperCase(Locale.ROOT);
        return n.contains("INNSKUDD") || n.contains("INNBETAL") || n.contains("DEPOSIT") ||
               n.contains("UTTAK") || n.contains("UTBETAL") || n.contains("WITHDRAW") ||
               n.contains("OVERF") || n.contains("TRANSFER") || n.contains("TRUSTLY");
    }

    private static double resolveTradeCashDelta(String transactionType, double amount,
                                                double quantity, double price, double totalFees) {
        if (isBuyLikeTransaction(transactionType, amount)) {
            double buyCashOut = Math.abs(amount);
            if (buyCashOut <= 1e-9 && Math.abs(quantity) > 1e-9 && price > 0.0) {
                buyCashOut = Math.abs(quantity) * price + Math.max(totalFees, 0.0);
            }
            return -buyCashOut;
        } else {
            double sellCashIn = Math.abs(amount);
            if (sellCashIn <= 1e-9 && Math.abs(quantity) > 1e-9 && price > 0.0) {
                sellCashIn = Math.max(0.0, (Math.abs(quantity) * price) - Math.max(totalFees, 0.0));
            }
            return sellCashIn;
        }
    }

    private static double resolveCashAccountEventDelta(String transactionType, double amount, double totalFees) {
        String n = transactionType.toUpperCase(Locale.ROOT);
        if (n.contains("INNSKUDD") || n.contains("INNBETAL") || n.contains("DEPOSIT") ||
            n.contains("TILBAKEBETALING") || n.contains("REFUND")) {
            return Math.abs(amount) > 1e-9 ? Math.abs(amount) : Math.max(0.0, totalFees);
        }
        if (n.contains("UTTAK") || n.contains("UTBETAL") || n.contains("WITHDRAW")) {
            return Math.abs(amount) > 1e-9 ? -Math.abs(amount) : -Math.max(0.0, totalFees);
        }
        if (n.contains("OVERF") || n.contains("TRANSFER") || n.contains("TRUSTLY")) {
            return amount;
        }
        if (n.contains("GEBYR") || n.contains("FEE") || n.contains("KOSTNAD") || n.contains("RENTE") || n.contains("OVERBEL")) {
            return amount;
        }
        return amount;
    }
}
