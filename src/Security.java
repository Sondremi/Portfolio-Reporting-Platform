import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.ArrayDeque;
import java.util.ArrayList;
import java.util.Comparator;
import java.text.DecimalFormat;
import java.text.DecimalFormatSymbols;
import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.net.URI;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLEncoder;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.Locale;

public class Security {
    private static final DateTimeFormatter CSV_DATE_OUTPUT = DateTimeFormatter.ofPattern("dd.MM.yyyy");
    private static final double EPSILON = 0.0000001;

    private static class SearchCandidate {
        final String symbol;
        final String exchange;
        final String quoteType;
        final String currency;
        final String shortName;
        final String longName;
        final double regularMarketPrice;

        SearchCandidate(String symbol, String exchange, String quoteType, String currency,
                        String shortName, String longName, double regularMarketPrice) {
            this.symbol = symbol;
            this.exchange = exchange;
            this.quoteType = quoteType;
            this.currency = currency;
            this.shortName = shortName;
            this.longName = longName;
            this.regularMarketPrice = regularMarketPrice;
        }
    }

    public enum AssetType {
        STOCK,
        FUND,
        UNKNOWN
    }

    private final String name;
    private final String isin;
    private String resolvedSecurityName = "";
    private String ticker = "";
    private AssetType assetType = AssetType.UNKNOWN;

    private final ArrayDeque<BuyLot> buyLots = new ArrayDeque<>();
    private final ArrayList<SaleTrade> saleTrades = new ArrayList<>();

    private double unitsOwned = 0.0;
    private double dividends = 0.0;
    private double realizedGain = 0.0;
    private double realizedCostBasis = 0.0;
    private double realizedSalesValue = 0.0;
    private double latestPrice = 0.0;
    private String currencyCode = "NOK";
    private LocalDate firstHoldingDate = null;

    private static class BuyLot {
        double remainingUnits;
        double unitCost;

        BuyLot(double remainingUnits, double unitCost) {
            this.remainingUnits = remainingUnits;
            this.unitCost = unitCost;
        }
    }

    public static class SaleTrade {
        private final LocalDate tradeDate;
        private final double units;
        private final double unitPrice;
        private final double saleValue;
        private final double costBasis;
        private final double gainLoss;
        private final double returnPct;

        SaleTrade(LocalDate tradeDate, double units, double unitPrice, double saleValue,
                  double costBasis, double gainLoss, double returnPct) {
            this.tradeDate = tradeDate;
            this.units = units;
            this.unitPrice = unitPrice;
            this.saleValue = saleValue;
            this.costBasis = costBasis;
            this.gainLoss = gainLoss;
            this.returnPct = returnPct;
        }

        public LocalDate getTradeDate() { return tradeDate; }
        public String getTradeDateAsCsv() { return tradeDate.equals(LocalDate.MIN) ? "" : tradeDate.format(CSV_DATE_OUTPUT); }
        public double getUnits() { return units; }
        public double getUnitPrice() { return unitPrice; }
        public double getSaleValue() { return saleValue; }
        public double getCostBasis() { return costBasis; }
        public double getGainLoss() { return gainLoss; }
        public double getReturnPct() { return returnPct; }
    }

    public Security(String n, String i) {
        name = n;
        isin = i;
        setTicker();
    }

    public String getName() { return name; }
    public String getDisplayName() {
        return resolvedSecurityName == null || resolvedSecurityName.isBlank() ? name : resolvedSecurityName;
    }
    public String getTicker() { return ticker; }
    public String getIsin() { return isin; }
    public AssetType getAssetType() { return assetType; }
    public String getCurrencyCode() { return currencyCode == null || currencyCode.isBlank() ? "NOK" : currencyCode; }

    public String getAverageCostAsText() { return formatNumber(getAverageCost(), 2); }
    public String getUnitsOwnedAsText() { return formatNumber(unitsOwned, 2); }

    public String getDividendsAsText() { return formatNumber(dividends, 2); }
    public String getLatestPriceAsText() {
        if (latestPrice <= EPSILON) {
            return "-";
        }
        return formatNumber(latestPrice, 2);
    }

    public String getRealizedGainAsText() { return formatNumber(realizedGain, 2); }

    public String getRealizedReturnPctAsText() {
        if (Math.abs(realizedCostBasis) < EPSILON) {
            return formatNumber(0.0, 2);
        }
        double percent = (realizedGain / realizedCostBasis) * 100.0;
        return formatNumber(percent, 2);
    }

    private static String formatNumber(double value, int decimals) {
        DecimalFormatSymbols symbols = DecimalFormatSymbols.getInstance(Locale.US);
        symbols.setGroupingSeparator(' ');
        symbols.setDecimalSeparator('.');

        DecimalFormat format = new DecimalFormat();
        format.setDecimalFormatSymbols(symbols);
        format.setGroupingUsed(true);
        format.setMinimumFractionDigits(decimals);
        format.setMaximumFractionDigits(decimals);
        return format.format(value);
    }

    public double getRealizedSalesValue() { return realizedSalesValue; }
    public double getRealizedCostBasis() { return realizedCostBasis; }
    public double getRealizedGain() { return realizedGain; }
    public double getUnitsOwned() { return unitsOwned; }
    public double getDividends() { return dividends; }
    public double getLatestPrice() { return latestPrice; }
    public LocalDate getFirstHoldingDate() { return firstHoldingDate; }

    public boolean isFullyRealized() {
        return Math.abs(unitsOwned) < EPSILON;
    }

    public boolean hasSales() {
        return !saleTrades.isEmpty();
    }

    public ArrayList<SaleTrade> getSaleTradesSortedByDate() {
        ArrayList<SaleTrade> sorted = new ArrayList<>(saleTrades);
        sorted.sort(Comparator.comparing(SaleTrade::getTradeDate));
        return sorted;
    }

    public double getTotalSoldUnits() {
        return saleTrades.stream().mapToDouble(SaleTrade::getUnits).sum();
    }

    public double getAverageCost() {
        if (Math.abs(unitsOwned) < EPSILON) {
            return 0.0;
        }

        double remainingCost = 0.0;
        for (BuyLot lot : buyLots) {
            remainingCost += lot.remainingUnits * lot.unitCost;
        }
        return remainingCost / unitsOwned;
    }

    public void addDividend(double amount) {
        dividends += amount;
    }

    public void addTransaction(String tradeDateText, String transactionType, double amount,
                               double quantity, double price, double reportedResult,
                               double totalFees) {
        double units = Math.abs(quantity);
        if (units < EPSILON) {
            return;
        }

        boolean isBuy = isBuyTransaction(transactionType, amount);
        LocalDate tradeDate = parseDate(tradeDateText);

        if (isBuy) {
            if (isReinvestTransaction(transactionType)) {
                // Reinvested dividends add units but do not represent new external capital.
                registerBuy(units, 0.0, 0.0, 0.0, tradeDate);
            } else {
                registerBuy(units, amount, price, totalFees, tradeDate);
            }
        } else {
            registerSale(tradeDate, units, amount, price, totalFees, reportedResult);
        }
    }

    public void addZeroCostUnits(double quantity) {
        addZeroCostUnits(quantity, null);
    }

    public void addZeroCostUnits(double quantity, String tradeDateText) {
        double units = Math.abs(quantity);
        if (units < EPSILON) {
            return;
        }

        LocalDate tradeDate = parseDate(tradeDateText);
        registerBuy(units, 0.0, 0.0, 0.0, tradeDate);
    }

    public void applyCostRefund(double refundAmount) {
        if (refundAmount <= EPSILON || unitsOwned <= EPSILON || buyLots.isEmpty()) {
            return;
        }

        double totalCost = 0.0;
        for (BuyLot lot : buyLots) {
            totalCost += lot.remainingUnits * lot.unitCost;
        }

        if (totalCost <= EPSILON) {
            return;
        }

        double adjustedTotalCost = Math.max(0.0, totalCost - refundAmount);
        double costScale = adjustedTotalCost / totalCost;
        for (BuyLot lot : buyLots) {
            lot.unitCost *= costScale;
        }
    }

    public void reconcileUnitsFromCorporateAction(double targetUnitsRaw) {
        double targetUnits = Math.abs(targetUnitsRaw);
        if (targetUnits <= EPSILON || unitsOwned <= EPSILON) {
            return;
        }

        if (Math.abs(unitsOwned - targetUnits) <= EPSILON) {
            return;
        }

        double scale = targetUnits / unitsOwned;
        if (scale <= EPSILON) {
            return;
        }

        // Scale units while preserving total cost basis for each lot.
        for (BuyLot lot : buyLots) {
            if (lot.remainingUnits <= EPSILON) {
                continue;
            }

            lot.remainingUnits *= scale;
            if (lot.unitCost > EPSILON) {
                lot.unitCost /= scale;
            }
        }

        unitsOwned = targetUnits;
    }

    public void updateAssetTypeFromHint(String securityTypeHint, String securityNameHint) {
        String fromType = securityTypeHint == null ? "" : securityTypeHint.trim().toUpperCase();

        if (fromType.contains("FOND") || fromType.contains("MUTUAL")) {
            assetType = AssetType.FUND;
            return;
        }

        if (fromType.contains("AKSJE") || fromType.contains("STOCK") || fromType.contains("ETF") || fromType.contains("EQUITY")) {
            assetType = AssetType.STOCK;
            return;
        }

        if (assetType != AssetType.UNKNOWN) {
            return;
        }
    }

    private boolean isBuyTransaction(String transactionType, double amount) {
        String normalized = transactionType == null ? "" : transactionType.toUpperCase();
        if (normalized.contains("KJØP") || normalized.contains("KJOP") || normalized.contains("BUY")
                || normalized.contains("REINVEST")) {
            return true;
        }
        if (normalized.contains("SALG") || normalized.contains("SELL")) {
            return false;
        }
        return amount < 0;
    }

    private boolean isReinvestTransaction(String transactionType) {
        if (transactionType == null || transactionType.isBlank()) {
            return false;
        }

        String normalized = transactionType.toUpperCase(Locale.ROOT);
        return normalized.contains("REINVEST");
    }

    private void registerBuy(double units, double amount, double price, double totalFees, LocalDate tradeDate) {
        double cashOut = Math.abs(amount);
        if (cashOut < EPSILON && price > 0) {
            cashOut = units * price + Math.max(totalFees, 0.0);
        }

        double unitCost = cashOut / units;
        buyLots.addLast(new BuyLot(units, unitCost));
        unitsOwned += units;

        if (tradeDate != null && !tradeDate.equals(LocalDate.MIN)
                && (firstHoldingDate == null || tradeDate.isBefore(firstHoldingDate))) {
            firstHoldingDate = tradeDate;
        }
    }

    private void registerSale(LocalDate tradeDate, double units, double amount, double price, double totalFees, double reportedResult) {
        double saleValue = Math.abs(amount);
        if (saleValue < EPSILON && price > 0) {
            saleValue = units * price - Math.max(totalFees, 0.0);
        }

        double availableUnitsBeforeSale = buyLots.stream().mapToDouble(lot -> lot.remainingUnits).sum();
        boolean fifoCoveredSale = availableUnitsBeforeSale + EPSILON >= units;

        double costBasis = consumeLotsUsingFifo(units);
        double gainLoss = saleValue - costBasis;
        boolean zeroCostFifoSale = fifoCoveredSale && costBasis <= EPSILON;

        // Broker-provided sale result is usually the most accurate source across partial histories.
        if (Math.abs(reportedResult) > EPSILON && !zeroCostFifoSale) {
            gainLoss = reportedResult;
            costBasis = saleValue - gainLoss;
        }

        // If we cannot resolve cost basis from FIFO and broker did not provide result,
        // keep gain/loss neutral instead of treating full sale value as profit.
        // If FIFO fully covered the sale, a zero cost basis can be valid (e.g. corporate action allotment).
        if (Math.abs(reportedResult) <= EPSILON && costBasis <= EPSILON && saleValue > EPSILON && !fifoCoveredSale) {
            costBasis = saleValue;
            gainLoss = 0.0;
        }

        if (costBasis < 0) {
            costBasis = 0.0;
        }

        double returnPct = costBasis > EPSILON ? (gainLoss / costBasis) * 100.0 : (gainLoss > EPSILON ? 100.0 : 0.0);

        saleTrades.add(new SaleTrade(tradeDate, units, price, saleValue, costBasis, gainLoss, returnPct));

        unitsOwned -= units;
        if (Math.abs(unitsOwned) < EPSILON) {
            unitsOwned = 0.0;
        }

        realizedSalesValue += saleValue;
        realizedCostBasis += costBasis;
        realizedGain += gainLoss;
    }

    private double consumeLotsUsingFifo(double unitsToSell) {
        double remainingUnitsToSell = unitsToSell;
        double costBasis = 0.0;

        while (remainingUnitsToSell > EPSILON && !buyLots.isEmpty()) {
            BuyLot oldestBuy = buyLots.peekFirst();
            if (oldestBuy == null) {
                break;
            }

            double unitsFromLot = Math.min(remainingUnitsToSell, oldestBuy.remainingUnits);
            costBasis += unitsFromLot * oldestBuy.unitCost;
            oldestBuy.remainingUnits -= unitsFromLot;
            remainingUnitsToSell -= unitsFromLot;

            if (oldestBuy.remainingUnits <= EPSILON) {
                buyLots.removeFirst();
            }
        }

        return costBasis;
    }

    private LocalDate parseDate(String tradeDateText) {
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
                // Try the next known format.
            }
        }
        return LocalDate.MIN;
    }

    private void setTicker() {
        if (isin == null || isin.isEmpty()) {
            ticker = "";
            latestPrice = 0.0;
            currencyCode = "NOK";
            return;
        }

        try {
            String url = "https://query2.finance.yahoo.com/v1/finance/search?q=" +
                         URLEncoder.encode(isin, "UTF-8") + "&quotesCount=10";

            String response = httpGetRequest(url);

            ArrayList<SearchCandidate> candidates = extractSearchCandidates(response);
            SearchCandidate candidate = chooseBestCandidate(candidates);
            candidate = refineCandidateBySymbolListings(candidate);
            candidate = preferOsloListingWhenAvailable(candidate);

            if (candidate != null && candidate.symbol != null && !candidate.symbol.isBlank()) {
                String candidateTicker = candidate.symbol;
                if (!candidateTicker.contains(".") && candidate.exchange != null && !candidate.exchange.isBlank()) {
                    String exchangeSuffix = getExchangeSuffix(candidate.exchange);
                    candidateTicker = candidateTicker + (exchangeSuffix.isEmpty() ? "" : "." + exchangeSuffix);
                }

                ticker = candidateTicker;
                currencyCode = normalizeCurrencyCode(candidate.currency, candidateTicker);
                updateAssetTypeFromQuoteType(candidate.quoteType);
                updateAssetTypeFromTicker(ticker);
                updateResolvedName(candidate);
                latestPrice = candidate.regularMarketPrice > EPSILON
                        ? candidate.regularMarketPrice
                        : fetchLatestPrice(ticker);
            } else {
                ticker = "";
                latestPrice = 0.0;
                currencyCode = "NOK";
            }
        } catch (Exception e) {
            System.err.println("Yahoo Finance ISIN lookup failed: " + e.getMessage());
            ticker = "";
            latestPrice = 0.0;
            currencyCode = "NOK";
        }
    }

    private String normalizeCurrencyCode(String candidateCurrency, String symbol) {
        if (candidateCurrency != null && !candidateCurrency.isBlank()) {
            return candidateCurrency.trim().toUpperCase(Locale.ROOT);
        }

        String inferred = inferCurrencyFromSymbol(symbol);
        if (!inferred.isBlank()) {
            return inferred;
        }

        return "NOK";
    }

    private String inferCurrencyFromSymbol(String symbol) {
        if (symbol == null || symbol.isBlank()) {
            return "";
        }

        String normalized = symbol.toUpperCase(Locale.ROOT);
        if (normalized.endsWith(".OL") || normalized.endsWith(".IR")) {
            return "NOK";
        }
        if (normalized.endsWith(".L")) {
            return "GBP";
        }
        if (normalized.endsWith(".PA") || normalized.endsWith(".DE")) {
            return "EUR";
        }
        if (normalized.endsWith(".TO")) {
            return "CAD";
        }
        if (normalized.endsWith(".HK")) {
            return "HKD";
        }
        if (normalized.endsWith(".AX")) {
            return "AUD";
        }
        if (normalized.endsWith(".T")) {
            return "JPY";
        }

        return "";
    }

    private SearchCandidate preferOsloListingWhenAvailable(SearchCandidate baseCandidate) {
        if (baseCandidate == null || baseCandidate.symbol == null || baseCandidate.symbol.isBlank()) {
            return baseCandidate;
        }

        String quoteType = baseCandidate.quoteType == null ? "" : baseCandidate.quoteType.toUpperCase(Locale.ROOT);
        if (!(quoteType.contains("EQUITY") || quoteType.contains("STOCK") || assetType == AssetType.STOCK)) {
            return baseCandidate;
        }

        if (baseCandidate.symbol.contains(".") || baseCandidate.symbol.contains("=")) {
            return baseCandidate;
        }

        String osloSymbol = baseCandidate.symbol + ".OL";
        double osloPrice = fetchLatestPrice(osloSymbol);
        if (osloPrice <= EPSILON) {
            return baseCandidate;
        }

        return new SearchCandidate(
            osloSymbol,
            "oslo",
            baseCandidate.quoteType,
            "NOK",
            baseCandidate.shortName,
            baseCandidate.longName,
            osloPrice
        );
    }

    private void updateResolvedName(SearchCandidate candidate) {
        if (candidate == null) {
            return;
        }

        String preferredName = candidate.longName;
        if (preferredName == null || preferredName.isBlank()) {
            preferredName = candidate.shortName;
        }
        if (preferredName == null || preferredName.isBlank()) {
            return;
        }

        String normalized = preferredName.trim();
        if (normalized.isBlank()) {
            return;
        }

        if (candidate.symbol != null && normalized.equalsIgnoreCase(candidate.symbol)) {
            return;
        }
        if (normalized.equalsIgnoreCase(name)) {
            return;
        }

        // Favor exchange-derived names for stock-like instruments where CSV often has only symbol.
        if (assetType == AssetType.STOCK || assetType == AssetType.UNKNOWN) {
            resolvedSecurityName = normalized;
        }
    }

    private SearchCandidate refineCandidateBySymbolListings(SearchCandidate baseCandidate) {
        if (baseCandidate == null || baseCandidate.symbol == null || baseCandidate.symbol.isBlank()) {
            return baseCandidate;
        }

        String baseSymbol = baseCandidate.symbol.split("\\.")[0].trim();
        if (baseSymbol.isBlank() || baseSymbol.contains("=")) {
            return baseCandidate;
        }

        try {
            String url = "https://query2.finance.yahoo.com/v1/finance/search?q="
                    + URLEncoder.encode(baseSymbol, "UTF-8")
                    + "&quotesCount=20";

            String response = httpGetRequest(url);
            ArrayList<SearchCandidate> listingCandidates = extractSearchCandidates(response);
            if (listingCandidates.isEmpty()) {
                return baseCandidate;
            }

            ArrayList<SearchCandidate> symbolMatches = new ArrayList<>();
            String baseUpper = baseSymbol.toUpperCase(Locale.ROOT);
            for (SearchCandidate candidate : listingCandidates) {
                if (candidate == null || candidate.symbol == null || candidate.symbol.isBlank()) {
                    continue;
                }

                String candidateSymbol = candidate.symbol.toUpperCase(Locale.ROOT);
                if (candidateSymbol.equals(baseUpper) || candidateSymbol.startsWith(baseUpper + ".")) {
                    symbolMatches.add(candidate);
                }
            }

            if (symbolMatches.isEmpty()) {
                return baseCandidate;
            }

            SearchCandidate refinedCandidate = chooseBestCandidate(symbolMatches);
            if (refinedCandidate == null) {
                return baseCandidate;
            }

            return scoreCandidate(refinedCandidate) > scoreCandidate(baseCandidate)
                    ? refinedCandidate
                    : baseCandidate;
        } catch (Exception ignored) {
            return baseCandidate;
        }
    }

    private void updateAssetTypeFromQuoteType(String quoteType) {
        if (quoteType == null || quoteType.isBlank()) {
            return;
        }

        String normalized = quoteType.toUpperCase(Locale.ROOT);
        if (normalized.contains("ETF") || normalized.contains("MUTUAL") || normalized.contains("FUND")) {
            if (assetType == AssetType.UNKNOWN) {
                assetType = AssetType.FUND;
            }
            return;
        }

        if (normalized.contains("EQUITY") || normalized.contains("STOCK")) {
            if (assetType == AssetType.UNKNOWN) {
                assetType = AssetType.STOCK;
            }
        }
    }

    private void updateAssetTypeFromTicker(String tickerValue) {
        if (tickerValue == null || tickerValue.isBlank() || assetType != AssetType.UNKNOWN) {
            return;
        }

        String normalized = tickerValue.toUpperCase(Locale.ROOT);
        if (normalized.endsWith(".IR") || normalized.startsWith("0P")) {
            assetType = AssetType.FUND;
            return;
        }

        assetType = AssetType.STOCK;
    }

    private ArrayList<SearchCandidate> extractSearchCandidates(String response) {
        ArrayList<SearchCandidate> candidates = new ArrayList<>();
        if (response == null || response.isBlank()) {
            return candidates;
        }

        String marker = "\"quotes\":";
        int quotesIndex = response.indexOf(marker);
        if (quotesIndex < 0) {
            return candidates;
        }

        int arrayStart = response.indexOf('[', quotesIndex + marker.length());
        if (arrayStart < 0) {
            return candidates;
        }

        int depth = 0;
        int objectStart = -1;
        boolean inString = false;
        boolean escaped = false;

        for (int i = arrayStart + 1; i < response.length(); i++) {
            char ch = response.charAt(i);

            if (inString) {
                if (escaped) {
                    escaped = false;
                } else if (ch == '\\') {
                    escaped = true;
                } else if (ch == '"') {
                    inString = false;
                }
                continue;
            }

            if (ch == '"') {
                inString = true;
                continue;
            }

            if (ch == '{') {
                if (depth == 0) {
                    objectStart = i;
                }
                depth++;
                continue;
            }

            if (ch == '}') {
                if (depth > 0) {
                    depth--;
                }
                if (depth == 0 && objectStart >= 0) {
                    String object = response.substring(objectStart, i + 1);
                    SearchCandidate candidate = parseSearchCandidate(object);
                    if (candidate != null) {
                        candidates.add(candidate);
                    }
                    objectStart = -1;
                }
                continue;
            }

            if (ch == ']' && depth == 0) {
                break;
            }
        }

        return candidates;
    }

    private SearchCandidate parseSearchCandidate(String quoteObject) {
        String symbol = extractValue(quoteObject, "symbol");
        if (symbol == null || symbol.isBlank()) {
            return null;
        }

        String exchange = extractValue(quoteObject, "exchange");
        String quoteType = extractValue(quoteObject, "quoteType");
        String currency = extractValue(quoteObject, "currency");
        String shortName = extractValue(quoteObject, "shortname");
        String longName = extractValue(quoteObject, "longname");
        double regularMarketPrice = extractNumericValue(quoteObject, "regularMarketPrice");

        return new SearchCandidate(symbol, exchange, quoteType, currency, shortName, longName, regularMarketPrice);
    }

    private SearchCandidate chooseBestCandidate(ArrayList<SearchCandidate> candidates) {
        SearchCandidate bestCandidate = null;
        int bestScore = Integer.MIN_VALUE;

        for (SearchCandidate candidate : candidates) {
            int score = scoreCandidate(candidate);
            if (score > bestScore) {
                bestScore = score;
                bestCandidate = candidate;
            }
        }

        return bestCandidate;
    }

    private int scoreCandidate(SearchCandidate candidate) {
        if (candidate == null || candidate.symbol == null || candidate.symbol.isBlank()) {
            return Integer.MIN_VALUE;
        }

        int score = 0;
        String quoteType = candidate.quoteType == null ? "" : candidate.quoteType.toUpperCase(Locale.ROOT);
        String currency = candidate.currency == null ? "" : candidate.currency.toUpperCase(Locale.ROOT);
        String symbol = candidate.symbol.toUpperCase(Locale.ROOT);
        String exchangeSuffix = getExchangeSuffix(candidate.exchange == null ? "" : candidate.exchange);

        if (symbol.contains("=")) {
            score -= 100;
        }

        if (assetType == AssetType.FUND) {
            if (quoteType.contains("ETF") || quoteType.contains("MUTUAL") || quoteType.contains("FUND")) {
                score += 60;
            }
            if (quoteType.contains("EQUITY") || quoteType.contains("STOCK")) {
                score -= 20;
            }
        } else if (assetType == AssetType.STOCK) {
            if (quoteType.contains("EQUITY") || quoteType.contains("STOCK")) {
                score += 40;
            }
            if (quoteType.contains("ETF") || quoteType.contains("MUTUAL") || quoteType.contains("FUND")) {
                score -= 25;
            }
        } else {
            if (quoteType.contains("EQUITY") || quoteType.contains("STOCK")) {
                score += 20;
            }
            if (quoteType.contains("ETF") || quoteType.contains("MUTUAL") || quoteType.contains("FUND")) {
                score += 15;
            }
        }

        if ("NOK".equals(currency)) {
            score += 30;
        }

        if ("OL".equals(exchangeSuffix)) {
            score += 25;
        }

        if (symbol.endsWith(".OL")) {
            score += 20;
        }

        if (candidate.regularMarketPrice > EPSILON) {
            score += 5;
        }

        return score;
    }

    private double fetchLatestPrice(String symbol) {
        if (symbol == null || symbol.isBlank()) {
            return 0.0;
        }

        try {
            String quoteUrl = "https://query1.finance.yahoo.com/v8/finance/chart/"
                    + URLEncoder.encode(symbol, "UTF-8")
                    + "?interval=1d&range=1d";
            String quoteResponse = httpGetRequest(quoteUrl);
            double regularMarketPrice = extractNumericValue(quoteResponse, "regularMarketPrice");
            if (regularMarketPrice > EPSILON) {
                return regularMarketPrice;
            }

            return extractNumericValue(quoteResponse, "chartPreviousClose");
        } catch (Exception ignored) {
            return 0.0;
        }
    }

    private String getExchangeSuffix(String exchangeName) {
        return switch (exchangeName.toLowerCase()) {
            case "oslo" -> "OL";
            case "new york stock exchange", "nyse" -> "NYSE";
            case "nasdaq" -> "";
            case "london" -> "L";
            case "xetra", "frankfurt" -> "DE";
            case "paris" -> "PA";
            case "tokyo" -> "T";
            case "hong kong" -> "HK";
            case "sydney" -> "AX";
            case "toronto" -> "TO";
            default -> "";
        };
    }

    private String extractValue(String json, String key) {
        Pattern pattern = Pattern.compile("\\\"" + key + "\\\":\\\"?(.*?)(\\\"|,|})");
        Matcher matcher = pattern.matcher(json);
        if (matcher.find()) {
            String value = matcher.group(1);
            return value.equals("null") ? null : value;
        }
        return null;
    }

    private double extractNumericValue(String json, String key) {
        Pattern pattern = Pattern.compile("\\\"" + key + "\\\":(-?\\d+(?:\\.\\d+)?)");
        Matcher matcher = pattern.matcher(json);
        if (matcher.find()) {
            try {
                return Double.parseDouble(matcher.group(1));
            } catch (NumberFormatException ignored) {
                return 0.0;
            }
        }
        return 0.0;
    }

    private String httpGetRequest(String urlString) throws Exception {
        URL url = URI.create(urlString).toURL();
        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        conn.setRequestMethod("GET");
        conn.setRequestProperty("User-Agent", "Mozilla/5.0");
        conn.setConnectTimeout(5000);
        conn.setReadTimeout(5000);

        StringBuilder response = new StringBuilder();
        try (BufferedReader reader = new BufferedReader(
                new InputStreamReader(conn.getInputStream()))) {
            String line;
            while ((line = reader.readLine()) != null) {
                response.append(line);
            }
        }
        return response.toString();
    }
}