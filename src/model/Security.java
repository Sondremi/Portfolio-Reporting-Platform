package model;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.ArrayDeque;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.text.DecimalFormat;
import java.text.DecimalFormatSymbols;
import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
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
    private static final long YAHOO_AUTH_TTL_MS = 10 * 60 * 1000L;
    private static final Map<String, String> FUND_QUOTE_SUMMARY_CACHE = new LinkedHashMap<>();
    private static final Map<String, LinkedHashMap<String, Double>> FUND_SECTOR_WEIGHT_CACHE = new LinkedHashMap<>();
    private static final Map<String, LinkedHashMap<String, Double>> FUND_REGION_WEIGHT_CACHE = new LinkedHashMap<>();
    private static final Map<String, String> HOLDING_SYMBOL_REGION_CACHE = new LinkedHashMap<>();

    private static String yahooAuthCookieHeader = "";
    private static String yahooAuthCrumb = "";
    private static long yahooAuthExpiresAtMs = 0L;

    private static class SearchCandidate {
        final String symbol;
        final String exchange;
        final String exchangeDisplay;
        final String quoteType;
        final String sector;
        final String industry;
        final String currency;
        final String shortName;
        final String longName;
        final double regularMarketPrice;

        SearchCandidate(String symbol, String exchange, String exchangeDisplay, String quoteType,
                        String sector, String industry, String currency,
                        String shortName, String longName, double regularMarketPrice) {
            this.symbol = symbol;
            this.exchange = exchange;
            this.exchangeDisplay = exchangeDisplay;
            this.quoteType = quoteType;
            this.sector = sector;
            this.industry = industry;
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
    private String resolvedSector = "";
    private String resolvedRegion = "";
    private final LinkedHashMap<String, Double> resolvedSectorWeights = new LinkedHashMap<>();
    private final LinkedHashMap<String, Double> resolvedRegionWeights = new LinkedHashMap<>();
    private String ticker = "";
    private AssetType assetType = AssetType.UNKNOWN;

    private final ArrayDeque<BuyLot> buyLots = new ArrayDeque<>();
    private final ArrayList<SaleTrade> saleTrades = new ArrayList<>();
    private final ArrayList<PriceObservation> priceObservations = new ArrayList<>();
    private final ArrayList<DividendEvent> currentDividendEvents = new ArrayList<>();
    private final ArrayList<DividendEvent> allDividendEvents = new ArrayList<>();
    private final ArrayList<ReinvestEvent> currentReinvestEvents = new ArrayList<>();

    private double unitsOwned = 0.0;
    private double dividends = 0.0;
    private double realizedGain = 0.0;
    private double realizedCostBasis = 0.0;
    private double realizedSalesValue = 0.0;
    private double latestPrice = 0.0;
    private double previousClose = 0.0;
    private String currencyCode = "NOK";
    private LocalDate firstHoldingDate = null;

    private static class BuyLot {
        LocalDate tradeDate;
        double remainingUnits;
        double unitCost;

        BuyLot(LocalDate tradeDate, double remainingUnits, double unitCost) {
            this.tradeDate = tradeDate;
            this.remainingUnits = remainingUnits;
            this.unitCost = unitCost;
        }
    }

    private static class PriceObservation {
        private final LocalDate tradeDate;
        private final double unitPrice;

        private PriceObservation(LocalDate tradeDate, double unitPrice) {
            this.tradeDate = tradeDate;
            this.unitPrice = unitPrice;
        }
    }

    private static final class LatestQuote {
        private final double latestPrice;
        private final double previousClose;

        private LatestQuote(double latestPrice, double previousClose) {
            this.latestPrice = latestPrice;
            this.previousClose = previousClose;
        }
    }

    private static class ReinvestEvent {
        private final LocalDate tradeDate;

        private ReinvestEvent(LocalDate tradeDate) {
            this.tradeDate = tradeDate;
        }
    }

    public static class CurrentHoldingLot {
        private final LocalDate tradeDate;
        private final double units;
        private final double unitCost;

        CurrentHoldingLot(LocalDate tradeDate, double units, double unitCost) {
            this.tradeDate = tradeDate;
            this.units = units;
            this.unitCost = unitCost;
        }

        public LocalDate getTradeDate() { return tradeDate; }
        public double getUnits() { return units; }
        public double getUnitCost() { return unitCost; }
        public double getCostBasis() { return units * unitCost; }
    }

    public static class DividendEvent {
        private final LocalDate tradeDate;
        private final double units;
        private final double amount;

        DividendEvent(LocalDate tradeDate, double units, double amount) {
            this.tradeDate = tradeDate;
            this.units = units;
            this.amount = amount;
        }

        public LocalDate getTradeDate() { return tradeDate; }
        public double getUnits() { return units; }
        public double getAmount() { return amount; }
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
    public String getResolvedSector() { return resolvedSector; }
    public String getResolvedRegion() { return resolvedRegion; }
    public Map<String, Double> getResolvedSectorWeights() { return new LinkedHashMap<>(resolvedSectorWeights); }
    public Map<String, Double> getResolvedRegionWeights() { return new LinkedHashMap<>(resolvedRegionWeights); }
    public String getTicker() { return ticker; }
    public String getIsin() { return isin; }
    public AssetType getAssetType() { return assetType; }
    public String getCurrencyCode() { return currencyCode == null || currencyCode.isBlank() ? "NOK" : currencyCode; }

    public void setCurrencyCodeFromTransaction(String transactionCurrency) {
        if (transactionCurrency == null || transactionCurrency.isBlank()) {
            return;
        }

        String normalized = transactionCurrency.trim().toUpperCase(Locale.ROOT);
        if (!normalized.matches("[A-Z]{3}")) {
            return;
        }

        currencyCode = normalized;
    }

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
    public double getPreviousClose() { return previousClose; }
    public LocalDate getFirstHoldingDate() { return firstHoldingDate; }
    public boolean hasDayChangePct() {
        return previousClose > EPSILON && latestPrice > EPSILON;
    }
    public double getDayChangePct() {
        if (!hasDayChangePct()) {
            return 0.0;
        }
        return ((latestPrice / previousClose) - 1.0) * 100.0;
    }

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

    public double getLatestObservedTradePriceOnOrBefore(LocalDate date) {
        if (date == null || priceObservations.isEmpty()) {
            return 0.0;
        }

        LocalDate bestDate = null;
        double bestPrice = 0.0;
        for (PriceObservation observation : priceObservations) {
            if (observation == null || observation.tradeDate == null || observation.tradeDate.equals(LocalDate.MIN)) {
                continue;
            }
            if (observation.unitPrice <= EPSILON || observation.tradeDate.isAfter(date)) {
                continue;
            }
            if (bestDate == null || observation.tradeDate.isAfter(bestDate)) {
                bestDate = observation.tradeDate;
                bestPrice = observation.unitPrice;
            }
        }

        return bestPrice;
    }

    public double getEarliestObservedTradePriceAfter(LocalDate date) {
        if (date == null || priceObservations.isEmpty()) {
            return 0.0;
        }

        LocalDate bestDate = null;
        double bestPrice = 0.0;
        for (PriceObservation observation : priceObservations) {
            if (observation == null || observation.tradeDate == null || observation.tradeDate.equals(LocalDate.MIN)) {
                continue;
            }
            if (observation.unitPrice <= EPSILON || !observation.tradeDate.isAfter(date)) {
                continue;
            }
            if (bestDate == null || observation.tradeDate.isBefore(bestDate)) {
                bestDate = observation.tradeDate;
                bestPrice = observation.unitPrice;
            }
        }

        return bestPrice;
    }

    public double getClosestObservedTradePriceAround(LocalDate date) {
        if (date == null || priceObservations.isEmpty()) {
            return 0.0;
        }

        long bestDistance = Long.MAX_VALUE;
        LocalDate bestDate = null;
        double bestPrice = 0.0;

        for (PriceObservation observation : priceObservations) {
            if (observation == null || observation.tradeDate == null || observation.tradeDate.equals(LocalDate.MIN)) {
                continue;
            }
            if (observation.unitPrice <= EPSILON) {
                continue;
            }

            long distance = Math.abs(java.time.temporal.ChronoUnit.DAYS.between(observation.tradeDate, date));
            if (distance < bestDistance) {
                bestDistance = distance;
                bestDate = observation.tradeDate;
                bestPrice = observation.unitPrice;
                continue;
            }

            if (distance == bestDistance && bestDate != null
                    && observation.tradeDate.isBefore(bestDate)
                    && !observation.tradeDate.isAfter(date)) {
                bestDate = observation.tradeDate;
                bestPrice = observation.unitPrice;
            }
        }

        return bestPrice;
    }

    public double getTotalSoldUnits() {
        return saleTrades.stream().mapToDouble(SaleTrade::getUnits).sum();
    }

    public double getAverageCost() {
        if (Math.abs(unitsOwned) < EPSILON) {
            return 0.0;
        }

        double adjustedCostBasis = getAdjustedCurrentCostBasis();
        return adjustedCostBasis / unitsOwned;
    }

    public void addDividend(double amount) {
        dividends += amount;
    }

    public void addDividend(double amount, String tradeDateText, double unitsAtDividend) {
        dividends += amount;
        LocalDate tradeDate = parseDate(tradeDateText);
        DividendEvent event = new DividendEvent(tradeDate, Math.max(0.0, unitsAtDividend), amount);
        currentDividendEvents.add(event);
        allDividendEvents.add(event);
    }

    public ArrayList<CurrentHoldingLot> getCurrentHoldingLotsSortedByDate() {
        ArrayList<CurrentHoldingLot> lots = new ArrayList<>();
        for (BuyLot lot : buyLots) {
            if (lot.remainingUnits <= EPSILON) {
                continue;
            }
            lots.add(new CurrentHoldingLot(lot.tradeDate, lot.remainingUnits, lot.unitCost));
        }
        lots.sort(Comparator.comparing(CurrentHoldingLot::getTradeDate));
        return lots;
    }

    public ArrayList<DividendEvent> getCurrentDividendEventsSortedByDate() {
        ArrayList<DividendEvent> events = new ArrayList<>(currentDividendEvents);
        events.sort(Comparator.comparing(DividendEvent::getTradeDate));
        return events;
    }

    public ArrayList<DividendEvent> getAllDividendEventsSortedByDate() {
        ArrayList<DividendEvent> events = new ArrayList<>(allDividendEvents);
        events.sort(Comparator.comparing(DividendEvent::getTradeDate));
        return events;
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
                currentReinvestEvents.add(new ReinvestEvent(tradeDate));
            }
            registerBuy(units, amount, price, totalFees, tradeDate);
        } else {
            registerSale(tradeDate, units, amount, price, totalFees, reportedResult);
        }
    }

    private boolean isReinvestTransaction(String transactionType) {
        if (transactionType == null || transactionType.isBlank()) {
            return false;
        }

        String normalized = transactionType.toUpperCase(Locale.ROOT);
        return normalized.contains("REINVEST");
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

    public void addCorporateActionUnits(double quantity, String tradeDateText, double totalCostBasis) {
        double units = Math.abs(quantity);
        if (units < EPSILON) {
            return;
        }

        LocalDate tradeDate = parseDate(tradeDateText);
        double safeCostBasis = Math.max(0.0, totalCostBasis);
        registerBuy(units, safeCostBasis, 0.0, 0.0, tradeDate);
    }

    public void alignCurrentCostBasisToTotal(double targetTotalCostBasis) {
        if (targetTotalCostBasis <= EPSILON || buyLots.isEmpty()) {
            return;
        }

        double currentTotalCost = 0.0;
        for (BuyLot lot : buyLots) {
            currentTotalCost += lot.remainingUnits * lot.unitCost;
        }
        if (currentTotalCost <= EPSILON) {
            return;
        }

        double scale = targetTotalCostBasis / currentTotalCost;
        if (!Double.isFinite(scale) || scale <= EPSILON) {
            return;
        }

        for (BuyLot lot : buyLots) {
            lot.unitCost *= scale;
        }
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

    private void registerBuy(double units, double amount, double price, double totalFees, LocalDate tradeDate) {
        double cashOut = Math.abs(amount);
        if (cashOut < EPSILON && price > 0) {
            cashOut = units * price + Math.max(totalFees, 0.0);
        }

        double unitCost = cashOut / units;
        buyLots.addLast(new BuyLot(tradeDate, units, unitCost));
        unitsOwned += units;

        if (tradeDate != null && !tradeDate.equals(LocalDate.MIN)
                && (firstHoldingDate == null || tradeDate.isBefore(firstHoldingDate))) {
            firstHoldingDate = tradeDate;
        }

        // Persist observed transaction prices for historical fallback valuation.
        if (price > EPSILON) {
            priceObservations.add(new PriceObservation(tradeDate, price));
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
        if (price > EPSILON) {
            priceObservations.add(new PriceObservation(tradeDate, price));
        }

        unitsOwned -= units;
        if (Math.abs(unitsOwned) < EPSILON) {
            unitsOwned = 0.0;
            currentDividendEvents.clear();
            currentReinvestEvents.clear();
        }

        realizedSalesValue += saleValue;
        realizedCostBasis += costBasis;
        realizedGain += gainLoss;
    }

    private double getAdjustedCurrentCostBasis() {
        double rawCostBasis = 0.0;
        for (BuyLot lot : buyLots) {
            rawCostBasis += lot.remainingUnits * lot.unitCost;
        }

        double adjustment = getReinvestedDividendAdjustment();
        return Math.max(0.0, rawCostBasis - adjustment);
    }

    private double getReinvestedDividendAdjustment() {
        if (currentReinvestEvents.isEmpty() || currentDividendEvents.isEmpty()) {
            return 0.0;
        }

        LinkedHashSet<LocalDate> reinvestDates = new LinkedHashSet<>();
        for (ReinvestEvent event : currentReinvestEvents) {
            if (event == null || event.tradeDate == null || event.tradeDate.equals(LocalDate.MIN)) {
                continue;
            }
            reinvestDates.add(event.tradeDate);
        }
        if (reinvestDates.isEmpty()) {
            return 0.0;
        }

        double dividendAdjustment = 0.0;
        for (DividendEvent event : currentDividendEvents) {
            if (event == null || event.tradeDate == null || event.tradeDate.equals(LocalDate.MIN)) {
                continue;
            }
            if (!reinvestDates.contains(event.tradeDate)) {
                continue;
            }
            dividendAdjustment += Math.max(0.0, event.amount);
        }

        return dividendAdjustment;
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
            previousClose = 0.0;
            currencyCode = "NOK";
            resolvedSectorWeights.clear();
            resolvedRegionWeights.clear();
            applyGeneralTickerAndNameFallback();
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
                applyCandidateMetadata(candidate);
            } else {
                ticker = "";
                latestPrice = 0.0;
                previousClose = 0.0;
                currencyCode = "NOK";
                resolvedSector = "Other";
                resolvedRegion = "Global";
                resolvedSectorWeights.clear();
                resolvedRegionWeights.clear();
            }
        } catch (Exception e) {
            System.err.println("Yahoo Finance ISIN lookup failed: " + e.getMessage());
            ticker = "";
            latestPrice = 0.0;
            previousClose = 0.0;
            currencyCode = "NOK";
            resolvedSector = "Other";
            resolvedRegion = "Global";
            resolvedSectorWeights.clear();
            resolvedRegionWeights.clear();
        }

        applyGeneralTickerAndNameFallback();
    }

    private void applyCandidateMetadata(SearchCandidate candidate) {
        if (candidate == null || candidate.symbol == null || candidate.symbol.isBlank()) {
            return;
        }

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
        updateResolvedClassification(candidate);
        LatestQuote quote = fetchLatestQuote(ticker);
        latestPrice = candidate.regularMarketPrice > EPSILON
                ? candidate.regularMarketPrice
                : quote.latestPrice;
        previousClose = quote.previousClose > EPSILON ? quote.previousClose : 0.0;
    }

    private void applyGeneralTickerAndNameFallback() {
        if ((ticker == null || ticker.isBlank()) && looksLikeTicker(name)) {
            SearchCandidate byNameCandidate = searchCandidateByQuery(name);
            if (byNameCandidate != null) {
                applyCandidateMetadata(byNameCandidate);
            }
        }

        if ((ticker == null || ticker.isBlank()) && name != null && !name.isBlank()) {
            String symbolLikeName = name.trim().toUpperCase(Locale.ROOT);
            if (looksLikeTicker(symbolLikeName)) {
                String inferredSuffix = inferExchangeSuffixFromIsin(isin);
                if (!symbolLikeName.contains(".") && !inferredSuffix.isBlank()) {
                    symbolLikeName = symbolLikeName + "." + inferredSuffix;
                }
                ticker = symbolLikeName;
            }
        }

        if (ticker != null && !ticker.isBlank()) {
            currencyCode = normalizeCurrencyCode(currencyCode, ticker);
            updateAssetTypeFromTicker(ticker);
        }
    }

    private SearchCandidate searchCandidateByQuery(String query) {
        if (query == null || query.isBlank()) {
            return null;
        }

        try {
            String normalizedQuery = query.trim();
            String url = "https://query2.finance.yahoo.com/v1/finance/search?q="
                    + URLEncoder.encode(normalizedQuery, "UTF-8")
                    + "&quotesCount=10";

            String response = httpGetRequest(url);
            ArrayList<SearchCandidate> candidates = extractSearchCandidates(response);
            if (candidates.isEmpty()) {
                return null;
            }

            SearchCandidate candidate = null;
            if (looksLikeTicker(normalizedQuery)) {
                candidate = chooseBestHoldingCandidate(candidates, normalizedQuery.toUpperCase(Locale.ROOT));
            }

            if (candidate == null) {
                return null;
            }

            candidate = refineCandidateBySymbolListings(candidate);
            candidate = preferOsloListingWhenAvailable(candidate);
            return candidate;
        } catch (Exception ignored) {
            return null;
        }
    }

    private boolean looksLikeTicker(String value) {
        if (value == null || value.isBlank()) {
            return false;
        }

        String trimmed = value.trim();
        if (trimmed.contains(" ")) {
            return false;
        }
        if (!trimmed.matches("[A-Za-z0-9.=\\-]{1,20}")) {
            return false;
        }

        boolean hasLetter = false;
        for (int i = 0; i < trimmed.length(); i++) {
            if (Character.isLetter(trimmed.charAt(i))) {
                hasLetter = true;
                break;
            }
        }
        return hasLetter;
    }

    private String inferExchangeSuffixFromIsin(String isinValue) {
        if (isinValue == null || isinValue.length() < 2) {
            return "";
        }

        String country = isinValue.substring(0, 2).toUpperCase(Locale.ROOT);
        return switch (country) {
            case "NO" -> "OL";
            case "SE" -> "ST";
            case "DK" -> "CO";
            case "FI" -> "HE";
            default -> "";
        };
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
            baseCandidate.exchangeDisplay,
            baseCandidate.quoteType,
            baseCandidate.sector,
            baseCandidate.industry,
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

    private void updateResolvedClassification(SearchCandidate candidate) {
        if (candidate == null) {
            return;
        }

        resolvedSectorWeights.clear();
        resolvedRegionWeights.clear();
        String quoteType = candidate.quoteType == null ? "" : candidate.quoteType.toUpperCase(Locale.ROOT);
        boolean fundLike = quoteType.contains("ETF") || quoteType.contains("MUTUAL")
                || quoteType.contains("FUND") || assetType == AssetType.FUND;

        if (fundLike) {
            LinkedHashMap<String, Double> weights = fetchFundSectorWeights(firstNonBlank(ticker, candidate.symbol));
            if (!weights.isEmpty()) {
                resolvedSectorWeights.putAll(weights);
            }

            LinkedHashMap<String, Double> regionWeights = fetchFundRegionWeights(firstNonBlank(ticker, candidate.symbol));
            if (!regionWeights.isEmpty()) {
                resolvedRegionWeights.putAll(regionWeights);
            }
        }

        String sectorValue = candidate.sector;
        if (sectorValue == null || sectorValue.isBlank()) {
            sectorValue = candidate.industry;
        }
        if (sectorValue == null || sectorValue.isBlank()) {
            sectorValue = fetchSectorFromSearchNav(candidate.longName);
        }
        if (sectorValue == null || sectorValue.isBlank()) {
            sectorValue = fetchSectorFromSearchNav(name);
        }
        if (sectorValue == null || sectorValue.isBlank()) {
            sectorValue = fetchSectorFromSearchNav(candidate.symbol);
        }

        if ((sectorValue == null || sectorValue.isBlank()) && !resolvedSectorWeights.isEmpty()) {
            sectorValue = findDominantSector(resolvedSectorWeights);
        }

        String normalizedSector = normalizeClassificationLabel(sectorValue);
        if (normalizedSector.isBlank()) {
            normalizedSector = "Other";
        }
        resolvedSector = normalizedSector;

        String normalizedRegion = "";
        if (!resolvedRegionWeights.isEmpty()) {
            normalizedRegion = findDominantSector(resolvedRegionWeights);
        }

        if (normalizedRegion.isBlank()) {
            normalizedRegion = resolveCountryFromListing(candidate);
        }

        if (normalizedRegion.isBlank()) {
            String regionValue = candidate.exchangeDisplay;
            if (regionValue == null || regionValue.isBlank()) {
                regionValue = candidate.exchange;
            }
            normalizedRegion = mapMarketToMacroRegion(regionValue);
        }
        if (normalizedRegion.isBlank()) {
            normalizedRegion = "Global";
        }
        resolvedRegion = normalizedRegion;
    }

    private LinkedHashMap<String, Double> fetchFundRegionWeights(String symbol) {
        if (symbol == null || symbol.isBlank()) {
            return new LinkedHashMap<>();
        }

        String normalizedSymbol = symbol.trim().toUpperCase(Locale.ROOT);
        LinkedHashMap<String, Double> cached = FUND_REGION_WEIGHT_CACHE.get(normalizedSymbol);
        if (cached != null && !cached.isEmpty()) {
            return new LinkedHashMap<>(cached);
        }

        String response = getFundQuoteSummaryResponse(normalizedSymbol);
        LinkedHashMap<String, Double> weights = parseRegionWeightingsFromTopHoldings(response);
        if (!weights.isEmpty()) {
            FUND_REGION_WEIGHT_CACHE.put(normalizedSymbol, new LinkedHashMap<>(weights));
        }
        return weights;
    }

    private LinkedHashMap<String, Double> parseRegionWeightingsFromTopHoldings(String response) {
        LinkedHashMap<String, Double> weights = new LinkedHashMap<>();
        if (response == null || response.isBlank()) {
            return weights;
        }

        Pattern holdingPattern = Pattern.compile(
                "\\\"symbol\\\"\\s*:\\s*\\\"([^\\\"]+)\\\".*?\\\"holdingPercent\\\"\\s*:\\s*\\{\\s*\\\"raw\\\"\\s*:\\s*(-?\\d+(?:\\.\\d+)?(?:[Ee][+-]?\\d+)?)",
                Pattern.DOTALL
        );

        Matcher matcher = holdingPattern.matcher(response);
        double total = 0.0;
        while (matcher.find()) {
            String holdingSymbol = matcher.group(1);
            String rawWeightText = matcher.group(2);
            if (holdingSymbol == null || holdingSymbol.isBlank()) {
                continue;
            }

            double rawWeight;
            try {
                rawWeight = Double.parseDouble(rawWeightText);
            } catch (NumberFormatException ignored) {
                continue;
            }

            if (!Double.isFinite(rawWeight) || rawWeight <= 0.0) {
                continue;
            }

            String regionLabel = resolveRegionForHoldingSymbol(holdingSymbol);
            if (regionLabel.isBlank()) {
                regionLabel = "Global";
            }

            weights.merge(regionLabel, rawWeight, Double::sum);
            total += rawWeight;
        }

        if (total <= 0.0) {
            return new LinkedHashMap<>();
        }

        for (Map.Entry<String, Double> entry : new ArrayList<>(weights.entrySet())) {
            weights.put(entry.getKey(), entry.getValue() / total);
        }
        return weights;
    }

    private String resolveRegionForHoldingSymbol(String symbol) {
        if (symbol == null || symbol.isBlank()) {
            return "";
        }

        String normalizedSymbol = symbol.trim().toUpperCase(Locale.ROOT);
        String cached = HOLDING_SYMBOL_REGION_CACHE.get(normalizedSymbol);
        if (cached != null && !cached.isBlank()) {
            return cached;
        }

        try {
            String url = "https://query2.finance.yahoo.com/v1/finance/search?q="
                    + URLEncoder.encode(normalizedSymbol, "UTF-8")
                    + "&quotesCount=8&newsCount=0";
            String response = httpGetRequest(url);
            ArrayList<SearchCandidate> candidates = extractSearchCandidates(response);
            SearchCandidate best = chooseBestHoldingCandidate(candidates, normalizedSymbol);
            if (best == null) {
                return "";
            }

            String region = resolveCountryFromListing(best);
            if (region == null || region.isBlank()) {
                region = mapMarketToMacroRegion(firstNonBlank(best.exchangeDisplay, best.exchange));
            }
            if (region == null) {
                region = "";
            }

            if (!region.isBlank()) {
                HOLDING_SYMBOL_REGION_CACHE.put(normalizedSymbol, region);
            }
            return region;
        } catch (Exception ignored) {
            return "";
        }
    }

    private SearchCandidate chooseBestHoldingCandidate(ArrayList<SearchCandidate> candidates, String symbol) {
        if (candidates == null || candidates.isEmpty()) {
            return null;
        }

        String normalizedSymbol = symbol == null ? "" : symbol.toUpperCase(Locale.ROOT);
        SearchCandidate exact = null;
        SearchCandidate startsWith = null;

        for (SearchCandidate candidate : candidates) {
            if (candidate == null || candidate.symbol == null || candidate.symbol.isBlank()) {
                continue;
            }

            String candidateSymbol = candidate.symbol.toUpperCase(Locale.ROOT);
            if (candidateSymbol.equals(normalizedSymbol)) {
                exact = candidate;
                break;
            }

            if (startsWith == null && candidateSymbol.startsWith(normalizedSymbol + ".")) {
                startsWith = candidate;
            }
        }

        if (exact != null) {
            return exact;
        }
        if (startsWith != null) {
            return startsWith;
        }
        return null;
    }

    private LinkedHashMap<String, Double> fetchFundSectorWeights(String symbol) {
        if (symbol == null || symbol.isBlank()) {
            return new LinkedHashMap<>();
        }

        String normalizedSymbol = symbol.trim().toUpperCase(Locale.ROOT);
        LinkedHashMap<String, Double> cached = FUND_SECTOR_WEIGHT_CACHE.get(normalizedSymbol);
        if (cached != null && !cached.isEmpty()) {
            return new LinkedHashMap<>(cached);
        }

        String response = getFundQuoteSummaryResponse(normalizedSymbol);
        LinkedHashMap<String, Double> weights = parseSectorWeightings(response);
        if (!weights.isEmpty()) {
            FUND_SECTOR_WEIGHT_CACHE.put(normalizedSymbol, new LinkedHashMap<>(weights));
        }
        return weights;
    }

    private String getFundQuoteSummaryResponse(String symbol) {
        if (symbol == null || symbol.isBlank()) {
            return "";
        }

        String normalizedSymbol = symbol.trim().toUpperCase(Locale.ROOT);
        String cached = FUND_QUOTE_SUMMARY_CACHE.get(normalizedSymbol);
        if (cached != null && !cached.isBlank()) {
            return cached;
        }

        String baseUrl;
        try {
            baseUrl = "https://query2.finance.yahoo.com/v10/finance/quoteSummary/"
                    + URLEncoder.encode(normalizedSymbol, "UTF-8")
                    + "?modules=topHoldings,fundProfile";
        } catch (Exception ignored) {
            return "";
        }

        String response = fetchYahooQuoteSummaryWithAuth(baseUrl);
        if (response != null && !response.isBlank()) {
            FUND_QUOTE_SUMMARY_CACHE.put(normalizedSymbol, response);
        }
        return response;
    }

    private String findDominantSector(Map<String, Double> weights) {
        String dominant = "";
        double bestWeight = -1.0;
        for (Map.Entry<String, Double> entry : weights.entrySet()) {
            if (entry.getValue() != null && entry.getValue() > bestWeight) {
                bestWeight = entry.getValue();
                dominant = entry.getKey();
            }
        }
        return dominant;
    }

    private LinkedHashMap<String, Double> parseSectorWeightings(String response) {
        LinkedHashMap<String, Double> weights = new LinkedHashMap<>();
        if (response == null || response.isBlank()) {
            return weights;
        }

        Pattern listPattern = Pattern.compile("\\\"sectorWeightings\\\"\\s*:\\s*\\[(.*?)]", Pattern.DOTALL);
        Matcher listMatcher = listPattern.matcher(response);
        if (!listMatcher.find()) {
            return weights;
        }

        String listBody = listMatcher.group(1);
        Pattern entryPattern = Pattern.compile(
                "\\{\\s*\\\"([a-z_]+)\\\"\\s*:\\s*\\{\\s*\\\"raw\\\"\\s*:\\s*(-?\\d+(?:\\.\\d+)?(?:[Ee][+-]?\\d+)?)",
                Pattern.DOTALL
        );
        Matcher entryMatcher = entryPattern.matcher(listBody);

        double total = 0.0;
        while (entryMatcher.find()) {
            String rawKey = entryMatcher.group(1);
            String rawValueText = entryMatcher.group(2);
            double rawValue;
            try {
                rawValue = Double.parseDouble(rawValueText);
            } catch (NumberFormatException ignored) {
                continue;
            }

            if (!Double.isFinite(rawValue) || rawValue <= 0.0) {
                continue;
            }

            String label = normalizeSectorWeightKey(rawKey);
            if (label.isBlank()) {
                continue;
            }

            weights.merge(label, rawValue, Double::sum);
            total += rawValue;
        }

        if (total <= 0.0) {
            weights.clear();
            return weights;
        }

        for (Map.Entry<String, Double> entry : new ArrayList<>(weights.entrySet())) {
            weights.put(entry.getKey(), entry.getValue() / total);
        }

        return weights;
    }

    private String normalizeSectorWeightKey(String rawKey) {
        if (rawKey == null || rawKey.isBlank()) {
            return "";
        }

        return switch (rawKey.toLowerCase(Locale.ROOT)) {
            case "other", "others" -> "Others";
            case "realestate" -> "Real Estate";
            case "consumer_cyclical" -> "Consumer Cyclical";
            case "consumer_defensive" -> "Consumer Defensive";
            case "basic_materials" -> "Basic Materials";
            case "communication_services" -> "Communication Services";
            case "financial_services" -> "Financial Services";
            default -> normalizeClassificationLabel(rawKey);
        };
    }

    private String fetchYahooQuoteSummaryWithAuth(String baseUrl) {
        try {
            if (!ensureYahooAuthSession()) {
                return "";
            }

            String response = doYahooAuthRequest(baseUrl);
            if (isYahooAuthError(response)) {
                clearYahooAuthSession();
                if (!ensureYahooAuthSession()) {
                    return "";
                }
                response = doYahooAuthRequest(baseUrl);
            }

            if (isYahooAuthError(response)) {
                return "";
            }
            return response;
        } catch (Exception ignored) {
            return "";
        }
    }

    private String doYahooAuthRequest(String baseUrl) throws Exception {
        String url = baseUrl + (baseUrl.contains("?") ? "&" : "?")
                + "crumb=" + URLEncoder.encode(yahooAuthCrumb, "UTF-8");

        Map<String, String> headers = new LinkedHashMap<>();
        headers.put("User-Agent", "Mozilla/5.0");
        headers.put("Cookie", yahooAuthCookieHeader);
        headers.put("Accept", "application/json,text/plain,*/*");
        headers.put("Referer", "https://finance.yahoo.com/");
        headers.put("Origin", "https://finance.yahoo.com");
        return httpGetRequest(url, headers);
    }

    private boolean ensureYahooAuthSession() {
        if (System.currentTimeMillis() < yahooAuthExpiresAtMs
                && yahooAuthCookieHeader != null && !yahooAuthCookieHeader.isBlank()
                && yahooAuthCrumb != null && !yahooAuthCrumb.isBlank()) {
            return true;
        }

        try {
            Map<String, String> cookieHeaders = fetchYahooCookies();
            if (cookieHeaders.isEmpty()) {
                return false;
            }

            StringBuilder cookieBuilder = new StringBuilder();
            for (Map.Entry<String, String> entry : cookieHeaders.entrySet()) {
                if (cookieBuilder.length() > 0) {
                    cookieBuilder.append("; ");
                }
                cookieBuilder.append(entry.getKey()).append("=").append(entry.getValue());
            }

            if (cookieBuilder.length() == 0) {
                return false;
            }

            Map<String, String> crumbHeaders = new LinkedHashMap<>();
            crumbHeaders.put("User-Agent", "Mozilla/5.0");
            crumbHeaders.put("Cookie", cookieBuilder.toString());

            String crumb = httpGetRequest("https://query1.finance.yahoo.com/v1/test/getcrumb", crumbHeaders).trim();
            if (crumb.isBlank() || crumb.startsWith("{") || crumb.startsWith("<") || crumb.contains("Edge:")) {
                return false;
            }

            yahooAuthCookieHeader = cookieBuilder.toString();
            yahooAuthCrumb = crumb;
            yahooAuthExpiresAtMs = System.currentTimeMillis() + YAHOO_AUTH_TTL_MS;
            return true;
        } catch (Exception ignored) {
            return false;
        }
    }

    private Map<String, String> fetchYahooCookies() throws Exception {
        URL url = URI.create("https://fc.yahoo.com").toURL();
        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        conn.setRequestMethod("GET");
        conn.setRequestProperty("User-Agent", "Mozilla/5.0");
        conn.setConnectTimeout(5000);
        conn.setReadTimeout(5000);

        try (InputStream input = conn.getInputStream()) {
            while (input.read() != -1) {
                // Consume response to complete the request.
            }
        } catch (IOException ignored) {
            // Continue; some responses still include cookies even if body read fails.
        }

        List<String> setCookieHeaders = new ArrayList<>();
        for (Map.Entry<String, List<String>> entry : conn.getHeaderFields().entrySet()) {
            if (entry.getKey() != null && "Set-Cookie".equalsIgnoreCase(entry.getKey()) && entry.getValue() != null) {
                setCookieHeaders.addAll(entry.getValue());
            }
        }
        Map<String, String> cookies = new LinkedHashMap<>();
        if (!setCookieHeaders.isEmpty()) {
            for (String header : setCookieHeaders) {
                if (header == null || header.isBlank()) {
                    continue;
                }

                String[] parts = header.split(";", 2);
                if (parts.length == 0 || !parts[0].contains("=")) {
                    continue;
                }

                String[] nameValue = parts[0].split("=", 2);
                if (nameValue.length != 2) {
                    continue;
                }

                String cookieName = nameValue[0].trim();
                String cookieValue = nameValue[1].trim();
                if (!cookieName.isBlank() && !cookieValue.isBlank()) {
                    cookies.put(cookieName, cookieValue);
                }
            }
        }

        conn.disconnect();
        return cookies;
    }

    private boolean isYahooAuthError(String response) {
        if (response == null || response.isBlank()) {
            return true;
        }

        return response.contains("Invalid Crumb")
                || response.contains("Invalid Cookie")
                || response.contains("Unauthorized")
                || response.contains("Too Many Requests")
                || response.contains("Edge:");
    }

    private void clearYahooAuthSession() {
        yahooAuthCookieHeader = "";
        yahooAuthCrumb = "";
        yahooAuthExpiresAtMs = 0L;
    }

    private String resolveCountryFromListing(SearchCandidate candidate) {
        if (candidate == null) {
            return "";
        }

        String symbol = candidate.symbol == null ? "" : candidate.symbol.toUpperCase(Locale.ROOT).trim();
        if (symbol.endsWith(".OL")) return "Norway";
        if (symbol.endsWith(".IR")) return "Ireland";
        if (symbol.endsWith(".L")) return "United Kingdom";
        if (symbol.endsWith(".DE")) return "Germany";
        if (symbol.endsWith(".PA")) return "France";
        if (symbol.endsWith(".AS")) return "Netherlands";
        if (symbol.endsWith(".ST")) return "Sweden";
        if (symbol.endsWith(".CO")) return "Denmark";
        if (symbol.endsWith(".HE")) return "Finland";
        if (symbol.endsWith(".SW")) return "Switzerland";
        if (symbol.endsWith(".MI")) return "Italy";
        if (symbol.endsWith(".MC")) return "Spain";
        if (symbol.endsWith(".TO")) return "Canada";
        if (symbol.endsWith(".AX")) return "Australia";
        if (symbol.endsWith(".HK")) return "Hong Kong";
        if (symbol.endsWith(".T")) return "Japan";

        String market = normalizeClassificationLabel(firstNonBlank(candidate.exchangeDisplay, candidate.exchange))
                .toLowerCase(Locale.ROOT);
        if (market.isBlank()) {
            return "";
        }

        if (containsAny(market, "oslo", "norway", "norwegian")) return "Norway";
        if (containsAny(market, "irish", "ireland", "dublin")) return "Ireland";
        if (containsAny(market, "new york", "nyse", "nasdaq", "united states", "usa", "us")) return "United States";
        if (containsAny(market, "london", "united kingdom", "uk")) return "United Kingdom";
        if (containsAny(market, "frankfurt", "xetra", "germany")) return "Germany";
        if (containsAny(market, "paris", "france")) return "France";
        if (containsAny(market, "amsterdam", "netherlands")) return "Netherlands";
        if (containsAny(market, "stockholm", "sweden")) return "Sweden";
        if (containsAny(market, "copenhagen", "denmark")) return "Denmark";
        if (containsAny(market, "helsinki", "finland")) return "Finland";
        if (containsAny(market, "swiss", "zurich", "switzerland")) return "Switzerland";
        if (containsAny(market, "milan", "italy")) return "Italy";
        if (containsAny(market, "madrid", "spain")) return "Spain";
        if (containsAny(market, "toronto", "canada")) return "Canada";
        if (containsAny(market, "sydney", "australia")) return "Australia";
        if (containsAny(market, "hong kong")) return "Hong Kong";
        if (containsAny(market, "tokyo", "japan")) return "Japan";

        return "";
    }

    private String firstNonBlank(String first, String second) {
        if (first != null && !first.isBlank()) {
            return first;
        }
        if (second != null && !second.isBlank()) {
            return second;
        }
        return "";
    }

    private String fetchSectorFromSearchNav(String query) {
        if (query == null || query.isBlank()) {
            return "";
        }

        try {
            String url = "https://query2.finance.yahoo.com/v1/finance/search?q="
                    + URLEncoder.encode(query, "UTF-8")
                    + "&quotesCount=5&newsCount=0";

            String response = httpGetRequest(url);
            return extractSectorNavName(response);
        } catch (Exception ignored) {
            return "";
        }
    }

    private String extractSectorNavName(String response) {
        if (response == null || response.isBlank()) {
            return "";
        }

        Pattern navPattern = Pattern.compile("\\\"navName\\\"\\s*:\\s*\\\"(.*?)\\\".*?\\\"navUrl\\\"\\s*:\\s*\\\"(.*?)\\\"", Pattern.DOTALL);
        Matcher matcher = navPattern.matcher(response);
        while (matcher.find()) {
            String navName = matcher.group(1);
            String navUrl = matcher.group(2);
            if (navName == null || navName.isBlank() || navUrl == null || navUrl.isBlank()) {
                continue;
            }

            if (navUrl.contains("/sectors/")) {
                return navName;
            }
        }

        return "";
    }

    private String mapMarketToMacroRegion(String marketValue) {
        String market = normalizeClassificationLabel(marketValue).toLowerCase(Locale.ROOT);
        if (market.isBlank()) {
            return "";
        }

        if (containsAny(market, "nasdaq", "nyse", "nysearca", "bats", "otc", "usa", "united states", "us")) {
            return "USA";
        }

        if (containsAny(market, "oslo", "irish", "ireland", "frankfurt", "xetra", "paris", "london", "euronext", "stockholm", "copenhagen", "helsinki", "amsterdam", "swiss", "vienna", "madrid", "milan", "europe")) {
            return "Europe";
        }

        if (containsAny(market, "tokyo", "japan", "hong kong", "shanghai", "shenzhen", "singapore", "korea", "india", "australia", "sydney", "asia")) {
            return "Asia-Pacific";
        }

        if (containsAny(market, "toronto", "canada")) {
            return "North America";
        }

        if (containsAny(market, "brazil", "mexico", "chile", "argentina", "latin")) {
            return "Latin America";
        }

        if (containsAny(market, "south africa", "africa")) {
            return "Africa";
        }

        if (containsAny(market, "saudi", "dubai", "abu dhabi", "qatar", "middle east")) {
            return "Middle East";
        }

        return "Global";
    }

    private boolean containsAny(String text, String... candidates) {
        if (text == null || text.isBlank()) {
            return false;
        }

        for (String candidate : candidates) {
            if (candidate != null && !candidate.isBlank() && text.contains(candidate)) {
                return true;
            }
        }
        return false;
    }

    private String normalizeClassificationLabel(String value) {
        if (value == null || value.isBlank()) {
            return "";
        }

        String normalized = value
                .replace("_", " ")
                .replace("-", " ")
                .trim();
        if (normalized.isBlank()) {
            return "";
        }

        String[] words = normalized.split("\\s+");
        StringBuilder output = new StringBuilder();
        for (String word : words) {
            if (word.isBlank()) {
                continue;
            }

            String lower = word.toLowerCase(Locale.ROOT);
            String pretty;
            if (lower.length() <= 3) {
                pretty = lower.toUpperCase(Locale.ROOT);
            } else {
                pretty = Character.toUpperCase(lower.charAt(0)) + lower.substring(1);
            }

            if (output.length() > 0) {
                output.append(' ');
            }
            output.append(pretty);
        }

        return output.toString().trim();
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
        String exchangeDisplay = extractValue(quoteObject, "exchDisp");
        String quoteType = extractValue(quoteObject, "quoteType");
        String sector = extractValue(quoteObject, "sectorDisp");
        if (sector == null || sector.isBlank()) {
            sector = extractValue(quoteObject, "sector");
        }
        String industry = extractValue(quoteObject, "industryDisp");
        if (industry == null || industry.isBlank()) {
            industry = extractValue(quoteObject, "industry");
        }
        String currency = extractValue(quoteObject, "currency");
        String shortName = extractValue(quoteObject, "shortname");
        String longName = extractValue(quoteObject, "longname");
        double regularMarketPrice = extractNumericValue(quoteObject, "regularMarketPrice");

        return new SearchCandidate(
                symbol,
                exchange,
                exchangeDisplay,
                quoteType,
                sector,
                industry,
                currency,
                shortName,
                longName,
                regularMarketPrice
        );
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
        return fetchLatestQuote(symbol).latestPrice;
    }

    private LatestQuote fetchLatestQuote(String symbol) {
        if (symbol == null || symbol.isBlank()) {
            return new LatestQuote(0.0, 0.0);
        }

        try {
            String quoteUrl = "https://query1.finance.yahoo.com/v8/finance/chart/"
                    + URLEncoder.encode(symbol, "UTF-8")
                    + "?interval=1d&range=1d";
            String quoteResponse = httpGetRequest(quoteUrl);
            double regularMarketPrice = extractNumericValue(quoteResponse, "regularMarketPrice");
            double chartPreviousClose = extractNumericValue(quoteResponse, "chartPreviousClose");
            double resolvedLatestPrice = regularMarketPrice > EPSILON
                    ? regularMarketPrice
                    : (chartPreviousClose > EPSILON ? chartPreviousClose : 0.0);
            double resolvedPreviousClose = chartPreviousClose > EPSILON ? chartPreviousClose : 0.0;
            return new LatestQuote(resolvedLatestPrice, resolvedPreviousClose);
        } catch (Exception ignored) {
            return new LatestQuote(0.0, 0.0);
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
        return httpGetRequest(urlString, null);
    }

    private String httpGetRequest(String urlString, Map<String, String> headers) throws Exception {
        URL url = URI.create(urlString).toURL();
        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        conn.setRequestMethod("GET");
        conn.setRequestProperty("User-Agent", "Mozilla/5.0");
        if (headers != null) {
            for (Map.Entry<String, String> header : headers.entrySet()) {
                if (header.getKey() == null || header.getValue() == null || header.getValue().isBlank()) {
                    continue;
                }
                conn.setRequestProperty(header.getKey(), header.getValue());
            }
        }
        conn.setConnectTimeout(5000);
        conn.setReadTimeout(5000);

        StringBuilder response = new StringBuilder();
        int statusCode = conn.getResponseCode();
        InputStream stream = (statusCode >= 200 && statusCode < 400)
                ? conn.getInputStream()
                : conn.getErrorStream();
        if (stream == null) {
            conn.disconnect();
            return "";
        }

        try (BufferedReader reader = new BufferedReader(new InputStreamReader(stream))) {
            String line;
            while ((line = reader.readLine()) != null) {
                response.append(line);
            }
        }
        conn.disconnect();
        return response.toString();
    }
}
