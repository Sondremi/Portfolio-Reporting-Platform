package csv;

import model.Events;
import model.Security;

import java.util.*;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;

public class TransactionStore {

    private final Map<String, String> renamedSecurityIsin = new LinkedHashMap<>();

    private final ArrayList<Security> securities = new ArrayList<>();
    private final Map<String, Security> securitiesByKey = new LinkedHashMap<>();
    private final Map<String, String> canonicalSecurityNameByIsin = new LinkedHashMap<>();

    private final ArrayList<Events.UnitEvent> unitEvents = new ArrayList<>();
    private final ArrayList<Events.CashEvent> cashEvents = new ArrayList<>();
    private final ArrayList<Events.PortfolioCashSnapshot> portfolioCashSnapshots = new ArrayList<>();

    private int loadedCsvFileCount = 0;
    private int loadedTransactionRowCount = 0;

    public Security getOrCreateSecurity(String name, String isin) {
        String key;
        if (isin == null || isin.isBlank()) {
            String normalizedName = normalizeSecurityNameKey(name);
            key = normalizedName.isBlank() ? "UNKNOWN_SECURITY" : "NAME:" + normalizedName;
        } else {
            key = isin.trim().toUpperCase(Locale.ROOT);
        }
        return securitiesByKey.computeIfAbsent(key, k -> {
            Security security = new Security(name, isin);
            securities.add(security);
            return security;
        });
    }

    private static String normalizeSecurityNameKey(String name) {
        if (name == null) {
            return "";
        }
        return name
                .trim()
                .replaceAll("\\s+", " ")
                .toUpperCase(Locale.ROOT);
    }

    public void addUnitEvent(Events.UnitEvent event) {
        unitEvents.add(event);
    }

    public void addCashEvent(Events.CashEvent event) {
        cashEvents.add(event);
    }

    public void addPortfolioCashSnapshot(Events.PortfolioCashSnapshot snapshot) {
        portfolioCashSnapshots.add(snapshot);
    }

    public void incrementTransactionRowCount() {
        loadedTransactionRowCount++;
    }

    public void setLoadedCsvFileCount(int count) {
        loadedCsvFileCount = Math.max(0, count);
    }

    public void rememberCanonicalSecurityName(String originalIsin, String canonicalIsin, String securityName) {
        if (securityName == null || securityName.isBlank() || canonicalIsin == null || canonicalIsin.isBlank()) {
            return;
        }
        String norm = canonicalIsin.trim().toUpperCase(Locale.ROOT);
        canonicalSecurityNameByIsin.putIfAbsent(norm, securityName);
    }

    public void rememberRenamedSecurityIsin(String oldIsin, String newIsin) {
        if (oldIsin == null || newIsin == null) {
            return;
        }

        String oldNorm = oldIsin.trim().toUpperCase(Locale.ROOT);
        String newNorm = newIsin.trim().toUpperCase(Locale.ROOT);
        if (oldNorm.isBlank() || newNorm.isBlank() || oldNorm.equals(newNorm)) {
            return;
        }

        renamedSecurityIsin.putIfAbsent(oldNorm, newNorm);
    }

    public List<Security> getSecurities() {
        return new ArrayList<>(securities);
    }

    // Resolves each security's market data (Yahoo ticker/price/classification) concurrently.
    // Call once after all CSVs are loaded and before building the report. Idempotent per security.
    public void resolveSecurityMarketData() {
        List<Security> pending = new ArrayList<>(securities);
        if (pending.isEmpty()) {
            return;
        }
        int workerCount = Math.max(1, Math.min(8, pending.size()));
        ExecutorService executor = Executors.newFixedThreadPool(workerCount);
        try {
            List<Future<?>> futures = new ArrayList<>();
            for (Security security : pending) {
                futures.add(executor.submit(security::resolveMarketData));
            }
            for (Future<?> future : futures) {
                try {
                    future.get();
                } catch (InterruptedException e) {
                    Thread.currentThread().interrupt();
                    break;
                } catch (ExecutionException e) {
                    // A single security failing to resolve must not abort the rest;
                    // it keeps its safe defaults (handled inside setTicker).
                }
            }
        } finally {
            executor.shutdown();
        }
    }

    public List<Events.UnitEvent> getUnitEvents() {
        return new ArrayList<>(unitEvents);
    }

    public List<Events.CashEvent> getCashEvents() {
        return new ArrayList<>(cashEvents);
    }

    public List<Events.PortfolioCashSnapshot> getPortfolioCashSnapshots() {
        return new ArrayList<>(portfolioCashSnapshots);
    }

    public Map<String, String> getCanonicalSecurityNameByIsin() {
        return new LinkedHashMap<>(canonicalSecurityNameByIsin);
    }

    public Map<String, String> getRenamedSecurityIsin() {
        return Collections.unmodifiableMap(renamedSecurityIsin);
    }

    public int getLoadedCsvFileCount() {
        return loadedCsvFileCount;
    }

    public int getLoadedTransactionRowCount() {
        return loadedTransactionRowCount;
    }

    public double getCurrentCashHoldings() {
        if (portfolioCashSnapshots.isEmpty()) {
            double cash = 0.0;
            for (Events.CashEvent event : cashEvents) {
                cash += event.cashDelta();
            }
            return cash;
        }

        LinkedHashMap<String, Events.PortfolioCashSnapshot> latestByPortfolio = new LinkedHashMap<>();
        for (Events.PortfolioCashSnapshot snapshot : portfolioCashSnapshots) {
            if (snapshot == null || snapshot.portfolioId() == null || snapshot.portfolioId().isBlank()) {
                continue;
            }

            Events.PortfolioCashSnapshot existing = latestByPortfolio.get(snapshot.portfolioId());
            if (existing == null
                    || snapshot.tradeDate().isAfter(existing.tradeDate())
                    || (snapshot.tradeDate().equals(existing.tradeDate()) && snapshot.sortId() >= existing.sortId())) {
                latestByPortfolio.put(snapshot.portfolioId(), snapshot);
            }
        }

        double authoritativePortfolioCash = 0.0;
        for (Events.PortfolioCashSnapshot snapshot : latestByPortfolio.values()) {
            authoritativePortfolioCash += snapshot.balance();
        }

        return authoritativePortfolioCash;
    }
}
