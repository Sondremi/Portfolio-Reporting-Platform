package report;

import java.util.LinkedHashMap;
import java.util.Map;

public class OverviewRow {

    public final String securityKey;
    public final String tickerText;
    public final String securityDisplayName;
    public final double dayChangePct;
    public final double previousClose;
    public final boolean hasDayChangePct;
    public final String assetType;
    public final String sectorLabel;
    public final Map<String, Double> sectorWeights;
    public final String regionLabel;
    public final Map<String, Double> regionWeights;
    public final String currencyCode;
    public final String realizedReturnPctText;
    public final double units;
    public final double averageCost;
    public final double latestPrice;
    public final double positionCostBasis;
    public final double realizedCostBasis;
    public final double historicalCostBasis;
    public final double marketValue;
    public final double unrealized;
    public final double unrealizedPct;
    public final double realized;
    public final double dividends;
    public final double totalReturn;
    public final double totalReturnPct;
    public final boolean hasPrice;

    public OverviewRow(
            String securityKey,
            String tickerText,
            String securityDisplayName,
            double dayChangePct,
            double previousClose,
            boolean hasDayChangePct,
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

        this.securityKey = securityKey != null ? securityKey : "";
        this.tickerText = tickerText != null ? tickerText : "-";
        this.securityDisplayName = securityDisplayName != null ? securityDisplayName : "-";
        this.dayChangePct = dayChangePct;
        this.previousClose = previousClose;
        this.hasDayChangePct = hasDayChangePct;
        this.assetType = assetType != null ? assetType : "UNKNOWN";
        this.sectorLabel = sectorLabel != null ? sectorLabel : "Other";
        this.sectorWeights = sectorWeights != null ? new LinkedHashMap<>(sectorWeights) : Map.of();
        this.regionLabel = regionLabel != null ? regionLabel : "Global";
        this.regionWeights = regionWeights != null ? new LinkedHashMap<>(regionWeights) : Map.of();
        this.currencyCode = currencyCode != null ? currencyCode : "NOK";
        this.realizedReturnPctText = realizedReturnPctText != null ? realizedReturnPctText : "0.00";
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
