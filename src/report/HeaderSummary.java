package report;

public class HeaderSummary {

    public final String generatedDate;
    public final int fileCount;
    public final int transactionCount;
    public final int holdingsCount;
    public final String totalCurrencyCode;
    public final double totalMarketValue;
    public final double cashHoldings;
    public final double totalReturn;
    public final double totalReturnPct;
    public final String bestLabel;
    public final double bestReturn;
    public final double bestReturnPct;
    public final String worstLabel;
    public final double worstReturn;
    public final double worstReturnPct;
    public final String sparklineSvg;

    public HeaderSummary(
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

        this.generatedDate = generatedDate != null ? generatedDate : "";
        this.fileCount = fileCount;
        this.transactionCount = transactionCount;
        this.holdingsCount = holdingsCount;
        this.totalCurrencyCode = totalCurrencyCode != null ? totalCurrencyCode : "NOK";
        this.totalMarketValue = totalMarketValue;
        this.cashHoldings = cashHoldings;
        this.totalReturn = totalReturn;
        this.totalReturnPct = totalReturnPct;
        this.bestLabel = bestLabel != null ? bestLabel : "-";
        this.bestReturn = bestReturn;
        this.bestReturnPct = bestReturnPct;
        this.worstLabel = worstLabel != null ? worstLabel : "-";
        this.worstReturn = worstReturn;
        this.worstReturnPct = worstReturnPct;
        this.sparklineSvg = sparklineSvg != null ? sparklineSvg : "";
    }
}