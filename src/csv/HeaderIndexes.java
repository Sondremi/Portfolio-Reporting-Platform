package csv;

public class HeaderIndexes {
    public int transactionId = -1;
    public int securityName = -1;
    public int securityType = -1;
    public int isin = -1;
    public int transactionType = -1;
    public int amount = -1;
    public int quantity = -1;
    public int price = -1;
    public int result = -1;
    public int totalFees = -1;
    public int tradeDate = -1;
    public int cancellationDate = -1;
    public int portfolioId = -1;
    public int cashBalance = -1;

    public boolean hasRequiredColumns() {
        return securityName >= 0 && transactionType >= 0;
    }
}