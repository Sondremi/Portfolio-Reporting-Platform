package finance;

import model.Security;

public final class YahooFinanceClient {

    private YahooFinanceClient() {
    }

    public static void enrichSecurity(Security security) {
        // Legacy enrichment is now handled inside the Security constructor.
    }
}
