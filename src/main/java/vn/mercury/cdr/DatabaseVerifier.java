package vn.mercury.cdr;

import java.math.BigDecimal;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;
import java.util.Set;

/** Executes only fixed, parameterized SELECT queries for post-simulation verification. */
final class DatabaseVerifier {
    private static final String IDENTIFIER = "[A-Za-z_][A-Za-z0-9_]*";
    private static final long AAX_CLIENT = 1130L;

    private DatabaseVerifier() { }

    static void testConnection(Config config) throws SQLException {
        try (Connection connection = open(config); Statement statement = connection.createStatement(); ResultSet ignored = statement.executeQuery("SELECT 1")) {
            // A successful SELECT proves the connection and permissions are usable.
        }
    }

    static Verification verify(Config config, String sessionId, String configuredClient,
                               BigDecimal expectedInew, BigDecimal expectedAax, Set<Long> claimedRawIds) throws SQLException {
        List<RawCdr> raws = readRawCdrs(config, sessionId);
        if (raws.isEmpty()) return Verification.missingRaw(sessionId);

        Verification firstFailure = null;
        for (RawCdr raw : raws) {
            if (claimedRawIds.contains(raw.id())) continue;
            Verification candidate = verifyDcdrs(config, raw, configuredClient, expectedInew, expectedAax);
            if (candidate.passed()) return candidate;
            if (firstFailure == null) firstFailure = candidate;
        }
        return firstFailure == null ? Verification.noUnclaimedRaw(sessionId) : firstFailure;
    }

    private static List<RawCdr> readRawCdrs(Config config, String sessionId) throws SQLException {
        String sql = "SELECT id, charged_quantity FROM " + config.schema() + ".c_cdr_inew_raw "
                + "WHERE session_id = ? ORDER BY id DESC";
        try (Connection connection = open(config); PreparedStatement statement = connection.prepareStatement(sql)) {
            statement.setString(1, sessionId);
            try (ResultSet result = statement.executeQuery()) {
                List<RawCdr> raws = new ArrayList<>();
                while (result.next()) {
                    raws.add(new RawCdr(result.getLong("id"), result.getBigDecimal("charged_quantity")));
                }
                return raws;
            }
        }
    }

    private static Verification verifyDcdrs(Config config, RawCdr raw, String configuredClient,
                                              BigDecimal expectedInew, BigDecimal expectedAax) throws SQLException {
        String sql = "SELECT client, amount FROM " + config.schema() + ".d_cdrs WHERE ref_id = ? ORDER BY id";
        List<Dcdr> dCdrs = new ArrayList<>();
        try (Connection connection = open(config); PreparedStatement statement = connection.prepareStatement(sql)) {
            statement.setLong(1, raw.id());
            try (ResultSet result = statement.executeQuery()) {
                while (result.next()) {
                    dCdrs.add(new Dcdr(result.getLong("client"), result.getBigDecimal("amount")));
                }
            }
        }

        boolean inewMatches = expectedInew == null || numbersMatch(expectedInew, raw.chargedQuantity());
        boolean anyNon1130Client = configuredClient == null || configuredClient.isBlank();
        long expectedClient = anyNon1130Client ? -1L : parseClientId(configuredClient);
        List<Dcdr> testClientRows = dCdrs.stream()
                .filter(row -> anyNon1130Client ? row.client() != AAX_CLIENT : row.client() == expectedClient)
                .toList();
        List<Dcdr> aaxRows = dCdrs.stream().filter(row -> row.client() == AAX_CLIENT).toList();
        BigDecimal actualAax = aaxRows.size() == 1 ? aaxRows.get(0).amount() : null;
        boolean aaxMatches = expectedAax == null || numbersMatch(expectedAax, actualAax);
        boolean passed = inewMatches && dCdrs.size() == 2 && testClientRows.size() == 1 && aaxRows.size() == 1 && aaxMatches;

        return new Verification(raw.id(), raw.chargedQuantity(), dCdrs.size(), testClientRows.size(), aaxRows.size(), actualAax,
                inewMatches, aaxMatches, passed, describe(raw, dCdrs.size(), testClientRows.size(), aaxRows.size(), actualAax,
                        inewMatches, aaxMatches, anyNon1130Client));
    }

    private static String describe(RawCdr raw, int totalDcdrs, int testClientRows, int aaxRows, BigDecimal actualAax,
                                   boolean inewMatches, boolean aaxMatches, boolean anyNon1130Client) {
        List<String> issues = new ArrayList<>();
        if (!inewMatches) issues.add("charged_quantity mismatch");
        if (totalDcdrs != 2) issues.add("expected 2 d_cdrs, found " + totalDcdrs);
        if (testClientRows != 1) issues.add("expected " + (anyNon1130Client ? "1 non-1130 client" : "1 configured client") + ", found " + testClientRows);
        if (aaxRows != 1) issues.add("expected 1 client 1130 row, found " + aaxRows);
        if (!aaxMatches) issues.add("AAX amount mismatch (actual " + display(actualAax) + ")");
        String summary = "iNew raw " + raw.id() + ", charged_quantity=" + display(raw.chargedQuantity())
                + ", d_cdrs=" + totalDcdrs + ", client 1130 amount=" + display(actualAax);
        return issues.isEmpty() ? "Verified: " + summary : "Verification failed: " + String.join("; ", issues) + ". " + summary;
    }

    private static boolean numbersMatch(BigDecimal expected, BigDecimal actual) {
        return expected != null && actual != null && expected.compareTo(actual) == 0;
    }

    private static String display(BigDecimal value) {
        return value == null ? "null" : value.stripTrailingZeros().toPlainString();
    }

    private static long parseClientId(String value) {
        try {
            return Long.parseLong(value.trim());
        } catch (NumberFormatException exception) {
            throw new IllegalArgumentException("Client ID must be a whole number for database verification.");
        }
    }

    private static Connection open(Config config) throws SQLException {
        Properties properties = new Properties();
        properties.setProperty("user", config.username());
        properties.setProperty("password", config.password());
        properties.setProperty("readOnly", "true");
        Connection connection = DriverManager.getConnection(config.jdbcUrl(), properties);
        connection.setReadOnly(true);
        return connection;
    }

    record Config(String host, int port, String database, String schema, String username, String password) {
        Config {
            if (host == null || host.isBlank()) throw new IllegalArgumentException("Database host is required.");
            if (port < 1 || port > 65535) throw new IllegalArgumentException("Database port must be between 1 and 65535.");
            if (database == null || database.isBlank()) throw new IllegalArgumentException("Database name is required.");
            if (schema == null || !schema.matches(IDENTIFIER)) throw new IllegalArgumentException("Database schema must contain only letters, digits, and underscores.");
            if (username == null || username.isBlank()) throw new IllegalArgumentException("Database username is required.");
            if (password == null || password.isBlank()) throw new IllegalArgumentException("Database password is required.");
        }

        String jdbcUrl() {
            return "jdbc:postgresql://" + host.trim() + ":" + port + "/" + database.trim() + "?sslmode=prefer";
        }
    }

    record Verification(Long rawId, BigDecimal chargedQuantity, int totalDcdrs, int testClientRows, int aaxRows,
                        BigDecimal actualAax, boolean inewMatches, boolean aaxMatches, boolean passed, String summary) {
        static Verification missingRaw(String sessionId) {
            return new Verification(null, null, 0, 0, 0, null, false, false, false,
                    "Verification pending: no c_cdr_inew_raw found for session ID " + sessionId + ".");
        }

        static Verification noUnclaimedRaw(String sessionId) {
            return new Verification(null, null, 0, 0, 0, null, false, false, false,
                    "Verification pending: all iNew raw rows for shared session ID " + sessionId + " are already matched to another Usage.");
        }
    }

    private record RawCdr(long id, BigDecimal chargedQuantity) { }
    private record Dcdr(long client, BigDecimal amount) { }
}
