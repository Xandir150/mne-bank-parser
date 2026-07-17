using System;

// ──────────────────────────────────────────────────────────────
// Classifies LoadResult errors into transient (retry with backoff)
// vs permanent (rename to .error, wait for a human).
// ──────────────────────────────────────────────────────────────
public static class ErrorClassifier
{
    // Substrings that indicate the 1C server/COM layer failed, not the data.
    // Matched case-insensitively against the whole error text.
    private static readonly string[] _transientMarkers =
    {
        "Сеанс завершен",            // session killed by admin / server restart
        "Сеанс отсутствует",
        "Соединение с сервером",     // server connection lost/reset
        "разорвано администратором",
        "RPC",                       // COM RPC failures
        "0x800706BA",                // RPC_S_SERVER_UNAVAILABLE
        "0x80010108",                // RPC_E_DISCONNECTED
        "0x80010105",                // RPC_E_SERVERFAULT
        "forcibly closed",           // TCP reset (english runtime message)
        "Превышено допустимое количество", // license/session limit — frees up later
    };

    /// <summary>True when the error is worth retrying automatically: the 1C
    /// connection/session broke, rather than the file data being wrong.
    /// "Connect failed:" / "COM error:" are the stable prefixes produced by
    /// Com1CConnector.LoadFile; data errors ("Account … not found …",
    /// "Bank account … not found …", "Doc #N: …") stay permanent.</summary>
    public static bool IsTransient(string? error)
    {
        if (string.IsNullOrEmpty(error)) return false;

        if (error.StartsWith("Connect failed:", StringComparison.Ordinal))
            return true;

        // "COM error:" wraps any exception from the load; only treat it as
        // transient when the inner text carries a connection/session marker —
        // a genuine data crash inside COM should not loop forever.
        bool comWrapped = error.StartsWith("COM error:", StringComparison.Ordinal);

        foreach (var marker in _transientMarkers)
        {
            if (error.IndexOf(marker, StringComparison.OrdinalIgnoreCase) >= 0)
                return true;
        }

        // A COM-wrapped error with no recognizable marker: retry anyway —
        // the session is torn down after it, so the retry runs on a fresh
        // connection; dedupe makes the retry idempotent.
        return comWrapped;
    }
}
