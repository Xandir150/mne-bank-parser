using System.Collections.Generic;

// ──────────────────────────────────────────────────────────────
// Configuration
// ──────────────────────────────────────────────────────────────
public class LoaderConfig
{
    public string Server { get; set; } = "navus-server";
    public string User { get; set; } = "";
    public string Password { get; set; } = "";
    public string OutputDir { get; set; } = "";
    public string LoadedDir { get; set; } = "";
    public string LogFile { get; set; } = "";
    /// <summary>Minimum log level for the file logger: Trace/Debug/Information/Warning/Error.</summary>
    public string LogLevel { get; set; } = "Information";
    public int ScanIntervalSec { get; set; } = 30;
    public bool PostDocuments { get; set; } = false;
    /// <summary>Working set threshold in MB. When exceeded after a file load the worker
    /// gracefully exits so service recovery restarts it fresh, releasing native COM buffers
    /// the .NET GC can't reclaim. Set to 0 to disable. Default 10 GB — high cap as a safety
    /// net only, not for routine recycling.</summary>
    public int MaxWorkingSetMB { get; set; } = 10240;
    /// <summary>Skip background DB rescan if cached mapping is younger than this.
    /// New banks/orgs added in 1C will only be discovered after this period. Default 12h.
    /// Set to 0 to always rescan on startup.</summary>
    public int RescanAfterHours { get; set; } = 12;
    /// <summary>How many times a file with a transient 1C error (session killed,
    /// connection lost) is retried before giving up (.error). 0 = old behavior:
    /// every error is permanent.</summary>
    public int MaxTransientAttempts { get; set; } = 8;
    /// <summary>Cap on the exponential backoff between transient retries, in scan
    /// cycles (backoff = ScanIntervalSec × min(2^attempt, this)). Default 32 ≈ 16 min.</summary>
    public int MaxRetryBackoffCycles { get; set; } = 32;
    /// <summary>Pick up accounts.config.json edits automatically (mtime check each
    /// scan cycle) — no service restart needed. false = load only at startup.</summary>
    public bool ConfigReloadEnabled { get; set; } = true;
    /// <summary>After the mapping changes, automatically re-queue .error files whose
    /// account became mapped (rename back to .txt).</summary>
    public bool AutoRequeueErrors { get; set; } = true;
    /// <summary>Max automatic re-queues per file (tracked in .error.info) — prevents
    /// .txt ↔ .error cycling when the mapped DB still lacks the bank account.</summary>
    public int MaxAutoRequeues { get; set; } = 3;
    /// <summary>What to do when a statement arrives for an account missing from the
    /// mapping: "off" = just .error; "suggest" = search all DBs, write a ready-made
    /// config snippet to discovered.accounts.txt; "append" = also add the entry to
    /// accounts.config.json automatically (single unambiguous hit only).</summary>
    public string AutoDiscovery { get; set; } = "suggest";
    /// <summary>How long a failed discovery (account found nowhere) suppresses
    /// repeat sweeps for the same account. Reset by config edits.</summary>
    public int DiscoveryNegativeCacheHours { get; set; } = 12;
    /// <summary>Max discovery sweeps per scan cycle (each sweep = 1 connect per DB).</summary>
    public int MaxDiscoverySweepsPerCycle { get; set; } = 1;
    /// <summary>Local hour (0-23) at which full discovery sweeps (a COM connect to every
    /// DB in <see cref="Databases"/>) are allowed to run — they block the scan loop for
    /// minutes, so by default they wait for an off-hours window instead of firing
    /// immediately for every newly-seen unknown account. discover.trigger still forces
    /// an immediate sweep regardless of this setting. -1 restores the legacy behavior
    /// (sweep right away, every cycle, for any pending unknown account).</summary>
    public int DiscoveryScheduleHour { get; set; } = -1;
    public List<string> Databases { get; set; } = new();
}

// ──────────────────────────────────────────────────────────────
// Account mapping: bank account number → (database, orgName)
// ──────────────────────────────────────────────────────────────
public class AccountMapping
{
    public string Database { get; set; } = "";
    public string OrgName { get; set; } = "";
    public string Inn { get; set; } = "";
}

// ──────────────────────────────────────────────────────────────
// Load result
// ──────────────────────────────────────────────────────────────
public class LoadResult
{
    public bool Success { get; set; }
    public int Created { get; set; }
    public int Skipped { get; set; }
    public string? Error { get; set; }
}
