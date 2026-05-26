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
