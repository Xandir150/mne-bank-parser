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
    public int ScanIntervalSec { get; set; } = 30;
    public bool PostDocuments { get; set; } = false;
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
