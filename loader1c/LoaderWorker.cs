using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Hosting;

// ──────────────────────────────────────────────────────────────
// Background worker service
// ──────────────────────────────────────────────────────────────
public class LoaderWorker : BackgroundService
{
    private readonly LoaderConfig _config;
    private readonly Com1CConnector _com;
    private readonly ILogger<LoaderWorker> _logger;

    public LoaderWorker(LoaderConfig config, Com1CConnector com,
        ILogger<LoaderWorker> logger)
    {
        _config = config;
        _com = com;
        _logger = logger;
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        _logger.LogInformation("Loader1C service starting...");

        // Use cached mapping for fast start; rescan databases in the background
        bool hasCached = _com.LoadCachedMapping();
        if (hasCached)
        {
            _logger.LogInformation("Using cached mapping, rescan will run in background");
            _ = Task.Run(() => { try { _com.ScanDatabases(); } catch (Exception ex) { _logger.LogError(ex, "Background rescan failed"); } }, stoppingToken);
        }
        else
        {
            _logger.LogInformation("No cached mapping, scanning databases...");
            await Task.Run(() => _com.ScanDatabases(), stoppingToken);
        }

        _logger.LogInformation("Watching {Dir} (interval: {Sec}s)",
            _config.OutputDir, _config.ScanIntervalSec);

        while (!stoppingToken.IsCancellationRequested)
        {
            try
            {
                ScanAndLoad();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error in scan cycle");
            }

            await Task.Delay(_config.ScanIntervalSec * 1000, stoppingToken);
        }
    }

    private void ScanAndLoad()
    {
        if (!Directory.Exists(_config.OutputDir)) return;

        var txtFiles = Directory.GetFiles(_config.OutputDir, "*.txt",
            SearchOption.AllDirectories);

        foreach (var filePath in txtFiles)
        {
            var fileName = Path.GetFileName(filePath);

            // Skip files that look like garbage (unknown_*)
            if (fileName.StartsWith("unknown_")) continue;

            // Only skip if .loaded marker exists — delete marker to re-process
            if (File.Exists(filePath + ".loaded")) continue;

            // Re-check file exists (may have been moved between GetFiles and now)
            if (!File.Exists(filePath)) continue;

            _logger.LogInformation("New file: {File}", filePath);

            try
            {
                var parsed = BankFileParser.Parse(filePath);
                if (parsed.Documents.Count == 0)
                {
                    _logger.LogWarning("No documents in {File}, skipping", fileName);
                    continue;
                }

                var result = _com.LoadFile(parsed);

                if (result.Created > 0)
                {
                    _logger.LogInformation("Loaded {File}: {Created} created, {Skipped} skipped",
                        fileName, result.Created, result.Skipped);

                    // Write .loaded marker (delete this file to re-process)
                    File.WriteAllText(filePath + ".loaded",
                        $"Loaded: {DateTime.Now:yyyy-MM-dd HH:mm:ss}, " +
                        $"Created: {result.Created}, Skipped: {result.Skipped}");

                    // Move to loaded dir (preserve subfolder structure)
                    var relDir = Path.GetRelativePath(_config.OutputDir,
                        Path.GetDirectoryName(filePath)!);
                    var destDir = Path.Combine(_config.LoadedDir, relDir);
                    Directory.CreateDirectory(destDir);
                    var destPath = Path.Combine(destDir, fileName);
                    if (File.Exists(destPath))
                        destPath = Path.Combine(destDir,
                            Path.GetFileNameWithoutExtension(fileName) +
                            "_" + DateTime.Now.ToString("yyyyMMddHHmmss") +
                            Path.GetExtension(fileName));
                    File.Move(filePath, destPath);
                }
                else if (result.Skipped > 0 && result.Error == null)
                {
                    // All documents already exist in 1C — mark as loaded, don't retry
                    _logger.LogInformation("All {Skipped} docs in {File} already exist, marking as loaded",
                        result.Skipped, fileName);
                    File.WriteAllText(filePath + ".loaded",
                        $"AllSkipped: {DateTime.Now:yyyy-MM-dd HH:mm:ss}, " +
                        $"Skipped: {result.Skipped}");
                }
                else if (result.Error != null)
                {
                    // Real error — rename to .error so we don't retry forever
                    _logger.LogWarning("Error loading {File}: {Error}", fileName, result.Error);
                    try { File.Move(filePath, filePath + ".error"); } catch { }
                }
                else
                {
                    // Created=0, Skipped=0, no error — empty result, skip
                    _logger.LogWarning("No result for {File}, skipping", fileName);
                    File.WriteAllText(filePath + ".loaded",
                        $"Empty: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error processing {File}", filePath);
            }
        }
    }
}
