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
    private readonly HashSet<string> _processedFiles = new();

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

        // Always scan databases on startup to refresh mapping
        await Task.Run(() => _com.ScanDatabases(), stoppingToken);

        // Track already processed files
        var loadedDir = _config.LoadedDir;
        Directory.CreateDirectory(loadedDir);

        // Scan loaded dir to know what's already done
        if (Directory.Exists(loadedDir))
        {
            foreach (var f in Directory.GetFiles(loadedDir, "*.txt",
                         SearchOption.AllDirectories))
                _processedFiles.Add(Path.GetFileName(f));
        }

        // Also check for .loaded marker files in output dir
        if (Directory.Exists(_config.OutputDir))
        {
            foreach (var f in Directory.GetFiles(_config.OutputDir, "*.txt.loaded",
                         SearchOption.AllDirectories))
            {
                // marker "xxx.txt.loaded" means "xxx.txt" was processed
                var originalName = Path.GetFileNameWithoutExtension(f); // strips .loaded → xxx.txt
                _processedFiles.Add(originalName);
            }
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
            if (_processedFiles.Contains(fileName)) continue;

            // Skip files that look like garbage (unknown_*)
            if (fileName.StartsWith("unknown_")) continue;

            // Skip if .loaded marker exists (already processed)
            if (File.Exists(filePath + ".loaded")) continue;

            _logger.LogInformation("New file: {File}", filePath);

            try
            {
                var parsed = BankFileParser.Parse(filePath);
                if (parsed.Documents.Count == 0)
                {
                    _logger.LogWarning("No documents in {File}, skipping", fileName);
                    _processedFiles.Add(fileName);
                    continue;
                }

                var result = _com.LoadFile(parsed);

                if (result.Created > 0)
                {
                    _logger.LogInformation("Loaded {File}: {Created} documents created",
                        fileName, result.Created);

                    // Write .loaded marker FIRST (survives crash/restart)
                    File.WriteAllText(filePath + ".loaded",
                        $"Loaded: {DateTime.Now:yyyy-MM-dd HH:mm:ss}, " +
                        $"Documents: {result.Created}");

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
                else
                {
                    _logger.LogWarning("Failed to load {File}: {Error}",
                        fileName, result.Error);
                }

                _processedFiles.Add(fileName);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error processing {File}", filePath);
                _processedFiles.Add(fileName); // don't retry endlessly
            }
        }
    }
}
