using System;
using System.IO;
using System.Text;
using Microsoft.Extensions.Logging;

// ──────────────────────────────────────────────────────────────
// File logger with persistent stream (no per-line file open/close).
// ──────────────────────────────────────────────────────────────
public class FileLoggerProvider : ILoggerProvider
{
    private readonly string _filePath;
    private readonly LogLevel _minLevel;
    private readonly StreamWriter _writer;
    private readonly object _lock = new();
    private bool _disposed;

    public FileLoggerProvider(string filePath, LogLevel minLevel = LogLevel.Information)
    {
        _filePath = filePath;
        _minLevel = minLevel;
        var dir = Path.GetDirectoryName(filePath);
        if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
        var fs = new FileStream(filePath, FileMode.Append, FileAccess.Write,
            FileShare.Read, bufferSize: 4096, useAsync: false);
        _writer = new StreamWriter(fs, new UTF8Encoding(false)) { AutoFlush = true };
    }

    public ILogger CreateLogger(string categoryName) =>
        new FileLogger(this, categoryName);

    internal void Write(string line)
    {
        lock (_lock)
        {
            if (_disposed) return;
            try { _writer.WriteLine(line); } catch { }
        }
    }

    internal LogLevel MinLevel => _minLevel;

    public void Dispose()
    {
        lock (_lock)
        {
            if (_disposed) return;
            _disposed = true;
            try { _writer.Flush(); } catch { }
            try { _writer.Dispose(); } catch { }
        }
    }
}

public class FileLogger : ILogger
{
    private readonly FileLoggerProvider _provider;
    private readonly string _category;

    public FileLogger(FileLoggerProvider provider, string category)
    {
        _provider = provider;
        _category = category;
    }

    public IDisposable? BeginScope<TState>(TState state) where TState : notnull => null;
    public bool IsEnabled(LogLevel logLevel) => logLevel >= _provider.MinLevel;

    public void Log<TState>(LogLevel logLevel, EventId eventId, TState state,
        Exception? exception, Func<TState, Exception?, string> formatter)
    {
        if (!IsEnabled(logLevel)) return;
        var msg = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{logLevel}] {_category}: {formatter(state, exception)}";
        if (exception != null)
            msg += Environment.NewLine + exception;
        _provider.Write(msg);
    }
}
