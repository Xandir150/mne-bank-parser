using System;
using System.IO;
using System.Text;
using Microsoft.Extensions.Logging;

// ──────────────────────────────────────────────────────────────
// Simple file logger (no external dependencies)
// ──────────────────────────────────────────────────────────────
public class FileLoggerProvider : ILoggerProvider
{
    private readonly string _filePath;
    private readonly object _lock = new();

    public FileLoggerProvider(string filePath) => _filePath = filePath;

    public ILogger CreateLogger(string categoryName) =>
        new FileLogger(_filePath, categoryName, _lock);

    public void Dispose() { }
}

public class FileLogger : ILogger
{
    private readonly string _filePath;
    private readonly string _category;
    private readonly object _lock;

    public FileLogger(string filePath, string category, object lockObj)
    {
        _filePath = filePath;
        _category = category;
        _lock = lockObj;
    }

    public IDisposable? BeginScope<TState>(TState state) where TState : notnull => null;
    public bool IsEnabled(LogLevel logLevel) => logLevel >= LogLevel.Information;

    public void Log<TState>(LogLevel logLevel, EventId eventId, TState state,
        Exception? exception, Func<TState, Exception?, string> formatter)
    {
        if (!IsEnabled(logLevel)) return;
        var msg = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{logLevel}] {_category}: {formatter(state, exception)}";
        if (exception != null)
            msg += Environment.NewLine + exception;
        lock (_lock)
        {
            File.AppendAllText(_filePath, msg + Environment.NewLine, Encoding.UTF8);
        }
    }
}
