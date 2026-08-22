namespace DeviceInterface.Webscan;

/// <summary>
/// Coordinates Webscan record acceptance with session shutdown.
///
/// A file watcher callback can arrive on a thread-pool thread just as the
/// session is stopping. Admission and invalidation must therefore be one
/// atomic operation: work admitted before invalidation is drained, while work
/// observed after invalidation is rejected.
/// </summary>
public sealed class WebscanAcceptanceTracker
{
    private readonly object _lock = new();
    private readonly HashSet<Task> _inFlight = [];
    private int _sessionGeneration;

    public void BeginSession(int generation)
    {
        lock (_lock)
            _sessionGeneration = generation;
    }

    public bool IsCurrent(int generation)
    {
        lock (_lock)
            return generation == _sessionGeneration;
    }

    public bool TryAdmit(
        int generation,
        Func<bool> isSessionRunning,
        Func<Task> acceptance)
    {
        Task acceptanceTask;
        lock (_lock)
        {
            if (generation != _sessionGeneration || !isSessionRunning())
                return false;

            acceptanceTask = Task.Run(acceptance);
            _inFlight.Add(acceptanceTask);
        }

        _ = acceptanceTask.ContinueWith(
            completed =>
            {
                lock (_lock)
                    _inFlight.Remove(completed);
            },
            TaskScheduler.Default);
        return true;
    }

    /// <summary>
    /// Invalidates callbacks and returns the exact set of work admitted before
    /// invalidation. Callers must await the returned tasks before closing the
    /// session-owned writer.
    /// </summary>
    public Task[] InvalidateAndCapture()
    {
        lock (_lock)
        {
            _sessionGeneration++;
            return _inFlight.ToArray();
        }
    }
}