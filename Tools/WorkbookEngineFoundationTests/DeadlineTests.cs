using System;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using Abovo.WorkbookEngines;

static class DeadlineTests
{
    sealed class HeldBackend : IWorkbookCalculationBackend
    {
        internal readonly ManualResetEventSlim Release = new ManualResetEventSlim();
        internal readonly TaskCompletionSource<bool> Entered = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        internal readonly TaskCompletionSource<bool> Disposed = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        internal bool HoldOpen;
        internal int Reads, Owner;
        public string Name => "Held native operation";
        public string Version => "1";
        public void OpenReadOnly(string path, WorkbookEngineOptions options)
        {
            Owner = Thread.CurrentThread.ManagedThreadId;
            if (HoldOpen) Hold();
        }
        void Hold() { Entered.TrySetResult(true); Release.Wait(); }
        public void Calculate(WorkbookCalculationKind kind) => Hold();
        public WorkbookValueBlock Read(WorkbookReadArea area) { Reads++; return new WorkbookValueBlock(area, new object[,] { { 1d } }); }
        public void Dispose()
        {
            if (Thread.CurrentThread.ManagedThreadId != Owner) throw new Exception("Wrong disposal thread");
            Disposed.TrySetResult(true);
        }
    }
    static int assertions;
    static void Check(bool condition, string message)
    {
        if (!condition) throw new Exception(message);
        Console.WriteLine("DEADLINE PASS " + (++assertions) + " " + message);
    }
    static async Task ExpectTimeout(Task operation)
    {
        if (await Task.WhenAny(operation, Task.Delay(5000)) != operation) throw new Exception("Watchdog itself hung");
        try { await operation; } catch (TimeoutException) { return; }
        throw new Exception("Expected native-operation timeout");
    }
    internal static async Task Run(string path)
    {
        var policy = new WorkbookEngineOptions(WorkbookEnginePreference.Automatic, false, false, 250);
        var backend = new HeldBackend();
        int factories = 0;
        var session = await WorkbookCalculationSession.OpenAsync(path, policy, default(CancellationToken), _ => { factories++; return backend; });
        try
        {
            var operation = session.CalculateAndReadAsync(0, WorkbookCalculationKind.Full, new[] { new WorkbookReadArea("Data", 0, 0, 1, 1) });
            await backend.Entered.Task;
            await ExpectTimeout(operation);
            Check(!backend.Disposed.Task.IsCompleted && session.NativeCleanupCompletion != null,
                "timeout quarantines but never disposes an in-flight native object");
            await ExpectTimeout(session.CloseAsync());
            Check(!session.NativeCleanupCompletion.IsCompleted, "close has a bounded wait and keeps real cleanup pending");
            bool refused = false;
            try { session.InvalidateResults(); } catch (ObjectDisposedException) { refused = true; }
            Check(refused && factories == 1, "timed-out session refuses reuse without silent fallback");
        }
        finally { backend.Release.Set(); await session.NativeCleanupCompletion; }
        Check(backend.Disposed.Task.IsCompleted && backend.Reads == 0, "late calculation cannot read/publish; cleanup resumes on the owner");

        backend = new HeldBackend { HoldOpen = true }; factories = 0;
        var opening = WorkbookCalculationSession.OpenAsync(path, policy, default(CancellationToken), _ => { factories++; return backend; });
        try
        {
            await backend.Entered.Task; await ExpectTimeout(opening);
            Check(factories == 1 && !backend.Disposed.Task.IsCompleted, "hung opening returns promptly and does not start a competing fallback");
        }
        finally { backend.Release.Set(); }
        if (await Task.WhenAny(backend.Disposed.Task, Task.Delay(5000)) != backend.Disposed.Task) throw new Exception("Late opening did not clean up");
        Check(factories == 1, "late opening cannot initiate DevExpress after timeout");
        Console.WriteLine("DEADLINE ASSERTIONS=" + assertions);
    }
}
