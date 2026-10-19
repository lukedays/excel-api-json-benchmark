using System.Diagnostics;
using System.Text;

namespace CsAddIn;

/// <summary>Cronômetro de fases da última carga, lido pela função CsLastTimings.</summary>
public static class PhaseTimer
{
    private static readonly Stopwatch Clock = new();
    private static readonly StringBuilder Report = new();
    private static TimeSpan _lastMark;

    public static string LastReport { get; private set; } = "";

    public static void Start()
    {
        Report.Clear();
        _lastMark = TimeSpan.Zero;
        Clock.Restart();
    }

    public static void Mark(string phase)
    {
        var now = Clock.Elapsed;
        Report.Append($"{phase} {(now - _lastMark).TotalSeconds:F3} ");
        _lastMark = now;
    }

    public static void Stop() => LastReport = Report.ToString().TrimEnd();
}
