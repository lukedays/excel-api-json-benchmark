using Microsoft.Office.Interop.Excel;

namespace CsAddIn;

/// <summary>Desliga tela, eventos e cálculo automático enquanto vivo; restaura o estado anterior no Dispose.</summary>
public sealed class QuietExcel : IDisposable
{
    private readonly Application _app;
    private readonly bool _screenUpdating;
    private readonly bool _enableEvents;
    private readonly XlCalculation _calculation;

    private QuietExcel(Application app)
    {
        _app = app;
        _screenUpdating = app.ScreenUpdating;
        _enableEvents = app.EnableEvents;
        _calculation = app.Calculation;
        app.ScreenUpdating = false;
        app.EnableEvents = false;
        app.Calculation = XlCalculation.xlCalculationManual;
    }

    public static QuietExcel Enter(Application app) => new(app);

    public void Dispose()
    {
        _app.Calculation = _calculation;
        _app.EnableEvents = _enableEvents;
        _app.ScreenUpdating = _screenUpdating;
    }
}
