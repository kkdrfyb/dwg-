namespace DwgOfficeToolbox.App.Models;

public sealed class DwgScanTarget
{
    public string DwgPath { get; init; } = string.Empty;
    public string DxfPath { get; init; } = string.Empty;
    public string OutputRoot { get; init; } = string.Empty;
}
