namespace DwgOfficeToolbox.App.Models;

public sealed class DwgScanTarget
{
    public required string DwgPath { get; init; }
    public required string DxfPath { get; init; }
    public required string OutputRoot { get; init; }
}
