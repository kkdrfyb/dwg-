namespace DwgOfficeToolbox.App.Models;

public sealed class InputItem
{
    public string Path { get; init; } = string.Empty;
    public string Type { get; init; } = string.Empty;
    public bool IsFolder { get; init; }
}
