namespace DwgOfficeToolbox.App.Models;

public sealed class InputItem
{
    public required string Path { get; init; }
    public required string Type { get; init; }
    public bool IsFolder { get; init; }
}
