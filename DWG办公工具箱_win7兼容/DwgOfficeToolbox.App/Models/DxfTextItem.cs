namespace DwgOfficeToolbox.App.Models;

public sealed class DxfTextItem
{
    public string ObjectType { get; init; } = string.Empty;
    public string Layer { get; init; } = string.Empty;
    public string Text { get; init; } = string.Empty;
}
