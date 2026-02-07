namespace DwgOfficeToolbox.App.Models;

public sealed class DxfTextItem
{
    public required string ObjectType { get; init; }
    public required string Layer { get; init; }
    public required string Text { get; init; }
}
