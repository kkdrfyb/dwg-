namespace DwgOfficeToolbox.App.Models;

public sealed class MatchResult
{
    public string FileName { get; init; } = string.Empty;
    public string ObjectType { get; init; } = string.Empty;
    public string Layer { get; init; } = string.Empty;
    public string Keyword { get; init; } = string.Empty;
    public string Content { get; init; } = string.Empty;
    public string FilePath { get; init; } = string.Empty;
    public string DwgPath { get; init; } = string.Empty;
}
