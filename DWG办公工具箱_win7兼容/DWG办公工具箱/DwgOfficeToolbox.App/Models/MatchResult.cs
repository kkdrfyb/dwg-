namespace DwgOfficeToolbox.App.Models;

public sealed class MatchResult
{
    public required string FileName { get; init; }
    public required string ObjectType { get; init; }
    public required string Layer { get; init; }
    public required string Keyword { get; init; }
    public required string Content { get; init; }
    public required string FilePath { get; init; }
    public required string DwgPath { get; init; }
}
