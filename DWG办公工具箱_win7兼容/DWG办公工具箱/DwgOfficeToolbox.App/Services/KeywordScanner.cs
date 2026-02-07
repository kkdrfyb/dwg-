using System.IO;
using DwgOfficeToolbox.App.Models;

namespace DwgOfficeToolbox.App.Services;

public static class KeywordScanner
{
    public static List<MatchResult> ScanStructured(
        string filePath,
        string dwgPath,
        IEnumerable<DxfTextItem> items,
        IReadOnlyList<string> keywords)
    {
        var results = new List<MatchResult>();
        var fileName = Path.GetFileName(filePath);

        if (keywords.Count == 0)
        {
            foreach (var item in items)
            {
                results.Add(new MatchResult
                {
                    FileName = fileName,
                    FilePath = filePath,
                    DwgPath = dwgPath,
                    ObjectType = item.ObjectType,
                    Layer = item.Layer,
                    Keyword = "全部",
                    Content = item.Text
                });
            }
            return results;
        }

        foreach (var item in items)
        {
            if (string.IsNullOrWhiteSpace(item.Text))
            {
                continue;
            }

            foreach (var keyword in keywords)
            {
                if (item.Text.Contains(keyword, StringComparison.OrdinalIgnoreCase))
                {
                    results.Add(new MatchResult
                    {
                    FileName = fileName,
                    FilePath = filePath,
                    DwgPath = dwgPath,
                    ObjectType = item.ObjectType,
                    Layer = item.Layer,
                    Keyword = keyword,
                    Content = item.Text
                });
                }
            }
        }

        return results;
    }

    public static List<MatchResult> ScanPlainText(
        string filePath,
        string dwgPath,
        IReadOnlyList<string> keywords,
        Func<IReadOnlyList<string>, HashSet<string>> matchFunc)
    {
        var results = new List<MatchResult>();
        var fileName = Path.GetFileName(filePath);

        if (keywords.Count == 0)
        {
            results.Add(new MatchResult
            {
                FileName = fileName,
                FilePath = filePath,
                DwgPath = dwgPath,
                ObjectType = "未知",
                Layer = "-",
                Keyword = "全部",
                Content = "(纯文本匹配)"
            });
            return results;
        }

        var matched = matchFunc(keywords);
        foreach (var keyword in matched)
        {
            results.Add(new MatchResult
            {
                FileName = fileName,
                FilePath = filePath,
                DwgPath = dwgPath,
                ObjectType = "未知",
                Layer = "-",
                Keyword = keyword,
                Content = "(纯文本匹配)"
            });
        }

        return results;
    }
}
