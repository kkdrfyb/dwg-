using System.IO;
using ClosedXML.Excel;
using DwgOfficeToolbox.App.Models;
using System.Text;

namespace DwgOfficeToolbox.App.Services;

public static class ExportService
{
    public static void ExportCsv(string filePath, IEnumerable<MatchResult> results)
    {
        var sb = new StringBuilder();
        sb.AppendLine("文件名,对象类型,图层,关键字,匹配内容");
        foreach (var r in results)
        {
            sb.AppendLine($"{Escape(r.FileName)},{Escape(r.ObjectType)},{Escape(r.Layer)},{Escape(r.Keyword)},{Escape(r.Content)}");
        }

        var data = Encoding.UTF8.GetPreamble().Concat(Encoding.UTF8.GetBytes(sb.ToString())).ToArray();
        File.WriteAllBytes(filePath, data);
    }

    public static void ExportXlsx(string filePath, IEnumerable<MatchResult> results)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("扫描结果");
        ws.Cell(1, 1).Value = "文件名";
        ws.Cell(1, 2).Value = "对象类型";
        ws.Cell(1, 3).Value = "图层";
        ws.Cell(1, 4).Value = "关键字";
        ws.Cell(1, 5).Value = "匹配内容";

        var row = 2;
        foreach (var r in results)
        {
            ws.Cell(row, 1).Value = r.FileName;
            ws.Cell(row, 2).Value = r.ObjectType;
            ws.Cell(row, 3).Value = r.Layer;
            ws.Cell(row, 4).Value = r.Keyword;
            ws.Cell(row, 5).Value = r.Content;
            row++;
        }

        ws.Columns().AdjustToContents();
        wb.SaveAs(filePath);
    }

    private static string Escape(string value)
    {
        var v = value?.Replace("\"", "\"\"") ?? string.Empty;
        return $"\"{v}\"";
    }
}
