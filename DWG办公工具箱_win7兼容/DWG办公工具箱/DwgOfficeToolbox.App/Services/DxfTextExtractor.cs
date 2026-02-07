using System.Text.RegularExpressions;
using DwgOfficeToolbox.App.Models;
using netDxf;
using netDxf.Entities;

namespace DwgOfficeToolbox.App.Services;

public static class DxfTextExtractor
{
    private static readonly Regex MTextStackRegex = new(@"\\S([^;]*);", RegexOptions.Compiled);
    private static readonly Regex MTextParamRegex = new(@"\\[ACHQTFW][^;]*;", RegexOptions.Compiled);
    private static readonly Regex MTextSimpleRegex = new(@"\\[LlOoKk]", RegexOptions.Compiled);
    private static readonly Regex MTextUnicodeRegex = new(@"\\U\+([0-9A-Fa-f]{4})", RegexOptions.Compiled);

    public static List<DxfTextItem> Extract(string dxfPath)
    {
        var doc = DxfDocument.Load(dxfPath);
        var results = new List<DxfTextItem>();

        foreach (var text in doc.Entities.Texts)
        {
            AddIfNotEmpty(results, "TEXT", text.Layer?.Name, Normalize(text.Value));
        }

        foreach (var mtext in doc.Entities.MTexts)
        {
            AddIfNotEmpty(results, "MTEXT", mtext.Layer?.Name, Normalize(GetMTextValue(mtext)));
        }

        foreach (var insert in doc.Entities.Inserts)
        {
            ExtractInsertText(results, insert);
        }

        return results;
    }

    private static void ExtractInsertText(List<DxfTextItem> results, Insert insert)
    {
        var insertLayer = insert.Layer?.Name;
        var block = insert.Block;

        if (!string.IsNullOrWhiteSpace(block?.Name))
        {
            AddIfNotEmpty(results, "INSERT", insertLayer, Normalize(block.Name));
        }

        if (insert.Attributes != null && insert.Attributes.Count > 0)
        {
            foreach (var attr in insert.Attributes)
            {
                var layerName = attr.Layer?.Name ?? insertLayer;
                AddIfNotEmpty(results, "ATTRIB", layerName, Normalize(attr.Value));
            }
        }

        if (block == null)
        {
            return;
        }

        foreach (var blockEntity in block.Entities)
        {
            switch (blockEntity)
            {
                case Text text:
                    AddIfNotEmpty(results, "BLOCK_TEXT", text.Layer?.Name ?? insertLayer, Normalize(text.Value));
                    break;
                case MText mtext:
                    AddIfNotEmpty(results, "BLOCK_MTEXT", mtext.Layer?.Name ?? insertLayer, Normalize(GetMTextValue(mtext)));
                    break;
            }
        }

        if ((insert.Attributes == null || insert.Attributes.Count == 0) && block.AttributeDefinitions != null)
        {
            foreach (var attdef in block.AttributeDefinitions.Values)
            {
                var layerName = attdef.Layer?.Name ?? insertLayer;
                AddIfNotEmpty(results, "ATTDEF", layerName, Normalize(attdef.Value));
            }
        }
    }

    private static void AddIfNotEmpty(List<DxfTextItem> results, string type, string? layer, string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return;
        }

        results.Add(new DxfTextItem
        {
            ObjectType = type,
            Layer = string.IsNullOrWhiteSpace(layer) ? "-" : layer,
            Text = text
        });
    }

    private static string Normalize(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        var text = value;
        text = text.Replace("\\P", " ")
            .Replace("\\X", " ")
            .Replace("\\N", " ")
            .Replace("\\n", " ")
            .Replace("\\~", " ");

        text = MTextStackRegex.Replace(text, "$1");
        text = MTextParamRegex.Replace(text, string.Empty);
        text = MTextSimpleRegex.Replace(text, string.Empty);
        text = MTextUnicodeRegex.Replace(text, match =>
        {
            if (int.TryParse(match.Groups[1].Value, System.Globalization.NumberStyles.HexNumber, null, out var code))
            {
                return char.ConvertFromUtf32(code);
            }
            return string.Empty;
        });

        text = text.Replace("{", string.Empty).Replace("}", string.Empty);
        text = text.Replace("\\\\", "\\");
        return text.Trim();
    }

    private static string GetMTextValue(MText mtext)
    {
        var plainProp = typeof(MText).GetProperty("PlainText");
        if (plainProp != null && plainProp.PropertyType == typeof(string))
        {
            var plain = plainProp.GetValue(mtext) as string;
            if (!string.IsNullOrWhiteSpace(plain))
            {
                return plain;
            }
        }

        var textProp = typeof(MText).GetProperty("Text");
        if (textProp != null && textProp.PropertyType == typeof(string))
        {
            var text = textProp.GetValue(mtext) as string;
            if (!string.IsNullOrWhiteSpace(text))
            {
                return text;
            }
        }

        return mtext.Value;
    }
}
