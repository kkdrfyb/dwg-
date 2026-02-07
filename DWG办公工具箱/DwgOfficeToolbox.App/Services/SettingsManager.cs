using System.IO;
using System.Text.Json;

namespace DwgOfficeToolbox.App.Services;

public sealed class AppSettings
{
    public string OdaExePath { get; set; } = string.Empty;
    public string OutputFolderName { get; set; } = "output";
    public string OutputVersion { get; set; } = "ACAD2018";
    public string OutputFormat { get; set; } = "DXF";
    public string InputFilter { get; set; } = "*.dwg";
    public int MaxConvertConcurrency { get; set; } = 4;
    public int LargeFileThresholdMB { get; set; } = 50;
    public int MaxParseConcurrency { get; set; } = 6;
    public bool EnableTextCache { get; set; } = true;
    public bool KeepConvertedDxf { get; set; } = true;
    public bool ClearOutputBeforeConvert { get; set; } = false;
}

public static class SettingsManager
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true
    };

    public static string AppDataDir =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CadTextSearch");

    public static string SettingsPath => Path.Combine(AppDataDir, "settings.json");

    public static AppSettings Load()
    {
        Directory.CreateDirectory(AppDataDir);
        if (!File.Exists(SettingsPath))
        {
            var settings = CreateDefault();
            Save(settings);
            return settings;
        }

        try
        {
            var json = File.ReadAllText(SettingsPath);
            var settings = JsonSerializer.Deserialize<AppSettings>(json);
            if (settings == null)
            {
                settings = CreateDefault();
                Save(settings);
            }
            return Normalize(settings);
        }
        catch
        {
            var settings = CreateDefault();
            Save(settings);
            return settings;
        }
    }

    public static void Save(AppSettings settings)
    {
        Directory.CreateDirectory(AppDataDir);
        var json = JsonSerializer.Serialize(settings, JsonOptions);
        File.WriteAllText(SettingsPath, json);
    }

    private static AppSettings Normalize(AppSettings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.OdaExePath) || !File.Exists(settings.OdaExePath))
        {
            settings.OdaExePath = GetDefaultOdaPath();
        }

        if (string.IsNullOrWhiteSpace(settings.OutputFolderName))
        {
            settings.OutputFolderName = "output";
        }

        if (string.IsNullOrWhiteSpace(settings.OutputVersion))
        {
            settings.OutputVersion = "ACAD2018";
        }

        if (string.IsNullOrWhiteSpace(settings.OutputFormat))
        {
            settings.OutputFormat = "DXF";
        }

        if (string.IsNullOrWhiteSpace(settings.InputFilter))
        {
            settings.InputFilter = "*.dwg";
        }

        if (settings.MaxConvertConcurrency <= 0)
        {
            settings.MaxConvertConcurrency = 4;
        }

        if (settings.LargeFileThresholdMB <= 0)
        {
            settings.LargeFileThresholdMB = 50;
        }

        if (settings.MaxParseConcurrency <= 0)
        {
            settings.MaxParseConcurrency = 6;
        }

        return settings;
    }

    private static AppSettings CreateDefault()
    {
        return new AppSettings
        {
            OdaExePath = GetDefaultOdaPath()
        };
    }

    private static string GetDefaultOdaPath()
    {
        return Path.Combine(AppContext.BaseDirectory, "ODAFileConverter", "ODAFileConverter.exe");
    }
}
