using System.IO;
using System.Windows;
using DwgOfficeToolbox.App.Services;

namespace DwgOfficeToolbox.App;

public partial class SettingsWindow : Window
{
    public AppSettings Settings { get; private set; }

    public SettingsWindow(AppSettings settings)
    {
        InitializeComponent();
        Settings = new AppSettings
        {
            OdaExePath = settings.OdaExePath,
            OutputFolderName = settings.OutputFolderName,
            OutputVersion = settings.OutputVersion,
            OutputFormat = settings.OutputFormat,
            InputFilter = settings.InputFilter,
            MaxConvertConcurrency = settings.MaxConvertConcurrency,
            LargeFileThresholdMB = settings.LargeFileThresholdMB,
            MaxParseConcurrency = settings.MaxParseConcurrency,
            EnableTextCache = settings.EnableTextCache,
            KeepConvertedDxf = settings.KeepConvertedDxf,
            ClearOutputBeforeConvert = settings.ClearOutputBeforeConvert
        };
        DataContext = Settings;
    }

    private void BrowseOdaPath_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "ODAFileConverter.exe|ODAFileConverter.exe",
            Title = "选择 ODAFileConverter.exe"
        };
        if (dialog.ShowDialog() == true)
        {
            Settings.OdaExePath = dialog.FileName;
        }
    }

    private void SaveButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = true;
        Close();
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}
