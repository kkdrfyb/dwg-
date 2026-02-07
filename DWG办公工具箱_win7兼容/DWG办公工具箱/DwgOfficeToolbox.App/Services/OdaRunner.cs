using System.Diagnostics;

namespace DwgOfficeToolbox.App.Services;

public static class OdaRunner
{
    public static async Task<int> ConvertAsync(
        string odaExePath,
        string inputFolder,
        string outputFolder,
        string outputVersion,
        string outputFormat,
        string inputFilter,
        bool recurse,
        bool audit,
        CancellationToken cancellationToken,
        Action<string>? log)
    {
        var recurseFlag = recurse ? "1" : "0";
        var auditFlag = audit ? "1" : "0";
        var args = $"\"{inputFolder}\" \"{outputFolder}\" {outputVersion} {outputFormat} {recurseFlag} {auditFlag} \"{inputFilter}\"";

        var psi = new ProcessStartInfo
        {
            FileName = odaExePath,
            Arguments = args,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            WindowStyle = ProcessWindowStyle.Hidden
        };

        using var process = new Process { StartInfo = psi, EnableRaisingEvents = true };
        var tcs = new TaskCompletionSource<int>(TaskCreationOptions.RunContinuationsAsynchronously);

        process.OutputDataReceived += (_, e) =>
        {
            if (!string.IsNullOrWhiteSpace(e.Data))
            {
                log?.Invoke(e.Data);
            }
        };

        process.ErrorDataReceived += (_, e) =>
        {
            if (!string.IsNullOrWhiteSpace(e.Data))
            {
                log?.Invoke($"[ERR] {e.Data}");
            }
        };

        process.Exited += (_, _) =>
        {
            tcs.TrySetResult(process.ExitCode);
        };

        log?.Invoke($"[ODA] {odaExePath} {args}");
        if (!process.Start())
        {
            throw new InvalidOperationException("无法启动 ODAFileConverter.exe");
        }

        process.BeginOutputReadLine();
        process.BeginErrorReadLine();

        using var registration = cancellationToken.Register(() =>
        {
            try
            {
                if (!process.HasExited)
                {
                    process.Kill(entireProcessTree: true);
                }
            }
            catch
            {
                // ignored
            }
        });

        return await tcs.Task.ConfigureAwait(false);
    }
}
