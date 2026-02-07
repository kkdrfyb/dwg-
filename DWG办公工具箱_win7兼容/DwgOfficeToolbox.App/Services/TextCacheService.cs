using System.IO;
using Microsoft.Data.Sqlite;
using DwgOfficeToolbox.App.Models;

namespace DwgOfficeToolbox.App.Services;

public sealed class CacheUpdateResult
{
    public CacheUpdateResult(List<DwgScanTarget> plainTextTargets, int cached, int skipped, int failed)
    {
        PlainTextTargets = plainTextTargets;
        CachedCount = cached;
        SkippedCount = skipped;
        FailedCount = failed;
    }

    public List<DwgScanTarget> PlainTextTargets { get; }
    public int CachedCount { get; }
    public int SkippedCount { get; }
    public int FailedCount { get; }
}

public sealed class TextCacheService
{
    private readonly string _dbPath;

    public TextCacheService(string dbPath)
    {
        _dbPath = dbPath;
        Initialize();
    }

    public CacheUpdateResult UpdateCache(
        List<DwgScanTarget> targets,
        int largeFileThresholdMB,
        Action<string>? log,
        Action<int, int>? progress,
        CancellationToken ct)
    {
        var plainTextTargets = new List<DwgScanTarget>();
        var cachedCount = 0;
        var skippedCount = 0;
        var failedCount = 0;
        var total = targets.Count;
        var processed = 0;
        var thresholdBytes = largeFileThresholdMB * 1024L * 1024L;

        using var connection = OpenConnection();
        connection.Open();
        EnableForeignKeys(connection);

        foreach (var target in targets)
        {
            ct.ThrowIfCancellationRequested();

            var current = Interlocked.Increment(ref processed);
            progress?.Invoke(current, total);

            if (!File.Exists(target.DxfPath))
            {
                continue;
            }

            var info = new FileInfo(target.DxfPath);
            var meta = GetFileMeta(connection, target.DxfPath);

            if (info.Length > thresholdBytes)
            {
                UpsertMeta(connection, target.DxfPath, info, cached: 0, textCount: 0);
                plainTextTargets.Add(target);
                failedCount++;
                continue;
            }

            if (meta != null &&
                meta.Size == info.Length &&
                meta.LastWriteUtcTicks == info.LastWriteTimeUtc.Ticks)
            {
                if (meta.Cached == 0)
                {
                    plainTextTargets.Add(target);
                }
                else
                {
                    skippedCount++;
                }
                continue;
            }

            try
            {
                var items = DxfTextExtractor.Extract(target.DxfPath);
                if (items.Count == 0)
                {
                    log?.Invoke($"解析未提取到文字：{Path.GetFileName(target.DxfPath)}");
                    UpsertMeta(connection, target.DxfPath, info, cached: 0, textCount: 0);
                    plainTextTargets.Add(target);
                    failedCount++;
                    continue;
                }

                using var tx = connection.BeginTransaction();
                var fileId = UpsertMeta(connection, target.DxfPath, info, cached: 1, textCount: items.Count, tx);
                DeleteTextRows(connection, fileId, tx);
                InsertItems(connection, fileId, items, tx);
                tx.Commit();
                cachedCount++;
            }
            catch (Exception ex)
            {
                log?.Invoke($"缓存更新失败：{Path.GetFileName(target.DxfPath)}，{ex.Message}");
                UpsertMeta(connection, target.DxfPath, info, cached: 0, textCount: 0);
                plainTextTargets.Add(target);
                failedCount++;
            }
        }

        return new CacheUpdateResult(plainTextTargets, cachedCount, skippedCount, failedCount);
    }

    public List<MatchResult> Query(List<DwgScanTarget> targets, IReadOnlyList<string> keywords)
    {
        var results = new List<MatchResult>();
        if (targets.Count == 0)
        {
            return results;
        }

        var dwgMap = targets
            .GroupBy(t => t.DxfPath, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.First().DwgPath, StringComparer.OrdinalIgnoreCase);

        using var connection = OpenConnection();
        connection.Open();
        EnableForeignKeys(connection);

        var fileIds = GetFileIds(connection, targets.Select(t => t.DxfPath).Distinct(StringComparer.OrdinalIgnoreCase).ToList());
        if (fileIds.Count == 0)
        {
            return results;
        }

        if (keywords.Count == 0)
        {
            var rows = QueryAllRows(connection, fileIds);
            foreach (var row in rows)
            {
                results.Add(new MatchResult
                {
                    FileName = Path.GetFileName(row.Path),
                    FilePath = row.Path,
                    DwgPath = dwgMap.TryGetValue(row.Path, out var dwg) ? dwg : row.Path,
                    ObjectType = row.ObjectType,
                    Layer = row.Layer,
                    Keyword = "全部",
                    Content = row.Text
                });
            }
            return results;
        }

        foreach (var keyword in keywords)
        {
            var rows = QueryKeywordRows(connection, fileIds, keyword);
            foreach (var row in rows)
            {
                results.Add(new MatchResult
                {
                    FileName = Path.GetFileName(row.Path),
                    FilePath = row.Path,
                    DwgPath = dwgMap.TryGetValue(row.Path, out var dwg) ? dwg : row.Path,
                    ObjectType = row.ObjectType,
                    Layer = row.Layer,
                    Keyword = keyword,
                    Content = row.Text
                });
            }
        }

        return results;
    }

    private void Initialize()
    {
        var dir = Path.GetDirectoryName(_dbPath);
        if (!string.IsNullOrWhiteSpace(dir))
        {
            Directory.CreateDirectory(dir);
        }
        using var connection = OpenConnection();
        connection.Open();
        EnableForeignKeys(connection);

        using var cmd = connection.CreateCommand();
        cmd.CommandText = @"
CREATE TABLE IF NOT EXISTS dxf_file (
  id INTEGER PRIMARY KEY,
  path TEXT NOT NULL UNIQUE,
  size INTEGER NOT NULL,
  last_write_utc_ticks INTEGER NOT NULL,
  cached INTEGER NOT NULL,
  text_count INTEGER NOT NULL
);
CREATE TABLE IF NOT EXISTS dxf_text (
  id INTEGER PRIMARY KEY,
  file_id INTEGER NOT NULL,
  object_type TEXT NOT NULL,
  layer TEXT NOT NULL,
  text TEXT NOT NULL,
  FOREIGN KEY(file_id) REFERENCES dxf_file(id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_dxf_text_file ON dxf_text(file_id);
";
        cmd.ExecuteNonQuery();
    }

    private SqliteConnection OpenConnection()
    {
        return new SqliteConnection($"Data Source={_dbPath}");
    }

    private static void EnableForeignKeys(SqliteConnection connection)
    {
        using var cmd = connection.CreateCommand();
        cmd.CommandText = "PRAGMA foreign_keys = ON;";
        cmd.ExecuteNonQuery();
    }

    private static FileMeta? GetFileMeta(SqliteConnection connection, string path)
    {
        using var cmd = connection.CreateCommand();
        cmd.CommandText = "SELECT id, size, last_write_utc_ticks, cached FROM dxf_file WHERE path = @path;";
        cmd.Parameters.AddWithValue("@path", path);

        using var reader = cmd.ExecuteReader();
        if (!reader.Read())
        {
            return null;
        }

        return new FileMeta(
            reader.GetInt64(0),
            reader.GetInt64(1),
            reader.GetInt64(2),
            reader.GetInt32(3));
    }

    private static long UpsertMeta(SqliteConnection connection, string path, FileInfo info, int cached, int textCount)
    {
        using var cmd = connection.CreateCommand();
        cmd.CommandText = @"
INSERT INTO dxf_file(path, size, last_write_utc_ticks, cached, text_count)
VALUES(@path, @size, @ticks, @cached, @textCount)
ON CONFLICT(path) DO UPDATE SET
  size = excluded.size,
  last_write_utc_ticks = excluded.last_write_utc_ticks,
  cached = excluded.cached,
  text_count = excluded.text_count;
";
        cmd.Parameters.AddWithValue("@path", path);
        cmd.Parameters.AddWithValue("@size", info.Length);
        cmd.Parameters.AddWithValue("@ticks", info.LastWriteTimeUtc.Ticks);
        cmd.Parameters.AddWithValue("@cached", cached);
        cmd.Parameters.AddWithValue("@textCount", textCount);
        cmd.ExecuteNonQuery();

        using var getCmd = connection.CreateCommand();
        getCmd.CommandText = "SELECT id FROM dxf_file WHERE path = @path;";
        getCmd.Parameters.AddWithValue("@path", path);
        return (long)getCmd.ExecuteScalar()!;
    }

    private static long UpsertMeta(SqliteConnection connection, string path, FileInfo info, int cached, int textCount, SqliteTransaction tx)
    {
        using var cmd = connection.CreateCommand();
        cmd.Transaction = tx;
        cmd.CommandText = @"
INSERT INTO dxf_file(path, size, last_write_utc_ticks, cached, text_count)
VALUES(@path, @size, @ticks, @cached, @textCount)
ON CONFLICT(path) DO UPDATE SET
  size = excluded.size,
  last_write_utc_ticks = excluded.last_write_utc_ticks,
  cached = excluded.cached,
  text_count = excluded.text_count;
";
        cmd.Parameters.AddWithValue("@path", path);
        cmd.Parameters.AddWithValue("@size", info.Length);
        cmd.Parameters.AddWithValue("@ticks", info.LastWriteTimeUtc.Ticks);
        cmd.Parameters.AddWithValue("@cached", cached);
        cmd.Parameters.AddWithValue("@textCount", textCount);
        cmd.ExecuteNonQuery();

        using var getCmd = connection.CreateCommand();
        getCmd.Transaction = tx;
        getCmd.CommandText = "SELECT id FROM dxf_file WHERE path = @path;";
        getCmd.Parameters.AddWithValue("@path", path);
        return (long)getCmd.ExecuteScalar()!;
    }

    private static void DeleteTextRows(SqliteConnection connection, long fileId, SqliteTransaction tx)
    {
        using var cmd = connection.CreateCommand();
        cmd.Transaction = tx;
        cmd.CommandText = "DELETE FROM dxf_text WHERE file_id = @fileId;";
        cmd.Parameters.AddWithValue("@fileId", fileId);
        cmd.ExecuteNonQuery();
    }

    private static void InsertItems(SqliteConnection connection, long fileId, List<DxfTextItem> items, SqliteTransaction tx)
    {
        using var cmd = connection.CreateCommand();
        cmd.Transaction = tx;
        cmd.CommandText = @"
INSERT INTO dxf_text(file_id, object_type, layer, text)
VALUES(@fileId, @type, @layer, @text);
";
        var pFileId = cmd.CreateParameter();
        pFileId.ParameterName = "@fileId";
        var pType = cmd.CreateParameter();
        pType.ParameterName = "@type";
        var pLayer = cmd.CreateParameter();
        pLayer.ParameterName = "@layer";
        var pText = cmd.CreateParameter();
        pText.ParameterName = "@text";
        cmd.Parameters.Add(pFileId);
        cmd.Parameters.Add(pType);
        cmd.Parameters.Add(pLayer);
        cmd.Parameters.Add(pText);

        foreach (var item in items)
        {
            pFileId.Value = fileId;
            pType.Value = item.ObjectType;
            pLayer.Value = item.Layer;
            pText.Value = item.Text;
            cmd.ExecuteNonQuery();
        }
    }

    private static Dictionary<long, string> GetFileIds(SqliteConnection connection, List<string> paths)
    {
        var map = new Dictionary<long, string>();
        if (paths.Count == 0)
        {
            return map;
        }

        using var cmd = connection.CreateCommand();
        var paramNames = new List<string>();
        for (var i = 0; i < paths.Count; i++)
        {
            var name = "@p" + i;
            paramNames.Add(name);
            cmd.Parameters.AddWithValue(name, paths[i]);
        }

        cmd.CommandText = $"SELECT id, path FROM dxf_file WHERE path IN ({string.Join(",", paramNames)}) AND cached = 1;";
        using var reader = cmd.ExecuteReader();
        while (reader.Read())
        {
            map[reader.GetInt64(0)] = reader.GetString(1);
        }

        return map;
    }

    private static List<TextRow> QueryAllRows(SqliteConnection connection, Dictionary<long, string> fileIds)
    {
        var rows = new List<TextRow>();
        using var cmd = connection.CreateCommand();
        var paramNames = new List<string>();
        var index = 0;
        foreach (var id in fileIds.Keys)
        {
            var name = "@id" + index++;
            paramNames.Add(name);
            cmd.Parameters.AddWithValue(name, id);
        }

        cmd.CommandText = $@"
SELECT f.path, t.object_type, t.layer, t.text
FROM dxf_text t
JOIN dxf_file f ON f.id = t.file_id
WHERE t.file_id IN ({string.Join(",", paramNames)});
";
        using var reader = cmd.ExecuteReader();
        while (reader.Read())
        {
            rows.Add(new TextRow(reader.GetString(0), reader.GetString(1), reader.GetString(2), reader.GetString(3)));
        }

        return rows;
    }

    private static List<TextRow> QueryKeywordRows(SqliteConnection connection, Dictionary<long, string> fileIds, string keyword)
    {
        var rows = new List<TextRow>();
        using var cmd = connection.CreateCommand();
        var paramNames = new List<string>();
        var index = 0;
        foreach (var id in fileIds.Keys)
        {
            var name = "@id" + index++;
            paramNames.Add(name);
            cmd.Parameters.AddWithValue(name, id);
        }

        cmd.CommandText = $@"
SELECT f.path, t.object_type, t.layer, t.text
FROM dxf_text t
JOIN dxf_file f ON f.id = t.file_id
WHERE t.file_id IN ({string.Join(",", paramNames)})
  AND t.text LIKE '%' || @kw || '%' COLLATE NOCASE;
";
        cmd.Parameters.AddWithValue("@kw", keyword);

        using var reader = cmd.ExecuteReader();
        while (reader.Read())
        {
            rows.Add(new TextRow(reader.GetString(0), reader.GetString(1), reader.GetString(2), reader.GetString(3)));
        }

        return rows;
    }

    private sealed record FileMeta(long Id, long Size, long LastWriteUtcTicks, int Cached);
    private sealed record TextRow(string Path, string ObjectType, string Layer, string Text);
}
