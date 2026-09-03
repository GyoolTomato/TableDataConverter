using ClosedXML.Excel;
using Microsoft.Data.Sqlite;
using System.Globalization;
using System.Text.RegularExpressions;

namespace TableDataConverter;

internal sealed record TableImportResult(int TableCount, int RowCount, string BackupPath);

internal static partial class SqliteTableImporter
{
    public static bool CanImport(string fileName)
    {
        var match = TableFileNameRegex().Match(fileName);
        return match.Success && match.Groups[1].Value[0] is not ('0' or '9');
    }

    public static TableImportResult Import(string databasePath, IReadOnlyCollection<FileInfo> excelFiles)
    {
        if (!File.Exists(databasePath))
            throw new FileNotFoundException("DB 파일을 찾을 수 없습니다.", databasePath);
        if (excelFiles.Count == 0)
            throw new InvalidOperationException("반영할 테이블을 하나 이상 선택하세요.");

        var tables = excelFiles.Select(ReadTable).ToList();
        var backupPath = CreateBackup(databasePath);

        using var connection = new SqliteConnection(new SqliteConnectionStringBuilder
        {
            DataSource = databasePath,
            Mode = SqliteOpenMode.ReadWrite
        }.ToString());
        connection.Open();
        using var transaction = connection.BeginTransaction();

        try
        {
            foreach (var table in tables)
                ReplaceTable(connection, transaction, table);
            transaction.Commit();
        }
        catch
        {
            transaction.Rollback();
            throw;
        }

        return new TableImportResult(tables.Count, tables.Sum(table => table.Rows.Count), backupPath);
    }

    static ExcelTable ReadTable(FileInfo file)
    {
        using var stream = OpenSnapshot(file.FullName);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheet(1);
        var range = worksheet.RangeUsed()
            ?? throw new InvalidDataException($"{file.Name}: 데이터가 없습니다.");

        var columns = new List<ExcelColumn>();
        var arrayIndexes = new Dictionary<string, int>(StringComparer.Ordinal);
        var arrayTypes = new Dictionary<string, string>(StringComparer.Ordinal);

        for (var columnNumber = 1; columnNumber <= range.LastColumn().ColumnNumber(); columnNumber++)
        {
            var header = worksheet.Cell(2, columnNumber).GetText().Trim();
            if (string.IsNullOrEmpty(header) || header.StartsWith('.'))
                continue;

            var typeName = worksheet.Cell(3, columnNumber).GetText().Trim();
            var columnName = header;
            if (TryGetArrayName(header, out var arrayName))
            {
                if (!string.IsNullOrEmpty(typeName))
                    arrayTypes[arrayName] = typeName;
                if (!arrayTypes.TryGetValue(arrayName, out typeName))
                    throw new InvalidDataException($"{file.Name}: {header} 배열의 타입이 없습니다.");

                var index = arrayIndexes.GetValueOrDefault(arrayName);
                arrayIndexes[arrayName] = index + 1;
                columnName = $"{arrayName}_{index}";
            }

            if (string.IsNullOrEmpty(typeName))
                throw new InvalidDataException($"{file.Name}: {header} 열의 타입이 없습니다.");
            if (columns.Any(column => column.Name.Equals(columnName, StringComparison.OrdinalIgnoreCase)))
                throw new InvalidDataException($"{file.Name}: 중복 열 이름 {columnName}");

            columns.Add(new ExcelColumn(columnNumber, columnName, typeName, ToSqliteType(typeName)));
        }

        if (!columns.Any(column => column.Name.Equals("key", StringComparison.OrdinalIgnoreCase)))
            throw new InvalidDataException($"{file.Name}: key 열이 없습니다.");

        var rows = new List<object?[]>();
        var keys = new HashSet<string>(StringComparer.Ordinal);
        for (var rowNumber = 4; rowNumber <= range.LastRow().RowNumber(); rowNumber++)
        {
            if (columns.All(column => worksheet.Cell(rowNumber, column.SourceIndex).IsEmpty()))
                continue;

            var values = columns.Select(column => ReadValue(worksheet.Cell(rowNumber, column.SourceIndex), column.TypeName)).ToArray();
            var keyIndex = columns.FindIndex(column => column.Name.Equals("key", StringComparison.OrdinalIgnoreCase));
            var key = values[keyIndex]?.ToString();
            if (string.IsNullOrWhiteSpace(key))
                throw new InvalidDataException($"{file.Name}: {rowNumber}행의 key가 비어 있습니다.");
            if (!keys.Add(key))
                throw new InvalidDataException($"{file.Name}: 중복 key {key}");
            rows.Add(values);
        }

        return new ExcelTable(Path.GetFileNameWithoutExtension(file.Name), columns, rows);
    }

    static void ReplaceTable(SqliteConnection connection, SqliteTransaction transaction, ExcelTable table)
    {
        using (var drop = connection.CreateCommand())
        {
            drop.Transaction = transaction;
            drop.CommandText = $"DROP TABLE IF EXISTS {Quote(table.Name)}";
            drop.ExecuteNonQuery();
        }

        using (var create = connection.CreateCommand())
        {
            create.Transaction = transaction;
            create.CommandText = $"CREATE TABLE {Quote(table.Name)} ({string.Join(", ", table.Columns.Select(column => $"{Quote(column.Name)} {column.SqliteType}{(column.Name.Equals("key", StringComparison.OrdinalIgnoreCase) ? " PRIMARY KEY" : string.Empty)}"))})";
            create.ExecuteNonQuery();
        }

        using var insert = connection.CreateCommand();
        insert.Transaction = transaction;
        insert.CommandText = $"INSERT INTO {Quote(table.Name)} ({string.Join(", ", table.Columns.Select(column => Quote(column.Name)))}) VALUES ({string.Join(", ", table.Columns.Select((_, index) => $"$p{index}"))})";
        for (var index = 0; index < table.Columns.Count; index++)
            insert.Parameters.Add(new SqliteParameter($"$p{index}", null));

        foreach (var row in table.Rows)
        {
            for (var index = 0; index < row.Length; index++)
                insert.Parameters[index].Value = row[index] ?? DBNull.Value;
            insert.ExecuteNonQuery();
        }
    }

    static string CreateBackup(string databasePath)
    {
        var backupPath = Path.Combine(
            Path.GetDirectoryName(databasePath)!,
            $"{Path.GetFileNameWithoutExtension(databasePath)}_{DateTime.Now:yyyyMMdd_HHmmss}.backup.db");
        using var source = new SqliteConnection(new SqliteConnectionStringBuilder
        {
            DataSource = databasePath,
            Mode = SqliteOpenMode.ReadOnly
        }.ToString());
        using var destination = new SqliteConnection(new SqliteConnectionStringBuilder
        {
            DataSource = backupPath,
            Mode = SqliteOpenMode.ReadWriteCreate
        }.ToString());
        source.Open();
        destination.Open();
        source.BackupDatabase(destination);
        return backupPath;
    }

    static MemoryStream OpenSnapshot(string path)
    {
        using var source = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        var snapshot = new MemoryStream();
        source.CopyTo(snapshot);
        snapshot.Position = 0;
        return snapshot;
    }

    static object? ReadValue(IXLCell cell, string typeName)
    {
        if (cell.IsEmpty()) return null;
        return typeName.Trim().ToLowerInvariant() switch
        {
            "byte" => cell.GetValue<byte>(),
            "short" or "int16" => cell.GetValue<short>(),
            "int" or "int32" => cell.GetValue<int>(),
            "long" or "int64" => cell.GetValue<long>(),
            "float" or "single" => cell.GetValue<float>(),
            "double" => cell.GetValue<double>(),
            "decimal" => cell.GetValue<decimal>(),
            "bool" or "boolean" => cell.TryGetValue<bool>(out var value) ? value ? 1 : 0 : cell.GetValue<int>(),
            _ => cell.GetFormattedString(CultureInfo.InvariantCulture).Trim()
        };
    }

    static string ToSqliteType(string typeName) => typeName.Trim().ToLowerInvariant() switch
    {
        "byte" or "short" or "int16" or "int" or "int32" or "long" or "int64" or "bool" or "boolean" => "INTEGER",
        "float" or "single" or "double" or "decimal" => "REAL",
        _ => "TEXT"
    };

    static bool TryGetArrayName(string header, out string name)
    {
        name = header.Length >= 3 && header[0] == '[' && header[^1] == ']'
            ? header[1..^1].Trim()
            : string.Empty;
        return name.Length > 0;
    }

    static string Quote(string identifier) => $"\"{identifier.Replace("\"", "\"\"")}\"";

    [GeneratedRegex(@"^_(\d{3})_.+\.xlsx$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex TableFileNameRegex();

    sealed record ExcelColumn(int SourceIndex, string Name, string TypeName, string SqliteType);
    sealed record ExcelTable(string Name, List<ExcelColumn> Columns, List<object?[]> Rows);
}
