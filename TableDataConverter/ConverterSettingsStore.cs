using System.Text.Json;

namespace TableDataConverter;

internal sealed class ConverterSettings
{
    public string BytesOutputPath { get; set; } = string.Empty;
    public string ScriptOutputPath { get; set; } = string.Empty;
    public string DatabasePath { get; set; } = string.Empty;
}

internal static class ConverterSettingsStore
{
    static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true
    };

    public static string FilePath =>
        Path.Combine(AppContext.BaseDirectory, "TableDataConverter.settings.json");

    public static ConverterSettings Load()
    {
        if (!File.Exists(FilePath))
            return new ConverterSettings();

        var json = File.ReadAllText(FilePath);
        return JsonSerializer.Deserialize<ConverterSettings>(json, JsonOptions)
            ?? new ConverterSettings();
    }

    public static void Save(ConverterSettings settings)
    {
        var temporaryPath = FilePath + ".tmp";
        var json = JsonSerializer.Serialize(settings, JsonOptions);
        File.WriteAllText(temporaryPath, json);
        File.Move(temporaryPath, FilePath, true);
    }
}
