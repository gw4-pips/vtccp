namespace ExcelEngine.Schema;

using System.Text.Json;

/// <summary>
/// Manages available column schemas and the currently active schema.
/// Built-in schemas are always available. Custom schemas can be loaded from JSON.
/// </summary>
public sealed class ColumnSchemaManager
{
    private readonly Dictionary<string, ColumnSchema> _schemas = new(StringComparer.OrdinalIgnoreCase);
    private string _activeSchemaName = WebscanCompatibleSchema.SchemaName;

    public ColumnSchemaManager()
    {
        Register(WebscanCompatibleSchema.Build());
    }

    public void Register(ColumnSchema schema)
    {
        _schemas[schema.Name] = schema;
    }

    public ColumnSchema GetActive() => _schemas[_activeSchemaName];

    public void SetActive(string name)
    {
        if (!_schemas.ContainsKey(name))
            throw new ArgumentException($"Schema '{name}' is not registered.", nameof(name));
        _activeSchemaName = name;
    }

    public IReadOnlyList<string> AvailableSchemaNames => [.. _schemas.Keys];

    /// <summary>
    /// Load a custom schema from a JSON file and register it.
    /// JSON format: { "name": "...", "description": "...", "columns": [ { "fieldId": "...", "displayName": "...", ... } ] }
    /// </summary>
    public void LoadFromFile(string filePath)
    {
        var json = File.ReadAllText(filePath);
        var schema = JsonSerializer.Deserialize<ColumnSchema>(json,
            new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
            ?? throw new InvalidDataException($"Could not deserialize schema from {filePath}");
        Register(schema);
    }

    /// <summary>
    /// Validate a schema: check for duplicate field ids and empty display names.
    /// Returns list of validation messages (empty = valid).
    /// </summary>
    public static IReadOnlyList<string> Validate(ColumnSchema schema)
    {
        var errors = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var col in schema.Columns)
        {
            if (string.IsNullOrWhiteSpace(col.FieldId))
                errors.Add("A column has an empty FieldId.");
            else if (!seen.Add(col.FieldId))
                errors.Add($"Duplicate FieldId: '{col.FieldId}'");
            if (string.IsNullOrWhiteSpace(col.DisplayName))
                errors.Add($"Column '{col.FieldId}' has an empty DisplayName.");
        }
        return errors;
    }
}
