using System.Text.Json;
using System.Text.Json.Serialization;
using System.Collections.Generic;

namespace XlsxReview;

/// <summary>
/// JSON source generator context for trim-safe / AOT-compatible serialization.
/// </summary>
[JsonSerializable(typeof(EditManifest))]
[JsonSerializable(typeof(ProcessingResult))]
[JsonSerializable(typeof(ReadResult))]
[JsonSourceGenerationOptions(
    PropertyNameCaseInsensitive = true,
    WriteIndented = true
)]
public partial class XlsxReviewJsonContext : JsonSerializerContext { }

// ── Manifest models ──

/// <summary>
/// Root manifest model deserialized from the JSON input.
/// </summary>
public class EditManifest
{
    [JsonPropertyName("author")]
    public string? Author { get; set; }

    [JsonPropertyName("changes")]
    public List<Change>? Changes { get; set; }

    [JsonPropertyName("comments")]
    public List<CommentDef>? Comments { get; set; }
}

/// <summary>
/// A single spreadsheet change.
/// </summary>
public class Change
{
    [JsonPropertyName("type")]
    public string Type { get; set; } = "set_cell";

    [JsonPropertyName("sheet")]
    public string? Sheet { get; set; }

    [JsonPropertyName("cell")]
    public string? Cell { get; set; }

    [JsonPropertyName("value")]
    public string? Value { get; set; }

    [JsonPropertyName("format")]
    public string? Format { get; set; }

    [JsonPropertyName("formula")]
    public string? Formula { get; set; }

    [JsonPropertyName("after")]
    public JsonElement? After { get; set; }  // int for rows, string for columns

    [JsonPropertyName("row")]
    public int? Row { get; set; }

    [JsonPropertyName("column")]
    public string? Column { get; set; }

    [JsonPropertyName("name")]
    public string? Name { get; set; }

    [JsonPropertyName("from")]
    public string? From { get; set; }

    [JsonPropertyName("to")]
    public string? To { get; set; }
}

/// <summary>
/// A comment on a cell.
/// </summary>
public class CommentDef
{
    [JsonPropertyName("sheet")]
    public string Sheet { get; set; } = "";

    [JsonPropertyName("cell")]
    public string Cell { get; set; } = "";

    [JsonPropertyName("text")]
    public string Text { get; set; } = "";
}

// ── Result models ──

/// <summary>
/// Result of processing a single edit or comment.
/// </summary>
public class EditResult
{
    [JsonPropertyName("index")]
    public int Index { get; set; }

    [JsonPropertyName("type")]
    public string Type { get; set; } = "";

    [JsonPropertyName("success")]
    public bool Success { get; set; }

    [JsonPropertyName("message")]
    public string Message { get; set; } = "";
}

/// <summary>
/// Overall result summary for JSON output mode.
/// </summary>
public class ProcessingResult
{
    [JsonPropertyName("input")]
    public string Input { get; set; } = "";

    [JsonPropertyName("output")]
    public string? Output { get; set; }

    [JsonPropertyName("author")]
    public string Author { get; set; } = "";

    [JsonPropertyName("changes_attempted")]
    public int ChangesAttempted { get; set; }

    [JsonPropertyName("changes_succeeded")]
    public int ChangesSucceeded { get; set; }

    [JsonPropertyName("comments_attempted")]
    public int CommentsAttempted { get; set; }

    [JsonPropertyName("comments_succeeded")]
    public int CommentsSucceeded { get; set; }

    [JsonPropertyName("results")]
    public List<EditResult> Results { get; set; } = new();

    [JsonPropertyName("success")]
    public bool Success { get; set; }
}

// ── Read mode models ──

public class ReadResult
{
    [JsonPropertyName("sheets")]
    public List<SheetData> Sheets { get; set; } = new();
}

public class SheetData
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = "";

    [JsonPropertyName("rows")]
    public List<RowData> Rows { get; set; } = new();
}

public class RowData
{
    [JsonPropertyName("row")]
    public int Row { get; set; }

    [JsonPropertyName("cells")]
    public List<CellData> Cells { get; set; } = new();
}

public class CellData
{
    [JsonPropertyName("cell")]
    public string Cell { get; set; } = "";

    [JsonPropertyName("value")]
    public string? Value { get; set; }
}
