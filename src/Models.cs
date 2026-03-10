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
[JsonSerializable(typeof(XlsxDiffResult))]
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
    [JsonPropertyName("workbook")]
    public WorkbookInfo Workbook { get; set; } = new();

    [JsonPropertyName("sheets")]
    public List<SheetData> Sheets { get; set; } = new();

    [JsonPropertyName("warnings")]
    public List<ReadWarning> Warnings { get; set; } = new();
}

public class WorkbookInfo
{
    [JsonPropertyName("document_type")]
    public string DocumentType { get; set; } = "";

    [JsonPropertyName("sheet_count")]
    public int SheetCount { get; set; }

    [JsonPropertyName("worksheet_count")]
    public int WorksheetCount { get; set; }

    [JsonPropertyName("chartsheet_count")]
    public int ChartsheetCount { get; set; }

    [JsonPropertyName("dialogsheet_count")]
    public int DialogsheetCount { get; set; }

    [JsonPropertyName("hidden_sheet_count")]
    public int HiddenSheetCount { get; set; }

    [JsonPropertyName("very_hidden_sheet_count")]
    public int VeryHiddenSheetCount { get; set; }

    [JsonPropertyName("defined_name_count")]
    public int DefinedNameCount { get; set; }

    [JsonPropertyName("external_link_count")]
    public int ExternalLinkCount { get; set; }

    [JsonPropertyName("has_macros")]
    public bool HasMacros { get; set; }

    [JsonPropertyName("protected")]
    public bool Protected { get; set; }
}

public class ReadWarning
{
    [JsonPropertyName("scope")]
    public string Scope { get; set; } = "";

    [JsonPropertyName("target")]
    public string Target { get; set; } = "";

    [JsonPropertyName("message")]
    public string Message { get; set; } = "";
}

public class SheetData
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = "";

    [JsonPropertyName("kind")]
    public string Kind { get; set; } = "worksheet";

    [JsonPropertyName("visibility")]
    public string Visibility { get; set; } = "visible";

    [JsonPropertyName("row_count")]
    public int RowCount { get; set; }

    [JsonPropertyName("cell_count")]
    public int CellCount { get; set; }

    [JsonPropertyName("formula_count")]
    public int FormulaCount { get; set; }

    [JsonPropertyName("comment_count")]
    public int CommentCount { get; set; }

    [JsonPropertyName("table_count")]
    public int TableCount { get; set; }

    [JsonPropertyName("data_validation_count")]
    public int DataValidationCount { get; set; }

    [JsonPropertyName("conditional_format_count")]
    public int ConditionalFormatCount { get; set; }

    [JsonPropertyName("pivot_table_count")]
    public int PivotTableCount { get; set; }

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

    [JsonPropertyName("formula")]
    public string? Formula { get; set; }

    [JsonPropertyName("type")]
    public string? Type { get; set; }
}
