using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace XlsxReview;

// ── Top-level diff result ──────────────────────────────────────

public class XlsxDiffResult
{
    [JsonPropertyName("old_file")]
    public string OldFile { get; set; } = "";

    [JsonPropertyName("new_file")]
    public string NewFile { get; set; } = "";

    [JsonPropertyName("sheets_diff")]
    public SheetsDiff SheetsDiff { get; set; } = new();

    [JsonPropertyName("cell_changes")]
    public List<SheetCellChanges> CellChanges { get; set; } = new();

    [JsonPropertyName("formula_changes")]
    public List<SheetFormulaChanges> FormulaChanges { get; set; } = new();

    [JsonPropertyName("structure_diff")]
    public StructureDiff StructureDiff { get; set; } = new();

    [JsonPropertyName("summary")]
    public XlsxDiffSummary Summary { get; set; } = new();
}

// ── Sheet-level diff ───────────────────────────────────────────

public class SheetsDiff
{
    [JsonPropertyName("added")]
    public List<string> Added { get; set; } = new();

    [JsonPropertyName("deleted")]
    public List<string> Deleted { get; set; } = new();

    [JsonPropertyName("matched")]
    public List<string> Matched { get; set; } = new();
}

// ── Cell-level diff (per sheet) ────────────────────────────────

public class SheetCellChanges
{
    [JsonPropertyName("sheet")]
    public string Sheet { get; set; } = "";

    [JsonPropertyName("changes")]
    public List<CellChange> Changes { get; set; } = new();
}

public class CellChange
{
    [JsonPropertyName("cell")]
    public string Cell { get; set; } = "";

    [JsonPropertyName("type")]
    public string Type { get; set; } = "";  // "modified", "added", "deleted"

    [JsonPropertyName("old_value")]
    public string? OldValue { get; set; }

    [JsonPropertyName("new_value")]
    public string? NewValue { get; set; }
}

// ── Formula-level diff (per sheet) ─────────────────────────────

public class SheetFormulaChanges
{
    [JsonPropertyName("sheet")]
    public string Sheet { get; set; } = "";

    [JsonPropertyName("changes")]
    public List<FormulaChange> Changes { get; set; } = new();
}

public class FormulaChange
{
    [JsonPropertyName("cell")]
    public string Cell { get; set; } = "";

    [JsonPropertyName("type")]
    public string Type { get; set; } = "";  // "modified", "added", "deleted"

    [JsonPropertyName("old_formula")]
    public string? OldFormula { get; set; }

    [JsonPropertyName("new_formula")]
    public string? NewFormula { get; set; }
}

// ── Structure diff ─────────────────────────────────────────────

public class StructureDiff
{
    [JsonPropertyName("sheet_changes")]
    public List<SheetStructureChange> SheetChanges { get; set; } = new();
}

public class SheetStructureChange
{
    [JsonPropertyName("sheet")]
    public string Sheet { get; set; } = "";

    [JsonPropertyName("old_rows")]
    public int OldRows { get; set; }

    [JsonPropertyName("new_rows")]
    public int NewRows { get; set; }

    [JsonPropertyName("old_columns")]
    public int OldColumns { get; set; }

    [JsonPropertyName("new_columns")]
    public int NewColumns { get; set; }
}

// ── Summary ────────────────────────────────────────────────────

public class XlsxDiffSummary
{
    [JsonPropertyName("sheets_added")]
    public int SheetsAdded { get; set; }

    [JsonPropertyName("sheets_deleted")]
    public int SheetsDeleted { get; set; }

    [JsonPropertyName("cells_added")]
    public int CellsAdded { get; set; }

    [JsonPropertyName("cells_deleted")]
    public int CellsDeleted { get; set; }

    [JsonPropertyName("cells_modified")]
    public int CellsModified { get; set; }

    [JsonPropertyName("formulas_added")]
    public int FormulasAdded { get; set; }

    [JsonPropertyName("formulas_deleted")]
    public int FormulasDeleted { get; set; }

    [JsonPropertyName("formulas_modified")]
    public int FormulasModified { get; set; }

    [JsonPropertyName("structure_changes")]
    public int StructureChanges { get; set; }

    [JsonPropertyName("identical")]
    public bool Identical { get; set; }
}
