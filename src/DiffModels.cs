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

    [JsonPropertyName("metadata_diff")]
    public MetadataDiff MetadataDiff { get; set; } = new();

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

public class MetadataDiff
{
    [JsonPropertyName("sheet_visibility_changes")]
    public List<SheetVisibilityChange> SheetVisibilityChanges { get; set; } = new();

    [JsonPropertyName("sheet_protection_changes")]
    public List<SheetProtectionChange> SheetProtectionChanges { get; set; } = new();

    [JsonPropertyName("defined_name_changes")]
    public List<DefinedNameChange> DefinedNameChanges { get; set; } = new();

    [JsonPropertyName("workbook_protection_change")]
    public WorkbookProtectionChange WorkbookProtectionChange { get; set; } = new();
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

public class SheetVisibilityChange
{
    [JsonPropertyName("sheet")]
    public string Sheet { get; set; } = "";

    [JsonPropertyName("old_visibility")]
    public string OldVisibility { get; set; } = "visible";

    [JsonPropertyName("new_visibility")]
    public string NewVisibility { get; set; } = "visible";
}

public class SheetProtectionChange
{
    [JsonPropertyName("sheet")]
    public string Sheet { get; set; } = "";

    [JsonPropertyName("old_protected")]
    public bool OldProtected { get; set; }

    [JsonPropertyName("new_protected")]
    public bool NewProtected { get; set; }
}

public class DefinedNameChange
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = "";

    [JsonPropertyName("scope_sheet")]
    public string? ScopeSheet { get; set; }

    [JsonPropertyName("type")]
    public string Type { get; set; } = "";

    [JsonPropertyName("old_refers_to")]
    public string? OldRefersTo { get; set; }

    [JsonPropertyName("new_refers_to")]
    public string? NewRefersTo { get; set; }

    [JsonPropertyName("old_hidden")]
    public bool? OldHidden { get; set; }

    [JsonPropertyName("new_hidden")]
    public bool? NewHidden { get; set; }

    [JsonPropertyName("old_comment")]
    public string? OldComment { get; set; }

    [JsonPropertyName("new_comment")]
    public string? NewComment { get; set; }
}

public class WorkbookProtectionChange
{
    [JsonPropertyName("changed")]
    public bool Changed { get; set; }

    [JsonPropertyName("old")]
    public WorkbookProtectionInfo Old { get; set; } = new();

    [JsonPropertyName("new")]
    public WorkbookProtectionInfo New { get; set; } = new();
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

    [JsonPropertyName("sheet_visibility_changes")]
    public int SheetVisibilityChanges { get; set; }

    [JsonPropertyName("sheet_protection_changes")]
    public int SheetProtectionChanges { get; set; }

    [JsonPropertyName("defined_name_changes")]
    public int DefinedNameChanges { get; set; }

    [JsonPropertyName("workbook_protection_changes")]
    public int WorkbookProtectionChanges { get; set; }

    [JsonPropertyName("metadata_changes")]
    public int MetadataChanges { get; set; }

    [JsonPropertyName("identical")]
    public bool Identical { get; set; }
}
