using System;
using System.Collections.Generic;
using System.Linq;

namespace XlsxReview;

/// <summary>
/// Compares two SpreadsheetExtractions and produces a semantic XlsxDiffResult.
/// </summary>
public static class SpreadsheetDiffer
{
    public static XlsxDiffResult Diff(SpreadsheetExtraction oldDoc, SpreadsheetExtraction newDoc)
    {
        var result = new XlsxDiffResult
        {
            OldFile = oldDoc.FileName,
            NewFile = newDoc.FileName
        };

        // Sheet-level diff
        var oldSheetNames = new HashSet<string>(oldDoc.Sheets.Select(s => s.Name));
        var newSheetNames = new HashSet<string>(newDoc.Sheets.Select(s => s.Name));

        result.SheetsDiff.Added = newSheetNames.Except(oldSheetNames).OrderBy(s => s).ToList();
        result.SheetsDiff.Deleted = oldSheetNames.Except(newSheetNames).OrderBy(s => s).ToList();
        result.SheetsDiff.Matched = oldSheetNames.Intersect(newSheetNames).OrderBy(s => s).ToList();

        // Per-sheet diffs (matched sheets only)
        var oldSheetMap = oldDoc.Sheets.ToDictionary(s => s.Name);
        var newSheetMap = newDoc.Sheets.ToDictionary(s => s.Name);

        foreach (var sheetName in result.SheetsDiff.Matched)
        {
            var oldSheet = oldSheetMap[sheetName];
            var newSheet = newSheetMap[sheetName];

            // Cell value changes
            var cellChanges = DiffCells(sheetName, oldSheet, newSheet);
            if (cellChanges.Changes.Count > 0)
                result.CellChanges.Add(cellChanges);

            // Formula changes
            var formulaChanges = DiffFormulas(sheetName, oldSheet, newSheet);
            if (formulaChanges.Changes.Count > 0)
                result.FormulaChanges.Add(formulaChanges);

            // Structure changes
            if (oldSheet.MaxRow != newSheet.MaxRow || oldSheet.MaxColumn != newSheet.MaxColumn)
            {
                result.StructureDiff.SheetChanges.Add(new SheetStructureChange
                {
                    Sheet = sheetName,
                    OldRows = oldSheet.MaxRow,
                    NewRows = newSheet.MaxRow,
                    OldColumns = oldSheet.MaxColumn,
                    NewColumns = newSheet.MaxColumn
                });
            }
        }

        // Build summary
        int totalCellsAdded = result.CellChanges.Sum(sc => sc.Changes.Count(c => c.Type == "added"));
        int totalCellsDeleted = result.CellChanges.Sum(sc => sc.Changes.Count(c => c.Type == "deleted"));
        int totalCellsModified = result.CellChanges.Sum(sc => sc.Changes.Count(c => c.Type == "modified"));
        int totalFormulasAdded = result.FormulaChanges.Sum(sf => sf.Changes.Count(f => f.Type == "added"));
        int totalFormulasDeleted = result.FormulaChanges.Sum(sf => sf.Changes.Count(f => f.Type == "deleted"));
        int totalFormulasModified = result.FormulaChanges.Sum(sf => sf.Changes.Count(f => f.Type == "modified"));

        result.Summary = new XlsxDiffSummary
        {
            SheetsAdded = result.SheetsDiff.Added.Count,
            SheetsDeleted = result.SheetsDiff.Deleted.Count,
            CellsAdded = totalCellsAdded,
            CellsDeleted = totalCellsDeleted,
            CellsModified = totalCellsModified,
            FormulasAdded = totalFormulasAdded,
            FormulasDeleted = totalFormulasDeleted,
            FormulasModified = totalFormulasModified,
            StructureChanges = result.StructureDiff.SheetChanges.Count,
            Identical = result.SheetsDiff.Added.Count == 0
                     && result.SheetsDiff.Deleted.Count == 0
                     && totalCellsAdded == 0 && totalCellsDeleted == 0 && totalCellsModified == 0
                     && totalFormulasAdded == 0 && totalFormulasDeleted == 0 && totalFormulasModified == 0
                     && result.StructureDiff.SheetChanges.Count == 0
        };

        return result;
    }

    private static SheetCellChanges DiffCells(string sheetName,
        ExtractedSheet oldSheet, ExtractedSheet newSheet)
    {
        var changes = new SheetCellChanges { Sheet = sheetName };

        var allCellRefs = new HashSet<string>(oldSheet.Cells.Keys);
        allCellRefs.UnionWith(newSheet.Cells.Keys);

        foreach (var cellRef in allCellRefs.OrderBy(r => SortableCellRef(r)))
        {
            bool inOld = oldSheet.Cells.TryGetValue(cellRef, out var oldCell);
            bool inNew = newSheet.Cells.TryGetValue(cellRef, out var newCell);

            if (inOld && inNew)
            {
                // Both exist — compare values
                string? oldVal = oldCell!.Value;
                string? newVal = newCell!.Value;

                if (oldVal != newVal)
                {
                    changes.Changes.Add(new CellChange
                    {
                        Cell = cellRef,
                        Type = "modified",
                        OldValue = oldVal,
                        NewValue = newVal
                    });
                }
            }
            else if (!inOld && inNew)
            {
                // Added
                if (newCell!.Value != null || newCell.Formula != null)
                {
                    changes.Changes.Add(new CellChange
                    {
                        Cell = cellRef,
                        Type = "added",
                        OldValue = null,
                        NewValue = newCell.Value
                    });
                }
            }
            else if (inOld && !inNew)
            {
                // Deleted
                if (oldCell!.Value != null || oldCell.Formula != null)
                {
                    changes.Changes.Add(new CellChange
                    {
                        Cell = cellRef,
                        Type = "deleted",
                        OldValue = oldCell.Value,
                        NewValue = null
                    });
                }
            }
        }

        return changes;
    }

    private static SheetFormulaChanges DiffFormulas(string sheetName,
        ExtractedSheet oldSheet, ExtractedSheet newSheet)
    {
        var changes = new SheetFormulaChanges { Sheet = sheetName };

        var allCellRefs = new HashSet<string>(oldSheet.Cells.Keys);
        allCellRefs.UnionWith(newSheet.Cells.Keys);

        foreach (var cellRef in allCellRefs.OrderBy(r => SortableCellRef(r)))
        {
            bool inOld = oldSheet.Cells.TryGetValue(cellRef, out var oldCell);
            bool inNew = newSheet.Cells.TryGetValue(cellRef, out var newCell);

            string? oldFormula = inOld ? oldCell!.Formula : null;
            string? newFormula = inNew ? newCell!.Formula : null;

            if (oldFormula == null && newFormula == null) continue;

            if (oldFormula != null && newFormula != null)
            {
                if (oldFormula != newFormula)
                {
                    changes.Changes.Add(new FormulaChange
                    {
                        Cell = cellRef,
                        Type = "modified",
                        OldFormula = oldFormula,
                        NewFormula = newFormula
                    });
                }
            }
            else if (oldFormula == null && newFormula != null)
            {
                changes.Changes.Add(new FormulaChange
                {
                    Cell = cellRef,
                    Type = "added",
                    OldFormula = null,
                    NewFormula = newFormula
                });
            }
            else if (oldFormula != null && newFormula == null)
            {
                changes.Changes.Add(new FormulaChange
                {
                    Cell = cellRef,
                    Type = "deleted",
                    OldFormula = oldFormula,
                    NewFormula = null
                });
            }
        }

        return changes;
    }

    /// <summary>
    /// Produce a sortable key from a cell reference (e.g., "A1" → (1,1), "B10" → (10,2)).
    /// Sorts by row first, then column.
    /// </summary>
    private static (int row, int col) SortableCellRef(string cellRef)
    {
        int i = 0;
        while (i < cellRef.Length && char.IsLetter(cellRef[i])) i++;
        string colPart = cellRef[..i];
        int row = int.TryParse(cellRef[i..], out int r) ? r : 0;
        int col = 0;
        foreach (char c in colPart)
            col = col * 26 + (c - 'A' + 1);
        return (row, col);
    }

    // ── Human-readable output ──────────────────────────────────

    public static void PrintHumanReadable(XlsxDiffResult result)
    {
        Console.WriteLine();
        Console.WriteLine($"xlsx-review diff: {result.OldFile} → {result.NewFile}");
        Console.WriteLine(new string('═', 60));

        if (result.Summary.Identical)
        {
            Console.WriteLine("\n  Spreadsheets are identical.");
            return;
        }

        // Sheet-level changes
        if (result.SheetsDiff.Added.Count > 0 || result.SheetsDiff.Deleted.Count > 0)
        {
            Console.WriteLine("\nSheets");
            Console.WriteLine(new string('─', 40));
            foreach (var s in result.SheetsDiff.Added)
                Console.WriteLine($"  + \"{s}\" (added)");
            foreach (var s in result.SheetsDiff.Deleted)
                Console.WriteLine($"  - \"{s}\" (deleted)");
            foreach (var s in result.SheetsDiff.Matched)
                Console.WriteLine($"    \"{s}\" (matched)");
        }

        // Cell changes per sheet
        foreach (var sc in result.CellChanges)
        {
            Console.WriteLine($"\nCell Changes — {sc.Sheet} ({sc.Changes.Count} changes)");
            Console.WriteLine(new string('─', 40));

            foreach (var c in sc.Changes)
            {
                switch (c.Type)
                {
                    case "modified":
                        Console.WriteLine($"  ~ {c.Cell}: \"{Trunc(c.OldValue ?? "", 30)}\" → \"{Trunc(c.NewValue ?? "", 30)}\"");
                        break;
                    case "added":
                        Console.WriteLine($"  + {c.Cell}: \"{Trunc(c.NewValue ?? "", 40)}\"");
                        break;
                    case "deleted":
                        Console.WriteLine($"  - {c.Cell}: \"{Trunc(c.OldValue ?? "", 40)}\"");
                        break;
                }
            }
        }

        // Formula changes per sheet
        foreach (var sf in result.FormulaChanges)
        {
            Console.WriteLine($"\nFormula Changes — {sf.Sheet} ({sf.Changes.Count} changes)");
            Console.WriteLine(new string('─', 40));

            foreach (var f in sf.Changes)
            {
                switch (f.Type)
                {
                    case "modified":
                        Console.WriteLine($"  ~ {f.Cell}: ={f.OldFormula} → ={f.NewFormula}");
                        break;
                    case "added":
                        Console.WriteLine($"  + {f.Cell}: ={f.NewFormula}");
                        break;
                    case "deleted":
                        Console.WriteLine($"  - {f.Cell}: ={f.OldFormula}");
                        break;
                }
            }
        }

        // Structure changes
        if (result.StructureDiff.SheetChanges.Count > 0)
        {
            Console.WriteLine("\nStructure Changes");
            Console.WriteLine(new string('─', 40));
            foreach (var sc in result.StructureDiff.SheetChanges)
            {
                if (sc.OldRows != sc.NewRows)
                    Console.WriteLine($"  {sc.Sheet}: rows {sc.OldRows} → {sc.NewRows}");
                if (sc.OldColumns != sc.NewColumns)
                    Console.WriteLine($"  {sc.Sheet}: columns {sc.OldColumns} → {sc.NewColumns}");
            }
        }

        // Summary
        Console.WriteLine($"\nSummary: {result.Summary.SheetsAdded} sheets added, "
            + $"{result.Summary.SheetsDeleted} deleted, "
            + $"{result.Summary.CellsModified} cells modified, "
            + $"{result.Summary.CellsAdded} added, "
            + $"{result.Summary.CellsDeleted} deleted, "
            + $"{result.Summary.FormulasModified} formulas modified, "
            + $"{result.Summary.FormulasAdded} added, "
            + $"{result.Summary.FormulasDeleted} deleted, "
            + $"{result.Summary.StructureChanges} structure changes");
        Console.WriteLine();
    }

    private static string Trunc(string s, int max) =>
        s.Length <= max ? s : s[..max] + "…";
}
