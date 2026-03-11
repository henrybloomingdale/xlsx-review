using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using ThreadedComments = DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XlsxReview;

public class SpreadsheetEditor
{
    private readonly string _author;

    public SpreadsheetEditor(string author)
    {
        _author = author;
    }

    // ── Read Mode ──

    public ReadResult ReadSpreadsheet(string inputPath)
    {
        SpreadsheetPackagePreflight.Validate(inputPath);

        var result = new ReadResult();

        using var doc = SpreadsheetDocument.Open(inputPath, false);
        var workbookPart = doc.WorkbookPart
            ?? throw new Exception("Invalid spreadsheet: no workbook part");

        var sheets = workbookPart.Workbook.Sheets?.Elements<Sheet>().ToList() ?? new List<Sheet>();
        var sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable;
        result.Workbook = BuildWorkbookInfo(doc, workbookPart, sheets);

        foreach (var sheet in sheets)
        {
            var sheetData = new SheetData_Read
            {
                Name = sheet.Name?.Value ?? "Unnamed",
                Visibility = GetSheetVisibility(sheet)
            };

            var sheetPart = TryGetSheetPart(workbookPart, sheet);
            sheetData.Kind = sheetPart.Kind;

            if (sheetPart.Warning != null)
            {
                result.Warnings.Add(new ReadWarning
                {
                    Scope = "sheet",
                    Target = sheetData.Name,
                    Message = sheetPart.Warning
                });
            }

            if (sheetPart.WorksheetPart == null)
            {
                result.Sheets.Add(new SheetData
                {
                    Name = sheetData.Name,
                    Kind = sheetData.Kind,
                    Visibility = sheetData.Visibility,
                    RowCount = sheetData.RowCount,
                    CellCount = sheetData.CellCount,
                    Tables = sheetData.Tables,
                    DataValidations = sheetData.DataValidations,
                    ConditionalFormats = sheetData.ConditionalFormats,
                    Rows = sheetData.Rows
                });
                continue;
            }

            sheetData.CommentCount = GetCommentCount(sheetPart.WorksheetPart);
            sheetData.ThreadedCommentCount = GetThreadedCommentCount(sheetPart.WorksheetPart);
            sheetData.Protected = IsSheetProtected(sheetPart.WorksheetPart);
            sheetData.MergedRanges = GetMergedRanges(sheetPart.WorksheetPart);
            sheetData.MergedCellCount = sheetData.MergedRanges.Count;
            sheetData.FreezePaneCell = GetFreezePaneCell(sheetPart.WorksheetPart);
            sheetData.AutoFilterRange = GetAutoFilterRange(sheetPart.WorksheetPart);
            sheetData.Tables = GetTableInfos(sheetPart.WorksheetPart);
            sheetData.TableCount = sheetData.Tables.Count;
            sheetData.DataValidations = GetDataValidationInfos(sheetPart.WorksheetPart);
            sheetData.DataValidationCount = sheetData.DataValidations.Count;
            sheetData.ConditionalFormats = GetConditionalFormatInfos(sheetPart.WorksheetPart);
            sheetData.ConditionalFormatCount = sheetData.ConditionalFormats.Sum(x => x.RuleCount);
            sheetData.PivotTableCount = sheetPart.WorksheetPart.PivotTableParts.Count();

            var rows = sheetPart.WorksheetPart.Worksheet.Descendants<Row>();
            foreach (var row in rows)
            {
                var rowData = new RowData { Row = (int)(row.RowIndex?.Value ?? 0) };

                foreach (var cell in row.Elements<Cell>())
                {
                    string? cellRef = cell.CellReference?.Value;
                    if (cellRef == null) continue;

                    string? cellValue = GetCellValue(cell, sharedStrings);
                    string? formula = cell.CellFormula?.Text;
                    string? formulaKind = GetFormulaKind(cell);
                    string cellType = GetCellType(cell);

                    if (formula != null)
                    {
                        sheetData.FormulaCount++;
                        switch (formulaKind)
                        {
                            case "shared":
                                sheetData.SharedFormulaCount++;
                                break;
                            case "array":
                                sheetData.ArrayFormulaCount++;
                                break;
                            case "dataTable":
                                sheetData.DataTableFormulaCount++;
                                break;
                        }
                    }

                    rowData.Cells.Add(new CellData
                    {
                        Cell = cellRef,
                        Value = cellValue,
                        Formula = formula,
                        FormulaKind = formulaKind,
                        Type = cellType
                    });
                    sheetData.CellCount++;
                }

                if (rowData.Cells.Count > 0)
                {
                    sheetData.Rows.Add(rowData);
                    sheetData.RowCount++;
                }
            }

            result.Sheets.Add(new SheetData
            {
                Name = sheetData.Name,
                Kind = sheetData.Kind,
                Visibility = sheetData.Visibility,
                RowCount = sheetData.RowCount,
                CellCount = sheetData.CellCount,
                FormulaCount = sheetData.FormulaCount,
                SharedFormulaCount = sheetData.SharedFormulaCount,
                ArrayFormulaCount = sheetData.ArrayFormulaCount,
                DataTableFormulaCount = sheetData.DataTableFormulaCount,
                CommentCount = sheetData.CommentCount,
                ThreadedCommentCount = sheetData.ThreadedCommentCount,
                TableCount = sheetData.TableCount,
                DataValidationCount = sheetData.DataValidationCount,
                ConditionalFormatCount = sheetData.ConditionalFormatCount,
                PivotTableCount = sheetData.PivotTableCount,
                Protected = sheetData.Protected,
                MergedCellCount = sheetData.MergedCellCount,
                MergedRanges = sheetData.MergedRanges,
                FreezePaneCell = sheetData.FreezePaneCell,
                AutoFilterRange = sheetData.AutoFilterRange,
                Tables = sheetData.Tables,
                DataValidations = sheetData.DataValidations,
                ConditionalFormats = sheetData.ConditionalFormats,
                Rows = sheetData.Rows
            });
        }

        return result;
    }

    private static string? GetCellValue(Cell cell, SharedStringTable? sharedStrings)
    {
        // Handle InlineString cells
        if (cell.DataType?.Value == CellValues.InlineString)
        {
            return cell.InlineString?.Text?.Text ?? cell.InlineString?.InnerText;
        }

        if (cell.CellValue == null) return null;
        string value = cell.CellValue.Text;

        if (cell.DataType?.Value == CellValues.SharedString && sharedStrings != null)
        {
            if (int.TryParse(value, out int idx))
            {
                if (idx >= 0 && idx < sharedStrings.ChildElements.Count)
                {
                    var ssItem = sharedStrings.ChildElements[idx] as SharedStringItem;
                    return ssItem?.InnerText ?? value;
                }

                return value;
            }
        }
        else if (cell.DataType?.Value == CellValues.Boolean)
        {
            return value == "1" ? "TRUE" : "FALSE";
        }

        return value;
    }

    private static WorkbookInfo BuildWorkbookInfo(SpreadsheetDocument document, WorkbookPart workbookPart, List<Sheet> sheets)
    {
        var workbook = workbookPart.Workbook;
        var definedNames = BuildDefinedNames(workbook, sheets);
        var externalLinks = BuildExternalLinks(workbookPart, workbook);
        var workbookProtection = BuildWorkbookProtection(workbook.WorkbookProtection);
        var info = new WorkbookInfo
        {
            DocumentType = GetDocumentTypeName(document.DocumentType),
            SheetCount = sheets.Count,
            DefinedNameCount = definedNames.Count,
            DefinedNames = definedNames,
            ExternalLinkCount = externalLinks.Count,
            ExternalLinks = externalLinks,
            HasMacros = workbookPart.GetPartsOfType<VbaProjectPart>().Any(),
            Protected = workbookProtection.Enabled,
            WorkbookProtection = workbookProtection
        };

        foreach (var sheet in sheets)
        {
            string visibility = GetSheetVisibility(sheet);
            if (visibility == "hidden") info.HiddenSheetCount++;
            if (visibility == "veryHidden") info.VeryHiddenSheetCount++;

            var sheetPart = TryGetSheetPart(workbookPart, sheet);
            switch (sheetPart.Kind)
            {
                case "worksheet":
                    info.WorksheetCount++;
                    break;
                case "chartsheet":
                    info.ChartsheetCount++;
                    break;
                case "dialogsheet":
                    info.DialogsheetCount++;
                    break;
            }
        }

        return info;
    }

    private static string GetDocumentTypeName(SpreadsheetDocumentType documentType) => documentType switch
    {
        SpreadsheetDocumentType.Workbook => "workbook",
        SpreadsheetDocumentType.Template => "template",
        SpreadsheetDocumentType.MacroEnabledWorkbook => "macroEnabledWorkbook",
        SpreadsheetDocumentType.MacroEnabledTemplate => "macroEnabledTemplate",
        SpreadsheetDocumentType.AddIn => "addIn",
        _ => documentType.ToString()
    };

    private static string GetCellType(Cell cell)
    {
        if (cell.CellFormula != null) return "formula";
        if (cell.DataType?.Value == CellValues.InlineString) return "string";
        if (cell.DataType?.Value == CellValues.SharedString) return "string";
        if (cell.DataType?.Value == CellValues.Boolean) return "boolean";
        if (cell.CellValue != null) return "number";
        return "empty";
    }

    private static string? GetFormulaKind(Cell cell)
    {
        if (cell.CellFormula == null)
            return null;

        var formulaType = cell.CellFormula.FormulaType?.Value;
        if (formulaType == CellFormulaValues.Shared) return "shared";
        if (formulaType == CellFormulaValues.Array) return "array";
        if (formulaType == CellFormulaValues.DataTable) return "dataTable";
        return "normal";
    }

    internal static WorkbookProtectionInfo BuildWorkbookProtection(WorkbookProtection? protection)
    {
        if (protection == null)
            return new WorkbookProtectionInfo();

        bool lockStructure = protection.LockStructure?.Value ?? false;
        bool lockWindows = protection.LockWindows?.Value ?? false;
        bool lockRevision = protection.LockRevision?.Value ?? false;

        return new WorkbookProtectionInfo
        {
            Enabled = lockStructure || lockWindows || lockRevision || protection.HasAttributes,
            LockStructure = lockStructure,
            LockWindows = lockWindows,
            LockRevision = lockRevision
        };
    }

    internal static List<DefinedNameInfo> BuildDefinedNames(Workbook workbook, List<Sheet> sheets)
    {
        return workbook.DefinedNames?.Elements<DefinedName>()
            .Select(name =>
            {
                string? scopeSheet = null;
                uint? localSheetId = name.LocalSheetId?.Value;
                if (localSheetId.HasValue && localSheetId.Value < sheets.Count)
                    scopeSheet = sheets[(int)localSheetId.Value].Name?.Value;

                return new DefinedNameInfo
                {
                    Name = name.Name?.Value ?? "",
                    ScopeSheet = scopeSheet,
                    RefersTo = name.Text ?? name.InnerText,
                    Hidden = name.Hidden?.Value ?? false,
                    BuiltIn = (name.Name?.Value ?? "").StartsWith("_xlnm.", StringComparison.Ordinal),
                    Comment = name.Comment?.Value
                };
            })
            .ToList() ?? new List<DefinedNameInfo>();
    }

    private static List<ExternalLinkInfo> BuildExternalLinks(WorkbookPart workbookPart, Workbook workbook)
    {
        var links = new List<ExternalLinkInfo>();

        foreach (var reference in workbook.Elements<ExternalReferences>()
                     .SelectMany(x => x.Elements<ExternalReference>()))
        {
            string relId = reference.Id?.Value ?? "";
            string? target = null;
            string? relationshipType = null;

            if (!string.IsNullOrWhiteSpace(relId))
            {
                var linkPart = workbookPart.Parts
                    .FirstOrDefault(x => x.RelationshipId == relId)
                    .OpenXmlPart as ExternalWorkbookPart;
                var externalRelationship = linkPart?.ExternalRelationships.FirstOrDefault();
                target = externalRelationship?.Uri?.ToString();
                relationshipType = externalRelationship?.RelationshipType;
            }

            links.Add(new ExternalLinkInfo
            {
                RelationshipId = relId,
                Target = target,
                RelationshipType = relationshipType,
                Broken = string.IsNullOrWhiteSpace(relId) || string.IsNullOrWhiteSpace(target)
            });
        }

        return links;
    }

    private static int GetCommentCount(WorksheetPart worksheetPart)
    {
        return worksheetPart.WorksheetCommentsPart?.Comments?
            .GetFirstChild<CommentList>()?
            .Elements<Comment>()
            .Count() ?? 0;
    }

    private static int GetThreadedCommentCount(WorksheetPart worksheetPart)
    {
        return worksheetPart.WorksheetThreadedCommentsParts
            .SelectMany(part => part.ThreadedComments?.Elements<ThreadedComments.ThreadedComment>()
                ?? Enumerable.Empty<ThreadedComments.ThreadedComment>())
            .Count();
    }

    internal static bool IsSheetProtected(WorksheetPart worksheetPart)
    {
        return worksheetPart.Worksheet.GetFirstChild<SheetProtection>() != null;
    }

    private static List<TableInfo> GetTableInfos(WorksheetPart worksheetPart)
    {
        return worksheetPart.TableDefinitionParts
            .Select(part =>
            {
                var table = part.Table;
                return new TableInfo
                {
                    Name = table?.Name?.Value,
                    DisplayName = table?.DisplayName?.Value,
                    Reference = table?.Reference?.Value,
                    TotalsRowShown = table?.TotalsRowShown?.Value ?? false,
                    HeaderRowCount = table?.HeaderRowCount?.Value,
                    StyleName = table?.TableStyleInfo?.Name?.Value
                };
            })
            .ToList();
    }

    private static List<DataValidationInfo> GetDataValidationInfos(WorksheetPart worksheetPart)
    {
        return worksheetPart.Worksheet.Elements<DataValidations>()
            .SelectMany(x => x.Elements<DataValidation>())
            .Select(validation => new DataValidationInfo
            {
                Range = validation.SequenceOfReferences?.InnerText,
                Type = validation.Type?.Value.ToString(),
                Operator = validation.Operator?.Value.ToString(),
                AllowBlank = validation.AllowBlank?.Value ?? false,
                ShowInputMessage = validation.ShowInputMessage?.Value ?? false,
                ShowErrorMessage = validation.ShowErrorMessage?.Value ?? false,
                Formula1 = validation.Formula1?.InnerText,
                Formula2 = validation.Formula2?.InnerText
            })
            .ToList();
    }

    private static List<ConditionalFormatInfo> GetConditionalFormatInfos(WorksheetPart worksheetPart)
    {
        return worksheetPart.Worksheet.Elements<ConditionalFormatting>()
            .Select(format => new ConditionalFormatInfo
            {
                Range = format.SequenceOfReferences?.InnerText,
                RuleCount = format.Elements<ConditionalFormattingRule>().Count(),
                RuleTypes = format.Elements<ConditionalFormattingRule>()
                    .Select(rule => rule.Type?.Value.ToString() ?? "unknown")
                    .Distinct()
                    .ToList(),
                Priorities = format.Elements<ConditionalFormattingRule>()
                    .Select(rule => (int)(rule.Priority?.Value ?? 0))
                    .Where(priority => priority > 0)
                    .ToList()
            })
            .ToList();
    }

    private static List<string> GetMergedRanges(WorksheetPart worksheetPart)
    {
        return worksheetPart.Worksheet.Elements<MergeCells>()
            .SelectMany(x => x.Elements<MergeCell>())
            .Select(x => x.Reference?.Value)
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Cast<string>()
            .ToList();
    }

    private static string? GetFreezePaneCell(WorksheetPart worksheetPart)
    {
        return worksheetPart.Worksheet.GetFirstChild<SheetViews>()?
            .Elements<SheetView>()
            .Select(view => view.GetFirstChild<Pane>())
            .Where(pane => pane?.State?.Value == PaneStateValues.Frozen || pane?.State?.Value == PaneStateValues.FrozenSplit)
            .Select(pane => pane?.TopLeftCell?.Value)
            .FirstOrDefault(cell => !string.IsNullOrWhiteSpace(cell));
    }

    private static string? GetAutoFilterRange(WorksheetPart worksheetPart)
    {
        return worksheetPart.Worksheet.GetFirstChild<AutoFilter>()?.Reference?.Value;
    }

    internal static string GetSheetVisibility(Sheet sheet)
    {
        var state = sheet.State?.Value;
        if (state == SheetStateValues.Hidden) return "hidden";
        if (state == SheetStateValues.VeryHidden) return "veryHidden";
        return "visible";
    }

    // ── Edit Mode ──

    public ProcessingResult Process(string inputPath, string outputPath, EditManifest manifest, bool dryRun)
    {
        var result = new ProcessingResult
        {
            Input = Path.GetFileName(inputPath),
            Output = dryRun ? null : Path.GetFileName(outputPath),
            Author = _author
        };

        var changes = manifest.Changes ?? new List<Change>();
        var comments = manifest.Comments ?? new List<CommentDef>();

        result.ChangesAttempted = changes.Count;
        result.CommentsAttempted = comments.Count;

        if (dryRun)
        {
            // Validate manifest without modifying
            for (int i = 0; i < changes.Count; i++)
            {
                var c = changes[i];
                var validation = ValidateChange(c, i);
                result.Results.Add(validation);
                if (validation.Success) result.ChangesSucceeded++;
            }
            for (int i = 0; i < comments.Count; i++)
            {
                var cm = comments[i];
                bool valid = !string.IsNullOrEmpty(cm.Sheet) && !string.IsNullOrEmpty(cm.Cell) && !string.IsNullOrEmpty(cm.Text);
                result.Results.Add(new EditResult
                {
                    Index = i,
                    Type = "comment",
                    Success = valid,
                    Message = valid ? $"Comment on {cm.Sheet}!{cm.Cell} would be added" : "Missing required fields (sheet, cell, text)"
                });
                if (valid) result.CommentsSucceeded++;
            }
            result.Success = result.Results.All(r => r.Success);
            return result;
        }

        // Copy input to output
        SpreadsheetPackagePreflight.Validate(inputPath);
        File.Copy(inputPath, outputPath, overwrite: true);

        using var doc = SpreadsheetDocument.Open(outputPath, true);
        var workbookPart = doc.WorkbookPart
            ?? throw new Exception("Invalid spreadsheet: no workbook part");

        // Apply changes
        for (int i = 0; i < changes.Count; i++)
        {
            var c = changes[i];
            try
            {
                ApplyChange(workbookPart, c);
                result.Results.Add(new EditResult
                {
                    Index = i,
                    Type = c.Type,
                    Success = true,
                    Message = DescribeChange(c)
                });
                result.ChangesSucceeded++;
            }
            catch (Exception ex)
            {
                result.Results.Add(new EditResult
                {
                    Index = i,
                    Type = c.Type,
                    Success = false,
                    Message = ex.Message
                });
            }
        }

        // Apply comments
        for (int i = 0; i < comments.Count; i++)
        {
            var cm = comments[i];
            try
            {
                ApplyComment(workbookPart, cm);
                result.Results.Add(new EditResult
                {
                    Index = i,
                    Type = "comment",
                    Success = true,
                    Message = $"Comment added on {cm.Sheet}!{cm.Cell}"
                });
                result.CommentsSucceeded++;
            }
            catch (Exception ex)
            {
                result.Results.Add(new EditResult
                {
                    Index = i,
                    Type = "comment",
                    Success = false,
                    Message = ex.Message
                });
            }
        }

        doc.Save();
        result.Success = result.Results.All(r => r.Success);
        return result;
    }

    private EditResult ValidateChange(Change c, int index)
    {
        string type = c.Type;
        bool valid = type switch
        {
            "set_cell" => !string.IsNullOrEmpty(c.Sheet) && !string.IsNullOrEmpty(c.Cell) && c.Value != null,
            "set_formula" => !string.IsNullOrEmpty(c.Sheet) && !string.IsNullOrEmpty(c.Cell) && !string.IsNullOrEmpty(c.Formula),
            "insert_row" => !string.IsNullOrEmpty(c.Sheet) && c.After != null,
            "delete_row" => !string.IsNullOrEmpty(c.Sheet) && c.Row != null,
            "insert_column" => !string.IsNullOrEmpty(c.Sheet) && c.After != null,
            "delete_column" => !string.IsNullOrEmpty(c.Sheet) && !string.IsNullOrEmpty(c.Column),
            "add_sheet" => !string.IsNullOrEmpty(c.Name),
            "rename_sheet" => !string.IsNullOrEmpty(c.From) && !string.IsNullOrEmpty(c.To),
            "delete_sheet" => !string.IsNullOrEmpty(c.Name),
            "set_sheet_visibility" => !string.IsNullOrEmpty(c.Name ?? c.Sheet) && IsValidVisibility(c.Visibility),
            "set_defined_name" => !string.IsNullOrEmpty(c.Name) && !string.IsNullOrEmpty(c.RefersTo),
            "add_defined_name" => !string.IsNullOrEmpty(c.Name) && !string.IsNullOrEmpty(c.RefersTo),
            "delete_defined_name" => !string.IsNullOrEmpty(c.Name),
            "set_workbook_protection" => c.Enabled != null || c.LockStructure != null || c.LockWindows != null || c.LockRevision != null,
            "set_sheet_protection" => !string.IsNullOrEmpty(c.Sheet) && c.Enabled != null,
            "merge_cells" => !string.IsNullOrEmpty(c.Sheet) && !string.IsNullOrEmpty(c.Range),
            "unmerge_cells" => !string.IsNullOrEmpty(c.Sheet) && !string.IsNullOrEmpty(c.Range),
            "set_freeze_panes" => !string.IsNullOrEmpty(c.Sheet) && !string.IsNullOrEmpty(c.Cell),
            "clear_freeze_panes" => !string.IsNullOrEmpty(c.Sheet),
            "set_auto_filter" => !string.IsNullOrEmpty(c.Sheet) && !string.IsNullOrEmpty(c.Range),
            "clear_auto_filter" => !string.IsNullOrEmpty(c.Sheet),
            _ => false
        };

        return new EditResult
        {
            Index = index,
            Type = type,
            Success = valid,
            Message = valid ? $"{DescribeChange(c)} would succeed" : $"Invalid {type}: missing required fields"
        };
    }

    private string DescribeChange(Change c) => c.Type switch
    {
        "set_cell" => $"Set {c.Sheet}!{c.Cell} = \"{c.Value}\"",
        "set_formula" => $"Set formula {c.Sheet}!{c.Cell} = {c.Formula}",
        "insert_row" => $"Inserted row after {c.Sheet}!row {GetAfterInt(c.After)}",
        "delete_row" => $"Deleted {c.Sheet}!row {c.Row}",
        "insert_column" => $"Inserted column after {c.Sheet}!{GetAfterString(c.After)}",
        "delete_column" => $"Deleted {c.Sheet}!column {c.Column}",
        "add_sheet" => $"Added sheet \"{c.Name}\"",
        "rename_sheet" => $"Renamed sheet \"{c.From}\" → \"{c.To}\"",
        "delete_sheet" => $"Deleted sheet \"{c.Name}\"",
        "set_sheet_visibility" => $"Set sheet \"{GetSheetTarget(c)}\" visibility to {NormalizeVisibility(c.Visibility!)}",
        "set_defined_name" => $"Set defined name \"{c.Name}\" → {c.RefersTo}" + (string.IsNullOrEmpty(c.ScopeSheet) ? "" : $" (scope: {c.ScopeSheet})"),
        "add_defined_name" => $"Set defined name \"{c.Name}\" → {c.RefersTo}" + (string.IsNullOrEmpty(c.ScopeSheet) ? "" : $" (scope: {c.ScopeSheet})"),
        "delete_defined_name" => $"Deleted defined name \"{c.Name}\"" + (string.IsNullOrEmpty(c.ScopeSheet) ? "" : $" (scope: {c.ScopeSheet})"),
        "set_workbook_protection" => DescribeWorkbookProtectionChange(c),
        "set_sheet_protection" => $"Set sheet protection on {c.Sheet} to {(c.Enabled == true ? "enabled" : "disabled")}",
        "merge_cells" => $"Merged {c.Sheet}!{c.Range}",
        "unmerge_cells" => $"Unmerged {c.Sheet}!{c.Range}",
        "set_freeze_panes" => $"Set freeze panes on {c.Sheet} at {c.Cell}",
        "clear_freeze_panes" => $"Cleared freeze panes on {c.Sheet}",
        "set_auto_filter" => $"Set auto filter on {c.Sheet}!{c.Range}",
        "clear_auto_filter" => $"Cleared auto filter on {c.Sheet}",
        _ => $"Unknown change type: {c.Type}"
    };

    private void ApplyChange(WorkbookPart workbookPart, Change c)
    {
        switch (c.Type)
        {
            case "set_cell":
                SetCell(workbookPart, c.Sheet!, c.Cell!, c.Value!, c.Format);
                break;
            case "set_formula":
                SetFormula(workbookPart, c.Sheet!, c.Cell!, c.Formula!);
                break;
            case "insert_row":
                InsertRow(workbookPart, c.Sheet!, GetAfterInt(c.After));
                break;
            case "delete_row":
                DeleteRow(workbookPart, c.Sheet!, c.Row!.Value);
                break;
            case "insert_column":
                InsertColumn(workbookPart, c.Sheet!, GetAfterString(c.After));
                break;
            case "delete_column":
                DeleteColumn(workbookPart, c.Sheet!, c.Column!);
                break;
            case "add_sheet":
                AddSheet(workbookPart, c.Name!);
                break;
            case "rename_sheet":
                RenameSheet(workbookPart, c.From!, c.To!);
                break;
            case "delete_sheet":
                DeleteSheet(workbookPart, c.Name!);
                break;
            case "set_sheet_visibility":
                SetSheetVisibility(workbookPart, GetSheetTarget(c), c.Visibility!);
                break;
            case "set_defined_name":
            case "add_defined_name":
                SetDefinedName(workbookPart, c.Name!, c.RefersTo!, c.ScopeSheet, c.Hidden ?? false, c.Comment);
                break;
            case "delete_defined_name":
                DeleteDefinedName(workbookPart, c.Name!, c.ScopeSheet);
                break;
            case "set_workbook_protection":
                SetWorkbookProtection(workbookPart, c.Enabled, c.LockStructure, c.LockWindows, c.LockRevision);
                break;
            case "set_sheet_protection":
                SetSheetProtection(workbookPart, c.Sheet!, c.Enabled!.Value);
                break;
            case "merge_cells":
                MergeCells(workbookPart, c.Sheet!, c.Range!);
                break;
            case "unmerge_cells":
                UnmergeCells(workbookPart, c.Sheet!, c.Range!);
                break;
            case "set_freeze_panes":
                SetFreezePanes(workbookPart, c.Sheet!, c.Cell!);
                break;
            case "clear_freeze_panes":
                ClearFreezePanes(workbookPart, c.Sheet!);
                break;
            case "set_auto_filter":
                SetAutoFilter(workbookPart, c.Sheet!, c.Range!);
                break;
            case "clear_auto_filter":
                ClearAutoFilter(workbookPart, c.Sheet!);
                break;
            default:
                throw new Exception($"Unknown change type: {c.Type}");
        }
    }

    // ── Cell Operations ──

    private void SetCell(WorkbookPart workbookPart, string sheetName, string cellRef, string value, string? format)
    {
        var worksheetPart = GetWorksheetPart(workbookPart, sheetName);
        var sheetData = worksheetPart.Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>()
            ?? throw new Exception($"Sheet '{sheetName}' has no data element");

        var (colName, rowIndex) = ParseCellReference(cellRef);
        var row = EnsureRow(sheetData, rowIndex);
        var cell = EnsureCell(row, cellRef, colName);

        // Set value based on format
        if (format == "number" && double.TryParse(value, out _))
        {
            cell.DataType = null; // numeric
            cell.CellValue = new CellValue(value);
        }
        else
        {
            cell.DataType = new EnumValue<CellValues>(CellValues.InlineString);
            cell.CellValue = null;
            cell.InlineString = new InlineString(new Text(value));
        }

        // Apply yellow highlight
        ApplyYellowFill(workbookPart, cell);
    }

    private void SetFormula(WorkbookPart workbookPart, string sheetName, string cellRef, string formula)
    {
        var worksheetPart = GetWorksheetPart(workbookPart, sheetName);
        var sheetData = worksheetPart.Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>()
            ?? throw new Exception($"Sheet '{sheetName}' has no data element");

        var (colName, rowIndex) = ParseCellReference(cellRef);
        var row = EnsureRow(sheetData, rowIndex);
        var cell = EnsureCell(row, cellRef, colName);

        cell.CellFormula = new CellFormula(formula);
        cell.CellValue = null; // Excel will calculate
        cell.DataType = null;

        ApplyYellowFill(workbookPart, cell);
    }

    // ── Row Operations ──

    private void InsertRow(WorkbookPart workbookPart, string sheetName, int afterRow)
    {
        var worksheetPart = GetWorksheetPart(workbookPart, sheetName);
        var sheetData = worksheetPart.Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>()
            ?? throw new Exception($"Sheet '{sheetName}' has no data element");

        uint newRowIndex = (uint)(afterRow + 1);

        // Shift existing rows down
        var rowsToShift = sheetData.Elements<Row>()
            .Where(r => r.RowIndex?.Value >= newRowIndex)
            .OrderByDescending(r => r.RowIndex?.Value)
            .ToList();

        foreach (var row in rowsToShift)
        {
            uint oldIndex = row.RowIndex!.Value;
            uint newIndex = oldIndex + 1;
            row.RowIndex = new UInt32Value(newIndex);

            foreach (var cell in row.Elements<Cell>())
            {
                var (col, _) = ParseCellReference(cell.CellReference?.Value ?? "A1");
                cell.CellReference = new StringValue($"{col}{newIndex}");
            }
        }

        // Insert empty row
        var newRow = new Row { RowIndex = new UInt32Value(newRowIndex) };

        var insertBefore = sheetData.Elements<Row>()
            .FirstOrDefault(r => r.RowIndex?.Value >= newRowIndex);

        if (insertBefore != null)
            sheetData.InsertBefore(newRow, insertBefore);
        else
            sheetData.AppendChild(newRow);
    }

    private void DeleteRow(WorkbookPart workbookPart, string sheetName, int rowNumber)
    {
        var worksheetPart = GetWorksheetPart(workbookPart, sheetName);
        var sheetData = worksheetPart.Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>()
            ?? throw new Exception($"Sheet '{sheetName}' has no data element");

        uint targetIndex = (uint)rowNumber;
        var targetRow = sheetData.Elements<Row>()
            .FirstOrDefault(r => r.RowIndex?.Value == targetIndex);

        if (targetRow != null)
            sheetData.RemoveChild(targetRow);

        // Shift rows up
        var rowsToShift = sheetData.Elements<Row>()
            .Where(r => r.RowIndex?.Value > targetIndex)
            .OrderBy(r => r.RowIndex?.Value)
            .ToList();

        foreach (var row in rowsToShift)
        {
            uint oldIndex = row.RowIndex!.Value;
            uint newIndex = oldIndex - 1;
            row.RowIndex = new UInt32Value(newIndex);

            foreach (var cell in row.Elements<Cell>())
            {
                var (col, _) = ParseCellReference(cell.CellReference?.Value ?? "A1");
                cell.CellReference = new StringValue($"{col}{newIndex}");
            }
        }
    }

    // ── Column Operations ──

    private void InsertColumn(WorkbookPart workbookPart, string sheetName, string afterColumn)
    {
        var worksheetPart = GetWorksheetPart(workbookPart, sheetName);
        var sheetData = worksheetPart.Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>()
            ?? throw new Exception($"Sheet '{sheetName}' has no data element");

        int afterColIndex = ColumnNameToIndex(afterColumn);
        int newColIndex = afterColIndex + 1;

        foreach (var row in sheetData.Elements<Row>())
        {
            var cellsToShift = row.Elements<Cell>()
                .Select(cell =>
                {
                    var (col, rowIdx) = ParseCellReference(cell.CellReference?.Value ?? "A1");
                    return new { Cell = cell, ColIndex = ColumnNameToIndex(col), Col = col, RowIdx = rowIdx };
                })
                .Where(x => x.ColIndex >= newColIndex)
                .OrderByDescending(x => x.ColIndex)
                .ToList();

            foreach (var item in cellsToShift)
            {
                string newCol = IndexToColumnName(item.ColIndex + 1);
                item.Cell.CellReference = new StringValue($"{newCol}{item.RowIdx}");
            }
        }
    }

    private void DeleteColumn(WorkbookPart workbookPart, string sheetName, string columnName)
    {
        var worksheetPart = GetWorksheetPart(workbookPart, sheetName);
        var sheetData = worksheetPart.Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>()
            ?? throw new Exception($"Sheet '{sheetName}' has no data element");

        int targetColIndex = ColumnNameToIndex(columnName);

        foreach (var row in sheetData.Elements<Row>())
        {
            // Remove cells in the target column
            var cellsToRemove = row.Elements<Cell>()
                .Where(cell =>
                {
                    var (col, _) = ParseCellReference(cell.CellReference?.Value ?? "A1");
                    return ColumnNameToIndex(col) == targetColIndex;
                })
                .ToList();

            foreach (var cell in cellsToRemove)
                row.RemoveChild(cell);

            // Shift remaining cells left
            var cellsToShift = row.Elements<Cell>()
                .Select(cell =>
                {
                    var (col, rowIdx) = ParseCellReference(cell.CellReference?.Value ?? "A1");
                    return new { Cell = cell, ColIndex = ColumnNameToIndex(col), RowIdx = rowIdx };
                })
                .Where(x => x.ColIndex > targetColIndex)
                .OrderBy(x => x.ColIndex)
                .ToList();

            foreach (var item in cellsToShift)
            {
                string newCol = IndexToColumnName(item.ColIndex - 1);
                item.Cell.CellReference = new StringValue($"{newCol}{item.RowIdx}");
            }
        }
    }

    // ── Sheet Operations ──

    private void AddSheet(WorkbookPart workbookPart, string name)
    {
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new DocumentFormat.OpenXml.Spreadsheet.SheetData());

        var sheets = workbookPart.Workbook.GetFirstChild<Sheets>()
            ?? workbookPart.Workbook.AppendChild(new Sheets());

        uint sheetId = 1;
        if (sheets.Elements<Sheet>().Any())
            sheetId = sheets.Elements<Sheet>().Max(s => s.SheetId?.Value ?? 0) + 1;

        sheets.AppendChild(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = new UInt32Value(sheetId),
            Name = name
        });
    }

    private void RenameSheet(WorkbookPart workbookPart, string from, string to)
    {
        var sheets = workbookPart.Workbook.GetFirstChild<Sheets>()
            ?? throw new Exception("No sheets found in workbook");

        var sheet = sheets.Elements<Sheet>()
            .FirstOrDefault(s => s.Name?.Value == from)
            ?? throw new Exception($"Sheet '{from}' not found");

        sheet.Name = to;
    }

    private void DeleteSheet(WorkbookPart workbookPart, string name)
    {
        var sheets = workbookPart.Workbook.GetFirstChild<Sheets>()
            ?? throw new Exception("No sheets found in workbook");

        var sheet = sheets.Elements<Sheet>()
            .FirstOrDefault(s => s.Name?.Value == name)
            ?? throw new Exception($"Sheet '{name}' not found");

        string relId = sheet.Id?.Value ?? "";
        sheet.Remove();

        if (!string.IsNullOrEmpty(relId))
        {
            var part = workbookPart.GetPartById(relId);
            if (part != null) workbookPart.DeletePart(part);
        }
    }

    private void SetSheetVisibility(WorkbookPart workbookPart, string sheetName, string visibility)
    {
        var sheet = GetSheetElement(workbookPart, sheetName);
        string normalized = NormalizeVisibility(visibility);

        sheet.State = normalized switch
        {
            "visible" => null,
            "hidden" => SheetStateValues.Hidden,
            "veryHidden" => SheetStateValues.VeryHidden,
            _ => throw new Exception($"Unsupported visibility: {visibility}")
        };
    }

    private void SetDefinedName(WorkbookPart workbookPart, string name, string refersTo, string? scopeSheet, bool hidden, string? comment)
    {
        var definedNames = workbookPart.Workbook.DefinedNames;
        if (definedNames == null)
        {
            definedNames = new DefinedNames();
            workbookPart.Workbook.AppendChild(definedNames);
        }

        uint? localSheetId = scopeSheet != null ? GetSheetIndex(workbookPart, scopeSheet) : null;
        var definedName = FindDefinedName(definedNames, name, localSheetId);

        if (definedName == null)
        {
            definedName = new DefinedName { Name = name };
            definedNames.AppendChild(definedName);
        }

        definedName.Name = name;
        definedName.Text = refersTo;
        definedName.Hidden = hidden;
        definedName.Comment = string.IsNullOrWhiteSpace(comment) ? null : comment;

        if (localSheetId.HasValue)
            definedName.LocalSheetId = localSheetId.Value;
        else
            definedName.LocalSheetId = null;
    }

    private void DeleteDefinedName(WorkbookPart workbookPart, string name, string? scopeSheet)
    {
        var definedNames = workbookPart.Workbook.DefinedNames
            ?? throw new Exception("Workbook has no defined names");

        uint? localSheetId = scopeSheet != null ? GetSheetIndex(workbookPart, scopeSheet) : null;
        var definedName = FindDefinedName(definedNames, name, localSheetId)
            ?? throw new Exception($"Defined name '{name}' not found" + (scopeSheet == null ? "" : $" for sheet '{scopeSheet}'"));

        definedName.Remove();

        if (!definedNames.Elements<DefinedName>().Any())
            definedNames.Remove();
    }

    private void SetWorkbookProtection(WorkbookPart workbookPart, bool? enabled, bool? lockStructure, bool? lockWindows, bool? lockRevision)
    {
        if (enabled == false)
        {
            workbookPart.Workbook.WorkbookProtection?.Remove();
            return;
        }

        bool anyLockSpecified = lockStructure != null || lockWindows != null || lockRevision != null;
        bool effectiveLockStructure = lockStructure ?? (enabled == true && !anyLockSpecified);
        bool effectiveLockWindows = lockWindows ?? false;
        bool effectiveLockRevision = lockRevision ?? false;

        if (!effectiveLockStructure && !effectiveLockWindows && !effectiveLockRevision)
        {
            workbookPart.Workbook.WorkbookProtection?.Remove();
            return;
        }

        var protection = workbookPart.Workbook.WorkbookProtection;
        if (protection == null)
        {
            protection = new WorkbookProtection();
            workbookPart.Workbook.AppendChild(protection);
        }

        protection.LockStructure = effectiveLockStructure;
        protection.LockWindows = effectiveLockWindows;
        protection.LockRevision = effectiveLockRevision;
    }

    private void SetSheetProtection(WorkbookPart workbookPart, string sheetName, bool enabled)
    {
        var worksheetPart = GetWorksheetPart(workbookPart, sheetName);
        var worksheet = worksheetPart.Worksheet;
        var protection = worksheet.GetFirstChild<SheetProtection>();

        if (!enabled)
        {
            protection?.Remove();
            return;
        }

        if (protection == null)
        {
            protection = new SheetProtection();
            var sheetData = worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>();
            if (sheetData != null)
                worksheet.InsertAfter(protection, sheetData);
            else
                worksheet.AppendChild(protection);
        }

        protection.Sheet = true;
    }

    private void MergeCells(WorkbookPart workbookPart, string sheetName, string range)
    {
        var worksheet = GetWorksheetPart(workbookPart, sheetName).Worksheet;
        var mergedCells = worksheet.GetFirstChild<MergeCells>();

        if (mergedCells == null)
        {
            mergedCells = InsertWorksheetElementAfterPredecessors(
                worksheet,
                new MergeCells(),
                typeof(CustomSheetViews),
                typeof(DataConsolidate),
                typeof(SortState),
                typeof(AutoFilter),
                typeof(Scenarios),
                typeof(ProtectedRanges),
                typeof(SheetProtection),
                typeof(SheetCalculationProperties),
                typeof(DocumentFormat.OpenXml.Spreadsheet.SheetData));
        }

        if (mergedCells.Elements<MergeCell>().Any(x => x.Reference?.Value == range))
            return;

        mergedCells.AppendChild(new MergeCell { Reference = range });
        mergedCells.Count = (uint)mergedCells.Elements<MergeCell>().Count();
    }

    private void UnmergeCells(WorkbookPart workbookPart, string sheetName, string range)
    {
        var worksheet = GetWorksheetPart(workbookPart, sheetName).Worksheet;
        var mergedCells = worksheet.GetFirstChild<MergeCells>()
            ?? throw new Exception($"Sheet '{sheetName}' has no merged cells");

        var mergeCell = mergedCells.Elements<MergeCell>()
            .FirstOrDefault(x => x.Reference?.Value == range)
            ?? throw new Exception($"Merged range '{range}' not found on sheet '{sheetName}'");

        mergeCell.Remove();

        if (!mergedCells.Elements<MergeCell>().Any())
            mergedCells.Remove();
        else
            mergedCells.Count = (uint)mergedCells.Elements<MergeCell>().Count();
    }

    private void SetFreezePanes(WorkbookPart workbookPart, string sheetName, string cellRef)
    {
        var worksheet = GetWorksheetPart(workbookPart, sheetName).Worksheet;
        var (colName, rowIndex) = ParseCellReference(cellRef);
        int colIndex = ColumnNameToIndex(colName);

        double xSplit = Math.Max(0, colIndex - 1);
        double ySplit = Math.Max(0, rowIndex - 1);

        if (xSplit == 0 && ySplit == 0)
        {
            ClearFreezePanes(workbookPart, sheetName);
            return;
        }

        var sheetViews = GetOrCreateSheetViews(worksheet);
        var sheetView = sheetViews.Elements<SheetView>().FirstOrDefault();
        if (sheetView == null)
        {
            sheetView = new SheetView { WorkbookViewId = 0U };
            sheetViews.AppendChild(sheetView);
        }

        var pane = sheetView.GetFirstChild<Pane>();
        if (pane == null)
        {
            pane = new Pane();
            sheetView.PrependChild(pane);
        }

        pane.State = PaneStateValues.Frozen;
        pane.TopLeftCell = cellRef;
        pane.HorizontalSplit = xSplit > 0 ? xSplit : null;
        pane.VerticalSplit = ySplit > 0 ? ySplit : null;
    }

    private void ClearFreezePanes(WorkbookPart workbookPart, string sheetName)
    {
        var worksheet = GetWorksheetPart(workbookPart, sheetName).Worksheet;
        worksheet.GetFirstChild<SheetViews>()?
            .Elements<SheetView>()
            .FirstOrDefault()?
            .GetFirstChild<Pane>()?
            .Remove();
    }

    private void SetAutoFilter(WorkbookPart workbookPart, string sheetName, string range)
    {
        var worksheet = GetWorksheetPart(workbookPart, sheetName).Worksheet;
        var autoFilter = worksheet.GetFirstChild<AutoFilter>();
        if (autoFilter == null)
        {
            autoFilter = InsertWorksheetElementAfterPredecessors(
                worksheet,
                new AutoFilter(),
                typeof(Scenarios),
                typeof(ProtectedRanges),
                typeof(SheetProtection),
                typeof(SheetCalculationProperties),
                typeof(DocumentFormat.OpenXml.Spreadsheet.SheetData));
        }

        autoFilter.Reference = range;
    }

    private void ClearAutoFilter(WorkbookPart workbookPart, string sheetName)
    {
        var worksheet = GetWorksheetPart(workbookPart, sheetName).Worksheet;
        worksheet.GetFirstChild<AutoFilter>()?.Remove();
    }

    // ── Comments (Legacy Notes) ──

    private void ApplyComment(WorkbookPart workbookPart, CommentDef commentDef)
    {
        var worksheetPart = GetWorksheetPart(workbookPart, commentDef.Sheet);

        // Ensure VmlDrawingPart exists
        var vmlPart = worksheetPart.VmlDrawingParts.FirstOrDefault();
        if (vmlPart == null)
        {
            vmlPart = worksheetPart.AddNewPart<VmlDrawingPart>();
            // Initialize VML with XML namespace
            using var writer = new StreamWriter(vmlPart.GetStream(FileMode.Create));
            writer.Write("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\"></xml>");
        }

        // Ensure WorksheetCommentsPart exists
        var commentsPart = worksheetPart.WorksheetCommentsPart;
        if (commentsPart == null)
        {
            commentsPart = worksheetPart.AddNewPart<WorksheetCommentsPart>();
            commentsPart.Comments = new Comments(
                new Authors(new Author(_author)),
                new CommentList()
            );

            // Add legacyDrawing reference to worksheet
            var legacyDrawing = worksheetPart.Worksheet.GetFirstChild<LegacyDrawing>();
            if (legacyDrawing == null)
            {
                worksheetPart.Worksheet.AppendChild(new LegacyDrawing
                {
                    Id = worksheetPart.GetIdOfPart(vmlPart)
                });
            }
        }

        var comments = commentsPart.Comments;
        var authors = comments.GetFirstChild<Authors>()!;
        var commentList = comments.GetFirstChild<CommentList>()!;

        // Find or add author
        uint authorId = 0;
        var existingAuthor = authors.Elements<Author>()
            .Select((a, idx) => new { Author = a, Index = idx })
            .FirstOrDefault(x => x.Author.Text == _author);

        if (existingAuthor != null)
        {
            authorId = (uint)existingAuthor.Index;
        }
        else
        {
            authors.AppendChild(new Author(_author));
            authorId = (uint)(authors.Elements<Author>().Count() - 1);
        }

        // Create the comment
        var comment = new Comment
        {
            Reference = commentDef.Cell,
            AuthorId = new UInt32Value(authorId)
        };

        var commentText = new CommentText();
        var run = new DocumentFormat.OpenXml.Spreadsheet.Run();
        var runProps = new DocumentFormat.OpenXml.Spreadsheet.RunProperties();
        runProps.Append(new Bold());
        runProps.Append(new FontSize { Val = 9 });
        runProps.Append(new DocumentFormat.OpenXml.Spreadsheet.Color { Indexed = 81 });
        runProps.Append(new RunFont { Val = "Tahoma" });
        run.Append(runProps);
        run.Append(new Text(_author + ":") { Space = SpaceProcessingModeValues.Preserve });
        commentText.Append(run);

        var run2 = new DocumentFormat.OpenXml.Spreadsheet.Run();
        var runProps2 = new DocumentFormat.OpenXml.Spreadsheet.RunProperties();
        runProps2.Append(new FontSize { Val = 9 });
        runProps2.Append(new DocumentFormat.OpenXml.Spreadsheet.Color { Indexed = 81 });
        runProps2.Append(new RunFont { Val = "Tahoma" });
        run2.Append(runProps2);
        run2.Append(new Text("\n" + commentDef.Text) { Space = SpaceProcessingModeValues.Preserve });
        commentText.Append(run2);

        comment.Append(commentText);
        commentList.AppendChild(comment);

        // Add VML shape for the comment
        AddVmlShape(vmlPart, commentDef.Cell);
    }

    private void AddVmlShape(VmlDrawingPart vmlPart, string cellRef)
    {
        var (colName, rowIndex) = ParseCellReference(cellRef);
        int colIdx = ColumnNameToIndex(colName) - 1; // 0-based
        int rowIdx = (int)rowIndex - 1; // 0-based

        string existingVml;
        using (var reader = new StreamReader(vmlPart.GetStream(FileMode.Open)))
        {
            existingVml = reader.ReadToEnd();
        }

        string shape = $@"<v:shape type=""#_x0000_t202"" style=""position:absolute;margin-left:80pt;margin-top:5pt;width:108pt;height:60pt;z-index:1;visibility:hidden"" fillcolor=""#ffffe1"" o:insetmode=""auto"">
  <v:fill color2=""#ffffe1""/>
  <v:shadow on=""t"" color=""black"" obscured=""t""/>
  <v:textbox style=""mso-direction-alt:auto""/>
  <x:ClientData ObjectType=""Note"">
    <x:MoveWithCells/>
    <x:SizeWithCells/>
    <x:Anchor>{colIdx + 1}, 15, {rowIdx}, 10, {colIdx + 3}, 15, {rowIdx + 4}, 4</x:Anchor>
    <x:AutoFill>False</x:AutoFill>
    <x:Row>{rowIdx}</x:Row>
    <x:Column>{colIdx}</x:Column>
  </x:ClientData>
</v:shape>";

        // Insert before closing </xml> tag
        string updatedVml;
        if (existingVml.Contains("</xml>"))
        {
            updatedVml = existingVml.Replace("</xml>", shape + "</xml>");
        }
        else
        {
            updatedVml = $"<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\">{shape}</xml>";
        }

        using var writer = new StreamWriter(vmlPart.GetStream(FileMode.Create));
        writer.Write(updatedVml);
    }

    // ── Yellow Fill Highlight ──

    private void ApplyYellowFill(WorkbookPart workbookPart, Cell cell)
    {
        var stylesheet = EnsureStylesheet(workbookPart);

        // Find or create yellow fill
        var fills = stylesheet.Fills ?? (stylesheet.Fills = new Fills());
        int yellowFillIndex = -1;

        int fillIdx = 0;
        foreach (var fill in fills.Elements<Fill>())
        {
            var patternFill = fill.PatternFill;
            if (patternFill?.PatternType?.Value == PatternValues.Solid)
            {
                var fgColor = patternFill.ForegroundColor;
                if (fgColor?.Rgb?.Value == "FFFFFF00")
                {
                    yellowFillIndex = fillIdx;
                    break;
                }
            }
            fillIdx++;
        }

        if (yellowFillIndex < 0)
        {
            var newFill = new Fill(
                new PatternFill(
                    new ForegroundColor { Rgb = new HexBinaryValue("FFFFFF00") },
                    new BackgroundColor { Indexed = 64 }
                )
                { PatternType = PatternValues.Solid }
            );
            fills.AppendChild(newFill);
            fills.Count = new UInt32Value((uint)fills.Elements<Fill>().Count());
            yellowFillIndex = (int)fills.Count.Value - 1;
        }

        // Find or create cell format with yellow fill
        var cellFormats = stylesheet.CellFormats ?? (stylesheet.CellFormats = new CellFormats());
        int yellowFormatIndex = -1;

        int fmtIdx = 0;
        foreach (var cf in cellFormats.Elements<CellFormat>())
        {
            if (cf.FillId?.Value == (uint)yellowFillIndex && cf.ApplyFill?.Value == true)
            {
                yellowFormatIndex = fmtIdx;
                break;
            }
            fmtIdx++;
        }

        if (yellowFormatIndex < 0)
        {
            var newFormat = new CellFormat
            {
                FillId = new UInt32Value((uint)yellowFillIndex),
                ApplyFill = new BooleanValue(true)
            };
            cellFormats.AppendChild(newFormat);
            cellFormats.Count = new UInt32Value((uint)cellFormats.Elements<CellFormat>().Count());
            yellowFormatIndex = (int)cellFormats.Count.Value - 1;
        }

        cell.StyleIndex = new UInt32Value((uint)yellowFormatIndex);
    }

    private Stylesheet EnsureStylesheet(WorkbookPart workbookPart)
    {
        var stylesPart = workbookPart.WorkbookStylesPart;
        if (stylesPart == null)
        {
            stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = new Stylesheet(
                new Fonts(new Font()),
                new Fills(
                    new Fill(new PatternFill { PatternType = PatternValues.None }),
                    new Fill(new PatternFill { PatternType = PatternValues.Gray125 })
                ),
                new Borders(new Border()),
                new CellFormats(new CellFormat())
            );
        }
        return stylesPart.Stylesheet;
    }

    // ── Helpers ──

    private static T InsertWorksheetElementAfterPredecessors<T>(
        Worksheet worksheet,
        T element,
        params Type[] predecessorTypes) where T : OpenXmlElement
    {
        var anchor = worksheet.Elements()
            .LastOrDefault(existing => predecessorTypes.Any(type => type.IsAssignableFrom(existing.GetType())));

        if (anchor != null)
            worksheet.InsertAfter(element, anchor);
        else
            worksheet.PrependChild(element);

        return element;
    }

    private static SheetViews GetOrCreateSheetViews(Worksheet worksheet)
    {
        var sheetViews = worksheet.GetFirstChild<SheetViews>();
        if (sheetViews != null)
            return sheetViews;

        sheetViews = new SheetViews();
        var successor = worksheet.Elements()
            .FirstOrDefault(existing =>
                existing is SheetFormatProperties
                || existing is Columns
                || existing is DocumentFormat.OpenXml.Spreadsheet.SheetData
                || existing is SheetCalculationProperties
                || existing is SheetProtection
                || existing is ProtectedRanges
                || existing is Scenarios
                || existing is AutoFilter
                || existing is SortState
                || existing is DataConsolidate
                || existing is CustomSheetViews
                || existing is MergeCells);

        if (successor != null)
            worksheet.InsertBefore(sheetViews, successor);
        else
            worksheet.PrependChild(sheetViews);

        return sheetViews;
    }

    private WorksheetPart GetWorksheetPart(WorkbookPart workbookPart, string sheetName)
    {
        var sheet = GetSheetElement(workbookPart, sheetName);

        var part = workbookPart.GetPartById(sheet.Id?.Value ?? "");
        if (part is WorksheetPart worksheetPart)
            return worksheetPart;

        string partKind = part switch
        {
            ChartsheetPart => "chartsheet",
            DialogsheetPart => "dialogsheet",
            _ => part.GetType().Name
        };

        throw new Exception($"Sheet '{sheetName}' is a {partKind} and cannot be edited");
    }

    private static Sheet GetSheetElement(WorkbookPart workbookPart, string sheetName)
    {
        var sheets = workbookPart.Workbook.GetFirstChild<Sheets>()
            ?? throw new Exception("No sheets found in workbook");

        return sheets.Elements<Sheet>()
            .FirstOrDefault(s => s.Name?.Value == sheetName)
            ?? throw new Exception($"Sheet '{sheetName}' not found");
    }

    private static string GetSheetTarget(Change change)
    {
        return change.Sheet ?? change.Name ?? throw new Exception("Missing sheet target");
    }

    private static bool IsValidVisibility(string? visibility)
    {
        if (visibility == null)
            return false;

        return visibility == "visible"
            || visibility == "hidden"
            || visibility == "veryHidden";
    }

    private static string NormalizeVisibility(string visibility)
    {
        return visibility switch
        {
            "visible" => "visible",
            "hidden" => "hidden",
            "veryHidden" => "veryHidden",
            _ => throw new Exception($"Visibility must be one of: visible, hidden, veryHidden")
        };
    }

    private static uint GetSheetIndex(WorkbookPart workbookPart, string sheetName)
    {
        var sheets = workbookPart.Workbook.GetFirstChild<Sheets>()
            ?? throw new Exception("No sheets found in workbook");

        var sheetList = sheets.Elements<Sheet>().ToList();
        for (int i = 0; i < sheetList.Count; i++)
        {
            if (sheetList[i].Name?.Value == sheetName)
                return (uint)i;
        }

        throw new Exception($"Sheet '{sheetName}' not found");
    }

    private static DefinedName? FindDefinedName(DefinedNames definedNames, string name, uint? localSheetId)
    {
        return definedNames.Elements<DefinedName>()
            .FirstOrDefault(x =>
            {
                if (x.Name?.Value != name)
                    return false;

                uint? existingLocalSheetId = x.LocalSheetId?.Value;
                return existingLocalSheetId == localSheetId;
            });
    }

    private static string DescribeWorkbookProtectionChange(Change change)
    {
        if (change.Enabled == false)
            return "Disabled workbook protection";

        bool anyLockSpecified = change.LockStructure != null || change.LockWindows != null || change.LockRevision != null;
        bool lockStructure = change.LockStructure ?? (change.Enabled == true && !anyLockSpecified);
        bool lockWindows = change.LockWindows ?? false;
        bool lockRevision = change.LockRevision ?? false;

        return $"Set workbook protection (lockStructure={lockStructure.ToString().ToLowerInvariant()}, lockWindows={lockWindows.ToString().ToLowerInvariant()}, lockRevision={lockRevision.ToString().ToLowerInvariant()})";
    }

    private static WorksheetPart? TryGetWorksheetPart(WorkbookPart workbookPart, Sheet sheet)
    {
        string? relId = sheet.Id?.Value;
        if (string.IsNullOrWhiteSpace(relId))
            return null;

        try
        {
            return workbookPart.GetPartById(relId) as WorksheetPart;
        }
        catch
        {
            return null;
        }
    }

    private static ReadSheetPart TryGetSheetPart(WorkbookPart workbookPart, Sheet sheet)
    {
        string? relId = sheet.Id?.Value;
        if (string.IsNullOrWhiteSpace(relId))
        {
            return new ReadSheetPart
            {
                Kind = "unreadable",
                Warning = "Missing relationship id"
            };
        }

        try
        {
            var part = workbookPart.GetPartById(relId);
            return part switch
            {
                WorksheetPart worksheetPart => new ReadSheetPart { Kind = "worksheet", WorksheetPart = worksheetPart },
                ChartsheetPart => new ReadSheetPart { Kind = "chartsheet" },
                DialogsheetPart => new ReadSheetPart { Kind = "dialogsheet" },
                _ => new ReadSheetPart
                {
                    Kind = "unsupported",
                    Warning = $"Unsupported sheet part type: {part.GetType().Name}"
                }
            };
        }
        catch (Exception ex)
        {
            return new ReadSheetPart
            {
                Kind = "unreadable",
                Warning = ex.Message
            };
        }
    }

    private static Row EnsureRow(DocumentFormat.OpenXml.Spreadsheet.SheetData sheetData, uint rowIndex)
    {
        var row = sheetData.Elements<Row>()
            .FirstOrDefault(r => r.RowIndex?.Value == rowIndex);

        if (row == null)
        {
            row = new Row { RowIndex = new UInt32Value(rowIndex) };

            var insertBefore = sheetData.Elements<Row>()
                .FirstOrDefault(r => r.RowIndex?.Value > rowIndex);

            if (insertBefore != null)
                sheetData.InsertBefore(row, insertBefore);
            else
                sheetData.AppendChild(row);
        }

        return row;
    }

    private static Cell EnsureCell(Row row, string cellRef, string colName)
    {
        var cell = row.Elements<Cell>()
            .FirstOrDefault(c => c.CellReference?.Value == cellRef);

        if (cell == null)
        {
            cell = new Cell { CellReference = new StringValue(cellRef) };

            // Insert in correct column order
            int newColIdx = ColumnNameToIndex(colName);
            Cell? insertBefore = null;

            foreach (var existing in row.Elements<Cell>())
            {
                var (existCol, _) = ParseCellReference(existing.CellReference?.Value ?? "A1");
                if (ColumnNameToIndex(existCol) > newColIdx)
                {
                    insertBefore = existing;
                    break;
                }
            }

            if (insertBefore != null)
                row.InsertBefore(cell, insertBefore);
            else
                row.AppendChild(cell);
        }

        return cell;
    }

    private static (string colName, uint rowIndex) ParseCellReference(string cellRef)
    {
        var match = Regex.Match(cellRef, @"^([A-Z]+)(\d+)$");
        if (!match.Success)
            throw new Exception($"Invalid cell reference: {cellRef}");

        return (match.Groups[1].Value, uint.Parse(match.Groups[2].Value));
    }

    private static int ColumnNameToIndex(string colName)
    {
        int index = 0;
        foreach (char c in colName)
        {
            index = index * 26 + (c - 'A' + 1);
        }
        return index;
    }

    private static string IndexToColumnName(int index)
    {
        string result = "";
        while (index > 0)
        {
            index--;
            result = (char)('A' + (index % 26)) + result;
            index /= 26;
        }
        return result;
    }

    private static int GetAfterInt(JsonElement? after)
    {
        if (after == null) throw new Exception("'after' field is required");
        if (after.Value.ValueKind == JsonValueKind.Number)
            return after.Value.GetInt32();
        if (after.Value.ValueKind == JsonValueKind.String && int.TryParse(after.Value.GetString(), out int val))
            return val;
        throw new Exception($"'after' must be a number for row operations");
    }

    private static string GetAfterString(JsonElement? after)
    {
        if (after == null) throw new Exception("'after' field is required");
        if (after.Value.ValueKind == JsonValueKind.String)
            return after.Value.GetString() ?? throw new Exception("'after' is null");
        throw new Exception($"'after' must be a string (column letter) for column operations");
    }
}

/// <summary>Internal helper to avoid type collision with OpenXml SheetData.</summary>
internal class SheetData_Read
{
    public string Name { get; set; } = "";
    public string Kind { get; set; } = "worksheet";
    public string Visibility { get; set; } = "visible";
    public int RowCount { get; set; }
    public int CellCount { get; set; }
    public int FormulaCount { get; set; }
    public int SharedFormulaCount { get; set; }
    public int ArrayFormulaCount { get; set; }
    public int DataTableFormulaCount { get; set; }
    public int CommentCount { get; set; }
    public int ThreadedCommentCount { get; set; }
    public int TableCount { get; set; }
    public int DataValidationCount { get; set; }
    public int ConditionalFormatCount { get; set; }
    public int PivotTableCount { get; set; }
    public bool Protected { get; set; }
    public int MergedCellCount { get; set; }
    public List<string> MergedRanges { get; set; } = new();
    public string? FreezePaneCell { get; set; }
    public string? AutoFilterRange { get; set; }
    public List<TableInfo> Tables { get; set; } = new();
    public List<DataValidationInfo> DataValidations { get; set; } = new();
    public List<ConditionalFormatInfo> ConditionalFormats { get; set; } = new();
    public List<RowData> Rows { get; set; } = new();
}

internal class ReadSheetPart
{
    public string Kind { get; set; } = "worksheet";
    public WorksheetPart? WorksheetPart { get; set; }
    public string? Warning { get; set; }
}
