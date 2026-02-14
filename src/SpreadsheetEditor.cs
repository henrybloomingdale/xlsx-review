using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
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
        var result = new ReadResult();

        using var doc = SpreadsheetDocument.Open(inputPath, false);
        var workbookPart = doc.WorkbookPart
            ?? throw new Exception("Invalid spreadsheet: no workbook part");

        var sheets = workbookPart.Workbook.Sheets?.Elements<Sheet>() ?? Enumerable.Empty<Sheet>();
        var sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable;

        foreach (var sheet in sheets)
        {
            var sheetData = new SheetData_Read
            {
                Name = sheet.Name?.Value ?? "Unnamed"
            };

            var worksheetPart = (WorksheetPart?)workbookPart.GetPartById(sheet.Id?.Value ?? "");
            if (worksheetPart == null) continue;

            var rows = worksheetPart.Worksheet.Descendants<Row>();
            foreach (var row in rows)
            {
                var rowData = new RowData { Row = (int)(row.RowIndex?.Value ?? 0) };

                foreach (var cell in row.Elements<Cell>())
                {
                    string? cellRef = cell.CellReference?.Value;
                    if (cellRef == null) continue;

                    string? cellValue = GetCellValue(cell, sharedStrings);

                    rowData.Cells.Add(new CellData
                    {
                        Cell = cellRef,
                        Value = cellValue
                    });
                }

                if (rowData.Cells.Count > 0)
                    sheetData.Rows.Add(rowData);
            }

            result.Sheets.Add(new SheetData
            {
                Name = sheetData.Name,
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
                var ssItem = sharedStrings.ElementAt(idx);
                return ssItem.InnerText;
            }
        }
        else if (cell.DataType?.Value == CellValues.Boolean)
        {
            return value == "1" ? "TRUE" : "FALSE";
        }

        return value;
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

    private WorksheetPart GetWorksheetPart(WorkbookPart workbookPart, string sheetName)
    {
        var sheets = workbookPart.Workbook.GetFirstChild<Sheets>()
            ?? throw new Exception("No sheets found in workbook");

        var sheet = sheets.Elements<Sheet>()
            .FirstOrDefault(s => s.Name?.Value == sheetName)
            ?? throw new Exception($"Sheet '{sheetName}' not found");

        return (WorksheetPart)workbookPart.GetPartById(sheet.Id?.Value ?? "");
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
    public List<RowData> Rows { get; set; } = new();
}
