using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XlsxReview;

/// <summary>
/// Extracted representation of a spreadsheet for diff and textconv.
/// </summary>
public class SpreadsheetExtraction
{
    public string FileName { get; set; } = "";
    public List<ExtractedSheet> Sheets { get; set; } = new();
}

public class ExtractedSheet
{
    public string Name { get; set; } = "";
    public int MaxRow { get; set; }
    public int MaxColumn { get; set; }
    public Dictionary<string, ExtractedCell> Cells { get; set; } = new();
}

public class ExtractedCell
{
    public string Reference { get; set; } = "";
    public string? Value { get; set; }
    public string? Formula { get; set; }
    public string? CellType { get; set; }  // "string", "number", "boolean", "date", "formula", "empty"
    public string? NumberFormat { get; set; }
    public bool Bold { get; set; }
    public string? FontName { get; set; }
    public string? FontColor { get; set; }
}

/// <summary>
/// Extracts all data from a .xlsx file for diff and textconv operations.
/// </summary>
public static class SpreadsheetExtractor
{
    public static SpreadsheetExtraction Extract(string path)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException($"File not found: {path}");

        var result = new SpreadsheetExtraction
        {
            FileName = Path.GetFileName(path)
        };

        using var doc = SpreadsheetDocument.Open(path, false);
        var workbookPart = doc.WorkbookPart
            ?? throw new Exception("Invalid spreadsheet: no workbook part");

        var sheets = workbookPart.Workbook.Sheets?.Elements<Sheet>() ?? Enumerable.Empty<Sheet>();
        var sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable;

        // Build number format map from stylesheet
        var numberFormats = BuildNumberFormatMap(workbookPart);

        foreach (var sheet in sheets)
        {
            var extracted = new ExtractedSheet
            {
                Name = sheet.Name?.Value ?? "Unnamed"
            };

            var worksheetPart = workbookPart.GetPartById(sheet.Id?.Value ?? "") as WorksheetPart;
            if (worksheetPart == null) continue;

            var rows = worksheetPart.Worksheet.Descendants<Row>();
            int maxRow = 0;
            int maxCol = 0;

            foreach (var row in rows)
            {
                int rowIdx = (int)(row.RowIndex?.Value ?? 0);
                if (rowIdx > maxRow) maxRow = rowIdx;

                foreach (var cell in row.Elements<Cell>())
                {
                    string? cellRef = cell.CellReference?.Value;
                    if (cellRef == null) continue;

                    var (colName, _) = ParseCellReference(cellRef);
                    int colIdx = ColumnNameToIndex(colName);
                    if (colIdx > maxCol) maxCol = colIdx;

                    var extractedCell = ExtractCell(cell, sharedStrings, numberFormats, workbookPart);
                    extractedCell.Reference = cellRef;
                    extracted.Cells[cellRef] = extractedCell;
                }
            }

            extracted.MaxRow = maxRow;
            extracted.MaxColumn = maxCol;

            result.Sheets.Add(extracted);
        }

        return result;
    }

    private static ExtractedCell ExtractCell(Cell cell, SharedStringTable? sharedStrings,
        Dictionary<uint, string> numberFormats, WorkbookPart workbookPart)
    {
        var result = new ExtractedCell();

        // Extract formula
        if (cell.CellFormula != null)
        {
            result.Formula = cell.CellFormula.Text;
            result.CellType = "formula";
        }

        // Extract value
        if (cell.DataType?.Value == CellValues.InlineString)
        {
            result.Value = cell.InlineString?.Text?.Text ?? cell.InlineString?.InnerText;
            result.CellType ??= "string";
        }
        else if (cell.CellValue != null)
        {
            string rawValue = cell.CellValue.Text;

            if (cell.DataType?.Value == CellValues.SharedString && sharedStrings != null)
            {
                if (int.TryParse(rawValue, out int idx))
                {
                    var ssItem = sharedStrings.ElementAt(idx);
                    result.Value = ssItem.InnerText;
                    result.CellType ??= "string";
                }
            }
            else if (cell.DataType?.Value == CellValues.Boolean)
            {
                result.Value = rawValue == "1" ? "TRUE" : "FALSE";
                result.CellType ??= "boolean";
            }
            else
            {
                result.Value = rawValue;
                result.CellType ??= "number";
            }
        }
        else
        {
            result.CellType ??= "empty";
        }

        // Extract number format
        if (cell.StyleIndex?.HasValue == true)
        {
            uint styleIdx = cell.StyleIndex.Value;
            var stylesPart = workbookPart.WorkbookStylesPart;
            if (stylesPart?.Stylesheet?.CellFormats != null)
            {
                var cellFormats = stylesPart.Stylesheet.CellFormats.Elements<CellFormat>().ToList();
                if ((int)styleIdx < cellFormats.Count)
                {
                    var cf = cellFormats[(int)styleIdx];
                    uint numFmtId = cf.NumberFormatId?.Value ?? 0;
                    if (numberFormats.TryGetValue(numFmtId, out string? fmt))
                        result.NumberFormat = fmt;
                    else if (numFmtId > 0)
                        result.NumberFormat = GetBuiltInNumberFormat(numFmtId);

                    // Font info
                    if (cf.FontId?.HasValue == true)
                    {
                        var fonts = stylesPart.Stylesheet.Fonts?.Elements<Font>().ToList();
                        if (fonts != null && (int)cf.FontId.Value < fonts.Count)
                        {
                            var font = fonts[(int)cf.FontId.Value];
                            result.Bold = font.Bold != null;
                            result.FontName = font.FontName?.Val?.Value;
                            result.FontColor = font.Color?.Rgb?.Value;
                        }
                    }
                }
            }
        }

        return result;
    }

    private static Dictionary<uint, string> BuildNumberFormatMap(WorkbookPart workbookPart)
    {
        var map = new Dictionary<uint, string>();
        var stylesPart = workbookPart.WorkbookStylesPart;
        if (stylesPart?.Stylesheet?.NumberingFormats == null)
            return map;

        foreach (var nf in stylesPart.Stylesheet.NumberingFormats.Elements<NumberingFormat>())
        {
            if (nf.NumberFormatId?.HasValue == true && nf.FormatCode?.Value != null)
                map[nf.NumberFormatId.Value] = nf.FormatCode.Value;
        }

        return map;
    }

    private static string? GetBuiltInNumberFormat(uint id) => id switch
    {
        1 => "0",
        2 => "0.00",
        3 => "#,##0",
        4 => "#,##0.00",
        9 => "0%",
        10 => "0.00%",
        11 => "0.00E+00",
        14 => "mm-dd-yy",
        15 => "d-mmm-yy",
        16 => "d-mmm",
        17 => "mmm-yy",
        18 => "h:mm AM/PM",
        19 => "h:mm:ss AM/PM",
        20 => "h:mm",
        21 => "h:mm:ss",
        22 => "m/d/yy h:mm",
        44 => "$#,##0.00",
        _ => null
    };

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
            index = index * 26 + (c - 'A' + 1);
        return index;
    }
}
