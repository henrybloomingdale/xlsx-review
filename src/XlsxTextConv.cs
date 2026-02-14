using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace XlsxReview;

/// <summary>
/// Produces a normalized text representation of a .xlsx file suitable for
/// use as a git textconv driver. This allows `git diff` to show meaningful
/// changes for Excel spreadsheets.
///
/// Output format:
/// - Metadata (sheet count)
/// - Per-sheet tabular output with cell references
/// - Formulas shown as =FORMULA (not computed values)
/// - Produces clean, diffable plain text
/// </summary>
public static class XlsxTextConv
{
    public static string Convert(SpreadsheetExtraction doc)
    {
        var sb = new StringBuilder();

        // ── Metadata ───────────────────────────────────────
        sb.AppendLine("=== METADATA ===");
        sb.AppendLine($"Sheets: {doc.Sheets.Count}");
        sb.AppendLine();

        // ── Per-sheet output ───────────────────────────────
        foreach (var sheet in doc.Sheets)
        {
            sb.AppendLine($"=== Sheet: {sheet.Name} ===");

            if (sheet.Cells.Count == 0)
            {
                sb.AppendLine("  (empty)");
                sb.AppendLine();
                continue;
            }

            // Determine dimensions
            int maxRow = sheet.MaxRow;
            int maxCol = sheet.MaxColumn;

            if (maxRow == 0 || maxCol == 0)
            {
                sb.AppendLine("  (empty)");
                sb.AppendLine();
                continue;
            }

            // Build column widths (minimum 10 chars, or content width + 2)
            var colWidths = new int[maxCol];
            for (int c = 1; c <= maxCol; c++)
            {
                string colName = IndexToColumnName(c);
                int maxWidth = colName.Length;

                for (int r = 1; r <= maxRow; r++)
                {
                    string cellRef = $"{colName}{r}";
                    if (sheet.Cells.TryGetValue(cellRef, out var cell))
                    {
                        string display = GetDisplayValue(cell);
                        if (display.Length > maxWidth)
                            maxWidth = display.Length;
                    }
                }

                colWidths[c - 1] = Math.Min(Math.Max(maxWidth, 3), 30);  // clamp 3..30
            }

            // Header row (column names)
            var headerParts = new List<string>();
            headerParts.Add("".PadRight(5));  // row number gutter
            for (int c = 1; c <= maxCol; c++)
            {
                string colName = IndexToColumnName(c);
                headerParts.Add(colName.PadRight(colWidths[c - 1]));
            }
            sb.Append("     | ");
            sb.AppendLine(string.Join(" | ", headerParts.Skip(1)) + " |");

            // Data rows
            for (int r = 1; r <= maxRow; r++)
            {
                // Check if this row has any data
                bool hasData = false;
                for (int c = 1; c <= maxCol; c++)
                {
                    string cellRef = $"{IndexToColumnName(c)}{r}";
                    if (sheet.Cells.ContainsKey(cellRef))
                    {
                        hasData = true;
                        break;
                    }
                }

                if (!hasData) continue;

                sb.Append($"{r,4} | ");
                var cellParts = new List<string>();
                for (int c = 1; c <= maxCol; c++)
                {
                    string cellRef = $"{IndexToColumnName(c)}{r}";
                    string display = "";
                    if (sheet.Cells.TryGetValue(cellRef, out var cell))
                        display = GetDisplayValue(cell);

                    // Truncate if too long
                    if (display.Length > colWidths[c - 1])
                        display = display[..(colWidths[c - 1] - 1)] + "…";

                    cellParts.Add(display.PadRight(colWidths[c - 1]));
                }
                sb.AppendLine(string.Join(" | ", cellParts) + " |");
            }

            sb.AppendLine();
        }

        return sb.ToString();
    }

    /// <summary>
    /// Get the display value for a cell, preferring formulas over computed values.
    /// </summary>
    private static string GetDisplayValue(ExtractedCell cell)
    {
        // Show formula if present (prefixed with =)
        if (!string.IsNullOrEmpty(cell.Formula))
            return $"={cell.Formula}";

        return cell.Value ?? "";
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
}
