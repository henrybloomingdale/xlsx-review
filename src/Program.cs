using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using XlsxReview;

class Program
{
    static int Main(string[] args)
    {
        // Parse arguments
        string? inputPath = null;
        string? manifestPath = null;
        string? outputPath = null;
        string? author = null;
        bool jsonOutput = false;
        bool dryRun = false;
        bool readMode = false;
        bool diffMode = false;
        bool textConvMode = false;
        bool gitSetup = false;
        bool showHelp = false;
        bool showVersion = false;
        var positionalArgs = new List<string>();

        for (int i = 0; i < args.Length; i++)
        {
            switch (args[i])
            {
                case "-v":
                case "--version":
                    showVersion = true;
                    break;
                case "-o":
                case "--output":
                    if (i + 1 < args.Length) outputPath = args[++i];
                    break;
                case "--author":
                    if (i + 1 < args.Length) author = args[++i];
                    break;
                case "--json":
                    jsonOutput = true;
                    break;
                case "--dry-run":
                    dryRun = true;
                    break;
                case "--read":
                    readMode = true;
                    break;
                case "--diff":
                    diffMode = true;
                    break;
                case "--textconv":
                    textConvMode = true;
                    break;
                case "--git-setup":
                    gitSetup = true;
                    break;
                case "-h":
                case "--help":
                    showHelp = true;
                    break;
                default:
                    if (!args[i].StartsWith("-"))
                        positionalArgs.Add(args[i]);
                    break;
            }
        }

        // Map positional args
        if (positionalArgs.Count >= 1) inputPath = positionalArgs[0];
        if (positionalArgs.Count >= 2) manifestPath = positionalArgs[1];

        if (showVersion)
        {
            Console.WriteLine($"xlsx-review {GetVersion()}");
            return 0;
        }

        // ── Git setup ──────────────────────────────────────────────
        if (gitSetup)
        {
            PrintGitSetup();
            return 0;
        }

        if (showHelp || (inputPath == null && !gitSetup))
        {
            PrintUsage();
            return showHelp ? 0 : 1;
        }

        // ── Diff mode ─────────────────────────────────────────────
        if (diffMode)
        {
            if (manifestPath == null)
            {
                Error("--diff requires two files: xlsx-review --diff old.xlsx new.xlsx");
                return 1;
            }

            if (!File.Exists(inputPath!))
            {
                Error($"Old file not found: {inputPath}");
                return 1;
            }
            if (!File.Exists(manifestPath))
            {
                Error($"New file not found: {manifestPath}");
                return 1;
            }

            try
            {
                var oldDoc = SpreadsheetExtractor.Extract(inputPath!);
                var newDoc = SpreadsheetExtractor.Extract(manifestPath);
                var diffResult = SpreadsheetDiffer.Diff(oldDoc, newDoc);

                if (jsonOutput)
                {
                    Console.WriteLine(JsonSerializer.Serialize(diffResult, XlsxReviewJsonContext.Default.XlsxDiffResult));
                }
                else
                {
                    SpreadsheetDiffer.PrintHumanReadable(diffResult);
                }
                return 0;
            }
            catch (Exception ex)
            {
                Error($"Diff failed: {ex.Message}");
                return 1;
            }
        }

        // ── TextConv mode ─────────────────────────────────────────
        if (textConvMode)
        {
            if (!File.Exists(inputPath!))
            {
                Error($"File not found: {inputPath}");
                return 1;
            }

            try
            {
                var extraction = SpreadsheetExtractor.Extract(inputPath!);
                Console.Write(XlsxTextConv.Convert(extraction));
                return 0;
            }
            catch (Exception ex)
            {
                Error($"TextConv failed: {ex.Message}");
                return 1;
            }
        }

        // Validate input file
        if (!File.Exists(inputPath))
        {
            Error($"Input file not found: {inputPath}");
            return 1;
        }

        // ── Read Mode ──
        if (readMode)
        {
            var editor = new SpreadsheetEditor(author ?? "Reviewer");
            ReadResult readResult;
            try
            {
                readResult = editor.ReadSpreadsheet(inputPath);
            }
            catch (Exception ex)
            {
                Error($"Failed to read spreadsheet: {ex.Message}");
                return 1;
            }

            if (jsonOutput)
            {
                Console.WriteLine(JsonSerializer.Serialize(readResult, XlsxReviewJsonContext.Default.ReadResult));
            }
            else
            {
                // Human-readable output
                foreach (var sheet in readResult.Sheets)
                {
                    Console.WriteLine($"\n📊 Sheet: {sheet.Name}");
                    Console.WriteLine(new string('─', 50));
                    foreach (var row in sheet.Rows)
                    {
                        var cells = string.Join(" | ", row.Cells.Select(c => $"{c.Cell}={c.Value ?? ""}"));
                        Console.WriteLine($"  Row {row.Row}: {cells}");
                    }
                }
                Console.WriteLine();
            }
            return 0;
        }

        // ── Edit Mode (original behavior) ─────────────────────────

        // Read manifest from file or stdin
        string manifestJson;
        if (manifestPath != null)
        {
            if (!File.Exists(manifestPath))
            {
                Error($"Manifest file not found: {manifestPath}");
                return 1;
            }
            manifestJson = File.ReadAllText(manifestPath);
        }
        else if (!Console.IsInputRedirected)
        {
            Error("No manifest file specified and no stdin input.\nUsage: xlsx-review <input.xlsx> <edits.json> -o <output.xlsx>");
            return 1;
        }
        else
        {
            manifestJson = Console.In.ReadToEnd();
        }

        // Default output path
        if (outputPath == null && !dryRun)
        {
            string dir = Path.GetDirectoryName(inputPath) ?? ".";
            string name = Path.GetFileNameWithoutExtension(inputPath);
            outputPath = Path.Combine(dir, $"{name}_edited.xlsx");
        }

        // Deserialize manifest (using source-generated context for trim/AOT safety)
        EditManifest manifest;
        try
        {
            manifest = JsonSerializer.Deserialize(manifestJson, XlsxReviewJsonContext.Default.EditManifest)
                ?? throw new Exception("Manifest deserialized to null");
        }
        catch (Exception ex)
        {
            Error($"Failed to parse manifest JSON: {ex.Message}");
            return 1;
        }

        // Resolve author (CLI flag > manifest > default)
        string effectiveAuthor = author ?? manifest.Author ?? "Reviewer";

        // Process
        var spreadsheetEditor = new SpreadsheetEditor(effectiveAuthor);
        ProcessingResult result;

        try
        {
            result = spreadsheetEditor.Process(inputPath, outputPath ?? "", manifest, dryRun);
        }
        catch (Exception ex)
        {
            Error($"Processing failed: {ex.Message}");
            return 1;
        }

        // Output
        if (jsonOutput)
        {
            Console.WriteLine(JsonSerializer.Serialize(result, XlsxReviewJsonContext.Default.ProcessingResult));
        }
        else
        {
            PrintHumanResult(result, dryRun);
        }

        return result.Success ? 0 : 1;
    }

    static void PrintUsage()
    {
        Console.Error.WriteLine(@"xlsx-review — Read, write, and diff Excel spreadsheets with full cell awareness

Usage:
  xlsx-review <input.xlsx> --read [--json]               Read spreadsheet contents
  xlsx-review <input.xlsx> <edits.json> [options]        Write cell changes/comments
  xlsx-review --diff <old.xlsx> <new.xlsx> [--json]      Semantic spreadsheet diff
  xlsx-review --textconv <file.xlsx>                     Git textconv (normalized text)
  xlsx-review --git-setup                                Print git configuration
  cat edits.json | xlsx-review <input.xlsx> [options]

Diff & Git Integration:
  --diff                 Compare two spreadsheets semantically (cells, formulas,
                         sheets, structure)
  --textconv             Output normalized tabular text for use as git diff textconv
  --git-setup            Print .gitattributes and .gitconfig setup instructions

Read/Write Options:
  --read                 Read mode: extract cell values from all sheets
  -o, --output <path>    Output file path (default: <input>_edited.xlsx)
  --author <name>        Author name for comments (overrides manifest author)
  --json                 Output results as JSON
  --dry-run              Validate manifest without modifying
  -v, --version          Show version
  -h, --help             Show this help

JSON Manifest Format:
  {
    ""author"": ""Reviewer Name"",
    ""changes"": [
      { ""type"": ""set_cell"", ""sheet"": ""Sheet1"", ""cell"": ""A1"", ""value"": ""New Value"" },
      { ""type"": ""set_formula"", ""sheet"": ""Sheet1"", ""cell"": ""C1"", ""formula"": ""=SUM(A1:B1)"" },
      { ""type"": ""insert_row"", ""sheet"": ""Sheet1"", ""after"": 5 },
      { ""type"": ""delete_row"", ""sheet"": ""Sheet1"", ""row"": 10 },
      { ""type"": ""add_sheet"", ""name"": ""Summary"" },
      { ""type"": ""rename_sheet"", ""from"": ""Sheet1"", ""to"": ""Data"" },
      { ""type"": ""delete_sheet"", ""name"": ""Old Sheet"" }
    ],
    ""comments"": [
      { ""sheet"": ""Sheet1"", ""cell"": ""A1"", ""text"": ""Review note"" }
    ]
  }");
    }

    static void PrintGitSetup()
    {
        Console.WriteLine(@"Git Integration for Excel Spreadsheets
══════════════════════════════════════

Add to your repository's .gitattributes:

  *.xlsx diff=xlsx

Add to your .gitconfig (global or per-repo):

  [diff ""xlsx""]
      textconv = xlsx-review --textconv

Now `git diff` will show meaningful content changes for .xlsx files,
including cell values, formulas, sheet structure, and metadata.

For two-file comparison outside git:

  xlsx-review --diff old.xlsx new.xlsx
  xlsx-review --diff old.xlsx new.xlsx --json
");
    }

    static void PrintHumanResult(ProcessingResult result, bool dryRun)
    {
        string mode = dryRun ? "[DRY RUN] " : "";
        Console.WriteLine($"\n{mode}xlsx-review results");
        Console.WriteLine(new string('─', 50));
        Console.WriteLine($"  Input:    {result.Input}");
        if (!dryRun && result.Output != null)
            Console.WriteLine($"  Output:   {result.Output}");
        Console.WriteLine($"  Author:   {result.Author}");
        Console.WriteLine($"  Changes:  {result.ChangesSucceeded}/{result.ChangesAttempted}");
        Console.WriteLine($"  Comments: {result.CommentsSucceeded}/{result.CommentsAttempted}");
        Console.WriteLine();

        foreach (var r in result.Results)
        {
            string icon = r.Success ? "✓" : "✗";
            Console.WriteLine($"  {icon} [{r.Type}] {r.Message}");
        }

        Console.WriteLine();
        if (result.Success)
            Console.WriteLine(dryRun ? "✅ All edits would succeed" : "✅ All edits applied successfully");
        else
            Console.WriteLine("⚠️  Some edits failed (see above)");
    }

    static string GetVersion()
    {
        var asm = System.Reflection.Assembly.GetExecutingAssembly();
        var ver = asm.GetName().Version;
        return ver != null ? $"{ver.Major}.{ver.Minor}.{ver.Build}" : "1.0.0";
    }

    static void Error(string msg) => Console.Error.WriteLine($"Error: {msg}");
}
