using System;
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
        bool showHelp = false;
        bool showVersion = false;

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
                case "-h":
                case "--help":
                    showHelp = true;
                    break;
                default:
                    if (!args[i].StartsWith("-"))
                    {
                        if (inputPath == null) inputPath = args[i];
                        else if (manifestPath == null) manifestPath = args[i];
                    }
                    break;
            }
        }

        if (showVersion)
        {
            Console.WriteLine($"xlsx-review {GetVersion()}");
            return 0;
        }

        if (showHelp || inputPath == null)
        {
            PrintUsage();
            return showHelp ? 0 : 1;
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

        // ── Edit Mode ──

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
        Console.Error.WriteLine(@"xlsx-review — Programmatic Excel (.xlsx) editing via JSON manifest

Usage:
  xlsx-review <input.xlsx> <edits.json> [options]
  xlsx-review <input.xlsx> --read [--json]
  cat edits.json | xlsx-review <input.xlsx> [options]

Options:
  -o, --output <path>    Output file path (default: <input>_edited.xlsx)
  --author <name>        Author name for comments (overrides manifest author)
  --json                 Output results as JSON
  --dry-run              Validate manifest without modifying
  --read                 Read spreadsheet contents (no manifest needed)
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
