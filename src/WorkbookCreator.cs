using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XlsxReview;

public class WorkbookCreator
{
    private const string DefaultTemplateLabel = "blank workbook (generated)";

    public CreateResult Create(
        string outputPath,
        EditManifest? manifest,
        string author,
        string? templatePath,
        bool dryRun)
    {
        string templateLabel = templatePath ?? DefaultTemplateLabel;

        if (templatePath != null)
        {
            if (!File.Exists(templatePath))
                throw new FileNotFoundException($"Template not found: {templatePath}");

            SpreadsheetPackagePreflight.Validate(templatePath);
        }

        var result = new CreateResult
        {
            Template = templateLabel,
            Output = dryRun ? null : outputPath,
            Populated = manifest != null,
            Success = true
        };

        if (dryRun && manifest == null)
            return result;

        string tempWorkbook = Path.Combine(Path.GetTempPath(), $"xlsx-review-template-{Guid.NewGuid()}.xlsx");
        CopyTemplate(templatePath, tempWorkbook);

        try
        {
            if (manifest != null)
            {
                if (!dryRun)
                {
                    var dir = Path.GetDirectoryName(outputPath);
                    if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                        Directory.CreateDirectory(dir);
                }

                var editor = new SpreadsheetEditor(author);
                var editResult = editor.Process(tempWorkbook, dryRun ? "" : outputPath, manifest, dryRun);

                result.ChangesAttempted = editResult.ChangesAttempted;
                result.ChangesSucceeded = editResult.ChangesSucceeded;
                result.CommentsAttempted = editResult.CommentsAttempted;
                result.CommentsSucceeded = editResult.CommentsSucceeded;
                result.Results = editResult.Results;
                result.Success = editResult.Success;
            }
            else if (!dryRun)
            {
                var dir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                File.Copy(tempWorkbook, outputPath, true);
            }
        }
        finally
        {
            if (File.Exists(tempWorkbook))
                File.Delete(tempWorkbook);
        }

        return result;
    }

    private static void CopyTemplate(string? templatePath, string destinationPath)
    {
        var dir = Path.GetDirectoryName(destinationPath);
        if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            Directory.CreateDirectory(dir);

        if (templatePath != null)
        {
            File.Copy(templatePath, destinationPath, true);
            return;
        }

        CreateBlankWorkbook(destinationPath);
    }

    private static void CreateBlankWorkbook(string outputPath)
    {
        using var document = SpreadsheetDocument.Create(outputPath, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new DocumentFormat.OpenXml.Spreadsheet.SheetData());

        var sheets = workbookPart.Workbook.AppendChild(new Sheets());
        sheets.AppendChild(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = "Sheet1"
        });

        worksheetPart.Worksheet.Save();
        workbookPart.Workbook.Save();
    }
}
