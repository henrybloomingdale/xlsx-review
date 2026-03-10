using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;

namespace XlsxReview;

internal static class SpreadsheetPackagePreflight
{
    private const int MaxEntryCount = 20000;
    private const long MaxTotalUncompressedBytes = 512L * 1024 * 1024;
    private const long MaxEntryUncompressedBytes = 256L * 1024 * 1024;
    private const long MaxXmlEntryBytes = 128L * 1024 * 1024;
    private const int MaxSharedStringTextLength = 262144;
    private const long MaxSharedStringsTotalTextLength = 8L * 1024 * 1024;
    private const long MinCompressionRatioCheckBytes = 1024L * 1024;
    private const double MaxCompressionRatio = 150.0;
    private const string RootRelationshipsPath = "_rels/.rels";
    private const string ContentTypesPath = "[Content_Types].xml";

    public static void Validate(string path)
    {
        try
        {
            using var stream = File.OpenRead(path);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
            ValidateArchive(archive);
        }
        catch (PreflightException)
        {
            throw;
        }
        catch (InvalidDataException ex)
        {
            throw new PreflightException(ex.Message, ex);
        }
        catch (IOException ex)
        {
            throw new PreflightException(ex.Message, ex);
        }
        catch (XmlException ex)
        {
            throw new PreflightException(ex.Message, ex);
        }
    }

    private static void ValidateArchive(ZipArchive archive)
    {
        if (archive.Entries.Count == 0)
            throw new PreflightException("Spreadsheet package is empty.");

        if (archive.Entries.Count > MaxEntryCount)
            throw new PreflightException($"Spreadsheet package contains too many parts ({archive.Entries.Count}).");

        var entries = new Dictionary<string, ZipArchiveEntry>(StringComparer.OrdinalIgnoreCase);
        long totalUncompressedBytes = 0;

        foreach (var entry in archive.Entries)
        {
            string partPath = NormalizePackagePath(entry.FullName);
            if (string.IsNullOrEmpty(partPath))
                continue;

            if (!entries.TryAdd(partPath, entry))
                throw new PreflightException($"Spreadsheet package contains duplicate part '{partPath}'.");

            long entryLength;
            long compressedLength;
            try
            {
                entryLength = entry.Length;
                compressedLength = entry.CompressedLength;
            }
            catch (InvalidDataException ex)
            {
                throw new PreflightException($"Spreadsheet part '{partPath}' is corrupted: {ex.Message}", ex);
            }

            totalUncompressedBytes = checked(totalUncompressedBytes + entryLength);
            if (totalUncompressedBytes > MaxTotalUncompressedBytes)
                throw new PreflightException("Spreadsheet package exceeds the safety limit for uncompressed size.");

            if (entryLength > MaxEntryUncompressedBytes)
                throw new PreflightException($"Spreadsheet part '{partPath}' exceeds the safety limit for uncompressed size.");

            if (compressedLength > 0 &&
                entryLength >= MinCompressionRatioCheckBytes &&
                (double)entryLength / compressedLength > MaxCompressionRatio)
            {
                throw new PreflightException($"Spreadsheet part '{partPath}' exceeds the safety compression ratio.");
            }

            if (!IsXmlPart(partPath))
                continue;

            if (entryLength > MaxXmlEntryBytes)
                throw new PreflightException($"XML part '{partPath}' exceeds the safety limit for XML size.");

            ProbeXmlRoot(entry, partPath);
        }

        if (!entries.TryGetValue(ContentTypesPath, out var contentTypesEntry))
            throw new PreflightException("Spreadsheet package is missing [Content_Types].xml.");

        if (!entries.TryGetValue(RootRelationshipsPath, out var rootRelationshipsEntry))
            throw new PreflightException("Spreadsheet package is missing _rels/.rels.");

        ValidateContentTypes(contentTypesEntry);
        string workbookPartPath = ResolveWorkbookPartPath(rootRelationshipsEntry);
        if (!entries.ContainsKey(workbookPartPath))
            throw new PreflightException($"Spreadsheet package is missing workbook part '{workbookPartPath}'.");

        ValidateWorkbookPart(entries[workbookPartPath], workbookPartPath);
        ValidateSharedStrings(entries);
        ValidateRelationshipTargets(entries);
    }

    private static void ValidateContentTypes(ZipArchiveEntry entry)
    {
        if (!string.Equals(GetRootElementName(entry, ContentTypesPath), "Types", StringComparison.Ordinal))
            throw new PreflightException("[Content_Types].xml does not have a Types root element.");
    }

    private static string ResolveWorkbookPartPath(ZipArchiveEntry entry)
    {
        bool sawRelationshipsRoot = false;
        string? target = null;

        try
        {
            using var reader = OpenXmlReader(entry, RootRelationshipsPath);
            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.Element)
                    continue;

                if (!sawRelationshipsRoot)
                {
                    if (!string.Equals(reader.LocalName, "Relationships", StringComparison.Ordinal))
                        throw new PreflightException("_rels/.rels does not have a Relationships root element.");

                    sawRelationshipsRoot = true;
                    continue;
                }

                if (!string.Equals(reader.LocalName, "Relationship", StringComparison.Ordinal))
                    continue;

                string? relationshipType = reader.GetAttribute("Type");
                if (relationshipType?.EndsWith("/officeDocument", StringComparison.OrdinalIgnoreCase) != true)
                    continue;

                target = reader.GetAttribute("Target");
                break;
            }
        }
        catch (XmlException ex)
        {
            throw new PreflightException($"XML part '{RootRelationshipsPath}' is invalid: {ex.Message}", ex);
        }
        catch (InvalidDataException ex)
        {
            throw new PreflightException($"XML part '{RootRelationshipsPath}' is unreadable: {ex.Message}", ex);
        }

        if (!sawRelationshipsRoot)
            throw new PreflightException("_rels/.rels does not have a Relationships root element.");

        if (string.IsNullOrWhiteSpace(target))
            throw new PreflightException("Spreadsheet package does not define a workbook target.");

        return ResolveRelationshipTarget(sourcePartPath: "", target);
    }

    private static void ValidateWorkbookPart(ZipArchiveEntry entry, string partPath)
    {
        if (!string.Equals(GetRootElementName(entry, partPath), "workbook", StringComparison.Ordinal))
            throw new PreflightException($"Workbook part '{partPath}' does not have a workbook root element.");
    }

    private static void ValidateRelationshipTargets(Dictionary<string, ZipArchiveEntry> entries)
    {
        foreach (var pair in entries.Where(pair => pair.Key.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)))
        {
            string relationshipsPath = pair.Key;
            string sourcePartPath = GetSourcePartPathFromRelationshipsPart(relationshipsPath);
            bool sawRelationshipsRoot = false;

            try
            {
                using var reader = OpenXmlReader(pair.Value, relationshipsPath);
                while (reader.Read())
                {
                    if (reader.NodeType != XmlNodeType.Element)
                        continue;

                    if (!sawRelationshipsRoot)
                    {
                        if (!string.Equals(reader.LocalName, "Relationships", StringComparison.Ordinal))
                        {
                            throw new PreflightException(
                                $"Relationship part '{relationshipsPath}' does not have a Relationships root element.");
                        }

                        sawRelationshipsRoot = true;
                        continue;
                    }

                    if (!string.Equals(reader.LocalName, "Relationship", StringComparison.Ordinal))
                        continue;

                    string? targetMode = reader.GetAttribute("TargetMode");
                    if (string.Equals(targetMode, "External", StringComparison.OrdinalIgnoreCase))
                        continue;

                    string? target = reader.GetAttribute("Target");
                    if (string.IsNullOrWhiteSpace(target))
                    {
                        throw new PreflightException(
                            $"Relationship part '{relationshipsPath}' contains a relationship without a target.");
                    }

                    _ = ResolveRelationshipTarget(sourcePartPath, target);
                }
            }
            catch (XmlException ex)
            {
                throw new PreflightException($"XML part '{relationshipsPath}' is invalid: {ex.Message}", ex);
            }
            catch (InvalidDataException ex)
            {
                throw new PreflightException($"XML part '{relationshipsPath}' is unreadable: {ex.Message}", ex);
            }

            if (!sawRelationshipsRoot)
            {
                throw new PreflightException(
                    $"Relationship part '{relationshipsPath}' does not have a Relationships root element.");
            }
        }
    }

    private static void ProbeXmlRoot(ZipArchiveEntry entry, string partPath)
    {
        try
        {
            using var stream = entry.Open();
            using var reader = XmlReader.Create(stream, CreateXmlReaderSettings());
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                    return;
            }
        }
        catch (XmlException ex)
        {
            throw new PreflightException($"XML part '{partPath}' is invalid: {ex.Message}", ex);
        }
        catch (InvalidDataException ex)
        {
            throw new PreflightException($"XML part '{partPath}' is unreadable: {ex.Message}", ex);
        }
    }

    private static string GetRootElementName(ZipArchiveEntry entry, string partPath)
    {
        try
        {
            using var reader = OpenXmlReader(entry, partPath);
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                    return reader.LocalName;
            }
        }
        catch (XmlException ex)
        {
            throw new PreflightException($"XML part '{partPath}' is invalid: {ex.Message}", ex);
        }
        catch (InvalidDataException ex)
        {
            throw new PreflightException($"XML part '{partPath}' is unreadable: {ex.Message}", ex);
        }

        throw new PreflightException($"XML part '{partPath}' does not contain a root element.");
    }

    private static void ValidateSharedStrings(Dictionary<string, ZipArchiveEntry> entries)
    {
        if (!entries.TryGetValue("xl/sharedStrings.xml", out var entry))
            return;

        try
        {
            long totalTextLength = 0;
            using var stream = entry.Open();
            using var reader = XmlReader.Create(stream, CreateXmlReaderSettings());

            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.Element || !string.Equals(reader.LocalName, "t", StringComparison.Ordinal))
                    continue;

                string text = reader.ReadElementContentAsString();
                if (text.Length > MaxSharedStringTextLength)
                {
                    throw new PreflightException(
                        $"Shared string table contains a text item longer than the safety limit ({text.Length} characters).");
                }

                totalTextLength += text.Length;
                if (totalTextLength > MaxSharedStringsTotalTextLength)
                    throw new PreflightException("Shared string table exceeds the safety limit for text content.");
            }
        }
        catch (PreflightException)
        {
            throw;
        }
        catch (XmlException ex)
        {
            throw new PreflightException($"Shared string table is invalid: {ex.Message}", ex);
        }
        catch (InvalidDataException ex)
        {
            throw new PreflightException($"Shared string table is unreadable: {ex.Message}", ex);
        }
    }

    private static XmlReaderSettings CreateXmlReaderSettings() => new()
    {
        DtdProcessing = DtdProcessing.Prohibit,
        XmlResolver = null,
        IgnoreComments = true,
        IgnoreWhitespace = true,
        MaxCharactersFromEntities = 0,
        MaxCharactersInDocument = 4L * 1024 * 1024
    };

    private static XmlReader OpenXmlReader(ZipArchiveEntry entry, string partPath)
    {
        try
        {
            return XmlReader.Create(entry.Open(), CreateXmlReaderSettings());
        }
        catch (XmlException ex)
        {
            throw new PreflightException($"XML part '{partPath}' is invalid: {ex.Message}", ex);
        }
        catch (InvalidDataException ex)
        {
            throw new PreflightException($"XML part '{partPath}' is unreadable: {ex.Message}", ex);
        }
    }

    private static bool IsXmlPart(string partPath) =>
        partPath.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) ||
        partPath.EndsWith(".rels", StringComparison.OrdinalIgnoreCase);

    private static string GetSourcePartPathFromRelationshipsPart(string relationshipsPath)
    {
        if (string.Equals(relationshipsPath, RootRelationshipsPath, StringComparison.OrdinalIgnoreCase))
            return string.Empty;

        int markerIndex = relationshipsPath.LastIndexOf("/_rels/", StringComparison.OrdinalIgnoreCase);
        if (markerIndex < 0 || !relationshipsPath.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
            throw new PreflightException($"Relationship part '{relationshipsPath}' has an invalid path.");

        string directory = relationshipsPath[..markerIndex];
        string fileName = relationshipsPath[(markerIndex + "/_rels/".Length)..];
        fileName = fileName[..^".rels".Length];
        return string.IsNullOrEmpty(directory) ? fileName : $"{directory}/{fileName}";
    }

    private static string ResolveRelationshipTarget(string sourcePartPath, string target)
    {
        string cleanTarget = target.Split('#', 2)[0].Split('?', 2)[0];
        if (cleanTarget.StartsWith("/", StringComparison.Ordinal))
            return NormalizePackagePath(cleanTarget);

        string baseDirectory = string.Empty;
        int lastSlash = sourcePartPath.LastIndexOf('/');
        if (lastSlash >= 0)
            baseDirectory = sourcePartPath[..(lastSlash + 1)];

        return NormalizePackagePath($"{baseDirectory}{cleanTarget}");
    }

    private static string NormalizePackagePath(string path)
    {
        var segments = new List<string>();
        foreach (var segment in path.Replace('\\', '/').Split('/', StringSplitOptions.RemoveEmptyEntries))
        {
            if (segment == ".")
                continue;

            if (segment == "..")
            {
                if (segments.Count == 0)
                    throw new PreflightException($"Package path escapes the archive root: {path}");

                segments.RemoveAt(segments.Count - 1);
                continue;
            }

            segments.Add(segment);
        }

        return string.Join('/', segments);
    }
}

internal sealed class PreflightException : Exception
{
    public PreflightException(string message)
        : base(message)
    {
    }

    public PreflightException(string message, Exception innerException)
        : base(message, innerException)
    {
    }
}
