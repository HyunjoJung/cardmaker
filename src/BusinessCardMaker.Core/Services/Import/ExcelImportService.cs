// Copyright (c) 2025 Business Card Maker Contributors
// Licensed under the Apache License, Version 2.0

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Buffers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using BusinessCardMaker.Core.Models;
using BusinessCardMaker.Core.Configuration;
using BusinessCardMaker.Core.Exceptions;

namespace BusinessCardMaker.Core.Services.Import;

/// <summary>
/// Excel import service (memory-only, no database persistence)
/// </summary>
public class ExcelImportService : IExcelImportService
{
    private readonly ILogger<ExcelImportService>? _logger;
    private readonly BusinessCardProcessingOptions _options;

    // Excel file signatures (magic bytes)
    private static readonly byte[] XlsxSignature = { 0x50, 0x4B, 0x03, 0x04 }; // PK.. (ZIP format)
    private static readonly byte[] XlsSignature = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 }; // OLE2 format

    public ExcelImportService(
        ILogger<ExcelImportService>? logger = null,
        IOptions<BusinessCardProcessingOptions>? options = null)
    {
        _logger = logger;
        _options = options?.Value ?? new BusinessCardProcessingOptions();
    }

    public async Task<ImportResult> ImportFromExcelAsync(Stream excelStream, IProgress<int>? progress = null)
    {
        // Input validation
        if (excelStream == null)
        {
            throw new ArgumentNullException(nameof(excelStream), "Excel stream cannot be null");
        }

        try
        {
            _logger?.LogInformation("Starting Excel import (memory-only)");

            // Copy to a buffer to avoid sync IO errors when the incoming stream disallows sync reads (e.g., ASP.NET Core uploads)
            if (excelStream.CanSeek)
            {
                excelStream.Position = 0;
            }

            using var buffer = new MemoryStream();
            var rental = ArrayPool<byte>.Shared.Rent(81920);
            try
            {
                int read;
                while ((read = await excelStream.ReadAsync(rental.AsMemory(0, rental.Length))) > 0)
                {
                    await buffer.WriteAsync(rental.AsMemory(0, read));
                }
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(rental);
            }
            buffer.Position = 0;

            // Validate Excel file after buffering
            ValidateExcelFile(buffer);
            buffer.Position = 0;

            using var document = SpreadsheetDocument.Open(buffer, false);
            var workbookPart = document.WorkbookPart!;
            var worksheetPart = workbookPart.WorksheetParts.First();
            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;

            // Find header row and build column mapping
            var (headerRowIndex, columnMapping) = FindHeaderAndBuildMapping(sheetData, workbookPart);

            if (columnMapping.Count == 0)
            {
                return ImportResult.CreateFailure(ErrorCodes.FormatError(ErrorCodes.InvalidHeaderFormat, "Could not find Name and Email columns."));
            }

            // Parse employee data
            var rows = sheetData.Elements<Row>()
                .Where(r => r.RowIndex != null && r.RowIndex.Value > headerRowIndex)
                .ToList();

            var employees = new List<Employee>();
            var warnings = new List<string>();
            int processed = 0;

            foreach (var row in rows)
            {
                try
                {
                    var employee = ParseEmployeeFromRow(row, columnMapping, workbookPart);
                    if (employee != null)
                    {
                        employees.Add(employee);
                    }
                }
                catch (Exception ex)
                {
                    _logger?.LogWarning(ex, "Failed to parse row {RowIndex}", row.RowIndex?.Value);
                    warnings.Add($"Row {row.RowIndex?.Value}: Parse failed - {ex.Message}");
                }

                processed++;
                if (rows.Count > 0)
                {
                    progress?.Report((processed * 100) / rows.Count);
                }
            }

            _logger?.LogInformation("Excel import completed. Total employees: {Total}", employees.Count);

            return await Task.FromResult(ImportResult.CreateSuccess(employees, warnings));
        }
        catch (OutOfMemoryException)
        {
            _logger?.LogError("Out of memory while importing Excel file");
            return ImportResult.CreateFailure(ErrorCodes.FormatError(ErrorCodes.OutOfMemory));
        }
        catch (IOException ex)
        {
            _logger?.LogError(ex, "IO error during Excel import");
            return ImportResult.CreateFailure(ErrorCodes.FormatError(ErrorCodes.IoError, ex.Message));
        }
        catch (UnauthorizedAccessException ex)
        {
            _logger?.LogError(ex, "Permission denied during Excel import");
            return ImportResult.CreateFailure(ErrorCodes.FormatError(ErrorCodes.PermissionDenied, ex.Message));
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("E0"))
        {
            _logger?.LogError(ex, "Excel validation error");
            return ImportResult.CreateFailure(ex.Message); // Already formatted with error code
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Unexpected error during Excel import");
            return ImportResult.CreateFailure(ErrorCodes.FormatError(ErrorCodes.ExcelProcessingFailed, ex.Message));
        }
    }

    public byte[] CreateTemplate()
    {
        using var stream = new MemoryStream();
        using (var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            var sheets = workbookPart.Workbook.AppendChild(new Sheets());
            var sheet = new Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "Employees"
            };
            sheets.Append(sheet);

            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;

            // Add header row
            var headerRow = new Row { RowIndex = 1 };
            var headers = new[]
            {
                "Name",
                "Name (English)",
                "Position",
                "Position (English)",
                "Email",
                "Mobile",
                "Extension",
                "Company (optional)",
                "Department (optional)"
            };

            uint colIndex = 1;
            foreach (var header in headers)
            {
                headerRow.Append(new Cell
                {
                    CellReference = GetCellReference(1, colIndex),
                    DataType = CellValues.String,
                    CellValue = new CellValue(header)
                });
                colIndex++;
            }
            sheetData.Append(headerRow);

            // Add sample row
            var sampleRow = new Row { RowIndex = 2 };
            var samples = new[]
            {
                "John Doe",
                "John Doe",
                "Team Leader",
                "Team Leader",
                "john@company.com",
                "010-1234-5678",
                "1234",
                "CardMaker Inc.",
                "Development"
            };

            colIndex = 1;
            foreach (var sample in samples)
            {
                sampleRow.Append(new Cell
                {
                    CellReference = GetCellReference(2, colIndex),
                    DataType = CellValues.String,
                    CellValue = new CellValue(sample)
                });
                colIndex++;
            }
            sheetData.Append(sampleRow);

            worksheetPart.Worksheet.Save();
            workbookPart.Workbook.Save();
        }

        stream.Position = 0;
        return stream.ToArray();
    }

    private (uint headerRowIndex, Dictionary<string, int> columnMapping) FindHeaderAndBuildMapping(SheetData sheetData, WorkbookPart workbookPart)
    {
        var columnMapping = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        uint headerRowIndex = 0;

        // Expected column names (flexible matching)
        var expectedColumns = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase)
        {
            { "Name", new[] { "name" } },
            { "NameEnglish", new[] { "name (english)", "name english", "english name" } },
            { "Company", new[] { "company" } },
            { "Department", new[] { "department" } },
            { "Position", new[] { "position", "title" } },
            { "PositionEnglish", new[] { "position (english)", "position english" } },
            { "Email", new[] { "email" } },
            { "Mobile", new[] { "mobile", "cell phone" } },
            { "Phone", new[] { "phone" } },
            { "Extension", new[] { "extension", "ext" } },
            { "Fax", new[] { "fax" } }
        };

        foreach (var row in sheetData.Elements<Row>().Take(_options.MaxHeaderSearchRows)) // Search configured number of rows
        {
            if (row.RowIndex == null) continue;

            var cellIndex = 0;
            var tempMapping = new Dictionary<string, int>();
            var allHeaders = new Dictionary<int, string>(); // Track all headers for custom fields

            foreach (var cell in row.Elements<Cell>())
            {
                var cellValue = GetCellValue(cell, workbookPart);
                if (string.IsNullOrWhiteSpace(cellValue)) continue;

                cellIndex++;
                allHeaders[cellIndex] = cellValue;

                // Try to match with expected columns
                var matched = false;
                foreach (var (key, aliases) in expectedColumns)
                {
                    if (aliases.Any(alias => cellValue.Contains(alias, StringComparison.OrdinalIgnoreCase)))
                    {
                        tempMapping[key] = cellIndex;
                        matched = true;
                        break;
                    }
                }

                // If not matched with known columns, add as custom field
                if (!matched)
                {
                    tempMapping[cellValue] = cellIndex;
                }
            }

            // If we found at least Name and Email, consider this the header row
            if (tempMapping.ContainsKey("Name") && tempMapping.ContainsKey("Email"))
            {
                headerRowIndex = row.RowIndex.Value;
                columnMapping = tempMapping;
                break;
            }
        }

        return (headerRowIndex, columnMapping);
    }

    private Employee? ParseEmployeeFromRow(Row row, Dictionary<string, int> columnMapping, WorkbookPart workbookPart)
    {
        var cells = row.Elements<Cell>().ToList();
        var cellValues = new Dictionary<int, string>();

        // Extract all cell values
        foreach (var cell in cells)
        {
            if (cell.CellReference == null) continue;
            var cellRef = cell.CellReference?.Value;
            if (string.IsNullOrWhiteSpace(cellRef))
            {
                continue;
            }

            var colIndex = GetColumnIndex(cellRef);
            cellValues[colIndex] = GetCellValue(cell, workbookPart);
        }

        // Get required fields
        var name = columnMapping.ContainsKey("Name") ? cellValues.GetValueOrDefault(columnMapping["Name"], "") : "";
        var email = columnMapping.ContainsKey("Email") ? cellValues.GetValueOrDefault(columnMapping["Email"], "") : "";
        var company = columnMapping.ContainsKey("Company") ? cellValues.GetValueOrDefault(columnMapping["Company"], "Default Company") : "Default Company";

        // Skip if no name
        if (string.IsNullOrWhiteSpace(name))
            return null;

        // Known field names (case-insensitive)
        var knownFields = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Name", "NameEnglish", "Company", "Department", "Position", "PositionEnglish",
            "Email", "Mobile", "Phone", "Extension", "Fax"
        };

        var employee = new Employee
        {
            Name = name.Trim(),
            NameEnglish = columnMapping.ContainsKey("NameEnglish") ? cellValues.GetValueOrDefault(columnMapping["NameEnglish"], "").Trim() : "",
            Company = company.Trim(),
            Department = columnMapping.ContainsKey("Department") ? cellValues.GetValueOrDefault(columnMapping["Department"], "").Trim() : "",
            Position = columnMapping.ContainsKey("Position") ? cellValues.GetValueOrDefault(columnMapping["Position"], "").Trim() : "",
            PositionEnglish = columnMapping.ContainsKey("PositionEnglish") ? cellValues.GetValueOrDefault(columnMapping["PositionEnglish"], "").Trim() : "",
            Email = email.Trim(),
            Mobile = columnMapping.ContainsKey("Mobile") ? cellValues.GetValueOrDefault(columnMapping["Mobile"], "").Trim() : "",
            Phone = columnMapping.ContainsKey("Phone") ? cellValues.GetValueOrDefault(columnMapping["Phone"], "").Trim() : "",
            Extension = columnMapping.ContainsKey("Extension") ? cellValues.GetValueOrDefault(columnMapping["Extension"], "").Trim() : "",
            Fax = columnMapping.ContainsKey("Fax") ? cellValues.GetValueOrDefault(columnMapping["Fax"], "").Trim() : ""
        };

        // Add custom fields for any unmapped columns
        foreach (var (fieldName, colIndex) in columnMapping)
        {
            if (!knownFields.Contains(fieldName))
            {
                var value = cellValues.GetValueOrDefault(colIndex, "").Trim();
                if (!string.IsNullOrEmpty(value))
                {
                    // Use lowercase field name for placeholder matching
                    employee.CustomFields[fieldName.ToLower()] = value;
                }
            }
        }

        // Auto-map position to English if not provided
        if (string.IsNullOrEmpty(employee.PositionEnglish) && !string.IsNullOrEmpty(employee.Position))
        {
            employee.PositionEnglish = employee.GetPositionEnglish();
        }

        return employee;
    }

    private string GetCellValue(Cell cell, WorkbookPart workbookPart)
    {
        if (cell.CellValue == null)
            return string.Empty;

        var value = cell.CellValue.Text;

        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
        {
            var sharedStringTable = workbookPart.SharedStringTablePart?.SharedStringTable;
            if (sharedStringTable != null && int.TryParse(value, out var index))
            {
                return sharedStringTable.ElementAt(index).InnerText;
            }
        }

        return value ?? string.Empty;
    }

    private int GetColumnIndex(string cellReference)
    {
        var letters = new string(cellReference.Where(char.IsLetter).ToArray());
        int index = 0;
        foreach (var c in letters)
        {
            index = index * 26 + (c - 'A' + 1);
        }
        return index;
    }

    private string GetCellReference(uint rowIndex, uint columnIndex)
    {
        string columnName = "";
        while (columnIndex > 0)
        {
            var modulo = (columnIndex - 1) % 26;
            columnName = Convert.ToChar('A' + modulo) + columnName;
            columnIndex = (columnIndex - modulo) / 26;
        }
        return columnName + rowIndex;
    }

    /// <summary>
    /// Validates Excel file (magic bytes and file size)
    /// </summary>
    private void ValidateExcelFile(Stream fileStream)
    {
        // Check file size
        if (fileStream.Length > _options.MaxExcelFileSizeBytes)
        {
            var fileSizeMB = fileStream.Length / (1024.0 * 1024.0);
            var maxSizeMB = _options.MaxExcelFileSizeMB;
            _logger?.LogWarning("Excel file exceeds maximum size: {FileSize:F2}MB > {MaxSize:F2}MB", fileSizeMB, maxSizeMB);
            throw new InvalidOperationException(ErrorCodes.FormatError(ErrorCodes.FileTooLarge, $"Excel file: {fileSizeMB:F1}MB > {maxSizeMB}MB"));
        }

        if (fileStream.Length == 0)
        {
            _logger?.LogWarning("Excel file is empty");
            throw new InvalidOperationException(ErrorCodes.FormatError(ErrorCodes.FileEmpty, "Excel file"));
        }

        // Check file signature (magic bytes)
        var buffer = new byte[8];
        var originalPosition = fileStream.Position;
        fileStream.Position = 0;
        var bytesRead = fileStream.Read(buffer, 0, buffer.Length);
        fileStream.Position = originalPosition;

        if (bytesRead < 4)
        {
            _logger?.LogWarning("Excel file is too small to be valid");
            throw new InvalidOperationException(ErrorCodes.FormatError(ErrorCodes.FileCorrupted, "Excel file"));
        }

        bool isXlsx = buffer[0] == XlsxSignature[0] &&
                      buffer[1] == XlsxSignature[1] &&
                      buffer[2] == XlsxSignature[2] &&
                      buffer[3] == XlsxSignature[3];

        bool isXls = bytesRead >= 8 &&
                     buffer[0] == XlsSignature[0] &&
                     buffer[1] == XlsSignature[1] &&
                     buffer[2] == XlsSignature[2] &&
                     buffer[3] == XlsSignature[3] &&
                     buffer[4] == XlsSignature[4] &&
                     buffer[5] == XlsSignature[5] &&
                     buffer[6] == XlsSignature[6] &&
                     buffer[7] == XlsSignature[7];

        if (!isXlsx && !isXls)
        {
            _logger?.LogWarning("Excel file has invalid signature. Bytes: {B0:X2} {B1:X2} {B2:X2} {B3:X2}",
                buffer[0], buffer[1], buffer[2], buffer[3]);
            throw new InvalidOperationException(ErrorCodes.FormatError(ErrorCodes.InvalidFileFormat, "Expected .xlsx or .xls file"));
        }

        _logger?.LogInformation("Excel file validated successfully ({FileSize} bytes, {FileType})",
            fileStream.Length, isXlsx ? "XLSX" : "XLS");
    }
}
