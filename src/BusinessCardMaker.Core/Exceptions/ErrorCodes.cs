// Copyright (c) 2025 Business Card Maker Contributors
// Licensed under the Apache License, Version 2.0

namespace BusinessCardMaker.Core.Exceptions;

/// <summary>
/// Structured error codes for business card processing
/// </summary>
public static class ErrorCodes
{
    // File Validation Errors (E001-E009)
    public const string InvalidFileFormat = "E001";
    public const string FileTooLarge = "E002";
    public const string FileEmpty = "E003";
    public const string FileCorrupted = "E004";
    public const string InvalidTemplateStructure = "E005";

    // Data Validation Errors (E010-E019)
    public const string NoEmployeesProvided = "E010";
    public const string InvalidHeaderFormat = "E011";
    public const string MissingRequiredColumn = "E012";
    public const string InvalidDataFormat = "E013";
    public const string BatchSizeExceeded = "E014";

    // System Resource Errors (E020-E029)
    public const string OutOfMemory = "E020";
    public const string IoError = "E021";
    public const string PermissionDenied = "E022";
    public const string DiskSpaceExhausted = "E023";

    // Processing Errors (E030-E039)
    public const string PowerPointProcessingFailed = "E030";
    public const string ExcelProcessingFailed = "E031";
    public const string QRCodeGenerationFailed = "E032";
    public const string TemplateGenerationFailed = "E033";
    public const string ZipCreationFailed = "E034";

    // Unknown/Unexpected Errors (E999)
    public const string UnexpectedError = "E999";

    /// <summary>
    /// Get user-friendly message for error code
    /// </summary>
    public static string GetErrorMessage(string errorCode, string? details = null)
    {
        var baseMessage = errorCode switch
        {
            InvalidFileFormat => "File is not a valid format.",
            FileTooLarge => "File size exceeds maximum allowed.",
            FileEmpty => "File is empty or has no content.",
            FileCorrupted => "File appears to be corrupted.",
            InvalidTemplateStructure => "Template file has invalid structure.",

            NoEmployeesProvided => "No employee data provided.",
            InvalidHeaderFormat => "Excel header format is invalid.",
            MissingRequiredColumn => "Required column is missing.",
            InvalidDataFormat => "Data format is invalid.",
            BatchSizeExceeded => "Batch size exceeds maximum allowed.",

            OutOfMemory => "File is too large to process in memory.",
            IoError => "Could not read or write file.",
            PermissionDenied => "Permission denied to access file.",
            DiskSpaceExhausted => "Insufficient disk space.",

            PowerPointProcessingFailed => "PowerPoint processing failed.",
            ExcelProcessingFailed => "Excel processing failed.",
            QRCodeGenerationFailed => "QR code generation failed.",
            TemplateGenerationFailed => "Template generation failed.",
            ZipCreationFailed => "ZIP file creation failed.",

            UnexpectedError => "An unexpected error occurred.",
            _ => "Unknown error."
        };

        return details != null ? $"{errorCode}: {baseMessage} {details}" : $"{errorCode}: {baseMessage}";
    }

    /// <summary>
    /// Get recovery suggestion for error code
    /// </summary>
    public static string? GetRecoverySuggestion(string errorCode)
    {
        return errorCode switch
        {
            InvalidFileFormat => "💡 Tip: Ensure the file is in .xlsx (Excel) or .pptx (PowerPoint) format.",
            FileTooLarge => "💡 Tip: Try reducing file size by removing unnecessary data or splitting into smaller files.",
            FileEmpty => "💡 Tip: Check that the file uploaded correctly and contains data.",
            FileCorrupted => "💡 Tip: Try opening the file in Excel/PowerPoint and re-saving it.",
            InvalidTemplateStructure => "💡 Tip: Use the provided template or ensure all required elements are present.",

            NoEmployeesProvided => "💡 Tip: Add at least one employee to the Excel file.",
            InvalidHeaderFormat => "💡 Tip: Use the provided Excel template for correct header format.",
            MissingRequiredColumn => "💡 Tip: Ensure 'Name' and 'Email' columns are present in the Excel file.",
            InvalidDataFormat => "💡 Tip: Check that data matches expected format (e.g., email addresses, phone numbers).",
            BatchSizeExceeded => "💡 Tip: Split your employee data into smaller batches.",

            OutOfMemory => "💡 Tip: Reduce file size or split data into smaller batches.",
            IoError => "💡 Tip: Check if the file is locked, corrupted, or on a network drive with connectivity issues.",
            PermissionDenied => "💡 Tip: Check file and folder permissions.",
            DiskSpaceExhausted => "💡 Tip: Free up disk space and try again.",

            PowerPointProcessingFailed => "💡 Tip: Verify the PowerPoint template is not password-protected or corrupted.",
            ExcelProcessingFailed => "💡 Tip: Verify the Excel file is not password-protected or corrupted.",
            QRCodeGenerationFailed => "💡 Tip: Ensure employee data includes valid contact information.",
            TemplateGenerationFailed => "💡 Tip: Try restarting the application.",
            ZipCreationFailed => "💡 Tip: Ensure sufficient disk space and write permissions.",

            _ => null
        };
    }

    /// <summary>
    /// Format complete error message with code, message, and suggestion
    /// </summary>
    public static string FormatError(string errorCode, string? details = null)
    {
        var message = GetErrorMessage(errorCode, details);
        var suggestion = GetRecoverySuggestion(errorCode);

        return suggestion != null ? $"{message}\n\n{suggestion}" : message;
    }
}
