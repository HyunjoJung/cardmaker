// Copyright (c) 2025 Business Card Maker Contributors
// Licensed under the Apache License, Version 2.0

using System.ComponentModel.DataAnnotations;

namespace BusinessCardMaker.Core.Configuration;

/// <summary>
/// Configuration options for business card processing
/// </summary>
public class BusinessCardProcessingOptions
{
    public const string SectionName = "BusinessCardProcessing";

    /// <summary>
    /// Maximum Excel file size in MB
    /// </summary>
    [Range(1, 100)]
    public int MaxExcelFileSizeMB { get; set; } = 10;

    /// <summary>
    /// Maximum PowerPoint template file size in MB
    /// </summary>
    [Range(1, 100)]
    public int MaxTemplateFileSizeMB { get; set; } = 50;

    /// <summary>
    /// Maximum number of rows to search for header in Excel
    /// </summary>
    [Range(1, 20)]
    public int MaxHeaderSearchRows { get; set; } = 5;

    /// <summary>
    /// Rate limit per minute per IP (0 to disable)
    /// </summary>
    [Range(0, 1000)]
    public int RateLimitPerMinute { get; set; } = 100;

    /// <summary>
    /// Maximum employees to process in a single batch
    /// </summary>
    [Range(1, 10000)]
    public int MaxBatchSize { get; set; } = 1000;

    // Calculated properties
    public long MaxExcelFileSizeBytes => MaxExcelFileSizeMB * 1024L * 1024L;
    public long MaxTemplateFileSizeBytes => MaxTemplateFileSizeMB * 1024L * 1024L;
}
