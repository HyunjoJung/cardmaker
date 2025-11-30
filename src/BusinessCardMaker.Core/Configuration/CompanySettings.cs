// Copyright (c) 2025 Business Card Maker Contributors
// Licensed under the Apache License, Version 2.0

using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace BusinessCardMaker.Core.Configuration;

public class CompanySettings
{
    public const string SectionName = "BusinessCard";

    public DatabaseSettings Database { get; set; } = new();
    public List<CompanyMapping> Companies { get; set; } = new();
    public ProcessingOptions Processing { get; set; } = new();
}

public class DatabaseSettings
{
    public string? Path { get; set; }
    public bool AutoBackup { get; set; } = true;
    public int BackupIntervalHours { get; set; } = 24;
}

public class CompanyMapping
{
    [Required]
    public string Name { get; set; } = string.Empty;

    [Required]
    public string Template { get; set; } = string.Empty;

    public string? EmailDomain { get; set; }
    public string? Code { get; set; }
}

public class ProcessingOptions
{
    [Range(1, 100)]
    public int MaxFileSizeMB { get; set; } = 10;

    [Range(1, 1000)]
    public int MaxBatchSize { get; set; } = 100;

    [Range(30, 600)]
    public int TimeoutSeconds { get; set; } = 300;

    [Range(1, 100)]
    public int RateLimitPerMinute { get; set; } = 30;
}
