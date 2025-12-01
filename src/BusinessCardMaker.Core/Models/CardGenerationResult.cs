// Copyright (c) 2025 Business Card Maker Contributors
// Licensed under the Apache License, Version 2.0

using System.Collections.Generic;

namespace BusinessCardMaker.Core.Models;

/// <summary>
/// Result of batch business card generation
/// </summary>
public class CardGenerationResult
{
    public bool Success { get; set; }
    public string? ErrorMessage { get; set; }
    public List<string> GeneratedFiles { get; set; } = new();
    public int SuccessCount { get; set; }
    public int FailedCount { get; set; }
    public string? ZipFilePath { get; set; }
    public List<string> Errors { get; set; } = new();

    public static CardGenerationResult CreateSuccess(List<string> generatedFiles, string zipFilePath)
    {
        return new CardGenerationResult
        {
            Success = true,
            GeneratedFiles = generatedFiles,
            SuccessCount = generatedFiles.Count,
            FailedCount = 0,
            ZipFilePath = zipFilePath
        };
    }

    public static CardGenerationResult CreateFailure(string errorMessage)
    {
        return new CardGenerationResult
        {
            Success = false,
            ErrorMessage = errorMessage
        };
    }
}
