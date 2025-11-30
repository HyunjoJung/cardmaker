// Copyright (c) 2025 Business Card Maker Contributors
// Licensed under the Apache License, Version 2.0

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using BusinessCardMaker.Core.Models;

namespace BusinessCardMaker.Core.Services.CardGenerator;

/// <summary>
/// Service for batch business card generation from PowerPoint templates
/// </summary>
public interface ICardGeneratorService
{
    /// <summary>
    /// Generates business cards for multiple employees and creates a zip file
    /// </summary>
    /// <param name="employees">List of employees to generate cards for</param>
    /// <param name="templateStream">PowerPoint template stream</param>
    /// <param name="progress">Progress reporter (0-100)</param>
    /// <returns>Generation result with zip file path</returns>
    Task<CardGenerationResult> GenerateBatchAsync(
        List<Employee> employees,
        Stream templateStream,
        IProgress<int>? progress = null);
}
