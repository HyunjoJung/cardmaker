// Copyright (c) 2025 Business Card Maker Contributors
// Licensed under the Apache License, Version 2.0

using System.Collections.Generic;

namespace BusinessCardMaker.Core.Models;

/// <summary>
/// Result of Excel import operation (memory-only)
/// </summary>
public class ImportResult
{
    public bool Success { get; set; }
    public string? ErrorMessage { get; set; }
    public List<Employee> Employees { get; set; } = new();
    public int TotalCount => Employees.Count;
    public List<string> Warnings { get; set; } = new();

    public static ImportResult CreateSuccess(List<Employee> employees, List<string>? warnings = null)
    {
        return new ImportResult
        {
            Success = true,
            Employees = employees,
            Warnings = warnings ?? new List<string>()
        };
    }

    public static ImportResult CreateFailure(string errorMessage)
    {
        return new ImportResult
        {
            Success = false,
            ErrorMessage = errorMessage,
            Employees = new List<Employee>()
        };
    }
}
