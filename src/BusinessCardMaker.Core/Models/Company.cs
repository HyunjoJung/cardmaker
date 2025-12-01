// Copyright (c) 2025 Business Card Maker Contributors
// Licensed under the Apache License, Version 2.0

using System;
using System.Collections.Generic;

namespace BusinessCardMaker.Core.Models;

/// <summary>
/// Represents a company (tenant) in the multi-tenant system
/// </summary>
public class Company
{
    public int Id { get; set; }

    /// <summary>
    /// Company name
    /// </summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// Email domain for this company (e.g., "companya.com")
    /// Used for automatic company assignment during registration
    /// </summary>
    public string EmailDomain { get; set; } = string.Empty;

    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;

    /// <summary>
    /// Whether this company is active
    /// </summary>
    public bool IsActive { get; set; } = true;

    /// <summary>
    /// Plan type: Free or Premium
    /// </summary>
    public string PlanType { get; set; } = "Free";

    /// <summary>
    /// Maximum number of employees allowed (enforced for Free plan)
    /// </summary>
    public int MaxEmployees { get; set; } = 100;

    /// <summary>
    /// Custom template path for this company (optional)
    /// </summary>
    public string? TemplatePath { get; set; }
}
