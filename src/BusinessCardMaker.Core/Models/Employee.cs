// Copyright (c) 2025 Business Card Maker Contributors
// Licensed under the Apache License, Version 2.0

using System;
using System.Collections.Generic;
using System.Linq;

namespace BusinessCardMaker.Core.Models;

/// <summary>
/// Represents an employee for business card generation (memory-only, no persistence)
/// </summary>
public class Employee
{
    public string Name { get; set; } = string.Empty;
    public string NameEnglish { get; set; } = string.Empty;
    public string Company { get; set; } = string.Empty;
    public string Department { get; set; } = string.Empty;
    public string Position { get; set; } = string.Empty;
    public string PositionEnglish { get; set; } = string.Empty;
    public string Email { get; set; } = string.Empty;
    public string Mobile { get; set; } = string.Empty;
    public string Phone { get; set; } = string.Empty;
    public string Extension { get; set; } = string.Empty;
    public string Fax { get; set; } = string.Empty;

    /// <summary>
    /// Custom fields for dynamic columns (e.g., LinkedIn, Hobby, Nickname, etc.)
    /// Key: lowercase field name, Value: field value
    /// </summary>
    public Dictionary<string, string> CustomFields { get; set; } = new();

    // Common position to English mapping (can be overridden in configuration)
    public static readonly Dictionary<string, string> DefaultPositionMapping = new()
    {
        { "대표", "CEO" },
        { "부대표", "Vice President" },
        { "감사", "Auditor" },
        { "상무", "Executive Director" },
        { "전무", "Managing Director" },
        { "이사", "Director" },
        { "본부장", "Division Manager" },
        { "그룹장", "Group Leader" },
        { "팀장", "Team Leader" },
        { "클럽장", "Club Leader" },
        { "수석", "Senior" },
        { "책임", "Principal" },
        { "차장부장", "Deputy General Manager" },
        { "매니저", "Manager" },
        { "대리", "Senior Account Executive" },
        { "사원", "Staff" },
        { "팀원", "Team Member" },
        { "인턴", "Intern" },
        { "디자이너", "Designer" },
        { "AE", "Account Executive" },
        { "차장", "Deputy Manager" },
        { "부장", "General Manager" },
        { "과장", "Section Manager" },
        { "주임", "Senior Staff" }
    };

    public string GetPositionEnglish(Dictionary<string, string>? customMapping = null)
    {
        // Use existing English position if set
        if (!string.IsNullOrEmpty(PositionEnglish))
            return PositionEnglish;

        if (string.IsNullOrEmpty(Position))
            return string.Empty;

        // Try custom mapping first, then default mapping
        var mapping = customMapping ?? DefaultPositionMapping;
        if (mapping.TryGetValue(Position, out var englishPosition))
        {
            PositionEnglish = englishPosition;
            return englishPosition;
        }

        // Return original if no mapping found
        return Position;
    }

    public string GetFormattedMobile()
    {
        if (string.IsNullOrEmpty(Mobile) || Mobile == "-" || Mobile == "0")
            return string.Empty;

        // Extract digits only
        var digits = new string(Mobile.Where(char.IsDigit).ToArray());

        // Format as: 10. 1234. 5678 for Korean mobile numbers (010-xxxx-xxxx)
        if (digits.Length == 11 && digits.StartsWith("010"))
        {
            return $"{digits.Substring(1, 2)}. {digits.Substring(3, 4)}. {digits.Substring(7, 4)}";
        }

        // Already formatted or other format
        return Mobile.Replace("010-", "10. ").Replace("-", ". ");
    }

    public string GetFormattedExtension()
    {
        if (string.IsNullOrEmpty(Extension) || Extension == "-" || Extension == "0")
            return string.Empty;

        // Format specific extension patterns (can be customized via configuration)
        if (Extension.StartsWith("3210-"))
        {
            return Extension.Replace("3210-", "2. 3210. ");
        }

        // 4-digit extension
        if (Extension.Length == 4 && Extension.All(char.IsDigit))
        {
            return $"2. 3210. {Extension}";
        }

        return Extension;
    }
}
