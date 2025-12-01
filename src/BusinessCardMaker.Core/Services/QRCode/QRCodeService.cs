// Copyright (c) 2025 Business Card Maker Contributors
// Licensed under the Apache License, Version 2.0

using System;
using System.IO;
using System.Text;
using BusinessCardMaker.Core.Models;
using QRCoder;

namespace BusinessCardMaker.Core.Services.QRCode;

/// <summary>
/// Service for generating QR codes with vCard contact information
/// </summary>
public class QRCodeService : IQRCodeService
{
    public byte[] GenerateQRCode(Employee employee)
    {
        // Generate vCard 3.0 format
        var vCard = GenerateVCard(employee);

        // Generate QR code
        using var qrGenerator = new QRCodeGenerator();
        using var qrCodeData = qrGenerator.CreateQrCode(vCard, QRCodeGenerator.ECCLevel.Q);
        using var qrCode = new PngByteQRCode(qrCodeData);

        // Generate PNG with 10 pixels per module
        return qrCode.GetGraphic(10);
    }

    private string GenerateVCard(Employee employee)
    {
        var vCard = new StringBuilder();

        vCard.AppendLine("BEGIN:VCARD");
        vCard.AppendLine("VERSION:3.0");

        // Full name
        if (!string.IsNullOrEmpty(employee.Name))
        {
            vCard.AppendLine($"FN:{employee.Name}");

            // Structured name (Family name;Given name;Additional names;Honorific prefixes;Honorific suffixes)
            vCard.AppendLine($"N:{employee.Name};;;;");
        }

        // Organization
        if (!string.IsNullOrEmpty(employee.Company))
        {
            vCard.AppendLine($"ORG:{employee.Company}");
        }

        // Title/Position
        if (!string.IsNullOrEmpty(employee.Position))
        {
            vCard.AppendLine($"TITLE:{employee.Position}");
        }

        // Email
        if (!string.IsNullOrEmpty(employee.Email))
        {
            vCard.AppendLine($"EMAIL;TYPE=WORK:{employee.Email}");
        }

        // Mobile phone
        if (!string.IsNullOrEmpty(employee.Mobile))
        {
            vCard.AppendLine($"TEL;TYPE=CELL:{employee.Mobile}");
        }

        // Office phone
        if (!string.IsNullOrEmpty(employee.Phone))
        {
            vCard.AppendLine($"TEL;TYPE=WORK:{employee.Phone}");
        }

        // Fax
        if (!string.IsNullOrEmpty(employee.Fax))
        {
            vCard.AppendLine($"TEL;TYPE=FAX:{employee.Fax}");
        }

        // Custom fields (e.g., LinkedIn, Website)
        foreach (var (fieldName, fieldValue) in employee.CustomFields)
        {
            if (fieldName.Equals("linkedin", StringComparison.OrdinalIgnoreCase) ||
                fieldName.Equals("website", StringComparison.OrdinalIgnoreCase) ||
                fieldName.Equals("url", StringComparison.OrdinalIgnoreCase))
            {
                vCard.AppendLine($"URL:{fieldValue}");
            }
            else
            {
                // Add as notes for other custom fields
                vCard.AppendLine($"NOTE:{fieldName}: {fieldValue}");
            }
        }

        vCard.AppendLine("END:VCARD");

        return vCard.ToString();
    }
}
