// Copyright (c) 2025 Business Card Maker Contributors
// Licensed under the Apache License, Version 2.0

using BusinessCardMaker.Core.Models;

namespace BusinessCardMaker.Core.Services.QRCode;

/// <summary>
/// Service for generating QR codes from employee information
/// </summary>
public interface IQRCodeService
{
    /// <summary>
    /// Generate a QR code image from employee vCard data
    /// </summary>
    /// <param name="employee">Employee information</param>
    /// <returns>PNG image bytes of the QR code</returns>
    byte[] GenerateQRCode(Employee employee);
}
