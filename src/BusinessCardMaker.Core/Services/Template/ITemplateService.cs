// Copyright (c) 2025 Business Card Maker Contributors
// Licensed under the Apache License, Version 2.0

namespace BusinessCardMaker.Core.Services.Template;

/// <summary>
/// Service for creating sample PowerPoint templates
/// </summary>
public interface ITemplateService
{
    /// <summary>
    /// Create a basic business card template with standard placeholders
    /// </summary>
    /// <returns>PPTX file as byte array</returns>
    byte[] CreateBasicTemplate();

    /// <summary>
    /// Create a business card template with QR code placeholder
    /// </summary>
    /// <returns>PPTX file as byte array</returns>
    byte[] CreateQRCodeTemplate();
}
