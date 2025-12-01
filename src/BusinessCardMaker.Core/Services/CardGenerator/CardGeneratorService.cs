// Copyright (c) 2025 Business Card Maker Contributors
// Licensed under the Apache License, Version 2.0

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using BusinessCardMaker.Core.Models;
using BusinessCardMaker.Core.Exceptions;
using BusinessCardMaker.Core.Services.QRCode;
using BusinessCardMaker.Core.Configuration;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace BusinessCardMaker.Core.Services.CardGenerator;

/// <summary>
/// Service for batch business card generation from PowerPoint templates
/// </summary>
public class CardGeneratorService : ICardGeneratorService
{
    private readonly ILogger<CardGeneratorService>? _logger;
    private readonly IQRCodeService? _qrCodeService;
    private readonly BusinessCardProcessingOptions _options;

    // PowerPoint file signature (ZIP format: PK..)
    private static readonly byte[] PptxSignature = { 0x50, 0x4B, 0x03, 0x04 };

    // Thread-safe picture ID counter
    private static int _pictureIdCounter = 1000;

    public CardGeneratorService(
        ILogger<CardGeneratorService>? logger = null,
        IQRCodeService? qrCodeService = null,
        IOptions<BusinessCardProcessingOptions>? options = null)
    {
        _logger = logger;
        _qrCodeService = qrCodeService;
        _options = options?.Value ?? new BusinessCardProcessingOptions();
    }

    public async Task<CardGenerationResult> GenerateBatchAsync(
        List<Employee> employees,
        Stream templateStream,
        IProgress<int>? progress = null)
    {
        // Input validation
        if (employees == null || employees.Count == 0)
        {
            return CardGenerationResult.CreateFailure(ErrorCodes.FormatError(ErrorCodes.NoEmployeesProvided));
        }

        if (templateStream == null)
        {
            throw new ArgumentNullException(nameof(templateStream), ErrorCodes.GetErrorMessage(ErrorCodes.InvalidFileFormat, "Template stream cannot be null"));
        }

        // Ensure we work with a seekable template stream (BrowserFileStream does not support Position)
        Stream seekableTemplateStream;
        try
        {
            if (templateStream.CanSeek)
            {
                templateStream.Seek(0, SeekOrigin.Begin);
                seekableTemplateStream = templateStream;
            }
            else
            {
                throw new NotSupportedException("Stream is not seekable");
            }
        }
        catch (NotSupportedException)
        {
            var memoryStream = new MemoryStream();
            await templateStream.CopyToAsync(memoryStream);
            memoryStream.Position = 0;
            seekableTemplateStream = memoryStream;
        }

        if (employees.Count > _options.MaxBatchSize)
        {
            return CardGenerationResult.CreateFailure(ErrorCodes.FormatError(ErrorCodes.BatchSizeExceeded, $"Provided {employees.Count} employees, maximum is {_options.MaxBatchSize}."));
        }

        var generatedFiles = new List<string>();
        var errors = new List<string>();
        string? tempDir = null;
        string? zipFilePath = null;

        try
        {
            _logger?.LogInformation("Starting batch card generation for {Count} employees", employees.Count);

            // Validate PowerPoint template file
            ValidatePowerPointFile(seekableTemplateStream);

            // Create temporary directory for this batch
            tempDir = Path.Combine(Path.GetTempPath(), $"BusinessCards_{Guid.NewGuid():N}");
            Directory.CreateDirectory(tempDir);

            // Save template to temp file (we'll use it for each employee)
            var templatePath = Path.Combine(tempDir, "template.pptx");
            using (var fileStream = File.Create(templatePath))
            {
                seekableTemplateStream.Position = 0;
                await seekableTemplateStream.CopyToAsync(fileStream);
            }

            _logger?.LogDebug("Template saved to {Path}", templatePath);

            // Process each employee
            int processed = 0;
            foreach (var employee in employees)
            {
                try
                {
                    var outputPath = await GenerateCardForEmployeeAsync(employee, templatePath, tempDir);
                    generatedFiles.Add(outputPath);
                    _logger?.LogDebug("Generated card for {Name}", employee.Name);
                }
                catch (Exception ex)
                {
                    _logger?.LogWarning(ex, "Failed to generate card for {Name}", employee.Name);
                    errors.Add($"{employee.Name}: {ex.Message}");
                }

                processed++;
                progress?.Report((processed * 90) / employees.Count); // Use up to 90% for generation
            }

            if (generatedFiles.Count == 0)
            {
                return CardGenerationResult.CreateFailure(ErrorCodes.FormatError(ErrorCodes.PowerPointProcessingFailed, "All business card generation failed."));
            }

            // Create zip file
            zipFilePath = Path.Combine(Path.GetTempPath(), $"BusinessCards_{DateTime.Now:yyyyMMdd_HHmmss}_{Guid.NewGuid():N}.zip");
            await CreateZipFileAsync(generatedFiles, zipFilePath, tempDir);
            _logger?.LogInformation("Created zip file: {ZipPath}", zipFilePath);

            progress?.Report(100);

            var result = new CardGenerationResult
            {
                Success = true,
                GeneratedFiles = generatedFiles,
                SuccessCount = generatedFiles.Count,
                FailedCount = errors.Count,
                ZipFilePath = zipFilePath,
                Errors = errors
            };

            return result;
        }
        catch (OutOfMemoryException)
        {
            _logger?.LogError("Out of memory while generating business cards");

            // Clean up zip file if it was created
            if (!string.IsNullOrEmpty(zipFilePath) && File.Exists(zipFilePath))
            {
                try { File.Delete(zipFilePath); } catch { }
            }

            return CardGenerationResult.CreateFailure(ErrorCodes.FormatError(ErrorCodes.OutOfMemory));
        }
        catch (IOException ex)
        {
            _logger?.LogError(ex, "IO error during business card generation");

            // Clean up zip file if it was created
            if (!string.IsNullOrEmpty(zipFilePath) && File.Exists(zipFilePath))
            {
                try { File.Delete(zipFilePath); } catch { }
            }

            return CardGenerationResult.CreateFailure(ErrorCodes.FormatError(ErrorCodes.IoError, ex.Message));
        }
        catch (UnauthorizedAccessException ex)
        {
            _logger?.LogError(ex, "Permission denied during business card generation");

            // Clean up zip file if it was created
            if (!string.IsNullOrEmpty(zipFilePath) && File.Exists(zipFilePath))
            {
                try { File.Delete(zipFilePath); } catch { }
            }

            return CardGenerationResult.CreateFailure(ErrorCodes.FormatError(ErrorCodes.PermissionDenied, ex.Message));
        }
        catch (BusinessCardException ex)
        {
            _logger?.LogError(ex, "Business card processing error");

            // Clean up zip file if it was created
            if (!string.IsNullOrEmpty(zipFilePath) && File.Exists(zipFilePath))
            {
                try { File.Delete(zipFilePath); } catch { }
            }

            return CardGenerationResult.CreateFailure(ex.Message); // Already formatted
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Unexpected error during business card generation");

            // Clean up zip file if it was created
            if (!string.IsNullOrEmpty(zipFilePath) && File.Exists(zipFilePath))
            {
                try { File.Delete(zipFilePath); } catch { }
            }

            return CardGenerationResult.CreateFailure(ErrorCodes.FormatError(ErrorCodes.UnexpectedError, ex.Message));
        }
        finally
        {
            // Clean up temp directory (but keep zip file)
            if (!string.IsNullOrEmpty(tempDir) && Directory.Exists(tempDir))
            {
                try
                {
                    Directory.Delete(tempDir, true);
                    _logger?.LogDebug("Cleaned up temp directory: {TempDir}", tempDir);
                }
                catch (Exception ex)
                {
                    _logger?.LogWarning(ex, "Failed to clean up temp directory: {TempDir}", tempDir);
                }
            }
        }
    }

    private async Task<string> GenerateCardForEmployeeAsync(Employee employee, string templatePath, string outputDir)
    {
        // Generate safe filename
        var safeFileName = $"{employee.Name}_{employee.Company}_{DateTime.Now:yyyyMMdd_HHmmss}.pptx"
            .Replace(" ", "_")
            .Replace("/", "_")
            .Replace("\\", "_")
            .Replace(":", "_");

        var outputPath = Path.Combine(outputDir, safeFileName);

        // Copy template to output path
        await Task.Run(() => File.Copy(templatePath, outputPath, true));

        // Process PowerPoint file
        await Task.Run(() => ProcessPowerPointFile(outputPath, employee));

        return outputPath;
    }

    private async Task CreateZipFileAsync(List<string> files, string zipFilePath, string baseDirectory)
    {
        await Task.Run(() =>
        {
            using var zipArchive = ZipFile.Open(zipFilePath, ZipArchiveMode.Create);

            foreach (var file in files)
            {
                var entryName = Path.GetFileName(file);
                zipArchive.CreateEntryFromFile(file, entryName, CompressionLevel.Optimal);
            }
        });
    }

    private int ProcessPowerPointFile(string filePath, Employee employee)
    {
        int totalReplacements = 0;

        try
        {
            using var presentationDocument = PresentationDocument.Open(filePath, true);
            var presentationPart = presentationDocument.PresentationPart;

            if (presentationPart?.Presentation.SlideIdList == null)
            {
                throw new BusinessCardException("Invalid PowerPoint file");
            }

            // Process all slides
            foreach (var slideId in presentationPart.Presentation.SlideIdList.Elements<SlideId>())
            {
                var slidePart = presentationPart.GetPartById(slideId.RelationshipId!) as SlidePart;
                if (slidePart?.Slide == null)
                {
                    continue;
                }

                totalReplacements += ReplaceTextInSlide(slidePart, employee, filePath);
            }

            // Save changes
            presentationPart.Presentation.Save();
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to process PowerPoint file: {FilePath}", filePath);
            throw new BusinessCardException($"PowerPoint processing failed: {ex.Message}", ex);
        }

        return totalReplacements;
    }

    private int ReplaceTextInSlide(SlidePart slidePart, Employee employee, string presentationPath)
    {
        int replacements = 0;

        var shapes = slidePart.Slide.CommonSlideData?.ShapeTree?.Elements<Shape>();
        if (shapes == null) return 0;

        var shapesToProcess = shapes.ToList();
        foreach (var shape in shapesToProcess)
        {
            var textBody = shape.TextBody;
            if (textBody == null) continue;

            var paragraphsToRemove = new List<A.Paragraph>();

            foreach (var paragraph in textBody.Elements<A.Paragraph>().ToList())
            {
                var runs = paragraph.Elements<A.Run>().ToList();
                if (runs.Count == 0) continue;

                var fullText = string.Join("", runs.Select(r => r.Text?.Text ?? ""));

                // Handle QR code placeholder
                if (fullText.Contains("{qrcode}", StringComparison.OrdinalIgnoreCase))
                {
                    if (_qrCodeService != null)
                    {
                        ReplaceShapeWithQRCode(slidePart, shape, employee);
                        replacements++;
                    }
                    else
                    {
                        // Remove QR code placeholder if service is not available
                        paragraphsToRemove.Add(paragraph);
                    }
                    continue;
                }

                // Remove paragraphs with empty fields
                if (ShouldRemoveParagraph(fullText, employee))
                {
                    paragraphsToRemove.Add(paragraph);
                    replacements++;
                    continue;
                }

                var replacedText = ReplaceEmployeeData(fullText, employee);

                if (fullText != replacedText)
                {
                    // Remove all runs and replace with a single new run
                    foreach (var run in runs.ToList())
                    {
                        run.Remove();
                    }

                    // Create new run with replaced text
                    var newRun = new A.Run();
                    var newText = new A.Text { Text = replacedText };

                    // Copy formatting from first run if available
                    var runProps = runs.FirstOrDefault()?.RunProperties;
                    if (runProps != null)
                    {
                        newRun.RunProperties = (A.RunProperties)runProps.CloneNode(true)!;
                    }

                    newRun.Append(newText);

                    // Adjust text box for long names (4+ characters)
                    if (fullText.Contains("{name}") && employee.Name.Length >= 4)
                    {
                        AdjustTextBoxForLongName(shape, employee.Name.Length);
                    }

                    // Insert before EndParagraphRunProperties if it exists
                    var endParaRPr = paragraph.Elements<A.EndParagraphRunProperties>().FirstOrDefault();
                    if (endParaRPr != null)
                    {
                        paragraph.InsertBefore(newRun, endParaRPr);
                    }
                    else
                    {
                        paragraph.Append(newRun);
                    }

                    replacements++;
                }
            }

            // Remove empty paragraphs
            foreach (var paragraph in paragraphsToRemove)
            {
                paragraph.Remove();
            }
        }

        slidePart.Slide.Save();
        return replacements;
    }

    private bool ShouldRemoveParagraph(string text, Employee employee)
    {
        // Remove line if extension is empty
        if (string.IsNullOrEmpty(employee.GetFormattedExtension()) &&
            (text.Contains("{extension}") || text.Contains("T. +82") || text.Contains("Ext.")))
            return true;

        // Remove line if mobile is empty
        if (string.IsNullOrEmpty(employee.GetFormattedMobile()) &&
            (text.Contains("{mobile}") || text.Contains("M.") || text.Contains("Mobile")))
            return true;

        // Remove line if fax is empty
        if (string.IsNullOrEmpty(employee.Fax) &&
            (text.Contains("{fax}") || text.Contains("F.") || text.Contains("Fax")))
            return true;

        return false;
    }

    private string ReplaceEmployeeData(string text, Employee employee)
    {
        if (string.IsNullOrEmpty(text)) return text;

        // Replace standard placeholders
        text = text
            .Replace("{name}", employee.Name)
            .Replace("{name_en}", employee.NameEnglish)
            .Replace("{position}", employee.Position)
            .Replace("{position_en}", employee.GetPositionEnglish())
            .Replace("{department}", employee.Department)
            .Replace("{email}", employee.Email)
            .Replace("{mobile}", employee.GetFormattedMobile())
            .Replace("{extension}", employee.GetFormattedExtension())
            .Replace("{phone}", employee.Phone)
            .Replace("{fax}", employee.Fax);

        // Replace custom field placeholders (e.g., {linkedin}, {hobby}, {nickname})
        foreach (var (fieldName, fieldValue) in employee.CustomFields)
        {
            var placeholder = $"{{{fieldName}}}";
            if (text.Contains(placeholder, StringComparison.OrdinalIgnoreCase))
            {
                text = text.Replace(placeholder, fieldValue, StringComparison.OrdinalIgnoreCase);
            }
        }

        return text;
    }

    private void AdjustTextBoxForLongName(Shape nameShape, int nameLength)
    {
        try
        {
            var spPr = nameShape.ShapeProperties;
            if (spPr?.Transform2D?.Extents == null || spPr.Transform2D.Offset == null)
                return;

            var extents = spPr.Transform2D.Extents;
            var offset = spPr.Transform2D.Offset;
            var currentWidth = extents.Cx ?? 0;
            var nameX = offset.X ?? 0;

            // 4 chars: 25% increase, 5+ chars: 40% increase
            double widthMultiplier = nameLength == 4 ? 1.25 : 1.40;
            long newWidth = (long)(currentWidth * widthMultiplier);
            long widthIncrease = newWidth - currentWidth;

            extents.Cx = newWidth;

            // Prevent text wrapping
            var textBody = nameShape.TextBody;
            if (textBody?.BodyProperties != null)
            {
                textBody.BodyProperties.Wrap = A.TextWrappingValues.None;
            }

            // Adjust English name position if needed
            AdjustEnglishNamePosition(nameShape, widthIncrease, nameX);
        }
        catch (Exception ex)
        {
            _logger?.LogWarning(ex, "Failed to adjust text box size for long name");
        }
    }

    private void ReplaceShapeWithQRCode(SlidePart slidePart, Shape shape, Employee employee)
    {
        try
        {
            if (_qrCodeService == null) return;

            // Generate QR code image bytes
            var qrCodeBytes = _qrCodeService.GenerateQRCode(employee);

            // Add image part to slide
            var imagePart = slidePart.AddImagePart(ImagePartType.Png);
            using (var stream = new MemoryStream(qrCodeBytes))
            {
                imagePart.FeedData(stream);
            }

            // Get the shape's position and size
            var spPr = shape.ShapeProperties;
            var transform = spPr?.Transform2D;
            var offset = transform?.Offset;
            var extents = transform?.Extents;

            long x = offset?.X ?? 0;
            long y = offset?.Y ?? 0;
            long width = extents?.Cx ?? 914400; // Default: 1 inch = 914400 EMUs
            long height = extents?.Cy ?? 914400;

            // QR codes must be square - use the smaller dimension to ensure it fits
            long size = Math.Min(width, height);
            width = size;
            height = size;

            // Create picture shape
            var relationshipId = slidePart.GetIdOfPart(imagePart);
            var picture = CreatePictureShape(relationshipId, x, y, width, height);

            // Replace the text shape with picture
            var shapeTree = slidePart.Slide.CommonSlideData?.ShapeTree;
            if (shapeTree != null)
            {
                shapeTree.InsertAfter(picture, shape);
                shape.Remove();
            }

            _logger?.LogDebug("Inserted QR code for {Name}", employee.Name);
        }
        catch (Exception ex)
        {
            _logger?.LogWarning(ex, "Failed to insert QR code for {Name}", employee.Name);
        }
    }

    private P.Picture CreatePictureShape(string relationshipId, long x, long y, long width, long height)
    {
        var picture = new P.Picture();

        // Non-visual picture properties (use thread-safe counter)
        var nvPicPr = new P.NonVisualPictureProperties(
            new P.NonVisualDrawingProperties
            {
                Id = (uint)Interlocked.Increment(ref _pictureIdCounter),
                Name = "QR Code"
            },
            new P.NonVisualPictureDrawingProperties(
                new A.PictureLocks { NoChangeAspect = true }),
            new DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties());

        // Picture fill
        var blipFill = new P.BlipFill(
            new A.Blip { Embed = relationshipId },
            new A.Stretch(new A.FillRectangle()));

        // Shape properties with position and size
        var shapeProperties = new P.ShapeProperties(
            new A.Transform2D(
                new A.Offset { X = x, Y = y },
                new A.Extents { Cx = width, Cy = height }),
            new A.PresetGeometry(new A.AdjustValueList())
            {
                Preset = A.ShapeTypeValues.Rectangle
            });

        picture.Append(nvPicPr);
        picture.Append(blipFill);
        picture.Append(shapeProperties);

        return picture;
    }

    private void AdjustEnglishNamePosition(Shape nameShape, long widthIncrease, long nameX)
    {
        try
        {
            var slide = nameShape.Ancestors<Slide>().FirstOrDefault();
            if (slide == null) return;

            var shapes = slide.CommonSlideData?.ShapeTree?.Elements<Shape>();
            if (shapes == null) return;

            foreach (var shape in shapes)
            {
                if (shape == nameShape) continue;

                var textBody = shape.TextBody;
                if (textBody == null) continue;

                var allText = string.Join("", textBody.Descendants<A.Text>().Select(t => t.Text));
                if (!allText.Contains("{name_en}")) continue;

                var englishSpPr = shape.ShapeProperties;
                if (englishSpPr?.Transform2D?.Offset == null) continue;

                var englishOffset = englishSpPr.Transform2D.Offset;
                var englishX = englishOffset.X ?? 0;

                // Move English name if it's to the right of the name field
                if (englishX > nameX)
                {
                    englishOffset.X = englishX + widthIncrease;
                    _logger?.LogDebug("Adjusted English name position: {OldX} → {NewX}", englishX, englishX + widthIncrease);
                }
                break;
            }
        }
        catch (Exception ex)
        {
            _logger?.LogWarning(ex, "Failed to adjust English name position");
        }
    }

    /// <summary>
    /// Validates PowerPoint template file (magic bytes and file size)
    /// </summary>
    private void ValidatePowerPointFile(Stream fileStream)
    {
        // Check file size
        if (fileStream.Length > _options.MaxTemplateFileSizeBytes)
        {
            var fileSizeMB = fileStream.Length / (1024.0 * 1024.0);
            var maxSizeMB = _options.MaxTemplateFileSizeMB;
            _logger?.LogWarning("Template file exceeds maximum size: {FileSize:F2}MB > {MaxSize:F2}MB", fileSizeMB, maxSizeMB);
            throw new BusinessCardException(ErrorCodes.FormatError(ErrorCodes.FileTooLarge, $"Template file: {fileSizeMB:F1}MB > {maxSizeMB}MB"));
        }

        if (fileStream.Length == 0)
        {
            _logger?.LogWarning("Template file is empty");
            throw new BusinessCardException(ErrorCodes.FormatError(ErrorCodes.FileEmpty, "Template file"));
        }

        // Check file signature (magic bytes)
        var buffer = new byte[4];
        var originalPosition = fileStream.Position;
        fileStream.Position = 0;
        var bytesRead = fileStream.Read(buffer, 0, buffer.Length);
        fileStream.Position = originalPosition;

        if (bytesRead < 4)
        {
            _logger?.LogWarning("Template file is too small to be valid PowerPoint");
            throw new BusinessCardException(ErrorCodes.FormatError(ErrorCodes.FileCorrupted, "PowerPoint template"));
        }

        bool isPptx = buffer[0] == PptxSignature[0] &&
                      buffer[1] == PptxSignature[1] &&
                      buffer[2] == PptxSignature[2] &&
                      buffer[3] == PptxSignature[3];

        if (!isPptx)
        {
            _logger?.LogWarning("Template file has invalid PowerPoint signature. Bytes: {B0:X2} {B1:X2} {B2:X2} {B3:X2}",
                buffer[0], buffer[1], buffer[2], buffer[3]);
            throw new BusinessCardException(ErrorCodes.FormatError(ErrorCodes.InvalidFileFormat, "Expected .pptx file"));
        }

        _logger?.LogInformation("Template file validated successfully ({FileSize} bytes)", fileStream.Length);
    }
}
