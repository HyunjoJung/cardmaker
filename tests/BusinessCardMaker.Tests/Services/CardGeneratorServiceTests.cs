using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using BusinessCardMaker.Core.Models;
using BusinessCardMaker.Core.Services.CardGenerator;
using BusinessCardMaker.Core.Services.QRCode;
using BusinessCardMaker.Core.Services.Template;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using Xunit;

namespace BusinessCardMaker.Tests.Services;

public class CardGeneratorServiceTests
{
    private const string PptxContentType = "application/vnd.openxmlformats-officedocument.presentationml.presentation";

    [Fact]
    public async Task GenerateBatchAsync_InsertsQrCodeImage_WhenPlaceholderExists()
    {
        // Arrange
        var templateBytes = new TemplateService().CreateQRCodeTemplate();
        var employee = new Employee
        {
            Name = "Alice",
            Company = "Contoso",
            Email = "alice@contoso.com",
            Mobile = "+1-555-0100"
        };

        var service = new CardGeneratorService(qrCodeService: new FakeQrCodeService());

        // Act
        var result = await service.GenerateBatchAsync(
            new List<Employee> { employee },
            new MemoryStream(templateBytes));

        // Assert
        Assert.True(result.Success, $"ErrorMessage: {result.ErrorMessage}, Errors: {string.Join("|", result.Errors ?? new List<string>())}");
        Assert.False(string.IsNullOrEmpty(result.ZipFilePath));
        var zipPath = result.ZipFilePath!;
        Assert.True(File.Exists(zipPath));

        using var pptStream = await ExtractSinglePresentationAsync(zipPath);
        using var ppt = PresentationDocument.Open(pptStream, false);
        var slidePart = ppt.PresentationPart!.SlideParts.First();

        var pictures = slidePart.Slide.Descendants<P.Picture>().ToList();
        var texts = slidePart.Slide.Descendants<A.Text>().Select(t => t.Text).ToList();

        Assert.True(pictures.Count > 0, $"Expected QR picture, found {pictures.Count}. Texts: {string.Join("|", texts)}");
        Assert.DoesNotContain(texts, t => t.Contains("{qrcode}", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public async Task GenerateBatchAsync_ReplacesCustomFieldPlaceholders()
    {
        // Arrange: start from basic template and convert one text box to a custom placeholder
        var templateBytes = new TemplateService().CreateBasicTemplate();
        templateBytes = ReplaceFirstText(templateBytes, "Your Company Address", "{linkedin}");

        var employee = new Employee
        {
            Name = "Bob",
            Company = "Fabrikam",
            Email = "bob@fabrikam.com",
            CustomFields = new Dictionary<string, string>
            {
                ["linkedin"] = "https://www.linkedin.com/in/bob"
            }
        };

        var service = new CardGeneratorService(qrCodeService: new FakeQrCodeService());

        // Act
        var result = await service.GenerateBatchAsync(
            new List<Employee> { employee },
            new MemoryStream(templateBytes));

        // Assert
        Assert.True(result.Success, string.Join(";", result.Errors));
        var zipPath = result.ZipFilePath!;
        using var pptStream = await ExtractSinglePresentationAsync(zipPath);
        using var ppt = PresentationDocument.Open(pptStream, false);
        var slide = ppt.PresentationPart!.SlideParts.First().Slide;
        var texts = slide.Descendants<A.Text>().Select(t => t.Text).ToList();

        Assert.Contains(texts, t => t.Contains("https://www.linkedin.com/in/bob", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(texts, t => t.Contains("{linkedin}", StringComparison.OrdinalIgnoreCase));
    }

    private static async Task<MemoryStream> ExtractSinglePresentationAsync(string zipPath)
    {
        using var zip = ZipFile.OpenRead(zipPath);
        var entry = zip.Entries.Single();
        using var entryStream = entry.Open();
        var ms = new MemoryStream();
        await entryStream.CopyToAsync(ms);
        ms.Position = 0;
        return ms;
    }

    private static byte[] ReplaceFirstText(byte[] pptxBytes, string findText, string replaceText)
    {
        using var ms = new MemoryStream();
        ms.Write(pptxBytes, 0, pptxBytes.Length);
        ms.Position = 0;
        using (var doc = PresentationDocument.Open(ms, true))
        {
            var textNode = doc.PresentationPart!
                .SlideParts
                .SelectMany(sp => sp.Slide.Descendants<A.Text>())
                .FirstOrDefault(t => t.Text.Contains(findText, StringComparison.OrdinalIgnoreCase));

            if (textNode != null)
            {
                textNode.Text = replaceText;
                doc.PresentationPart.Presentation.Save();
            }
        }
        return ms.ToArray();
    }

    private sealed class FakeQrCodeService : IQRCodeService
    {
        // 1x1 transparent PNG
        private static readonly byte[] PngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMB/6XqvucAAAAASUVORK5CYII=");

        public byte[] GenerateQRCode(Employee employee) => PngBytes;
    }
}
