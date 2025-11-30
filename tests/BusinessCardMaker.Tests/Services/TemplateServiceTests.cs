using System.Linq;
using BusinessCardMaker.Core.Services.Template;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace BusinessCardMaker.Tests.Services;

public class TemplateServiceTests
{
    [Theory]
    [InlineData(true, "{qrcode}")]
    [InlineData(false, "{name}")]
    public void TemplatesContainExpectedPlaceholders(bool withQr, string expectedPlaceholder)
    {
        var service = new TemplateService();
        var template = withQr ? service.CreateQRCodeTemplate() : service.CreateBasicTemplate();

        using var doc = PresentationDocument.Open(new MemoryStream(template), false);
        var texts = doc.PresentationPart!.SlideParts
            .SelectMany(sp => sp.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Text>())
            .Select(t => t.Text)
            .ToList();

        Assert.Contains(texts, t => t.Contains(expectedPlaceholder));
    }
}
