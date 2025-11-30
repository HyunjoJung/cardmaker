using System.Net;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc.Testing;
using Xunit;

namespace BusinessCardMaker.Tests.Web;

public class WebAppSmokeTests : IClassFixture<WebApplicationFactory<Program>>
{
    private readonly WebApplicationFactory<Program> _factory;

    public WebAppSmokeTests(WebApplicationFactory<Program> factory)
    {
        _factory = factory.WithWebHostBuilder(_ => { });
    }

    [Fact]
    public async Task Health_ReturnsOk_And_SecurityHeaders()
    {
        var client = _factory.CreateClient();

        var response = await client.GetAsync("/health");

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        Assert.True(response.Headers.Contains("Content-Security-Policy"));
        Assert.True(response.Headers.Contains("X-Content-Type-Options"));
        Assert.True(response.Headers.Contains("X-Frame-Options"));
    }

    [Fact]
    public async Task Home_ReturnsHtmlWithTitle()
    {
        var client = _factory.CreateClient();

        var response = await client.GetAsync("/");
        var body = await response.Content.ReadAsStringAsync();

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        Assert.Contains("CardMaker", body, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("/api/templates/basic")]
    [InlineData("/api/templates/qrcode")]
    public async Task TemplateEndpoints_ReturnPptx(string path)
    {
        var client = _factory.CreateClient();

        var response = await client.GetAsync(path);
        var bytes = await response.Content.ReadAsByteArrayAsync();

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        Assert.True(bytes.Length > 1000); // non-empty PPTX
        Assert.Equal("application/vnd.openxmlformats-officedocument.presentationml.presentation",
            response.Content.Headers.ContentType?.MediaType);
    }
}
