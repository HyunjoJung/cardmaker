using System;
using System.Text.RegularExpressions;
using Microsoft.Playwright;
using NUnit.Framework;

namespace BusinessCardMaker.E2ETests;

[TestFixture]
public class HomePageTests : BusinessCardPageTest
{
    [Test]
    public async Task HomePage_ShouldLoadSuccessfully()
    {
        await Page.GotoAsync(BaseUrl);
        await WaitForHomeInteractiveAsync();

        // Check page title
        await Expect(Page).ToHaveTitleAsync(new Regex("CardMaker"));

        // Check main heading
        var heading = Page.Locator("h1");
        await Expect(heading).ToContainTextAsync("CardMaker");
    }

    [Test]
    public async Task Step1_DownloadTemplateButton_ShouldBeVisible()
    {
        await Page.GotoAsync(BaseUrl);
        await WaitForHomeInteractiveAsync();

        // Check if download template button is visible
        var downloadButton = Page.Locator("button:has-text('Excel Template')");
        await Expect(downloadButton).ToBeVisibleAsync();
    }

    [Test]
    public async Task Step2_ExcelUploadButton_ShouldBeDisabledWithoutFile()
    {
        await Page.GotoAsync(BaseUrl);
        await WaitForHomeInteractiveAsync();

        // Generate button should be disabled without files selected
        var generateButton = Page.Locator("button:has-text('Generate Cards')");
        await Expect(generateButton).ToBeDisabledAsync();
    }

    [Test]
    public async Task Step3_GenerateButton_ShouldBeDisabledWithoutTemplate()
    {
        await Page.GotoAsync(BaseUrl);
        await WaitForHomeInteractiveAsync();

        // Generate button should be disabled without template file
        var generateButton = Page.Locator("button:has-text('Generate Cards')");
        await Expect(generateButton).ToBeDisabledAsync();
    }

    [Test]
    public async Task ExcelUpload_WithValidFile_ShouldShowSuccessMessage()
    {
        await Page.GotoAsync(BaseUrl);
        await WaitForHomeInteractiveAsync();
        await Page.Locator("#excelInput").WaitForAsync(new() { State = WaitForSelectorState.Visible, Timeout = 10000 });

        // Create test Excel file
        var excelBytes = TestHelpers.CreateTestEmployeeExcel();
        var tempExcelFile = Path.GetTempFileName() + ".xlsx";
        await File.WriteAllBytesAsync(tempExcelFile, excelBytes);

        try
        {
            // Upload the Excel file (auto-import triggers)
            var fileInput = Page.Locator("input[type='file']#excelInput");
            await fileInput.SetInputFilesAsync(tempExcelFile);

            // Click upload
            var uploadButton = Page.Locator("button:has-text('Upload & Process')");
            await uploadButton.ClickAsync();

            // Wait for success message or fail with details
            var resultAlert = await WaitForImportResultAsync(expectSuccess: true);
            var text = await resultAlert.InnerTextAsync();
            Assert.That(text.Contains("Loaded", StringComparison.OrdinalIgnoreCase), Is.True, $"Unexpected alert text: {text}");
        }
        finally
        {
            if (File.Exists(tempExcelFile))
            {
                File.Delete(tempExcelFile);
            }
        }
    }

    [Test]
    public async Task ExcelUpload_WithInvalidFile_ShouldShowErrorMessage()
    {
        await Page.GotoAsync(BaseUrl);
        await WaitForHomeInteractiveAsync();
        await Page.Locator("#excelInput").WaitForAsync(new() { State = WaitForSelectorState.Visible, Timeout = 10000 });

        // Create a temporary text file (not Excel)
        var tempFile = Path.GetTempFileName();
        await File.WriteAllTextAsync(tempFile, "This is not an Excel file");

        try
        {
            // Upload the invalid file (auto-import triggers)
            var fileInput = Page.Locator("input[type='file']#excelInput");
            await fileInput.SetInputFilesAsync(tempFile);

            var uploadButton = Page.Locator("button:has-text('Upload & Process')");
            await uploadButton.ClickAsync();

            // Wait for error message
            var resultAlert = await WaitForImportResultAsync(expectSuccess: false);
            var classes = await resultAlert.GetAttributeAsync("class") ?? string.Empty;
            Assert.That(classes.Contains("alert-danger"), Is.True, $"Expected danger alert, got classes: {classes}");
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Test]
    public async Task GitHubLink_ShouldBePresent()
    {
        await Page.GotoAsync(BaseUrl);
        await WaitForHomeInteractiveAsync();

        // Check for GitHub link (nav or footer)
        var githubLink = Page.Locator("a[href*='github.com']").First;
        await Expect(githubLink).ToBeVisibleAsync();
    }

    [Test]
    public async Task HowToUse_SectionShouldBeVisible()
    {
        await Page.GotoAsync(BaseUrl);
        await WaitForHomeInteractiveAsync();

        // Scroll down to see "How to Use" section
        await Page.EvaluateAsync("window.scrollTo(0, document.body.scrollHeight)");

        // Check for "How to Use" heading
        var howToSection = Page.Locator("h5:has-text('How to Use')");
        await Expect(howToSection).ToBeVisibleAsync(new() { Timeout = 10000 });
    }

    private async Task<IElementHandle> WaitForImportResultAsync(bool expectSuccess)
    {
        var resultAlert = await Page.WaitForSelectorAsync(".alert-success, .alert-danger",
            new() { State = WaitForSelectorState.Visible, Timeout = 60000 });

        if (resultAlert == null)
        {
            var body = await Page.ContentAsync();
            Assert.Fail($"Import result alert not found. Page content: {body}");
        }

        var classes = await resultAlert!.GetAttributeAsync("class") ?? string.Empty;
        if (expectSuccess && classes.Contains("alert-danger"))
        {
            var text = await resultAlert.InnerTextAsync();
            Assert.Fail($"Import failed: {text}");
        }

        return resultAlert!;
    }

    [Test]
    public async Task HealthCheck_ShouldReturnOK()
    {
        using var client = new HttpClient();
        var response = await client.GetAsync($"{BaseUrl}/health");

        Assert.That(response.IsSuccessStatusCode, Is.True);
        var content = await response.Content.ReadAsStringAsync();
        Assert.That(content, Is.EqualTo("Healthy"));
    }
}
