// Copyright (c) 2025 Business Card Maker Contributors
// Licensed under the Apache License, Version 2.0

using BusinessCardMaker.Core.Services.Import;
using BusinessCardMaker.Core.Services.CardGenerator;
using BusinessCardMaker.Core.Services.QRCode;
using BusinessCardMaker.Core.Services.Template;
using BusinessCardMaker.Core.Configuration;
using BusinessCardMaker.Web.Components;
using Microsoft.AspNetCore.RateLimiting;
using System.Threading.RateLimiting;

var builder = WebApplication.CreateBuilder(args);

// Add configuration with validation
builder.Services.AddOptions<BusinessCardProcessingOptions>()
    .Bind(builder.Configuration.GetSection(BusinessCardProcessingOptions.SectionName))
    .ValidateDataAnnotations()
    .ValidateOnStart();

var disableRateLimiter = builder.Configuration.GetValue<bool>("DisableRateLimiter") ||
                         builder.Environment.IsEnvironment("Development") ||
                         string.Equals(Environment.GetEnvironmentVariable("DISABLE_RATE_LIMITER"), "true", StringComparison.OrdinalIgnoreCase);

// Add memory cache and response caching
builder.Services.AddMemoryCache();
builder.Services.AddResponseCaching();

// Add services to the container
builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();

// Register stateless application services (memory-only, no database)
builder.Services.AddScoped<IExcelImportService, ExcelImportService>();
builder.Services.AddScoped<IQRCodeService, QRCodeService>();
builder.Services.AddScoped<ITemplateService, TemplateService>();
builder.Services.AddScoped<ICardGeneratorService, CardGeneratorService>();

// Add health checks
builder.Services.AddHealthChecks();

// Add rate limiting (configurable requests per minute per IP) unless disabled
var processingOptions = builder.Configuration
    .GetSection(BusinessCardProcessingOptions.SectionName)
    .Get<BusinessCardProcessingOptions>() ?? new BusinessCardProcessingOptions();

if (!disableRateLimiter)
{
    builder.Services.AddRateLimiter(options =>
    {
        options.GlobalLimiter = PartitionedRateLimiter.Create<HttpContext, string>(context =>
            RateLimitPartition.GetFixedWindowLimiter(
                partitionKey: context.Connection.RemoteIpAddress?.ToString() ?? "unknown",
                factory: partition => new FixedWindowRateLimiterOptions
                {
                    PermitLimit = processingOptions.RateLimitPerMinute,
                    Window = TimeSpan.FromMinutes(1),
                    QueueProcessingOrder = QueueProcessingOrder.OldestFirst,
                    QueueLimit = 0
                }));

        options.RejectionStatusCode = 429; // Too Many Requests
    });
}

var app = builder.Build();

// Configure the HTTP request pipeline
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error", createScopeForErrors: true);
    app.UseHsts();
}

app.UseStatusCodePagesWithReExecute("/not-found", createScopeForStatusCodePages: true);
app.UseHttpsRedirection();

// Add security headers
app.Use(async (context, next) =>
{
    // Content Security Policy
    context.Response.Headers.Append("Content-Security-Policy",
        "default-src 'self'; " +
        "script-src 'self' 'unsafe-inline' 'unsafe-eval'; " + // Blazor requires unsafe-eval
        "style-src 'self' 'unsafe-inline'; " +
        "img-src 'self' data:; " +
        "font-src 'self'; " +
        "connect-src 'self' wss: ws:; " + // WebSocket for Blazor SignalR
        "frame-ancestors 'none'; " +
        "base-uri 'self'; " +
        "form-action 'self'");

    // Additional security headers
    context.Response.Headers.Append("X-Content-Type-Options", "nosniff");
    context.Response.Headers.Append("X-Frame-Options", "DENY");
    context.Response.Headers.Append("X-XSS-Protection", "1; mode=block");
    context.Response.Headers.Append("Referrer-Policy", "strict-origin-when-cross-origin");
    context.Response.Headers.Append("Permissions-Policy", "geolocation=(), microphone=(), camera=()");

    await next();
});

// Enable rate limiting
if (!disableRateLimiter)
{
    app.UseRateLimiter();
}
app.UseAntiforgery();
app.UseResponseCaching();

// Map health check endpoint
app.MapHealthChecks("/health");

// Map template download endpoints (before Razor components to prevent routing conflicts)
app.MapGet("/api/templates/basic", (ITemplateService templateService) =>
{
    var template = templateService.CreateBasicTemplate();
    return Results.File(template, "application/vnd.openxmlformats-officedocument.presentationml.presentation", "BusinessCard_Basic_Template.pptx");
})
.WithName("DownloadBasicTemplate")
.DisableAntiforgery();

app.MapGet("/api/templates/qrcode", (ITemplateService templateService) =>
{
    var template = templateService.CreateQRCodeTemplate();
    return Results.File(template, "application/vnd.openxmlformats-officedocument.presentationml.presentation", "BusinessCard_QRCode_Template.pptx");
})
.WithName("DownloadQRCodeTemplate")
.DisableAntiforgery();

app.MapStaticAssets();
app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();

app.Run();

// Expose Program for WebApplicationFactory in tests
public partial class Program { }
