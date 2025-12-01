# CardMaker

Stateless, database-free business card generator that ingests employee data from Excel, merges it into PowerPoint templates with OpenXML, and generates QR codes for vCards. Built with Blazor Server for a fast, secure, and easy-to-deploy solution.

🌐 **Live Demo**: [cardmaker.hyunjo.uk](https://cardmaker.hyunjo.uk)

## Highlights
- Excel import (xlsx) with header auto-detection and a downloadable template; no database writes.
- PowerPoint card generation with formatting preservation via `DocumentFormat.OpenXml`; optional PDF conversion via LibreOffice.
- QR code generation for business cards with customizable vCard data.
- Security-first defaults in the web host (CSP, rate limiting, health checks).
- Docker- and compose-ready; Apache 2.0 licensed.

## Project Layout
```
src/
  BusinessCardMaker.Core/   # Core models, config DTOs, OpenXML services
  BusinessCardMaker.Web/    # Blazor Server host, DI setup, components, static assets
tests/
  BusinessCardMaker.Tests/  # xUnit test project (Moq available)
  BusinessCardMaker.E2ETests/ # Playwright E2E tests

Dockerfile.web             # Production image for the web app
docker-compose.yml         # Local container runner
```

## Quick Start
Prereqs: .NET 10 SDK, Node not required, Docker optional.

```bash
# Restore and build everything
dotnet build BusinessCardMaker.slnx

# Run the web app
dotnet run --project src/BusinessCardMaker.Web
# browse http://localhost:5255

# Tests (add tests under tests/BusinessCardMaker.Tests)
dotnet test BusinessCardMaker.slnx

# Docker
docker compose up --build         # or: docker build -f Dockerfile.web -t cardmaker:latest .
docker run -p 5049:5049 cardmaker:latest
```

## Templates
- Excel template downloads directly from the UI (Step 1).
- PowerPoint template is served as a static file: `wwwroot/templates/CardMaker_Template.pptx` (includes QR placeholder).

## Development Notes
- Core services live in `src/BusinessCardMaker.Core/Services` (`Import` for Excel ingestion, `CardGenerator` for PPT output). All services are DI-friendly.
- Web host wiring is in `src/BusinessCardMaker.Web/Program.cs`; components live under `Components/` with layouts and shared pieces separated.
- Configuration uses `appsettings*.json` with `BusinessCard` sections for processing limits and template mapping. Keep secrets out of source; prefer environment variables for overrides.
- Keep the stateless philosophy: no DB writes, no user accounts, clear data at session end.

## Contributing
Issues and PRs are welcome! Please keep changes small, tested, and documented.
