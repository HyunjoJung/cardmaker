# CardMaker

Stateless, memory-only business card generator that ingests employee data from Excel, merges it into PowerPoint templates with OpenXML, and serves a Blazor Server UI. Designed to be fast to run, easy to contribute, and safe for open source.

🌐 **Live Demo**: [cardmaker.hyunjo.uk](https://cardmaker.hyunjo.uk)

## Highlights
- Excel import (xlsx) with header auto-detection and a downloadable template; no database writes.
- PowerPoint card generation with formatting preservation via `DocumentFormat.OpenXml`; optional PDF conversion via LibreOffice.
- Security-first defaults in the web host (CSP, rate limiting, health checks).
- Docker- and compose-ready; Apache 2.0 licensed.

## Project Layout
```
src/
  BusinessCardMaker.Core/   # Core models, config DTOs, OpenXML services
  BusinessCardMaker.Web/    # Blazor Server host, DI setup, components, static assets
tests/
  BusinessCardMaker.Tests/  # xUnit test project (Moq available)

Dockerfile.web             # Production image for the web app
docker-compose.yml         # Local container runner
ppts/                      # PowerPoint templates (example placeholders)
AGENTS.md                  # Contributor guide
plan.md                    # Build plan and roadmap
```

## Quick Start
Prereqs: .NET 10 SDK, Node not required, Docker optional.

```bash
# Restore and build everything
dotnet build BusinessCardMaker.slnx

# Run the web app
dotnet run --project src/BusinessCardMaker.Web
# browse http://localhost:5000

# Tests (add tests under tests/BusinessCardMaker.Tests)
dotnet test BusinessCardMaker.slnx

# Docker
docker compose up --build         # or: docker build -f Dockerfile.web -t businesscardmaker:latest .
docker run -p 5050:8080 businesscardmaker:latest
```

## Development Notes
- Core services live in `src/BusinessCardMaker.Core/Services` (`Import` for Excel ingestion, `CardGenerator` for PPT output). All services are DI-friendly.
- Web host wiring is in `src/BusinessCardMaker.Web/Program.cs`; components live under `Components/` with layouts and shared pieces separated.
- Configuration uses `appsettings*.json` with `BusinessCard` sections for processing limits and template mapping. Keep secrets out of source; prefer environment variables for overrides.
- Keep the stateless philosophy: no DB writes, no user accounts, clear data at session end.

## Contributing
Read `AGENTS.md` for coding style, commands, and PR expectations. Roadmap items and UX targets are outlined in `plan.md`. Issues and PRs are welcome—keep changes small, tested, and documented.***
