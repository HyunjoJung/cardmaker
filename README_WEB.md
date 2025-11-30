# Business Card Maker - Web Version

Professional business card generator web application built with Blazor Server.

## Status

**Current Status**: Stateless refactor in progress — core import/generation services are ready, UI polish and tests in flight.

### Completed
- ✅ Core library with models, OpenXML import/generation services, configuration
- ✅ Blazor Server host with clean base UI and layouts
- ✅ Security headers, Rate Limiting, Health checks
- ✅ Docker & docker-compose configuration
- ✅ Apache 2.0 License

### In Progress / TODO
- ⏳ Blazor components for end-to-end stateless workflow (upload Excel, upload template, download ZIP)
- ⏳ Add focused unit tests and smoke tests
- ⏳ Complete documentation pass and contribution templates
- ⏳ Optional: PDF conversion via LibreOffice

See [plan.md](plan.md) for detailed migration plan.

## Quick Start

### Run Locally
```bash
dotnet run --project src/BusinessCardMaker.Web
```

Visit http://localhost:5000

### Build
```bash
dotnet build BusinessCardMaker.slnx
```

### Docker
```bash
# Using docker-compose
docker compose up --build

# Or build manually
docker build -f Dockerfile.web -t businesscardmaker:latest .
docker run -p 5050:8080 businesscardmaker:latest
```

## Project Structure

```
BusinessCardMaker/
├── src/
│   ├── BusinessCardMaker.Core/      # Business logic library (stateless services)
│   │   ├── Models/                  # Employee, results, configuration DTOs
│   │   ├── Services/                # Excel import & PPT generation
│   │   └── Exceptions/              # Custom exceptions
│   └── BusinessCardMaker.Web/       # Blazor Server web app
│       ├── Components/
│       │   ├── Pages/               # Blazor pages (stateless flow)
│       │   └── Layout/              # Layout components and shared UI
│       ├── Program.cs               # DI, middleware, security headers
│       └── appsettings.json         # Configuration defaults
│
├── tests/
│   └── BusinessCardMaker.Tests/     # Unit tests (xUnit, Moq)
│
├── docker-compose.yml               # Docker Compose configuration
├── Dockerfile.web                   # Web app Dockerfile
├── plan.md                          # Stateless build plan
└── LICENSE                          # Apache 2.0
```

## Tech Stack

- **Framework**: .NET 10.0
- **UI**: Blazor Server (interactive components)
- **PowerPoint**: DocumentFormat.OpenXml (cross-platform, no COM)
- **Security**: CSP headers, Rate Limiting, HSTS
- **Testing**: xUnit, Moq, Playwright

## Configuration

Edit `appsettings.json` (Blazor project). Example:

```json
{
  "BusinessCard": {
    "Companies": [
      {
        "Name": "Default",
        "Template": "ppts/default.pptx",
        "EmailDomain": "company.com"
      }
    ],
    "Processing": {
      "MaxFileSizeMB": 10,
      "MaxBatchSize": 100,
      "TimeoutSeconds": 300,
      "RateLimitPerMinute": 30
    }
  }
}
```

## Key Features (Current & Planned)

- **Excel Import**: Auto-detect headers, provide template download, memory-only parsing.
- **Card Generation**: OpenXML-based PPTX merge with placeholder replacement; optional ZIP output.
- **(Optional) PDF Export**: Via LibreOffice inside the container.
- **Security**: Rate limiting, CSP, HSTS, and health checks.
- **Responsive UI**: Blazor components with a single-page flow for upload → generate → download.

## Development

### Prerequisites
- .NET 10.0 SDK
- Docker (optional)

### Run Tests
```bash
dotnet test BusinessCardMaker.slnx
```

### Run with Watch
```bash
dotnet watch --project BusinessCardMaker.Web
```

## Migration from Windows Forms

This project is migrated from a Windows Forms desktop application to a modern web application.

**Why Web?**
- ✅ Cross-platform (Windows, Linux, macOS)
- ✅ No PowerPoint COM dependency
- ✅ Accessible from any browser
- ✅ Easier deployment & updates
- ✅ Better security & scalability

**Migration Progress**: See [plan.md](plan.md)

## Contributing

This is an open-source project. Contributions are welcome!

1. Fork the repository
2. Create a feature branch
3. Implement your changes (see plan.md for TODOs)
4. Write tests
5. Submit a pull request

## Disclaimer

This is an independent open-source project and is not affiliated with Microsoft Corporation.

"PowerPoint" and "Microsoft Office" are trademarks of Microsoft Corporation. This tool works with PowerPoint file formats using the open-source DocumentFormat.OpenXml library.

## License

Apache License 2.0 - see [LICENSE](LICENSE) for details.

## Author

Created by [HyunjoJung](https://github.com/HyunjoJung)

---

**Note**: This project is currently in active development. Core structure is complete, but service implementations are pending. See [plan.md](plan.md) for implementation roadmap.
