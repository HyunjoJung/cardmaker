// Copyright (c) 2025 Business Card Maker Contributors
// Licensed under the Apache License, Version 2.0

using System;
using System.IO;
using System.Threading.Tasks;
using BusinessCardMaker.Core.Models;

namespace BusinessCardMaker.Core.Services.Import;

public interface IExcelImportService
{
    Task<ImportResult> ImportFromExcelAsync(Stream excelStream, IProgress<int>? progress = null);
    byte[] CreateTemplate();
}
