// Copyright (c) 2025 Business Card Maker Contributors
// Licensed under the Apache License, Version 2.0

using System;

namespace BusinessCardMaker.Core.Exceptions;

public class BusinessCardException : Exception
{
    public BusinessCardException() { }

    public BusinessCardException(string message) : base(message) { }

    public BusinessCardException(string message, Exception innerException)
        : base(message, innerException) { }
}

public class ImportException : BusinessCardException
{
    public ImportException(string message) : base(message) { }

    public ImportException(string message, Exception innerException)
        : base(message, innerException) { }
}

public class CardGenerationException : BusinessCardException
{
    public CardGenerationException(string message) : base(message) { }

    public CardGenerationException(string message, Exception innerException)
        : base(message, innerException) { }
}
