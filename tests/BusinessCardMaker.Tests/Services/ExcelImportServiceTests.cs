using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using BusinessCardMaker.Core.Services.Import;
using Microsoft.Extensions.Logging.Abstractions;
using Xunit;

namespace BusinessCardMaker.Tests.Services;

public class ExcelImportServiceTests
{
    [Fact]
    public async Task ImportFromExcelAsync_AllowsAsyncOnlyStream()
    {
        var service = new ExcelImportService(new NullLogger<ExcelImportService>());
        var bytes = service.CreateTemplate();
        Assert.True(bytes.Length > 0);
        await using var asyncOnlyStream = new AsyncOnlyStream(bytes);

        // Sanity check: async read should produce data even though sync reads are disallowed
        var probeBuffer = new byte[1024];
        var probeRead = await asyncOnlyStream.ReadAsync(probeBuffer, 0, probeBuffer.Length);
        Assert.True(probeRead > 0);
        asyncOnlyStream.Position = 0;

        var result = await service.ImportFromExcelAsync(asyncOnlyStream);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.True(result.TotalCount > 0);
        Assert.False(string.IsNullOrWhiteSpace(result.Employees.Single().Name));
    }

    private sealed class AsyncOnlyStream : Stream
    {
        private readonly MemoryStream _inner;

        public AsyncOnlyStream(byte[] buffer)
        {
            _inner = new MemoryStream(buffer, writable: false);
        }

        public override bool CanRead => true;
        public override bool CanSeek => _inner.CanSeek;
        public override bool CanWrite => false;
        public override long Length => _inner.Length;
        public override long Position { get => _inner.Position; set => _inner.Position = value; }

        public override void Flush() => throw new NotSupportedException();

        public override int Read(byte[] buffer, int offset, int count) =>
            throw new NotSupportedException("Synchronous reads are not supported.");

        public override int Read(Span<byte> buffer) =>
            throw new NotSupportedException("Synchronous reads are not supported.");

        public override ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
        {
            var read = _inner.Read(buffer.Span);
            return ValueTask.FromResult(read);
        }

        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) =>
            Task.FromResult(_inner.Read(buffer, offset, count));

        public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);

        public override void SetLength(long value) => throw new NotSupportedException();

        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        public override void Write(ReadOnlySpan<byte> buffer) => throw new NotSupportedException();

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _inner.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
