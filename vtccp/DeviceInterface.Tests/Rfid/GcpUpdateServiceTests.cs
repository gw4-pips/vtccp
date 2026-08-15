using System.Net;
using System.Security.Cryptography;
using System.Text;
using DeviceInterface.Rfid.Gcp;
using Xunit;

namespace DeviceInterface.Tests.Rfid;

public sealed class GcpCryptoTests
{
    [Fact]
    public void EncryptDecrypt_RoundTrips()
    {
        byte[] key = RandomNumberGenerator.GetBytes(GcpCrypto.KeyLength);
        byte[] plain = Encoding.UTF8.GetBytes("<GCPPrefixFormatList date=\"2026-06-03\"/>");

        byte[] envelope = GcpCrypto.Encrypt(plain, key);
        Assert.Equal(plain, GcpCrypto.Decrypt(envelope, key));
    }

    [Fact]
    public void Decrypt_WrongKey_Throws()
    {
        byte[] envelope = GcpCrypto.Encrypt("payload"u8.ToArray(),
            RandomNumberGenerator.GetBytes(GcpCrypto.KeyLength));
        Assert.ThrowsAny<CryptographicException>(() =>
            GcpCrypto.Decrypt(envelope, RandomNumberGenerator.GetBytes(GcpCrypto.KeyLength)));
    }

    [Fact]
    public void Decrypt_BadMagic_Throws()
    {
        byte[] key = RandomNumberGenerator.GetBytes(GcpCrypto.KeyLength);
        byte[] envelope = GcpCrypto.Encrypt("payload"u8.ToArray(), key);
        envelope[0] ^= 0xFF;
        Assert.Throws<InvalidDataException>(() => GcpCrypto.Decrypt(envelope, key));
    }

    [Fact]
    public void Decrypt_TamperedCiphertext_Throws()
    {
        byte[] key = RandomNumberGenerator.GetBytes(GcpCrypto.KeyLength);
        byte[] envelope = GcpCrypto.Encrypt("payload"u8.ToArray(), key);
        envelope[^1] ^= 0xFF;
        Assert.ThrowsAny<CryptographicException>(() => GcpCrypto.Decrypt(envelope, key));
    }
}

public sealed class GcpUpdateServiceTests : IDisposable
{
    private const string Xml =
        "<GCPPrefixFormatList date=\"2026-07-01T00:00:00Z\">" +
        "<entry prefix=\"000000\" gcpLength=\"7\" />" +
        "</GCPPrefixFormatList>";

    private readonly string _dir = Directory.CreateTempSubdirectory("gcp-test").FullName;

    public void Dispose() { try { Directory.Delete(_dir, recursive: true); } catch { } }

    /// <summary>Fake HTTP handler serving /api/gcpMeta and /api/gcpTable.</summary>
    private sealed class FakeHandler(byte[] envelope, string metaJson) : HttpMessageHandler
    {
        public string? SeenToken;
        protected override Task<HttpResponseMessage> SendAsync(
            HttpRequestMessage request, CancellationToken ct)
        {
            SeenToken = request.Headers.TryGetValues("X-Device-Token", out var v) ? v.First() : null;
            HttpResponseMessage resp = request.RequestUri!.AbsolutePath.EndsWith("gcpMeta")
                ? new(HttpStatusCode.OK) { Content = new StringContent(metaJson) }
                : new(HttpStatusCode.OK) { Content = new ByteArrayContent(envelope) };
            return Task.FromResult(resp);
        }
    }

    private (GcpUpdateService service, FakeHandler handler) Build(string tableDate = "2026-07-01T00:00:00Z")
    {
        byte[] key = RandomNumberGenerator.GetBytes(GcpCrypto.KeyLength);
        string keyPath = Path.Combine(_dir, "gcpKey.bin");
        File.WriteAllBytes(keyPath, key);

        byte[] envelope = GcpCrypto.Encrypt(Encoding.UTF8.GetBytes(Xml), key);
        string sha = Convert.ToHexString(SHA256.HashData(envelope)).ToLowerInvariant();
        string metaJson = $"{{\"date\":\"{tableDate}\",\"sha256\":\"{sha}\"}}";

        var handler = new FakeHandler(envelope, metaJson);
        var service = new GcpUpdateService(new GcpUpdateOptions
        {
            ServiceUrl   = "https://example.test",
            DeviceToken  = "tok-1",
            LocalXmlPath = Path.Combine(_dir, "gcpPrefixes.xml"),
            KeyPath      = keyPath,
        }, new HttpClient(handler));
        return (service, handler);
    }

    [Fact]
    public async Task CheckNow_NoLocalFile_ReportsUpdateAvailable()
    {
        var (service, handler) = Build();
        var check = await service.CheckNowAsync();

        Assert.NotNull(check);
        Assert.True(check!.UpdateAvailable);
        Assert.Null(check.LocalDate);
        Assert.Equal("tok-1", handler.SeenToken);
    }

    [Fact]
    public async Task DownloadAndInstall_VerifiesDecryptsAndWrites()
    {
        var (service, _) = Build();
        var date = await service.DownloadAndInstallAsync();

        Assert.Equal(new DateTimeOffset(2026, 7, 1, 0, 0, 0, TimeSpan.Zero), date);
        string installed = Path.Combine(_dir, "gcpPrefixes.xml");
        Assert.True(File.Exists(installed));

        var table = GcpLengthTable.LoadFromFile(installed);
        Assert.Equal(1, table.EntryCount);
    }

    [Fact]
    public async Task CheckNow_AfterInstall_ReportsUpToDate()
    {
        var (service, _) = Build();
        await service.DownloadAndInstallAsync();
        var check = await service.CheckNowAsync();

        Assert.NotNull(check);
        Assert.False(check!.UpdateAvailable);
        Assert.Equal(check.ServerDate, check.LocalDate);
    }

    [Fact]
    public async Task DownloadAndInstall_HashMismatch_Throws()
    {
        byte[] key = RandomNumberGenerator.GetBytes(GcpCrypto.KeyLength);
        string keyPath = Path.Combine(_dir, "gcpKey.bin");
        File.WriteAllBytes(keyPath, key);
        byte[] envelope = GcpCrypto.Encrypt(Encoding.UTF8.GetBytes(Xml), key);
        // deliberately wrong hash
        string metaJson = "{\"date\":\"2026-07-01T00:00:00Z\",\"sha256\":\"" + new string('0', 64) + "\"}";

        var service = new GcpUpdateService(new GcpUpdateOptions
        {
            ServiceUrl = "https://example.test", DeviceToken = "tok-1",
            LocalXmlPath = Path.Combine(_dir, "gcpPrefixes.xml"), KeyPath = keyPath,
        }, new HttpClient(new FakeHandler(envelope, metaJson)));

        await Assert.ThrowsAsync<InvalidDataException>(() => service.DownloadAndInstallAsync());
        Assert.False(File.Exists(Path.Combine(_dir, "gcpPrefixes.xml")));
    }

    [Fact]
    public async Task CheckNow_NotConfigured_ReturnsNull()
    {
        var service = new GcpUpdateService(new GcpUpdateOptions
        {
            ServiceUrl = "", DeviceToken = "", LocalXmlPath = Path.Combine(_dir, "x.xml"),
        });
        Assert.Null(await service.CheckNowAsync());
    }
}
