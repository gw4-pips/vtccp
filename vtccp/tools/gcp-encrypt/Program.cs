// gcpEncrypt — prepares a GS1 GCP Prefix Format List for the Azure update service.
//
// Usage:
//   gcpEncrypt <input.xml> <output.xml.enc> [--key gcpKey.bin] [--meta gcpMeta.json]
//
// Steps:
//   1. Load (or generate) the 32-byte AES-256 key file (gcpKey.bin).
//   2. Encrypt the XML into the GCP1 envelope: "GCP1" | nonce(12) | tag(16) | ciphertext.
//   3. Compute SHA-256 of the *encrypted* file.
//   4. Extract the "date" attribute from the XML root element.
//   5. Write gcpMeta.json: { "date": "...", "sha256": "..." }.
//
// Upload the .enc and gcpMeta.json to the Azure blob container 'gcp-prefix-table/'.
// Wire format is documented in vtccp/architecture/gcp-update-service.md and MUST stay
// in sync with DeviceInterface/Rfid/Gcp/GcpCrypto.cs.

using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml;

const int KeyLength = 32, NonceLength = 12, TagLength = 16;
byte[] magic = "GCP1"u8.ToArray();

var positional = new List<string>();
string keyPath = "gcpKey.bin";
string? metaPath = null;

for (int i = 0; i < args.Length; i++)
{
    switch (args[i])
    {
        case "--key":  keyPath  = args[++i]; break;
        case "--meta": metaPath = args[++i]; break;
        case "-h" or "--help":
            Console.WriteLine("Usage: gcpEncrypt <input.xml> <output.xml.enc> [--key gcpKey.bin] [--meta gcpMeta.json]");
            return 0;
        default: positional.Add(args[i]); break;
    }
}

if (positional.Count != 2)
{
    Console.Error.WriteLine("Usage: gcpEncrypt <input.xml> <output.xml.enc> [--key gcpKey.bin] [--meta gcpMeta.json]");
    return 1;
}

string inputPath = positional[0], outputPath = positional[1];
metaPath ??= Path.Combine(Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? ".", "gcpMeta.json");

if (!File.Exists(inputPath))
{
    Console.Error.WriteLine($"error: input file not found: {inputPath}");
    return 1;
}

// ── 1. Key ──────────────────────────────────────────────────────────────────
byte[] key;
if (File.Exists(keyPath))
{
    key = File.ReadAllBytes(keyPath);
    if (key.Length != KeyLength)
    {
        Console.Error.WriteLine($"error: {keyPath} must be exactly {KeyLength} bytes (found {key.Length}).");
        return 1;
    }
    Console.WriteLine($"key    : {keyPath} (existing)");
}
else
{
    key = RandomNumberGenerator.GetBytes(KeyLength);
    File.WriteAllBytes(keyPath, key);
    Console.WriteLine($"key    : {keyPath} (GENERATED — deploy this beside every VtccpApp.exe and keep it safe)");
}

// ── 2. Read + validate XML, extract root "date" attribute ───────────────────
byte[] plaintext = File.ReadAllBytes(inputPath);
string? tableDate = null;
try
{
    using var stream = new MemoryStream(plaintext, writable: false);
    using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Ignore });
    while (reader.Read())
    {
        if (reader.NodeType != XmlNodeType.Element) continue;
        tableDate = reader.GetAttribute("date");
        break;
    }
}
catch (Exception ex)
{
    Console.Error.WriteLine($"error: input is not valid XML: {ex.Message}");
    return 1;
}
if (string.IsNullOrWhiteSpace(tableDate))
{
    Console.Error.WriteLine("error: root element has no 'date' attribute — is this a GCPPrefixFormatList file?");
    return 1;
}

// ── 3. Encrypt (GCP1 envelope) ──────────────────────────────────────────────
byte[] nonce  = RandomNumberGenerator.GetBytes(NonceLength);
byte[] tag    = new byte[TagLength];
byte[] cipher = new byte[plaintext.Length];
using (var gcm = new AesGcm(key, TagLength))
    gcm.Encrypt(nonce, plaintext, cipher, tag);

byte[] envelope = new byte[magic.Length + NonceLength + TagLength + cipher.Length];
magic.CopyTo(envelope, 0);
nonce.CopyTo(envelope, magic.Length);
tag.CopyTo(envelope, magic.Length + NonceLength);
cipher.CopyTo(envelope, magic.Length + NonceLength + TagLength);
File.WriteAllBytes(outputPath, envelope);

// ── 4–5. SHA-256 of the encrypted file + gcpMeta.json ───────────────────────
string sha256 = Convert.ToHexString(SHA256.HashData(envelope)).ToLowerInvariant();
var meta = new { date = tableDate, sha256 };
File.WriteAllText(metaPath,
    JsonSerializer.Serialize(meta, new JsonSerializerOptions { WriteIndented = true }) + Environment.NewLine,
    new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));

Console.WriteLine($"input  : {inputPath} ({plaintext.Length:N0} bytes)");
Console.WriteLine($"output : {outputPath} ({envelope.Length:N0} bytes)");
Console.WriteLine($"meta   : {metaPath}");
Console.WriteLine($"date   : {tableDate}");
Console.WriteLine($"sha256 : {sha256}");
Console.WriteLine();
Console.WriteLine("Next: upload the .enc and gcpMeta.json to the 'gcp-prefix-table' blob container.");
return 0;
