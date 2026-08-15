using System.Security.Cryptography;

namespace DeviceInterface.Rfid.Gcp;

/// <summary>
/// AES-256-GCM envelope used for the distributed GCP prefix table.
///
/// Wire format of a <c>.enc</c> file (documented in
/// <c>vtccp/architecture/gcp-update-service.md</c> — keep in sync with the
/// standalone <c>gcpEncrypt</c> CLI in <c>vtccp/tools/gcp-encrypt/</c>):
///
///   bytes 0–3   magic  "GCP1" (ASCII)
///   bytes 4–15  nonce  (12 bytes, random per file)
///   bytes 16–31 tag    (16-byte GCM authentication tag)
///   bytes 32–…  ciphertext
///
/// Key: <c>gcpKey.bin</c> — exactly 32 raw bytes, deployed beside the EXE.
/// </summary>
public static class GcpCrypto
{
    public const int KeyLength   = 32;
    public const int NonceLength = 12;
    public const int TagLength   = 16;

    private static readonly byte[] Magic = "GCP1"u8.ToArray();

    /// <summary>Header size preceding the ciphertext.</summary>
    public static int HeaderLength => Magic.Length + NonceLength + TagLength;

    /// <summary>Encrypts <paramref name="plaintext"/> into the GCP1 envelope.</summary>
    public static byte[] Encrypt(byte[] plaintext, byte[] key)
    {
        ValidateKey(key);
        byte[] nonce = RandomNumberGenerator.GetBytes(NonceLength);
        byte[] tag   = new byte[TagLength];
        byte[] cipher = new byte[plaintext.Length];

        using var gcm = new AesGcm(key, TagLength);
        gcm.Encrypt(nonce, plaintext, cipher, tag);

        byte[] output = new byte[HeaderLength + cipher.Length];
        Magic.CopyTo(output, 0);
        nonce.CopyTo(output, Magic.Length);
        tag.CopyTo(output, Magic.Length + NonceLength);
        cipher.CopyTo(output, HeaderLength);
        return output;
    }

    /// <summary>
    /// Decrypts a GCP1 envelope. Throws <see cref="InvalidDataException"/> on a
    /// malformed envelope and <see cref="CryptographicException"/> on tag mismatch
    /// (wrong key or tampered payload).
    /// </summary>
    public static byte[] Decrypt(byte[] envelope, byte[] key)
    {
        ValidateKey(key);
        if (envelope.Length < HeaderLength || !envelope.AsSpan(0, Magic.Length).SequenceEqual(Magic))
            throw new InvalidDataException("Not a GCP1 encrypted envelope.");

        var nonce  = envelope.AsSpan(Magic.Length, NonceLength);
        var tag    = envelope.AsSpan(Magic.Length + NonceLength, TagLength);
        var cipher = envelope.AsSpan(HeaderLength);
        byte[] plain = new byte[cipher.Length];

        using var gcm = new AesGcm(key, TagLength);
        gcm.Decrypt(nonce, cipher, tag, plain);
        return plain;
    }

    private static void ValidateKey(byte[] key)
    {
        if (key is not { Length: KeyLength })
            throw new ArgumentException($"GCP key must be exactly {KeyLength} bytes.", nameof(key));
    }
}
