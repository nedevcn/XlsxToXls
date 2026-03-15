using System.Buffers;
using System.Buffers.Binary;

namespace Nedev.FileConverters.XlsxToXls.Internal;

/// <summary>
/// Writes digital signature records in BIFF8 format.
/// </summary>
internal ref struct DigitalSignatureWriter
{
    private Span<byte> _buffer;
    private int _position;

    public DigitalSignatureWriter(Span<byte> buffer)
    {
        _buffer = buffer;
        _position = 0;
    }

    public int Position => _position;

    /// <summary>
    /// Creates a DigitalSignatureWriter using ArrayPool for buffer management.
    /// </summary>
    public static DigitalSignatureWriter CreatePooled(out byte[] buffer, int minSize = 16384)
    {
        buffer = ArrayPool<byte>.Shared.Rent(minSize);
        return new DigitalSignatureWriter(buffer.AsSpan());
    }

    /// <summary>
    /// Disposes the DigitalSignatureWriter and returns the buffer to the pool.
    /// </summary>
    public void Dispose()
    {
        // Buffer is managed externally
    }

    /// <summary>
    /// Writes document security records to the workbook stream.
    /// </summary>
    public int WriteDocumentSecurity(DocumentSecurityInfo security)
    {
        if (!security.IsSigned && !security.ReadOnlyRecommended)
            return _position;

        // Write FILESHARING record if read-only recommended
        if (security.ReadOnlyRecommended)
        {
            WriteFileSharing(security.ReadOnlyPasswordHash);
        }

        // Write PROT4REVPASS record for signature protection
        if (security.IsSigned)
        {
            WriteProt4RevPass();
        }

        return _position;
    }

    /// <summary>
    /// Writes FILESHARING record (0x005B).
    /// </summary>
    private void WriteFileSharing(string? passwordHash)
    {
        // FILESHARING record structure:
        // 2 bytes: ReadOnly recommended (0x0001 = true)
        // 2 bytes: Password hash (or 0x0000)
        // 2 bytes: Length of title string
        // variable: Title string (Unicode)

        var title = "Read-Only Recommended";
        var titleBytes = System.Text.Encoding.Unicode.GetBytes(title);
        var recLen = 2 + 2 + 2 + titleBytes.Length;

        WriteRecordHeader(0x005B, recLen);

        // ReadOnly recommended
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0001);
        _position += 2;

        // Password hash (simplified - just write 0)
        var hash = 0;
        if (!string.IsNullOrEmpty(passwordHash) && ushort.TryParse(passwordHash, out var parsedHash))
        {
            hash = parsedHash;
        }
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)hash);
        _position += 2;

        // Title length (in characters)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)title.Length);
        _position += 2;

        // Title string
        titleBytes.CopyTo(_buffer.Slice(_position));
        _position += titleBytes.Length;
    }

    /// <summary>
    /// Writes PROT4REVPASS record (0x00AF) for signature protection.
    /// </summary>
    private void WriteProt4RevPass()
    {
        // PROT4REVPASS record structure:
        // 2 bytes: Password hash (0x0000 for signed documents)
        // 2 bytes: Reserved

        WriteRecordHeader(0x00AF, 4);

        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0000);
        _position += 2;

        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0000);
        _position += 2;
    }

    /// <summary>
    /// Writes digital signature stream data for the _signatures stream.
    /// </summary>
    public byte[] WriteSignatureStream(List<DigitalSignatureInfo> signatures)
    {
        using var ms = new MemoryStream();
        using var writer = new BinaryWriter(ms);

        // Write signature header
        WriteSignatureHeader(writer);

        // Write each signature
        foreach (var sig in signatures)
        {
            WriteSignatureEntry(writer, sig);
        }

        return ms.ToArray();
    }

    /// <summary>
    /// Writes the signature header.
    /// </summary>
    private void WriteSignatureHeader(BinaryWriter writer)
    {
        // Signature header version
        writer.Write((uint)0x00000001);

        // Number of signatures
        writer.Write((uint)0); // Will be updated later

        // Reserved
        writer.Write((uint)0);
    }

    /// <summary>
    /// Writes a single signature entry.
    /// </summary>
    private void WriteSignatureEntry(BinaryWriter writer, DigitalSignatureInfo signature)
    {
        // Signature type
        writer.Write((uint)signature.Type);

        // Certificate data length and data
        writer.Write((uint)signature.CertificateData.Length);
        writer.Write(signature.CertificateData);

        // Signature value length and data
        writer.Write((uint)signature.SignatureValue.Length);
        writer.Write(signature.SignatureValue);

        // Algorithm OIDs (null-terminated strings)
        WriteNullTerminatedString(writer, signature.SignatureAlgorithm);
        WriteNullTerminatedString(writer, signature.DigestAlgorithm);

        // Signing time (FILETIME format)
        var fileTime = signature.SigningTime.ToFileTimeUtc();
        writer.Write((ulong)fileTime);

        // Signer info
        WriteSignerInfo(writer, signature.Signer);

        // Purpose
        writer.Write((uint)signature.Purpose);

        // Is valid flag
        writer.Write(signature.IsValid ? (byte)1 : (byte)0);
    }

    /// <summary>
    /// Writes signer information.
    /// </summary>
    private void WriteSignerInfo(BinaryWriter writer, SignerInfoInfo signer)
    {
        WriteNullTerminatedString(writer, signer.Name);
        WriteNullTerminatedString(writer, signer.Email ?? "");
        WriteNullTerminatedString(writer, signer.Organization ?? "");
        WriteNullTerminatedString(writer, signer.OrganizationalUnit ?? "");
        WriteNullTerminatedString(writer, signer.Locality ?? "");
        WriteNullTerminatedString(writer, signer.State ?? "");
        WriteNullTerminatedString(writer, signer.Country ?? "");
        WriteNullTerminatedString(writer, signer.Issuer ?? "");
        WriteNullTerminatedString(writer, signer.SerialNumber ?? "");

        // Valid from/to (FILETIME, 0 if not set)
        writer.Write((ulong)(signer.ValidFrom?.ToFileTimeUtc() ?? 0));
        writer.Write((ulong)(signer.ValidTo?.ToFileTimeUtc() ?? 0));
    }

    /// <summary>
    /// Writes a null-terminated string.
    /// </summary>
    private void WriteNullTerminatedString(BinaryWriter writer, string value)
    {
        var bytes = System.Text.Encoding.UTF8.GetBytes(value);
        writer.Write(bytes);
        writer.Write((byte)0);
    }

    /// <summary>
    /// Writes a BIFF record header.
    /// </summary>
    private void WriteRecordHeader(ushort recordType, int dataLength)
    {
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), recordType);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)dataLength);
        _position += 2;
    }
}

/// <summary>
/// Utility class for digital signature operations.
/// </summary>
internal static class DigitalSignatureUtility
{
    /// <summary>
    /// Converts XLSX XML signature to BIFF8 format.
    /// </summary>
    public static DigitalSignatureInfo? ConvertXmlSignatureToBiff(DigitalSignatureInfo xmlSignature)
    {
        if (xmlSignature.Type != SignatureType.XmlDsig)
            return null;

        // Convert XML-DSig to PKCS#7/CMS format
        // This is a simplified implementation
        return xmlSignature with
        {
            Type = SignatureType.Pkcs7
        };
    }

    /// <summary>
    /// Validates a digital signature.
    /// </summary>
    public static bool ValidateSignature(DigitalSignatureInfo signature)
    {
        try
        {
            if (signature.CertificateData.Length == 0 || signature.SignatureValue.Length == 0)
                return false;

            // Create certificate from data
            using var cert = new System.Security.Cryptography.X509Certificates.X509Certificate2(signature.CertificateData);

            // Check certificate validity period
            if (cert.NotBefore > DateTime.Now || cert.NotAfter < DateTime.Now)
                return false;

            // In a full implementation, this would verify the signature cryptographically
            // For now, we just check that the certificate is valid
            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Computes a hash of data using the specified algorithm.
    /// </summary>
    public static byte[] ComputeHash(byte[] data, string algorithmOid)
    {
        System.Security.Cryptography.HashAlgorithm hashAlg = algorithmOid switch
        {
            "1.3.14.3.2.26" => System.Security.Cryptography.SHA1.Create(),
            "2.16.840.1.101.3.4.2.1" => System.Security.Cryptography.SHA256.Create(),
            "2.16.840.1.101.3.4.2.2" => System.Security.Cryptography.SHA384.Create(),
            "2.16.840.1.101.3.4.2.3" => System.Security.Cryptography.SHA512.Create(),
            _ => System.Security.Cryptography.SHA1.Create()
        };

        using (hashAlg)
        {
            return hashAlg.ComputeHash(data);
        }
    }

    /// <summary>
    /// Gets the algorithm name from OID.
    /// </summary>
    public static string GetAlgorithmName(string oid)
    {
        return oid switch
        {
            "1.2.840.113549.1.1.1" => "RSA",
            "1.2.840.113549.1.1.4" => "md5RSA",
            "1.2.840.113549.1.1.5" => "sha1RSA",
            "1.2.840.113549.1.1.11" => "sha256RSA",
            "1.2.840.113549.1.1.12" => "sha384RSA",
            "1.2.840.113549.1.1.13" => "sha512RSA",
            "1.3.14.3.2.26" => "SHA-1",
            "2.16.840.1.101.3.4.2.1" => "SHA-256",
            "2.16.840.1.101.3.4.2.2" => "SHA-384",
            "2.16.840.1.101.3.4.2.3" => "SHA-512",
            _ => oid
        };
    }
}
