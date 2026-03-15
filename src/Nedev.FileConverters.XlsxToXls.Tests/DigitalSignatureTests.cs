using System;
using System.Buffers;
using System.Collections.Generic;
using Nedev.FileConverters.XlsxToXls.Internal;
using Xunit;

namespace Nedev.FileConverters.XlsxToXls.Tests;

/// <summary>
/// Tests for digital signature functionality.
/// </summary>
public class DigitalSignatureTests
{
    [Fact]
    public void DigitalSignature_DefaultValues()
    {
        var sig = new DigitalSignature();

        Assert.Equal(SignatureType.XmlDsig, sig.Type);
        Assert.Empty(sig.CertificateData);
        Assert.Empty(sig.SignatureValue);
        Assert.Equal("1.2.840.113549.1.1.5", sig.SignatureAlgorithm);
        Assert.Equal("1.3.14.3.2.26", sig.DigestAlgorithm);
        Assert.NotNull(sig.Signer);
        Assert.Equal(DateTime.MinValue, sig.SigningTime);
        Assert.Null(sig.Comments);
        Assert.False(sig.IsValid);
        Assert.Equal(SignaturePurpose.DocumentIntegrity, sig.Purpose);
    }

    [Fact]
    public void DigitalSignature_WithValues()
    {
        var signer = new SignerInfo
        {
            Name = "Test User",
            Email = "test@example.com",
            Organization = "Test Org"
        };

        var sig = new DigitalSignature
        {
            Type = SignatureType.Pkcs7,
            CertificateData = new byte[] { 1, 2, 3 },
            SignatureValue = new byte[] { 4, 5, 6 },
            SignatureAlgorithm = "1.2.840.113549.1.1.11",
            DigestAlgorithm = "2.16.840.1.101.3.4.2.1",
            Signer = signer,
            SigningTime = new DateTime(2024, 1, 15, 10, 30, 0),
            Comments = "Test signature",
            IsValid = true,
            Purpose = SignaturePurpose.Both
        };

        Assert.Equal(SignatureType.Pkcs7, sig.Type);
        Assert.Equal(new byte[] { 1, 2, 3 }, sig.CertificateData);
        Assert.Equal(new byte[] { 4, 5, 6 }, sig.SignatureValue);
        Assert.Equal("1.2.840.113549.1.1.11", sig.SignatureAlgorithm);
        Assert.Equal("2.16.840.1.101.3.4.2.1", sig.DigestAlgorithm);
        Assert.Equal("Test User", sig.Signer.Name);
        Assert.Equal("test@example.com", sig.Signer.Email);
        Assert.Equal("Test Org", sig.Signer.Organization);
        Assert.Equal(new DateTime(2024, 1, 15, 10, 30, 0), sig.SigningTime);
        Assert.Equal("Test signature", sig.Comments);
        Assert.True(sig.IsValid);
        Assert.Equal(SignaturePurpose.Both, sig.Purpose);
    }

    [Fact]
    public void SignerInfo_DefaultValues()
    {
        var signer = new SignerInfo();

        Assert.Empty(signer.Name);
        Assert.Null(signer.Email);
        Assert.Null(signer.Organization);
        Assert.Null(signer.OrganizationalUnit);
        Assert.Null(signer.Locality);
        Assert.Null(signer.State);
        Assert.Null(signer.Country);
        Assert.Null(signer.Issuer);
        Assert.Null(signer.SerialNumber);
        Assert.Null(signer.ValidFrom);
        Assert.Null(signer.ValidTo);
    }

    [Fact]
    public void SignerInfo_WithValues()
    {
        var signer = new SignerInfo
        {
            Name = "John Doe",
            Email = "john@example.com",
            Organization = "Acme Corp",
            OrganizationalUnit = "Engineering",
            Locality = "San Francisco",
            State = "CA",
            Country = "US",
            Issuer = "DigiCert Inc",
            SerialNumber = "1234567890",
            ValidFrom = new DateTime(2023, 1, 1),
            ValidTo = new DateTime(2024, 12, 31)
        };

        Assert.Equal("John Doe", signer.Name);
        Assert.Equal("john@example.com", signer.Email);
        Assert.Equal("Acme Corp", signer.Organization);
        Assert.Equal("Engineering", signer.OrganizationalUnit);
        Assert.Equal("San Francisco", signer.Locality);
        Assert.Equal("CA", signer.State);
        Assert.Equal("US", signer.Country);
        Assert.Equal("DigiCert Inc", signer.Issuer);
        Assert.Equal("1234567890", signer.SerialNumber);
        Assert.Equal(new DateTime(2023, 1, 1), signer.ValidFrom);
        Assert.Equal(new DateTime(2024, 12, 31), signer.ValidTo);
    }

    [Fact]
    public void SignatureLine_DefaultValues()
    {
        var line = new SignatureLine();

        Assert.NotNull(line.Id);
        Assert.NotEmpty(line.Id);
        Assert.Null(line.SuggestedSigner);
        Assert.Null(line.SuggestedSignerTitle);
        Assert.Null(line.SuggestedSignerEmail);
        Assert.Null(line.Instructions);
        Assert.False(line.IsSigned);
        Assert.True(line.ShowDate);
        Assert.Null(line.SignatureImage);
        Assert.NotNull(line.Position);
    }

    [Fact]
    public void SignatureLine_WithValues()
    {
        var position = new SignatureLinePosition
        {
            WorksheetName = "Sheet1",
            CellReference = "A1",
            Left = 10.5,
            Top = 20.0,
            Width = 200.0,
            Height = 60.0
        };

        var line = new SignatureLine
        {
            Id = "sig-line-1",
            SuggestedSigner = "Jane Doe",
            SuggestedSignerTitle = "Manager",
            SuggestedSignerEmail = "jane@example.com",
            Instructions = "Please sign here",
            IsSigned = true,
            ShowDate = false,
            SignatureImage = new byte[] { 1, 2, 3 },
            Position = position
        };

        Assert.Equal("sig-line-1", line.Id);
        Assert.Equal("Jane Doe", line.SuggestedSigner);
        Assert.Equal("Manager", line.SuggestedSignerTitle);
        Assert.Equal("jane@example.com", line.SuggestedSignerEmail);
        Assert.Equal("Please sign here", line.Instructions);
        Assert.True(line.IsSigned);
        Assert.False(line.ShowDate);
        Assert.Equal(new byte[] { 1, 2, 3 }, line.SignatureImage);
        Assert.Equal("Sheet1", line.Position.WorksheetName);
        Assert.Equal("A1", line.Position.CellReference);
        Assert.Equal(10.5, line.Position.Left);
        Assert.Equal(20.0, line.Position.Top);
        Assert.Equal(200.0, line.Position.Width);
        Assert.Equal(60.0, line.Position.Height);
    }

    [Fact]
    public void SignatureLinePosition_DefaultValues()
    {
        var pos = new SignatureLinePosition();

        Assert.Null(pos.WorksheetName);
        Assert.Null(pos.CellReference);
        Assert.Equal(0.0, pos.Left);
        Assert.Equal(0.0, pos.Top);
        Assert.Equal(200.0, pos.Width);
        Assert.Equal(60.0, pos.Height);
    }

    [Fact]
    public void DocumentSecurity_DefaultValues()
    {
        var security = new DocumentSecurity();

        Assert.False(security.IsSigned);
        Assert.False(security.ReadOnlyRecommended);
        Assert.Null(security.ReadOnlyPasswordHash);
        Assert.Empty(security.Signatures);
        Assert.Empty(security.SignatureLines);
    }

    [Fact]
    public void DocumentSecurity_WithValues()
    {
        var signatures = new List<DigitalSignature>
        {
            new DigitalSignature { Type = SignatureType.XmlDsig },
            new DigitalSignature { Type = SignatureType.Pkcs7 }
        };

        var lines = new List<SignatureLine>
        {
            new SignatureLine { SuggestedSigner = "Signer 1" },
            new SignatureLine { SuggestedSigner = "Signer 2" }
        };

        var security = new DocumentSecurity
        {
            IsSigned = true,
            ReadOnlyRecommended = true,
            ReadOnlyPasswordHash = "abc123",
            Signatures = signatures,
            SignatureLines = lines
        };

        Assert.True(security.IsSigned);
        Assert.True(security.ReadOnlyRecommended);
        Assert.Equal("abc123", security.ReadOnlyPasswordHash);
        Assert.Equal(2, security.Signatures.Count);
        Assert.Equal(2, security.SignatureLines.Count);
    }

    [Fact]
    public void DigitalSignatureInfo_Create()
    {
        var signer = new SignerInfoInfo(
            "Test User",
            "test@example.com",
            null, null, null, null, null, null, null, null, null);

        var sig = new DigitalSignatureInfo(
            SignatureType.Pkcs7,
            new byte[] { 1, 2, 3 },
            new byte[] { 4, 5, 6 },
            "1.2.840.113549.1.1.5",
            "1.3.14.3.2.26",
            signer,
            DateTime.Now,
            true,
            SignaturePurpose.DocumentIntegrity);

        Assert.Equal(SignatureType.Pkcs7, sig.Type);
        Assert.Equal("Test User", sig.Signer.Name);
        Assert.True(sig.IsValid);
    }

    [Fact]
    public void SignerInfoInfo_Create()
    {
        var signer = new SignerInfoInfo(
            "John Doe",
            "john@example.com",
            "Acme Corp",
            "Engineering",
            "San Francisco",
            "CA",
            "US",
            "DigiCert",
            "12345",
            new DateTime(2023, 1, 1),
            new DateTime(2024, 12, 31));

        Assert.Equal("John Doe", signer.Name);
        Assert.Equal("john@example.com", signer.Email);
        Assert.Equal("Acme Corp", signer.Organization);
        Assert.Equal("Engineering", signer.OrganizationalUnit);
        Assert.Equal(new DateTime(2023, 1, 1), signer.ValidFrom);
        Assert.Equal(new DateTime(2024, 12, 31), signer.ValidTo);
    }

    [Fact]
    public void SignatureLineInfo_Create()
    {
        var position = new SignatureLinePositionInfo("Sheet1", "A1", 10.0, 20.0, 200.0, 60.0);

        var line = new SignatureLineInfo(
            "sig-1",
            "Jane Doe",
            "Manager",
            "jane@example.com",
            "Sign here",
            true,
            true,
            null,
            position);

        Assert.Equal("sig-1", line.Id);
        Assert.Equal("Jane Doe", line.SuggestedSigner);
        Assert.True(line.IsSigned);
        Assert.Equal("Sheet1", line.Position.WorksheetName);
    }

    [Fact]
    public void DocumentSecurityInfo_Create()
    {
        var signatures = new List<DigitalSignatureInfo>();
        var lines = new List<SignatureLineInfo>();

        var security = new DocumentSecurityInfo(
            true,
            true,
            "password123",
            signatures,
            lines);

        Assert.True(security.IsSigned);
        Assert.True(security.ReadOnlyRecommended);
        Assert.Equal("password123", security.ReadOnlyPasswordHash);
    }

    [Theory]
    [InlineData(SignatureType.XmlDsig)]
    [InlineData(SignatureType.Pkcs7)]
    [InlineData(SignatureType.VbaProject)]
    public void SignatureType_EnumValues(SignatureType type)
    {
        // Just verify enum values can be assigned
        var sig = new DigitalSignature { Type = type };
        Assert.Equal(type, sig.Type);
    }

    [Theory]
    [InlineData(SignaturePurpose.DocumentIntegrity)]
    [InlineData(SignaturePurpose.SignerAuthentication)]
    [InlineData(SignaturePurpose.Both)]
    public void SignaturePurpose_EnumValues(SignaturePurpose purpose)
    {
        var sig = new DigitalSignature { Purpose = purpose };
        Assert.Equal(purpose, sig.Purpose);
    }

    [Fact]
    public void DigitalSignatureUtility_GetAlgorithmName()
    {
        Assert.Equal("RSA", DigitalSignatureUtility.GetAlgorithmName("1.2.840.113549.1.1.1"));
        Assert.Equal("md5RSA", DigitalSignatureUtility.GetAlgorithmName("1.2.840.113549.1.1.4"));
        Assert.Equal("sha1RSA", DigitalSignatureUtility.GetAlgorithmName("1.2.840.113549.1.1.5"));
        Assert.Equal("sha256RSA", DigitalSignatureUtility.GetAlgorithmName("1.2.840.113549.1.1.11"));
        Assert.Equal("sha384RSA", DigitalSignatureUtility.GetAlgorithmName("1.2.840.113549.1.1.12"));
        Assert.Equal("sha512RSA", DigitalSignatureUtility.GetAlgorithmName("1.2.840.113549.1.1.13"));
        Assert.Equal("SHA-1", DigitalSignatureUtility.GetAlgorithmName("1.3.14.3.2.26"));
        Assert.Equal("SHA-256", DigitalSignatureUtility.GetAlgorithmName("2.16.840.1.101.3.4.2.1"));
        Assert.Equal("SHA-384", DigitalSignatureUtility.GetAlgorithmName("2.16.840.1.101.3.4.2.2"));
        Assert.Equal("SHA-512", DigitalSignatureUtility.GetAlgorithmName("2.16.840.1.101.3.4.2.3"));
        Assert.Equal("unknown.oid", DigitalSignatureUtility.GetAlgorithmName("unknown.oid"));
    }

    [Fact]
    public void DigitalSignatureUtility_ComputeHash_SHA1()
    {
        var data = System.Text.Encoding.UTF8.GetBytes("Hello, World!");
        var hash = DigitalSignatureUtility.ComputeHash(data, "1.3.14.3.2.26");

        Assert.NotNull(hash);
        Assert.Equal(20, hash.Length); // SHA-1 produces 20 bytes
    }

    [Fact]
    public void DigitalSignatureUtility_ComputeHash_SHA256()
    {
        var data = System.Text.Encoding.UTF8.GetBytes("Hello, World!");
        var hash = DigitalSignatureUtility.ComputeHash(data, "2.16.840.1.101.3.4.2.1");

        Assert.NotNull(hash);
        Assert.Equal(32, hash.Length); // SHA-256 produces 32 bytes
    }

    [Fact]
    public void DigitalSignatureUtility_ComputeHash_Default()
    {
        var data = System.Text.Encoding.UTF8.GetBytes("Hello, World!");
        var hash = DigitalSignatureUtility.ComputeHash(data, "unknown.oid");

        Assert.NotNull(hash);
        Assert.Equal(20, hash.Length); // Default is SHA-1
    }

    [Fact]
    public void DigitalSignatureUtility_ValidateSignature_EmptyData()
    {
        var sig = new DigitalSignatureInfo(
            SignatureType.Pkcs7,
            Array.Empty<byte>(),
            Array.Empty<byte>(),
            "1.2.840.113549.1.1.5",
            "1.3.14.3.2.26",
            new SignerInfoInfo(string.Empty, null, null, null, null, null, null, null, null, null, null),
            DateTime.Now,
            false,
            SignaturePurpose.DocumentIntegrity);

        var isValid = DigitalSignatureUtility.ValidateSignature(sig);
        Assert.False(isValid);
    }

    [Fact]
    public void DigitalSignatureWriter_CreatePooled()
    {
        var writer = DigitalSignatureWriter.CreatePooled(out var buffer, 4096);

        Assert.NotNull(buffer);
        Assert.True(buffer.Length >= 4096);
        Assert.Equal(0, writer.Position);

        ArrayPool<byte>.Shared.Return(buffer);
    }

    [Fact]
    public void DigitalSignatureWriter_WriteDocumentSecurity_Empty()
    {
        var writer = DigitalSignatureWriter.CreatePooled(out var buffer, 4096);
        try
        {
            var security = new DocumentSecurityInfo(
                false, false, null,
                new List<DigitalSignatureInfo>(),
                new List<SignatureLineInfo>());

            var written = writer.WriteDocumentSecurity(security);

            // Should not write anything for empty security
            Assert.Equal(0, written);
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(buffer);
        }
    }

    [Fact]
    public void DigitalSignatureWriter_WriteDocumentSecurity_ReadOnly()
    {
        var writer = DigitalSignatureWriter.CreatePooled(out var buffer, 4096);
        try
        {
            var security = new DocumentSecurityInfo(
                false, true, null,
                new List<DigitalSignatureInfo>(),
                new List<SignatureLineInfo>());

            var written = writer.WriteDocumentSecurity(security);

            // Should write FILESHARING record
            Assert.True(written > 0);
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(buffer);
        }
    }

    [Fact]
    public void DigitalSignatureWriter_WriteDocumentSecurity_Signed()
    {
        var writer = DigitalSignatureWriter.CreatePooled(out var buffer, 4096);
        try
        {
            var security = new DocumentSecurityInfo(
                true, false, null,
                new List<DigitalSignatureInfo>(),
                new List<SignatureLineInfo>());

            var written = writer.WriteDocumentSecurity(security);

            // Should write PROT4REVPASS record
            Assert.True(written > 0);
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(buffer);
        }
    }

    [Fact]
    public void DigitalSignatureWriter_WriteSignatureStream()
    {
        var writer = DigitalSignatureWriter.CreatePooled(out var buffer, 4096);
        try
        {
            var signatures = new List<DigitalSignatureInfo>
            {
                new DigitalSignatureInfo(
                    SignatureType.Pkcs7,
                    new byte[] { 1, 2, 3 },
                    new byte[] { 4, 5, 6 },
                    "1.2.840.113549.1.1.5",
                    "1.3.14.3.2.26",
                    new SignerInfoInfo("Test", null, null, null, null, null, null, null, null, null, null),
                    DateTime.Now,
                    true,
                    SignaturePurpose.DocumentIntegrity)
            };

            var streamData = writer.WriteSignatureStream(signatures);

            Assert.NotNull(streamData);
            Assert.True(streamData.Length > 0);
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(buffer);
        }
    }

    [Fact]
    public void DigitalSignatureUtility_ConvertXmlSignatureToBiff_NonXml()
    {
        var sig = new DigitalSignatureInfo(
            SignatureType.Pkcs7,
            new byte[] { 1 },
            new byte[] { 2 },
            "1.2.840.113549.1.1.5",
            "1.3.14.3.2.26",
            new SignerInfoInfo(string.Empty, null, null, null, null, null, null, null, null, null, null),
            DateTime.Now,
            false,
            SignaturePurpose.DocumentIntegrity);

        var converted = DigitalSignatureUtility.ConvertXmlSignatureToBiff(sig);

        Assert.Null(converted);
    }

    [Fact]
    public void DigitalSignatureUtility_ConvertXmlSignatureToBiff_Xml()
    {
        var sig = new DigitalSignatureInfo(
            SignatureType.XmlDsig,
            new byte[] { 1 },
            new byte[] { 2 },
            "1.2.840.113549.1.1.5",
            "1.3.14.3.2.26",
            new SignerInfoInfo(string.Empty, null, null, null, null, null, null, null, null, null, null),
            DateTime.Now,
            false,
            SignaturePurpose.DocumentIntegrity);

        var converted = DigitalSignatureUtility.ConvertXmlSignatureToBiff(sig);

        Assert.NotNull(converted);
        Assert.Equal(SignatureType.Pkcs7, converted.Value.Type);
    }
}
