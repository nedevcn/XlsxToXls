using System.IO.Compression;
using System.Xml;

namespace Nedev.FileConverters.XlsxToXls.Internal;

/// <summary>
/// Reads digital signatures from XLSX files.
/// </summary>
internal static class DigitalSignatureReader
{
    /// <summary>
    /// Reads document security settings including digital signatures.
    /// </summary>
    public static DocumentSecurityInfo ReadDocumentSecurity(ZipArchive archive, Action<string>? log)
    {
        var signatures = new List<DigitalSignatureInfo>();
        var signatureLines = new List<SignatureLineInfo>();
        var isSigned = false;
        var readOnlyRecommended = false;
        string? readOnlyPasswordHash = null;

        try
        {
            // Read workbook properties for security settings
            var securityFromProps = ReadWorkbookProperties(archive, log);
            readOnlyRecommended = securityFromProps.ReadOnlyRecommended;
            readOnlyPasswordHash = securityFromProps.ReadOnlyPasswordHash;

            // Read XML signatures from xl/signatures.xml
            var xmlSignatures = ReadXmlSignatures(archive, log);
            signatures.AddRange(xmlSignatures);

            // Read signature lines from worksheets
            var lines = ReadSignatureLines(archive, log);
            signatureLines.AddRange(lines);

            isSigned = signatures.Count > 0;

            log?.Invoke($"[DigitalSignatureReader] Found {signatures.Count} signatures, {signatureLines.Count} signature lines");
        }
        catch (Exception ex)
        {
            log?.Invoke($"[DigitalSignatureReader] Error reading document security: {ex.Message}");
        }

        return new DocumentSecurityInfo(
            isSigned,
            readOnlyRecommended,
            readOnlyPasswordHash,
            signatures,
            signatureLines);
    }

    /// <summary>
    /// Reads workbook properties for security settings.
    /// </summary>
    private static (bool ReadOnlyRecommended, string? ReadOnlyPasswordHash) ReadWorkbookProperties(ZipArchive archive, Action<string>? log)
    {
        try
        {
            var entry = archive.GetEntry("xl/workbook.xml");
            if (entry == null) return (false, null);

            using var stream = entry.Open();
            using var reader = XmlReader.Create(stream);
            var ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "fileSharing" && reader.NamespaceURI == ns)
                {
                    var readOnlyRec = reader.GetAttribute("readOnlyRecommended") == "1";
                    var passwordHash = reader.GetAttribute("userName");
                    return (readOnlyRec, passwordHash);
                }
            }
        }
        catch (Exception ex)
        {
            log?.Invoke($"[DigitalSignatureReader] Error reading workbook properties: {ex.Message}");
        }

        return (false, null);
    }

    /// <summary>
    /// Reads XML digital signatures from xl/signatures.xml.
    /// </summary>
    private static List<DigitalSignatureInfo> ReadXmlSignatures(ZipArchive archive, Action<string>? log)
    {
        var signatures = new List<DigitalSignatureInfo>();

        try
        {
            // Check for signature origin
            var originEntry = archive.GetEntry("xl/_rels/origin.sigs.rels");
            if (originEntry == null)
            {
                log?.Invoke("[DigitalSignatureReader] No signature origin found");
                return signatures;
            }

            // Read signature relationships
            var sigPaths = new List<string>();
            using (var stream = originEntry.Open())
            using (var reader = XmlReader.Create(stream))
            {
                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "Relationship")
                    {
                        var target = reader.GetAttribute("Target");
                        if (!string.IsNullOrEmpty(target))
                        {
                            sigPaths.Add($"xl/signatures/{target}");
                        }
                    }
                }
            }

            // Read each signature
            foreach (var sigPath in sigPaths)
            {
                var sigEntry = archive.GetEntry(sigPath);
                if (sigEntry == null) continue;

                var signature = ReadSignatureXml(sigEntry, log);
                if (signature.HasValue)
                {
                    signatures.Add(signature.Value);
                }
            }
        }
        catch (Exception ex)
        {
            log?.Invoke($"[DigitalSignatureReader] Error reading XML signatures: {ex.Message}");
        }

        return signatures;
    }

    /// <summary>
    /// Reads a single signature XML.
    /// </summary>
    private static DigitalSignatureInfo? ReadSignatureXml(ZipArchiveEntry entry, Action<string>? log)
    {
        try
        {
            using var stream = entry.Open();
            using var reader = XmlReader.Create(stream);

            byte[]? certificateData = null;
            byte[]? signatureValue = null;
            var signatureAlgorithm = "1.2.840.113549.1.1.5"; // sha1RSA default
            var digestAlgorithm = "1.3.14.3.2.26"; // SHA-1 default
            SignerInfoInfo signer = new(string.Empty, null, null, null, null, null, null, null, null, null, null);
            var signingTime = DateTime.MinValue;
            var isValid = false;

            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    switch (reader.LocalName)
                    {
                        case "X509Certificate":
                            var certBase64 = reader.ReadElementContentAsString();
                            if (!string.IsNullOrEmpty(certBase64))
                            {
                                certificateData = Convert.FromBase64String(certBase64);
                            }
                            break;

                        case "SignatureValue":
                            var sigBase64 = reader.ReadElementContentAsString();
                            if (!string.IsNullOrEmpty(sigBase64))
                            {
                                signatureValue = Convert.FromBase64String(sigBase64);
                            }
                            break;

                        case "SigningTime":
                            var timeStr = reader.ReadElementContentAsString();
                            if (DateTime.TryParse(timeStr, out var parsedTime))
                            {
                                signingTime = parsedTime;
                            }
                            break;

                        case "X509IssuerName":
                            var issuer = reader.ReadElementContentAsString();
                            signer = signer with { Issuer = issuer };
                            break;

                        case "X509SerialNumber":
                            var serial = reader.ReadElementContentAsString();
                            signer = signer with { SerialNumber = serial };
                            break;
                    }
                }
            }

            if (certificateData == null || signatureValue == null)
            {
                log?.Invoke("[DigitalSignatureReader] Signature missing certificate or value");
                return null;
            }

            // Try to extract signer info from certificate
            signer = ExtractSignerInfoFromCertificate(certificateData, signer);

            return new DigitalSignatureInfo(
                SignatureType.XmlDsig,
                certificateData,
                signatureValue,
                signatureAlgorithm,
                digestAlgorithm,
                signer,
                signingTime,
                isValid,
                SignaturePurpose.DocumentIntegrity);
        }
        catch (Exception ex)
        {
            log?.Invoke($"[DigitalSignatureReader] Error reading signature XML: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Extracts signer information from certificate data.
    /// </summary>
    private static SignerInfoInfo ExtractSignerInfoFromCertificate(byte[] certificateData, SignerInfoInfo existing)
    {
        try
        {
            // Parse X.509 certificate to extract subject information
            // This is a simplified implementation
            var cert = new System.Security.Cryptography.X509Certificates.X509Certificate2(certificateData);

            var subject = cert.Subject;
            var name = cert.GetNameInfo(System.Security.Cryptography.X509Certificates.X509NameType.SimpleName, false);

            // Parse subject components
            string? email = null;
            string? organization = null;
            string? orgUnit = null;
            string? locality = null;
            string? state = null;
            string? country = null;

            var parts = subject.Split(',');
            foreach (var part in parts)
            {
                var trimmed = part.Trim();
                if (trimmed.StartsWith("E=") || trimmed.StartsWith("EMAILADDRESS="))
                {
                    email = trimmed.Substring(trimmed.IndexOf('=') + 1);
                }
                else if (trimmed.StartsWith("O="))
                {
                    organization = trimmed.Substring(2);
                }
                else if (trimmed.StartsWith("OU="))
                {
                    orgUnit = trimmed.Substring(3);
                }
                else if (trimmed.StartsWith("L="))
                {
                    locality = trimmed.Substring(2);
                }
                else if (trimmed.StartsWith("ST=") || trimmed.StartsWith("S="))
                {
                    state = trimmed.Substring(trimmed.IndexOf('=') + 1);
                }
                else if (trimmed.StartsWith("C="))
                {
                    country = trimmed.Substring(2);
                }
            }

            return existing with
            {
                Name = name ?? existing.Name,
                Email = email ?? existing.Email,
                Organization = organization ?? existing.Organization,
                OrganizationalUnit = orgUnit ?? existing.OrganizationalUnit,
                Locality = locality ?? existing.Locality,
                State = state ?? existing.State,
                Country = country ?? existing.Country,
                ValidFrom = cert.NotBefore,
                ValidTo = cert.NotAfter
            };
        }
        catch
        {
            // If certificate parsing fails, return existing info
            return existing;
        }
    }

    /// <summary>
    /// Reads signature lines from worksheets.
    /// </summary>
    private static List<SignatureLineInfo> ReadSignatureLines(ZipArchive archive, Action<string>? log)
    {
        var lines = new List<SignatureLineInfo>();

        try
        {
            // Signature lines are stored in worksheet drawings
            // This is a simplified implementation
            var drawingEntries = archive.Entries.Where(e => e.FullName.StartsWith("xl/drawings/") && e.FullName.EndsWith(".xml"));

            foreach (var entry in drawingEntries)
            {
                using var stream = entry.Open();
                using var reader = XmlReader.Create(stream);

                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "sp")
                    {
                        // Check if this is a signature line shape
                        var signatureLine = ReadSignatureLineShape(reader, log);
                        if (signatureLine.HasValue)
                        {
                            lines.Add(signatureLine.Value);
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            log?.Invoke($"[DigitalSignatureReader] Error reading signature lines: {ex.Message}");
        }

        return lines;
    }

    /// <summary>
    /// Reads a signature line shape from drawing XML.
    /// </summary>
    private static SignatureLineInfo? ReadSignatureLineShape(XmlReader reader, Action<string>? log)
    {
        try
        {
            var id = Guid.NewGuid().ToString();
            string? signer = null;
            string? title = null;
            string? email = null;
            string? instructions = null;
            var isSigned = false;
            var showDate = true;

            // Read signature line properties
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "sp")
                    break;

                if (reader.NodeType == XmlNodeType.Element)
                {
                    if (reader.LocalName == "signatureLine")
                    {
                        var isSignedAttr = reader.GetAttribute("isSigned");
                        isSigned = isSignedAttr == "1";

                        var showDateAttr = reader.GetAttribute("showSignDate");
                        showDate = showDateAttr != "0";

                        signer = reader.GetAttribute("suggestedSigner");
                        title = reader.GetAttribute("suggestedSignerTitle");
                        email = reader.GetAttribute("suggestedSignerEmail");
                    }
                }
            }

            // Only return if this is actually a signature line
            if (signer != null || title != null)
            {
                return new SignatureLineInfo(
                    id,
                    signer,
                    title,
                    email,
                    instructions,
                    isSigned,
                    showDate,
                    null,
                    new SignatureLinePositionInfo(null, null, 0, 0, 200, 60));
            }

            return null;
        }
        catch (Exception ex)
        {
            log?.Invoke($"[DigitalSignatureReader] Error reading signature line shape: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Reads VBA project signature if present.
    /// </summary>
    public static DigitalSignatureInfo? ReadVbaProjectSignature(ZipArchive archive, Action<string>? log)
    {
        try
        {
            // VBA project signature is in xl/vbaProjectSignature.bin
            var sigEntry = archive.GetEntry("xl/vbaProjectSignature.bin");
            if (sigEntry == null)
            {
                return null;
            }

            byte[] signatureData;
            using (var stream = sigEntry.Open())
            using (var ms = new MemoryStream())
            {
                stream.CopyTo(ms);
                signatureData = ms.ToArray();
            }

            // VBA signatures are PKCS#7/CMS format
            // This is a simplified implementation
            log?.Invoke($"[DigitalSignatureReader] Found VBA project signature ({signatureData.Length} bytes)");

            return new DigitalSignatureInfo(
                SignatureType.VbaProject,
                Array.Empty<byte>(),
                signatureData,
                "1.2.840.113549.1.7.2", // PKCS#7 signed data
                "1.3.14.3.2.26", // SHA-1
                new SignerInfoInfo(string.Empty, null, null, null, null, null, null, null, null, null, null),
                DateTime.MinValue,
                false,
                SignaturePurpose.DocumentIntegrity);
        }
        catch (Exception ex)
        {
            log?.Invoke($"[DigitalSignatureReader] Error reading VBA signature: {ex.Message}");
            return null;
        }
    }
}
