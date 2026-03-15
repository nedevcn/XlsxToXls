namespace Nedev.FileConverters.XlsxToXls.Internal;

/// <summary>
/// Digital signature information for workbook or VBA project.
/// </summary>
public sealed class DigitalSignature
{
    /// <summary>
    /// Gets or sets the signature type.
    /// </summary>
    public SignatureType Type { get; set; }

    /// <summary>
    /// Gets or sets the signature certificate data.
    /// </summary>
    public byte[] CertificateData { get; set; } = Array.Empty<byte>();

    /// <summary>
    /// Gets or sets the signature value (encrypted hash).
    /// </summary>
    public byte[] SignatureValue { get; set; } = Array.Empty<byte>();

    /// <summary>
    /// Gets or sets the signature algorithm OID.
    /// </summary>
    public string SignatureAlgorithm { get; set; } = "1.2.840.113549.1.1.5"; // sha1RSA

    /// <summary>
    /// Gets or sets the digest algorithm OID.
    /// </summary>
    public string DigestAlgorithm { get; set; } = "1.3.14.3.2.26"; // SHA-1

    /// <summary>
    /// Gets or sets the signer information.
    /// </summary>
    public SignerInfo Signer { get; set; } = new();

    /// <summary>
    /// Gets or sets the signing time.
    /// </summary>
    public DateTime SigningTime { get; set; }

    /// <summary>
    /// Gets or sets the signature comments.
    /// </summary>
    public string? Comments { get; set; }

    /// <summary>
    /// Gets or sets whether the signature is valid.
    /// </summary>
    public bool IsValid { get; set; }

    /// <summary>
    /// Gets or sets the signature purpose.
    /// </summary>
    public SignaturePurpose Purpose { get; set; } = SignaturePurpose.DocumentIntegrity;
}

/// <summary>
/// Signature types supported.
/// </summary>
public enum SignatureType
{
    /// <summary>
    /// XML Digital Signature (XLSX).
    /// </summary>
    XmlDsig,

    /// <summary>
    /// PKCS#7/CMS signature (XLS).
    /// </summary>
    Pkcs7,

    /// <summary>
    /// VBA project signature.
    /// </summary>
    VbaProject
}

/// <summary>
/// Signature purpose.
/// </summary>
public enum SignaturePurpose
{
    /// <summary>
    /// Ensure document integrity (content not modified).
    /// </summary>
    DocumentIntegrity,

    /// <summary>
    /// Authenticate the signer identity.
    /// </summary>
    SignerAuthentication,

    /// <summary>
    /// Both integrity and authentication.
    /// </summary>
    Both
}

/// <summary>
/// Signer information.
/// </summary>
public sealed class SignerInfo
{
    /// <summary>
    /// Gets or sets the signer name.
    /// </summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the signer email.
    /// </summary>
    public string? Email { get; set; }

    /// <summary>
    /// Gets or sets the organization.
    /// </summary>
    public string? Organization { get; set; }

    /// <summary>
    /// Gets or sets the organizational unit.
    /// </summary>
    public string? OrganizationalUnit { get; set; }

    /// <summary>
    /// Gets or sets the locality.
    /// </summary>
    public string? Locality { get; set; }

    /// <summary>
    /// Gets or sets the state or province.
    /// </summary>
    public string? State { get; set; }

    /// <summary>
    /// Gets or sets the country.
    /// </summary>
    public string? Country { get; set; }

    /// <summary>
    /// Gets or sets the certificate issuer.
    /// </summary>
    public string? Issuer { get; set; }

    /// <summary>
    /// Gets or sets the certificate serial number.
    /// </summary>
    public string? SerialNumber { get; set; }

    /// <summary>
    /// Gets or sets the certificate validity start date.
    /// </summary>
    public DateTime? ValidFrom { get; set; }

    /// <summary>
    /// Gets or sets the certificate validity end date.
    /// </summary>
    public DateTime? ValidTo { get; set; }
}

/// <summary>
/// Signature line properties (visible signature placeholder).
/// </summary>
public sealed class SignatureLine
{
    /// <summary>
    /// Gets or sets the signature line identifier.
    /// </summary>
    public string Id { get; set; } = Guid.NewGuid().ToString();

    /// <summary>
    /// Gets or sets the suggested signer name.
    /// </summary>
    public string? SuggestedSigner { get; set; }

    /// <summary>
    /// Gets or sets the suggested signer title.
    /// </summary>
    public string? SuggestedSignerTitle { get; set; }

    /// <summary>
    /// Gets or sets the suggested signer email.
    /// </summary>
    public string? SuggestedSignerEmail { get; set; }

    /// <summary>
    /// Gets or sets the instructions for the signer.
    /// </summary>
    public string? Instructions { get; set; }

    /// <summary>
    /// Gets or sets whether the signature line is signed.
    /// </summary>
    public bool IsSigned { get; set; }

    /// <summary>
    /// Gets or sets whether to show the date.
    /// </summary>
    public bool ShowDate { get; set; } = true;

    /// <summary>
    /// Gets or sets the signature image data.
    /// </summary>
    public byte[]? SignatureImage { get; set; }

    /// <summary>
    /// Gets or sets the signature line position.
    /// </summary>
    public SignatureLinePosition Position { get; set; } = new();
}

/// <summary>
/// Signature line position.
/// </summary>
public sealed class SignatureLinePosition
{
    /// <summary>
    /// Gets or sets the worksheet name.
    /// </summary>
    public string? WorksheetName { get; set; }

    /// <summary>
    /// Gets or sets the cell reference (e.g., "A1").
    /// </summary>
    public string? CellReference { get; set; }

    /// <summary>
    /// Gets or sets the left position in points.
    /// </summary>
    public double Left { get; set; }

    /// <summary>
    /// Gets or sets the top position in points.
    /// </summary>
    public double Top { get; set; }

    /// <summary>
    /// Gets or sets the width in points.
    /// </summary>
    public double Width { get; set; } = 200;

    /// <summary>
    /// Gets or sets the height in points.
    /// </summary>
    public double Height { get; set; } = 60;
}

/// <summary>
/// Document security settings.
/// </summary>
public sealed class DocumentSecurity
{
    /// <summary>
    /// Gets or sets whether the document is signed.
    /// </summary>
    public bool IsSigned { get; set; }

    /// <summary>
    /// Gets or sets whether the document is read-only recommended.
    /// </summary>
    public bool ReadOnlyRecommended { get; set; }

    /// <summary>
    /// Gets or sets the read-only password hash.
    /// </summary>
    public string? ReadOnlyPasswordHash { get; set; }

    /// <summary>
    /// Gets or sets the digital signatures.
    /// </summary>
    public List<DigitalSignature> Signatures { get; set; } = new();

    /// <summary>
    /// Gets or sets the signature lines.
    /// </summary>
    public List<SignatureLine> SignatureLines { get; set; } = new();
}

/// <summary>
/// Internal record for digital signature information.
/// </summary>
internal record struct DigitalSignatureInfo(
    SignatureType Type,
    byte[] CertificateData,
    byte[] SignatureValue,
    string SignatureAlgorithm,
    string DigestAlgorithm,
    SignerInfoInfo Signer,
    DateTime SigningTime,
    bool IsValid,
    SignaturePurpose Purpose);

/// <summary>
/// Internal record for signer information.
/// </summary>
internal record struct SignerInfoInfo(
    string Name,
    string? Email,
    string? Organization,
    string? OrganizationalUnit,
    string? Locality,
    string? State,
    string? Country,
    string? Issuer,
    string? SerialNumber,
    DateTime? ValidFrom,
    DateTime? ValidTo);

/// <summary>
/// Internal record for signature line information.
/// </summary>
internal record struct SignatureLineInfo(
    string Id,
    string? SuggestedSigner,
    string? SuggestedSignerTitle,
    string? SuggestedSignerEmail,
    string? Instructions,
    bool IsSigned,
    bool ShowDate,
    byte[]? SignatureImage,
    SignatureLinePositionInfo Position);

/// <summary>
/// Internal record for signature line position.
/// </summary>
internal record struct SignatureLinePositionInfo(
    string? WorksheetName,
    string? CellReference,
    double Left,
    double Top,
    double Width,
    double Height);

/// <summary>
/// Internal record for document security information.
/// </summary>
internal record struct DocumentSecurityInfo(
    bool IsSigned,
    bool ReadOnlyRecommended,
    string? ReadOnlyPasswordHash,
    List<DigitalSignatureInfo> Signatures,
    List<SignatureLineInfo> SignatureLines);
