<#
.SYNOPSIS
    Create a self-signed code-signing certificate for the RBLclass
    Phase 0 POC.

.DESCRIPTION
    Idempotent. If a certificate with the same subject already exists
    in Cert:\CurrentUser\My, the script reuses it instead of creating
    a duplicate. The certificate carries the Code Signing EKU
    (OID 1.3.6.1.5.5.7.3.3) as required by signtool / the real
    internal-PKI cert will later have.

    The certificate is also installed into Cert:\CurrentUser\Root so
    that Authenticode signatures it produces are trusted on this
    machine without warnings.

.OUTPUTS
    The certificate thumbprint is printed to stdout. Pass it to
    Install-HelloPstAddIn.ps1 via -Thumbprint.
#>

[CmdletBinding()]
param(
    [string] $Subject = "CN=RBLclass POC Dev",
    [int]    $YearsValid = 3
)

$ErrorActionPreference = "Stop"

$existing = Get-ChildItem Cert:\CurrentUser\My |
    Where-Object { $_.Subject -eq $Subject -and $_.HasPrivateKey } |
    Sort-Object NotAfter -Descending |
    Select-Object -First 1

if ($existing) {
    Write-Host "Reusing existing certificate: $($existing.Thumbprint) (expires $($existing.NotAfter))"
    $cert = $existing
} else {
    Write-Host "Creating new self-signed code-signing certificate: $Subject"
    $cert = New-SelfSignedCertificate `
        -Type CodeSigningCert `
        -Subject $Subject `
        -KeyUsage DigitalSignature `
        -KeyAlgorithm RSA -KeyLength 2048 `
        -CertStoreLocation "Cert:\CurrentUser\My" `
        -NotAfter (Get-Date).AddYears($YearsValid) `
        -TextExtension @("2.5.29.37={text}1.3.6.1.5.5.7.3.3")
}

$rootStore = Get-ChildItem Cert:\CurrentUser\Root |
    Where-Object { $_.Thumbprint -eq $cert.Thumbprint }
if (-not $rootStore) {
    Write-Host "Importing certificate into Cert:\CurrentUser\Root for local trust"
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store(
        "Root", "CurrentUser")
    $store.Open("ReadWrite")
    try { $store.Add($cert) } finally { $store.Close() }
} else {
    Write-Host "Certificate already trusted in Cert:\CurrentUser\Root"
}

Write-Host ""
Write-Host "Thumbprint: $($cert.Thumbprint)"
Write-Host ""
Write-Host "Pass this thumbprint to Install-HelloPstAddIn.ps1:"
Write-Host "  .\Install-HelloPstAddIn.ps1 -Thumbprint $($cert.Thumbprint)"
