<#
.SYNOPSIS
  Report how an encrypted (S/MIME) mail is built, without decrypting it.

.DESCRIPTION
  Investigation helper for GH #62 (experimental re-encryption, step 0).
  Reads a .msg file directly as an OLE compound file (Outlook and any
  encryption add-in are NOT involved), extracts its smime.p7m / smime.p7s
  part, and decodes only the CMS envelope: content type, algorithms and the
  certificate(s) it is addressed to. Nothing is decrypted: no subject, body or
  attachment content is read or printed.

  It then looks in the current user's certificate store for a certificate
  matching each recipient, which tells whether Windows can see the encryption
  certificate (needed to re-encrypt).

  To get a .msg: in Outlook, with the smart card REMOVED (so the mail stays
  encrypted), drag the encrypted mail from the message list into a folder in
  File Explorer.

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File .\Inspect-SmimeMsg.ps1 -Path "C:\Temp\mail.msg"
#>
param(
    [Parameter(Mandatory = $true)][string]$Path
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Security

if (-not ('RblMsgReader' -as [type])) {
Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using IStream = System.Runtime.InteropServices.ComTypes.IStream;
using STATSTG = System.Runtime.InteropServices.ComTypes.STATSTG;

public static class RblMsgReader
{
    [ComImport, Guid("0000000b-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IStorage
    {
        void CreateStream(string name, uint mode, uint r1, uint r2, out IStream stm);
        void OpenStream(string name, IntPtr r1, uint mode, uint r2, out IStream stm);
        void CreateStorage(string name, uint mode, uint r1, uint r2, out IStorage stg);
        void OpenStorage(string name, IStorage priority, uint mode, IntPtr exclude, uint r, out IStorage stg);
        void CopyTo(uint ciidExclude, IntPtr rgiidExclude, IntPtr snbExclude, IStorage dest);
        void MoveElementTo(string name, IStorage dest, string newName, uint flags);
        void Commit(uint flags);
        void Revert();
        void EnumElements(uint r1, IntPtr r2, uint r3, out IEnumSTATSTG e);
        void DestroyElement(string name);
        void RenameElement(string oldName, string newName);
        void SetElementTimes(string name, IntPtr c, IntPtr a, IntPtr m);
        void SetClass(ref Guid clsid);
        void SetStateBits(uint bits, uint mask);
        void Stat(out STATSTG stat, uint flag);
    }

    [ComImport, Guid("0000000d-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IEnumSTATSTG
    {
        [PreserveSig] int Next(uint celt, [Out, MarshalAs(UnmanagedType.LPArray)] STATSTG[] rgelt, out uint fetched);
        void Skip(uint celt);
        void Reset();
        void Clone(out IEnumSTATSTG e);
    }

    [DllImport("ole32.dll")]
    static extern int StgOpenStorage([MarshalAs(UnmanagedType.LPWStr)] string name, IStorage priority,
                                     uint mode, IntPtr exclude, uint reserved, out IStorage stg);

    const uint STGM_READ = 0x0, STGM_SHARE_DENY_WRITE = 0x20, STGM_SHARE_EXCLUSIVE = 0x10;

    static byte[] ReadStream(IStorage stg, string name)
    {
        IStream s;
        try { stg.OpenStream(name, IntPtr.Zero, STGM_READ | STGM_SHARE_EXCLUSIVE, 0, out s); }
        catch { return null; }
        try
        {
            STATSTG st; s.Stat(out st, 1);
            var buf = new byte[st.cbSize];
            s.Read(buf, buf.Length, IntPtr.Zero);
            return buf;
        }
        finally { Marshal.ReleaseComObject(s); }
    }

    static string ReadString(IStorage stg, string tag)
    {
        var b = ReadStream(stg, "__substg1.0_" + tag + "001F");
        if (b != null) return System.Text.Encoding.Unicode.GetString(b).TrimEnd('\0');
        b = ReadStream(stg, "__substg1.0_" + tag + "001E");
        return b == null ? null : System.Text.Encoding.Default.GetString(b).TrimEnd('\0');
    }

    public sealed class Part { public string FileName; public string MimeTag; public byte[] Data; }

    public static string MessageClass;

    public static List<Part> Read(string path)
    {
        IStorage root;
        int hr = StgOpenStorage(path, null, STGM_READ | STGM_SHARE_DENY_WRITE, IntPtr.Zero, 0, out root);
        if (hr != 0) throw new COMException("Not an Outlook .msg file (StgOpenStorage failed)", hr);
        var parts = new List<Part>();
        try
        {
            MessageClass = ReadString(root, "001A");
            for (int i = 0; i < 64; i++)
            {
                IStorage att;
                try { root.OpenStorage("__attach_version1.0_#" + i.ToString("X8"), null, STGM_READ | STGM_SHARE_EXCLUSIVE, IntPtr.Zero, 0, out att); }
                catch { break; }
                try
                {
                    parts.Add(new Part {
                        FileName = ReadString(att, "3707") ?? ReadString(att, "3704"),
                        MimeTag = ReadString(att, "370E"),
                        Data = ReadStream(att, "__substg1.0_37010102") });
                }
                finally { Marshal.ReleaseComObject(att); }
            }
        }
        finally { Marshal.ReleaseComObject(root); }
        return parts;
    }
}
"@
}

function Describe-Oid($oid) {
    $names = @{
        '1.2.840.113549.1.7.1' = 'data'; '1.2.840.113549.1.7.2' = 'signedData'; '1.2.840.113549.1.7.3' = 'envelopedData'
        '1.2.840.113549.1.1.1' = 'RSA (PKCS#1 v1.5)'; '1.2.840.113549.1.1.7' = 'RSA-OAEP'
        '2.16.840.1.101.3.4.1.2' = 'AES-128-CBC'; '2.16.840.1.101.3.4.1.22' = 'AES-192-CBC'; '2.16.840.1.101.3.4.1.42' = 'AES-256-CBC'
        '1.2.840.113549.3.7' = '3DES-CBC'; '1.2.840.113549.3.2' = 'RC2-CBC'
        '2.16.840.1.101.3.4.2.1' = 'SHA-256'; '1.3.14.3.2.26' = 'SHA-1'; '1.3.132.1.11.1' = 'ECDH (SHA-256 KDF)'
    }
    $n = $names[$oid.Value]; if (-not $n) { $n = $oid.FriendlyName }
    "$($oid.Value)$(if ($n) { " ($n)" })"
}

$full = (Resolve-Path $Path).Path
"File            : $full"
$parts = [RblMsgReader]::Read($full)
"Message class   : $([RblMsgReader]::MessageClass)"
"Parts           : $($parts.Count)"
foreach ($p in $parts) { "  - name='$($p.FileName)' mime='$($p.MimeTag)' size=$(if ($p.Data) { $p.Data.Length } else { 0 })" }

$payload = $parts | Where-Object { $_.FileName -match '^smime\.p7[ms]$' -or $_.MimeTag -match 'pkcs7' } | Select-Object -First 1
if (-not $payload -or -not $payload.Data) {
    ""
    "RESULT: no S/MIME payload part in this .msg."
    "        If the mail is encrypted, it was probably saved while decrypted (smart card inserted):"
    "        remove the card, restart Outlook if needed, and drag the mail out again."
    return
}

""
"S/MIME payload  : '$($payload.FileName)' ($($payload.Data.Length) bytes)"
$env = New-Object System.Security.Cryptography.Pkcs.EnvelopedCms
try {
    $env.Decode($payload.Data)
} catch {
    $signed = New-Object System.Security.Cryptography.Pkcs.SignedCms
    try {
        $signed.Decode($payload.Data)
        "RESULT: standard CMS SignedData (signed, not encrypted)."
        foreach ($s in $signed.SignerInfos) { "  signer: $($s.Certificate.Subject)  digest=$(Describe-Oid $s.DigestAlgorithm)" }
    } catch {
        "RESULT: NOT standard CMS (neither EnvelopedData nor SignedData)."
        "        Decoder message: $($_.Exception.Message)"
        "        First bytes: $([BitConverter]::ToString($payload.Data, 0, [Math]::Min(16, $payload.Data.Length)))"
    }
    return
}

"RESULT: standard CMS EnvelopedData (S/MIME encrypted)."
"Content cipher  : $(Describe-Oid $env.ContentEncryptionAlgorithm.Oid) keyLength=$($env.ContentEncryptionAlgorithm.KeyLength)"
"Recipients      : $($env.RecipientInfos.Count)"
$store = New-Object System.Security.Cryptography.X509Certificates.X509Store 'My', 'CurrentUser'
$store.Open('ReadOnly')
try {
    $i = 0
    foreach ($r in $env.RecipientInfos) {
        $i++
        "  [$i] type=$($r.Type) keyEncryption=$(Describe-Oid $r.KeyEncryptionAlgorithm.Oid)"
        $rid = $r.RecipientIdentifier
        $match = $null
        if ($rid.Type -eq 'IssuerAndSerialNumber') {
            "      issuer=$($rid.Value.IssuerName)  serial=$($rid.Value.SerialNumber)"
            $match = $store.Certificates | Where-Object { $_.Issuer -eq $rid.Value.IssuerName -and $_.SerialNumber -eq $rid.Value.SerialNumber } | Select-Object -First 1
        } elseif ($rid.Type -eq 'SubjectKeyIdentifier') {
            "      subjectKeyIdentifier=$($rid.Value)"
            $match = $store.Certificates | Where-Object { ($_.Extensions | Where-Object { $_.Oid.Value -eq '2.5.29.14' } | ForEach-Object { $_.SubjectKeyIdentifier }) -eq $rid.Value } | Select-Object -First 1
        }
        if ($match) {
            "      IN WINDOWS STORE (CurrentUser\My): subject='$($match.Subject)' hasPrivateKey=$($match.HasPrivateKey) notAfter=$($match.NotAfter.ToString('yyyy-MM-dd'))"
        } else {
            "      not found in CurrentUser\My"
        }
    }
} finally { $store.Close() }
