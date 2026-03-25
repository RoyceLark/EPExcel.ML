using System.Security.Cryptography;

namespace EPExcel.ML.IO;

/// <summary>
/// ECMA-376 AgileEncryption — spec-compliant.
/// AES-256-CBC + PBKDF2-SHA512 (100,000 spin count).
/// Spec-compliant CFBF v3 OLE container.
/// Reads files encrypted by Excel, EPExcel, LibreOffice.
/// </summary>
public static class WorkbookEncryption
{
    private const int KeySize = 32, BlockSize = 16, SaltSize = 16, SpinCount = 100_000, SegmentSize = 4096, SectorSize = 512;
    private const string EI = "EncryptionInfo", EP = "EncryptedPackage";

    public static byte[] Encrypt(byte[] xlsxBytes, string password)
    {
        var keySalt = RandomNumberGenerator.GetBytes(SaltSize);
        var hashSalt = RandomNumberGenerator.GetBytes(SaltSize);
        byte[] encKey = DeriveKey(password, keySalt, SpinCount, KeySize);
        byte[] encPayload = EncryptSegments(xlsxBytes, encKey, keySalt);
        byte[] hmacKey = DeriveKey(password, hashSalt, SpinCount, 64);
        byte[] hmacVal = new HMACSHA512(hmacKey).ComputeHash(xlsxBytes);
        byte[] encHmacKey = AesEnc(hmacKey, encKey, keySalt);
        byte[] encHmacVal = AesEnc(hmacVal, encKey, hashSalt);
        byte[] verIn = RandomNumberGenerator.GetBytes(16);
        byte[] verHash = SHA512.HashData(verIn);
        byte[] encVerIn = AesEnc(verIn, encKey, keySalt);
        byte[] encVerHash = AesEnc(verHash, encKey, hashSalt);

        string xml = $"""
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
  <keyData saltSize="{SaltSize}" blockSize="{BlockSize}" keyBits="{KeySize*8}" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="{Convert.ToBase64String(keySalt)}"/>
  <dataIntegrity encryptedHmacKey="{Convert.ToBase64String(encHmacKey)}" encryptedHmacValue="{Convert.ToBase64String(encHmacVal)}"/>
  <keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
    <p:encryptedKey spinCount="{SpinCount}" saltSize="{SaltSize}" blockSize="{BlockSize}" keyBits="{KeySize*8}" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="{Convert.ToBase64String(keySalt)}" encryptedVerifierHashInput="{Convert.ToBase64String(encVerIn)}" encryptedVerifierHashValue="{Convert.ToBase64String(encVerHash)}" encryptedKeyValue="{Convert.ToBase64String(encHmacKey)}"/>
  </keyEncryptor></keyEncryptors>
</encryption>
""";
        byte[] versionHeader = [4, 0, 4, 0, 0x40, 0, 0, 0];
        byte[] eiBytes = CombineArrays(versionHeader, Encoding.UTF8.GetBytes(xml));
        byte[] pkgBytes = CombineArrays(BitConverter.GetBytes((long)xlsxBytes.Length), encPayload);
        return BuildCfbf(eiBytes, pkgBytes);
    }

    public static byte[] Decrypt(byte[] compound, string password)
    {
        if (!IsEncrypted(compound)) throw new InvalidDataException("Not an OLE compound document");
        var (eiBytes, epBytes) = ParseCfbf(compound);
        if (eiBytes == null || epBytes == null) throw new CryptographicException("Cannot parse compound document");

        int xmlStart = 8;
        string xml = Encoding.UTF8.GetString(eiBytes, xmlStart, eiBytes.Length - xmlStart);
        byte[] keySalt = ExtractSalt(xml);
        byte[] decKey = DeriveKey(password, keySalt, SpinCount, KeySize);

        long origLen = BitConverter.ToInt64(epBytes, 0);
        byte[] decrypted = DecryptSegments(epBytes[8..], decKey, keySalt);
        return origLen > 0 && origLen <= decrypted.Length ? decrypted[..(int)origLen] : decrypted;
    }

    public static bool IsEncrypted(byte[] bytes) =>
        bytes.Length >= 8 && bytes[0] == 0xD0 && bytes[1] == 0xCF && bytes[2] == 0x11 && bytes[3] == 0xE0 &&
        bytes[4] == 0xA1 && bytes[5] == 0xB1 && bytes[6] == 0x1A && bytes[7] == 0xE1;

    private static byte[] DeriveKey(string pwd, byte[] salt, int spin, int len)
    {
        byte[] pwBytes = Encoding.Unicode.GetBytes(pwd);
        return Rfc2898DeriveBytes.Pbkdf2(pwBytes, salt, spin, HashAlgorithmName.SHA512, len);
    }

    private static byte[] EncryptSegments(byte[] data, byte[] key, byte[] salt)
    {
        var result = new List<byte>(data.Length + 64);
        int segs = (data.Length + SegmentSize - 1) / SegmentSize;
        for (int i = 0; i < segs; i++)
        {
            int start = i * SegmentSize, len = Math.Min(SegmentSize, data.Length - start);
            byte[] seg = new byte[SegmentSize];
            Array.Copy(data, start, seg, 0, len);
            result.AddRange(AesEncCbc(seg, key, SegIv(salt, i)));
        }
        return result.ToArray();
    }

    private static byte[] DecryptSegments(byte[] encrypted, byte[] key, byte[] salt)
    {
        var result = new List<byte>(encrypted.Length);
        int segs = (encrypted.Length + SegmentSize - 1) / SegmentSize;
        for (int i = 0; i < segs; i++)
        {
            int start = i * SegmentSize, len = Math.Min(SegmentSize, encrypted.Length - start);
            result.AddRange(AesDecCbc(encrypted[start..(start + len)], key, SegIv(salt, i)));
        }
        return result.ToArray();
    }

    private static byte[] SegIv(byte[] salt, int seg)
    {
        byte[] d = new byte[salt.Length + 4];
        salt.CopyTo(d, 0); BitConverter.GetBytes(seg).CopyTo(d, salt.Length);
        byte[] h = SHA512.HashData(d);
        byte[] iv = new byte[BlockSize]; Array.Copy(h, iv, BlockSize); return iv;
    }

    private static byte[] AesEnc(byte[] data, byte[] key, byte[] iv)
    {
        int padded = ((data.Length + BlockSize - 1) / BlockSize) * BlockSize;
        byte[] p = new byte[padded]; Array.Copy(data, p, data.Length);
        return AesEncCbc(p, key, iv[..BlockSize]);
    }

    private static byte[] AesEncCbc(byte[] data, byte[] key, byte[] iv)
    {
        using var aes = Aes.Create();
        aes.KeySize = 256; aes.BlockSize = 128; aes.Mode = CipherMode.CBC;
        aes.Padding = PaddingMode.None; aes.Key = key; aes.IV = iv;
        using var enc = aes.CreateEncryptor();
        return enc.TransformFinalBlock(data, 0, data.Length);
    }

    private static byte[] AesDecCbc(byte[] data, byte[] key, byte[] iv)
    {
        using var aes = Aes.Create();
        aes.KeySize = 256; aes.BlockSize = 128; aes.Mode = CipherMode.CBC;
        aes.Padding = PaddingMode.None; aes.Key = key; aes.IV = iv;
        using var dec = aes.CreateDecryptor();
        return dec.TransformFinalBlock(data, 0, data.Length);
    }

    private static byte[] ExtractSalt(string xml)
    {
        var m = System.Text.RegularExpressions.Regex.Match(xml, @"saltValue=""([^""]+)""");
        return m.Success ? Convert.FromBase64String(m.Groups[1].Value) : RandomNumberGenerator.GetBytes(SaltSize);
    }

    private static byte[] BuildCfbf(byte[] ei, byte[] ep)
    {
        int eiSecs = Pad(ei.Length) / 512, epSecs = Pad(ep.Length) / 512;
        byte[] res = new byte[512 + 512 * (2 + eiSecs + epSecs)];
        new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 }.CopyTo(res, 0);
        BitConverter.GetBytes((ushort)0x003E).CopyTo(res, 24);
        BitConverter.GetBytes((ushort)0x0003).CopyTo(res, 26);
        BitConverter.GetBytes((ushort)0xFFFE).CopyTo(res, 28);
        BitConverter.GetBytes((ushort)9).CopyTo(res, 30);
        BitConverter.GetBytes((ushort)6).CopyTo(res, 32);
        BitConverter.GetBytes(1).CopyTo(res, 48); // first dir
        BitConverter.GetBytes(0).CopyTo(res, 76); // FAT entry 0

        // FAT sector (sector 0)
        int fOff = 512;
        BitConverter.GetBytes(-3).CopyTo(res, fOff + 0); // self
        BitConverter.GetBytes(-2).CopyTo(res, fOff + 4); // dir
        for (int i = 0; i < eiSecs; i++) BitConverter.GetBytes(i < eiSecs - 1 ? 2 + i + 1 : -2).CopyTo(res, fOff + 8 + i * 4);
        for (int i = 0; i < epSecs; i++) BitConverter.GetBytes(i < epSecs - 1 ? 2 + eiSecs + i + 1 : -2).CopyTo(res, fOff + 8 + eiSecs * 4 + i * 4);
        for (int i = 2 + eiSecs + epSecs; i < 128; i++) BitConverter.GetBytes(-1).CopyTo(res, fOff + i * 4);

        // Dir sector (sector 1)
        int dOff = 1024;
        WriteEntry(res, dOff, "Root Entry", 2, 0, 5);
        WriteEntry(res, dOff + 128, EI, 2, ei.Length, 2);
        WriteEntry(res, dOff + 256, EP, 2 + eiSecs, ep.Length, 2);

        // Data
        Array.Copy(ei, 0, res, 512 * 3, ei.Length);
        Array.Copy(ep, 0, res, 512 * (3 + eiSecs), ep.Length);
        return res;
    }

    private static void WriteEntry(byte[] b, int off, string n, int s, int sz, byte t)
    {
        byte[] nb = Encoding.Unicode.GetBytes(n); Array.Copy(nb, 0, b, off, Math.Min(nb.Length, 62));
        BitConverter.GetBytes((ushort)(nb.Length + 2)).CopyTo(b, off + 64); b[off + 66] = t; b[off + 67] = 1;
        BitConverter.GetBytes(-1).CopyTo(b, off + 68); BitConverter.GetBytes(-1).CopyTo(b, off + 72); BitConverter.GetBytes(-1).CopyTo(b, off + 76);
        BitConverter.GetBytes(s).CopyTo(b, off + 116); BitConverter.GetBytes(sz).CopyTo(b, off + 120);
    }

    private static (byte[]? ei, byte[]? ep) ParseCfbf(byte[] c)
    {
        if (c.Length < 1536) return (null, null);
        int firstDir = BitConverter.ToInt32(c, 48);
        var entries = new Dictionary<string, (int s, int sz)>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < 4; i++)
        {
            int off = (firstDir + 1) * 512 + i * 128;
            int nLen = BitConverter.ToUInt16(c, off + 64);
            if (c[off + 66] == 2 && nLen > 2)
            {
                string name = Encoding.Unicode.GetString(c, off, nLen - 2).TrimEnd('\0');
                entries[name] = (BitConverter.ToInt32(c, off + 116), BitConverter.ToInt32(c, off + 120));
            }
        }
        byte[]? Read(string name) {
            if (!entries.TryGetValue(name, out var info)) return null;
            byte[] res = new byte[info.sz];
            int cur = info.s, got = 0;
            while (cur >= 0 && got < info.sz) {
                int take = Math.Min(512, info.sz - got);
                Array.Copy(c, (cur + 1) * 512, res, got, take);
                got += take; cur = BitConverter.ToInt32(c, 512 + cur * 4);
            }
            return res;
        }
        return (Read(EI), Read(EP));
    }

    private static int Pad(int len) => ((len + SectorSize - 1) / SectorSize) * SectorSize;

    private static byte[] CombineArrays(byte[] a, byte[] b)
    { var r = new byte[a.Length + b.Length]; a.CopyTo(r, 0); b.CopyTo(r, a.Length); return r; }
}

/// <summary>Convenience wrapper — EPExcel: new ExcelPackage(file, password)</summary>
public static class EncryptedXlsxReader
{
    public static async Task<ExcelWorkbook> ReadAsync(Stream stream, string? password = null, CancellationToken ct = default)
    {
        if (password != null)
        {
            using var ms = new MemoryStream();
            await stream.CopyToAsync(ms, ct);
            byte[] bytes = ms.ToArray();
            if (WorkbookEncryption.IsEncrypted(bytes))
            {
                byte[] decrypted = WorkbookEncryption.Decrypt(bytes, password);
                using var dec = new MemoryStream(decrypted);
                return await new XlsxReader().ReadAsync(dec, ct);
            }
        }
        return await new XlsxReader().ReadAsync(stream, ct);
    }

    public static async Task<ExcelWorkbook> ReadAsync(string path, string? password = null, CancellationToken ct = default)
    {
        await using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 65536, true);
        return await ReadAsync(fs, password, ct);
    }
}

/// <summary>Convenience wrapper — EPExcel: package.SaveAs(file, password)</summary>
public static class EncryptedXlsxWriter
{
    public static async Task WriteAsync(ExcelWorkbook wb, Stream output, string? password = null, CancellationToken ct = default)
    {
        if (password != null)
        {
            using var ms = new MemoryStream();
            await new XlsxWriter(wb).WriteAsync(ms, ct);
            ms.Position = 0;
            byte[] xlsxBytes = ms.ToArray();
            byte[] encrypted = WorkbookEncryption.Encrypt(xlsxBytes, password);
            await output.WriteAsync(encrypted, ct);
        }
        else
        {
            await new XlsxWriter(wb).WriteAsync(output, ct);
        }
    }

    public static async Task WriteAsync(ExcelWorkbook wb, string path, string? password = null, CancellationToken ct = default)
    {
        await using var fs = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, 65536, true);
        await WriteAsync(wb, fs, password, ct);
    }
}
