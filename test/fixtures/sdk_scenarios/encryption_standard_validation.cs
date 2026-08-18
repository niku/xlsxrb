// Decrypts an encrypted XLSX file and validates it with OpenXmlValidator
var bytes = File.ReadAllBytes(XlsxPath);

if (bytes.Length < 512 || bytes[0] != 0xD0 || bytes[1] != 0xCF || bytes[2] != 0x11 || bytes[3] != 0xE0)
{
    throw new Exception("Not a valid CFB file");
}

int sectorSize = 512;
int miniSectorSize = 64;
int firstDirSector = BitConverter.ToInt32(bytes, 0x30);
int firstMiniFatSector = BitConverter.ToInt32(bytes, 0x3C);

var fat = new List<uint>();
for (int i = 0; i < 109; i++)
{
    uint sec = BitConverter.ToUInt32(bytes, 0x4C + (i * 4));
    if (sec == 0xFFFFFFFF || sec == 0xFFFFFFFE) break;
    int offset = 512 + (int)(sec * sectorSize);
    for (int j = 0; j < sectorSize / 4; j++)
    {
        fat.Add(BitConverter.ToUInt32(bytes, offset + (j * 4)));
    }
}

var miniFat = new List<uint>();
if (firstMiniFatSector != -2 && firstMiniFatSector != -1)
{
    uint curr = (uint)firstMiniFatSector;
    while (curr != 0xFFFFFFFE && curr != 0xFFFFFFFF && curr < fat.Count)
    {
        int offset = 512 + (int)(curr * sectorSize);
        for (int j = 0; j < sectorSize / 4; j++)
        {
            miniFat.Add(BitConverter.ToUInt32(bytes, offset + (j * 4)));
        }
        curr = fat[(int)curr];
    }
}

var dirData = new List<byte>();
uint dirSec = (uint)firstDirSector;
while (dirSec != 0xFFFFFFFE && dirSec != 0xFFFFFFFF && dirSec < fat.Count)
{
    int offset = 512 + (int)(dirSec * sectorSize);
    dirData.AddRange(bytes.Skip(offset).Take(sectorSize));
    dirSec = fat[(int)dirSec];
}

byte[] ReadRegularStream(uint startSector, long size)
{
    if (startSector == 0xFFFFFFFE || size <= 0) return Array.Empty<byte>();
    var res = new List<byte>();
    uint s = startSector;
    while (s != 0xFFFFFFFE && s != 0xFFFFFFFF && s < fat.Count)
    {
        int offset = 512 + (int)(s * sectorSize);
        res.AddRange(bytes.Skip(offset).Take(sectorSize));
        s = fat[(int)s];
    }
    return res.Take((int)size).ToArray();
}

byte[] ReadMiniStream(uint startSector, long size, byte[] miniStreamContainer)
{
    if (startSector == 0xFFFFFFFE || size <= 0 || miniStreamContainer == null) return Array.Empty<byte>();
    var res = new List<byte>();
    uint s = startSector;
    while (s != 0xFFFFFFFE && s != 0xFFFFFFFF && s < miniFat.Count)
    {
        int offset = (int)(s * miniSectorSize);
        res.AddRange(miniStreamContainer.Skip(offset).Take(miniSectorSize));
        s = miniFat[(int)s];
    }
    return res.Take((int)size).ToArray();
}

// 1. Read Root Entry to get Mini Stream container
byte[] miniStreamContainer = null;
int numEntries = dirData.Count / 128;
for (int i = 0; i < numEntries; i++)
{
    int entryOffset = i * 128;
    byte type = dirData[entryOffset + 0x42];
    if (type == 5) // Root Entry
    {
        uint startSector = BitConverter.ToUInt32(dirData.ToArray(), entryOffset + 0x74);
        long size = BitConverter.ToInt64(dirData.ToArray(), entryOffset + 0x78);
        miniStreamContainer = ReadRegularStream(startSector, size);
        break;
    }
}

// 2. Read EncryptionInfo & EncryptedPackage
byte[] encInfo = null;
byte[] encPackage = null;

for (int i = 0; i < numEntries; i++)
{
    int entryOffset = i * 128;
    int nameLen = BitConverter.ToInt16(dirData.ToArray(), entryOffset + 0x40);
    byte type = dirData[entryOffset + 0x42];
    if (type != 2) continue; // Stream only

    string name = "";
    if (nameLen > 2)
    {
        name = System.Text.Encoding.Unicode.GetString(dirData.ToArray(), entryOffset, nameLen - 2);
    }

    uint startSector = BitConverter.ToUInt32(dirData.ToArray(), entryOffset + 0x74);
    long size = BitConverter.ToInt64(dirData.ToArray(), entryOffset + 0x78);

    byte[] streamData = (size < 4096 && miniStreamContainer != null && miniStreamContainer.Length > 0)
        ? ReadMiniStream(startSector, size, miniStreamContainer)
        : ReadRegularStream(startSector, size);

    if (name.Equals("EncryptionInfo", StringComparison.OrdinalIgnoreCase))
    {
        encInfo = streamData;
    }
    else if (name.Equals("EncryptedPackage", StringComparison.OrdinalIgnoreCase))
    {
        encPackage = streamData;
    }
}

if (encInfo == null || encPackage == null || encPackage.Length < 8)
{
    throw new Exception("Failed to read valid EncryptionInfo or EncryptedPackage from CFB");
}

string password = "OpenXmlSecretPass123";
int headerSize = BitConverter.ToInt32(encInfo, 8);
int verifierOffset = 12 + headerSize;
int saltSize = BitConverter.ToInt32(encInfo, verifierOffset);
byte[] salt = encInfo.Skip(verifierOffset + 4).Take(saltSize).ToArray();

byte[] derivedKey;
using (var sha1 = SHA1.Create())
{
    byte[] pwBytes = System.Text.Encoding.Unicode.GetBytes(password);
    byte[] h = sha1.ComputeHash(salt.Concat(pwBytes).ToArray());

    for (int i = 0; i < 50000; i++)
    {
        byte[] iBytes = BitConverter.GetBytes(i);
        h = sha1.ComputeHash(iBytes.Concat(h).ToArray());
    }

    byte[] x = sha1.ComputeHash(h.Concat(BitConverter.GetBytes(0)).ToArray());

    byte[] buf1 = new byte[64];
    byte[] buf2 = new byte[64];
    Array.Copy(x, buf1, x.Length);
    Array.Copy(x, buf2, x.Length);
    for (int i = 0; i < 64; i++)
    {
        buf1[i] ^= 0x36;
        buf2[i] ^= 0x5C;
    }

    byte[] k1 = sha1.ComputeHash(buf1);
    byte[] k2 = sha1.ComputeHash(buf2);
    derivedKey = k1.Concat(k2).Take(16).ToArray();
}

long totalSize = BitConverter.ToInt64(encPackage, 0);
byte[] cipherData = encPackage.Skip(8).ToArray();
byte[] plainZip;

using (var aes = Aes.Create())
{
    aes.Key = derivedKey;
    aes.Mode = CipherMode.ECB;
    aes.Padding = PaddingMode.None;

    using (var decryptor = aes.CreateDecryptor())
    {
        byte[] decryptedPackage = decryptor.TransformFinalBlock(cipherData, 0, cipherData.Length);
        plainZip = decryptedPackage.Take((int)totalSize).ToArray();
    }
}

using (var zipStream = new MemoryStream(plainZip))
using (var doc = SpreadsheetDocument.Open(zipStream, false))
{
    var validator = new OpenXmlValidator();
    var errors = validator.Validate(doc);
    if (errors.Any())
    {
        var errorMsg = string.Join("\n", errors.Select(e => $"{e.Description} at {e.Path?.XPath}"));
        throw new Exception($"OpenXmlValidator found schema errors in decrypted document:\n{errorMsg}");
    }

    var sheet = doc.WorkbookPart.WorksheetParts.First().Worksheet;
    var sheetData = sheet.GetFirstChild<SheetData>();
    var rows = sheetData.Elements<Row>().ToList();

    if (rows.Count < 2)
    {
        throw new Exception($"Expected at least 2 rows, found {rows.Count}");
    }
}
