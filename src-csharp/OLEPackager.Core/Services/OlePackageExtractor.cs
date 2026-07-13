using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace OLEPackager.Core
{
    public sealed class OlePackageExtractor
    {
        private const int EndOfChain = -2;
        private const int FreeSector = -1;

        public string ExtractZip(byte[] oleBytes, string baseName)
        {
            byte[] nativeBytes = ExtractOle10NativeStream(oleBytes);
            int zipStart = FindZipStart(nativeBytes);
            if (zipStart < 0)
            {
                throw new InvalidDataException("所选 OLE 对象中找不到 zip 数据。");
            }

            int zipLength = nativeBytes.Length - zipStart;
            if (zipStart >= 4)
            {
                uint declaredLength = ReadUInt32(nativeBytes, zipStart - 4);
                if (declaredLength > 0 && declaredLength <= int.MaxValue && zipStart + declaredLength <= nativeBytes.Length)
                {
                    zipLength = (int)declaredLength;
                }
            }

            byte[] zipBytes = CopyRange(nativeBytes, zipStart, zipLength);
            string zipPath = BuildTemporaryZipPath(baseName);
            File.WriteAllBytes(zipPath, zipBytes);

            try
            {
                using (FileStream stream = new FileStream(zipPath, FileMode.Open, FileAccess.Read, FileShare.Read))
                using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read, false, Encoding.UTF8))
                {
                    int ignored = archive.Entries.Count;
                }
            }
            catch
            {
                File.Delete(zipPath);
                throw new InvalidDataException("从所选 OLE 对象恢复出的 zip 数据无效。");
            }

            return zipPath;
        }

        private static byte[] ExtractOle10NativeStream(byte[] oleBytes)
        {
            if (oleBytes == null || oleBytes.Length < 512)
            {
                throw new InvalidDataException("嵌入的 OLE 数据太小。");
            }

            byte[] signature = { 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1 };
            for (int index = 0; index < signature.Length; index++)
            {
                if (oleBytes[index] != signature[index])
                {
                    throw new InvalidDataException("嵌入的 OLE 数据不是复合文档格式。");
                }
            }

            int sectorSize = CheckedPowerOfTwo(ReadUInt16(oleBytes, 30), "OLE 扇区");
            int miniSectorSize = CheckedPowerOfTwo(ReadUInt16(oleBytes, 32), "OLE Mini 扇区");
            int numberOfFatSectors = CheckedInt(ReadUInt32(oleBytes, 44), "FAT 扇区数量");
            int firstDirectorySector = ReadInt32(oleBytes, 48);
            int miniStreamCutoff = CheckedInt(ReadUInt32(oleBytes, 56), "Mini 流阈值");
            int firstMiniFatSector = ReadInt32(oleBytes, 60);

            int[] fat = BuildFatTable(oleBytes, sectorSize, numberOfFatSectors);
            byte[] directoryBytes = ReadRegularStream(oleBytes, fat, firstDirectorySector, sectorSize, -1);

            int rootStart;
            int rootSize;
            int nativeStart;
            int nativeSize;
            FindDirectoryStreams(directoryBytes, out rootStart, out rootSize, out nativeStart, out nativeSize);

            if (nativeSize < miniStreamCutoff)
            {
                byte[] rootStream = ReadRegularStream(oleBytes, fat, rootStart, sectorSize, rootSize);
                int[] miniFat = BuildMiniFatTable(oleBytes, fat, firstMiniFatSector, sectorSize);
                return ReadMiniStream(rootStream, miniFat, nativeStart, miniSectorSize, nativeSize);
            }

            return ReadRegularStream(oleBytes, fat, nativeStart, sectorSize, nativeSize);
        }

        private static int[] BuildFatTable(byte[] oleBytes, int sectorSize, int numberOfFatSectors)
        {
            if (numberOfFatSectors <= 0)
            {
                throw new InvalidDataException("OLE 复合文档没有 FAT 扇区。");
            }

            List<int> fatSectorIds = new List<int>();
            for (int index = 0; index < 109; index++)
            {
                int sectorId = ReadInt32(oleBytes, 76 + index * 4);
                if (sectorId >= 0)
                {
                    fatSectorIds.Add(sectorId);
                }
            }

            int difatSector = ReadInt32(oleBytes, 68);
            int numberOfDifatSectors = CheckedInt(ReadUInt32(oleBytes, 72), "DIFAT 扇区数量");
            int idsPerDifatSector = sectorSize / 4 - 1;
            for (int index = 0; index < numberOfDifatSectors && difatSector >= 0; index++)
            {
                byte[] sector = GetSector(oleBytes, difatSector, sectorSize);
                for (int item = 0; item < idsPerDifatSector; item++)
                {
                    int sectorId = ReadInt32(sector, item * 4);
                    if (sectorId >= 0)
                    {
                        fatSectorIds.Add(sectorId);
                    }
                }

                difatSector = ReadInt32(sector, sectorSize - 4);
            }

            if (fatSectorIds.Count < numberOfFatSectors)
            {
                throw new InvalidDataException("OLE FAT 扇区列表不完整。");
            }

            int entriesPerSector = sectorSize / 4;
            int[] fat = new int[checked(numberOfFatSectors * entriesPerSector)];
            int position = 0;
            for (int index = 0; index < numberOfFatSectors; index++)
            {
                byte[] sector = GetSector(oleBytes, fatSectorIds[index], sectorSize);
                for (int offset = 0; offset < sectorSize; offset += 4)
                {
                    fat[position++] = ReadInt32(sector, offset);
                }
            }

            return fat;
        }

        private static int[] BuildMiniFatTable(byte[] oleBytes, int[] fat, int firstMiniFatSector, int sectorSize)
        {
            if (firstMiniFatSector < 0)
            {
                return new int[0];
            }

            byte[] miniFatBytes = ReadRegularStream(oleBytes, fat, firstMiniFatSector, sectorSize, -1);
            int[] miniFat = new int[miniFatBytes.Length / 4];
            for (int index = 0; index < miniFat.Length; index++)
            {
                miniFat[index] = ReadInt32(miniFatBytes, index * 4);
            }

            return miniFat;
        }

        private static byte[] ReadRegularStream(
            byte[] oleBytes,
            int[] fat,
            int startSector,
            int sectorSize,
            int streamSize)
        {
            List<int> chain = GetSectorChain(fat, startSector);
            if (chain.Count == 0)
            {
                return new byte[0];
            }

            int resultSize = checked(chain.Count * sectorSize);
            byte[] result = new byte[resultSize];
            int position = 0;
            foreach (int sectorId in chain)
            {
                byte[] sector = GetSector(oleBytes, sectorId, sectorSize);
                Buffer.BlockCopy(sector, 0, result, position, sectorSize);
                position += sectorSize;
            }

            return streamSize >= 0 && streamSize < result.Length
                ? CopyRange(result, 0, streamSize)
                : result;
        }

        private static byte[] ReadMiniStream(
            byte[] rootStream,
            int[] miniFat,
            int startMiniSector,
            int miniSectorSize,
            int streamSize)
        {
            List<int> chain = GetSectorChain(miniFat, startMiniSector);
            if (chain.Count == 0 || streamSize <= 0)
            {
                return new byte[0];
            }

            int resultSize = checked(chain.Count * miniSectorSize);
            byte[] result = new byte[resultSize];
            int position = 0;
            foreach (int miniSectorId in chain)
            {
                int sourceOffset = checked(miniSectorId * miniSectorSize);
                EnsureRange(rootStream, sourceOffset, miniSectorSize);
                Buffer.BlockCopy(rootStream, sourceOffset, result, position, miniSectorSize);
                position += miniSectorSize;
            }

            return streamSize < result.Length ? CopyRange(result, 0, streamSize) : result;
        }

        private static List<int> GetSectorChain(int[] table, int startSector)
        {
            List<int> chain = new List<int>();
            HashSet<int> visited = new HashSet<int>();
            int sectorId = startSector;

            while (sectorId >= 0)
            {
                if (sectorId >= table.Length)
                {
                    throw new InvalidDataException("OLE 扇区链超出 FAT 范围。");
                }

                if (!visited.Add(sectorId))
                {
                    throw new InvalidDataException("OLE 扇区链存在循环。");
                }

                chain.Add(sectorId);
                sectorId = table[sectorId];
                if (sectorId == EndOfChain || sectorId == FreeSector)
                {
                    break;
                }
            }

            return chain;
        }

        private static void FindDirectoryStreams(
            byte[] directoryBytes,
            out int rootStart,
            out int rootSize,
            out int nativeStart,
            out int nativeSize)
        {
            rootStart = -1;
            rootSize = 0;
            nativeStart = -1;
            nativeSize = 0;

            for (int offset = 0; offset + 128 <= directoryBytes.Length; offset += 128)
            {
                int nameLength = ReadUInt16(directoryBytes, offset + 64);
                int entryType = directoryBytes[offset + 66];
                string entryName = ReadDirectoryName(directoryBytes, offset, nameLength);

                if (entryType == 5)
                {
                    rootStart = ReadInt32(directoryBytes, offset + 116);
                    rootSize = ReadStreamSize(directoryBytes, offset + 120);
                }

                if (entryType == 2 && entryName.IndexOf("Ole10Native", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    nativeStart = ReadInt32(directoryBytes, offset + 116);
                    nativeSize = ReadStreamSize(directoryBytes, offset + 120);
                }
            }

            if (rootStart < 0)
            {
                throw new InvalidDataException("OLE 复合文档缺少根存储流。");
            }

            if (nativeStart < 0)
            {
                throw new InvalidDataException("所选 OLE 对象不包含 Ole10Native 包流。");
            }
        }

        private static string ReadDirectoryName(byte[] bytes, int offset, int byteCount)
        {
            if (byteCount <= 2 || byteCount > 64 || (byteCount & 1) != 0)
            {
                return string.Empty;
            }

            EnsureRange(bytes, offset, byteCount);
            return Encoding.Unicode.GetString(bytes, offset, byteCount - 2);
        }

        private static int ReadStreamSize(byte[] bytes, int offset)
        {
            EnsureRange(bytes, offset, 8);
            ulong value = BitConverter.ToUInt64(bytes, offset);
            if (value > int.MaxValue)
            {
                throw new InvalidDataException("OLE 流过大，当前版本无法处理。");
            }

            return (int)value;
        }

        private static int FindZipStart(byte[] bytes)
        {
            for (int index = 0; index <= bytes.Length - 4; index++)
            {
                if (bytes[index] == 0x50
                    && bytes[index + 1] == 0x4b
                    && ((bytes[index + 2] == 0x03 && bytes[index + 3] == 0x04)
                        || (bytes[index + 2] == 0x05 && bytes[index + 3] == 0x06)))
                {
                    return index;
                }
            }

            return -1;
        }

        private static byte[] GetSector(byte[] bytes, int sectorId, int sectorSize)
        {
            if (sectorId < 0)
            {
                throw new InvalidDataException("OLE 扇区编号无效。");
            }

            int offset = checked((sectorId + 1) * sectorSize);
            return CopyRange(bytes, offset, sectorSize);
        }

        private static byte[] CopyRange(byte[] bytes, int offset, int count)
        {
            EnsureRange(bytes, offset, count);
            byte[] result = new byte[count];
            Buffer.BlockCopy(bytes, offset, result, 0, count);
            return result;
        }

        private static int CheckedPowerOfTwo(int exponent, string description)
        {
            if (exponent < 0 || exponent > 20)
            {
                throw new InvalidDataException(description + "大小无效。");
            }

            return 1 << exponent;
        }

        private static int CheckedInt(uint value, string description)
        {
            if (value > int.MaxValue)
            {
                throw new InvalidDataException(description + "过大。");
            }

            return (int)value;
        }

        private static int ReadUInt16(byte[] bytes, int offset)
        {
            EnsureRange(bytes, offset, 2);
            return bytes[offset] | bytes[offset + 1] << 8;
        }

        private static uint ReadUInt32(byte[] bytes, int offset)
        {
            EnsureRange(bytes, offset, 4);
            return (uint)(bytes[offset]
                | bytes[offset + 1] << 8
                | bytes[offset + 2] << 16
                | bytes[offset + 3] << 24);
        }

        private static int ReadInt32(byte[] bytes, int offset)
        {
            return unchecked((int)ReadUInt32(bytes, offset));
        }

        private static void EnsureRange(byte[] bytes, int offset, int count)
        {
            if (bytes == null || offset < 0 || count < 0 || offset > bytes.Length - count)
            {
                throw new InvalidDataException("OLE 复合文档的数据范围无效。");
            }
        }

        private static string BuildTemporaryZipPath(string baseName)
        {
            foreach (char invalidCharacter in Path.GetInvalidFileNameChars())
            {
                baseName = baseName.Replace(invalidCharacter, '_');
            }

            string directory = Path.Combine(Path.GetTempPath(), "OLE Packager");
            Directory.CreateDirectory(directory);
            return Path.Combine(directory, baseName + "_" + Guid.NewGuid().ToString("N") + ".zip");
        }
    }
}
