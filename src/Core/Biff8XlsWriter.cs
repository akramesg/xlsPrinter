using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace XlsPrinter.Core
{
    /// <summary>
    /// Pure-managed Excel 97-2003 (.xls) binary writer implementing BIFF8 inside an OLE2 Compound Document.
    /// Zero external dependencies; works on Windows XP / 7 / 10 without Excel installed.
    /// Full UTF-16LE Unicode support for Arabic characters, Right-to-Left (RTL) sheet view,
    /// automatic number detection, column auto-fit, and bold formatting.
    /// </summary>
    public class Biff8XlsWriter
    {
        // BIFF8 Record IDs
        private const ushort REC_BOF = 0x0809;
        private const ushort REC_EOF = 0x000A;
        private const ushort REC_CODEPAGE = 0x0042;
        private const ushort REC_WINDOW1 = 0x003D;
        private const ushort REC_FONT = 0x0031;
        private const ushort REC_XF = 0x00E0;
        private const ushort REC_STYLE = 0x0293;
        private const ushort REC_BOUNDSHEET = 0x0085;
        private const ushort REC_SST = 0x00FC;
        private const ushort REC_LABELSST = 0x00FD;
        private const ushort REC_NUMBER = 0x0203;
        private const ushort REC_DIMENSIONS = 0x0200;
        private const ushort REC_COLINFO = 0x007D;
        private const ushort REC_WINDOW2 = 0x023E;

        // XF (Extended Format) indices:
        // 0..14: Style XFs (required by BIFF8)
        // 15: Default Cell XF (Regular, General)
        // 16: Bold Cell XF
        // 17: RTL Cell XF (Regular, Right-to-Left reading order)
        // 18: Bold RTL Cell XF
        private const ushort XF_DEFAULT = 15;
        private const ushort XF_BOLD = 16;
        private const ushort XF_RTL = 17;
        private const ushort XF_BOLD_RTL = 18;

        public bool RightToLeftSheet { get; set; }
        public bool AutoParseNumbers { get; set; }
        public string SheetName { get; set; }

        public Biff8XlsWriter()
        {
            RightToLeftSheet = false;
            AutoParseNumbers = true;
            SheetName = "PrintOutput";
        }

        public void SaveToXls(ReconstructedTable table, string filePath)
        {
            byte[] workbookStream = BuildBiff8WorkbookStream(table);
            byte[] ole2File = WrapInOle2CompoundDocument(workbookStream);
            File.WriteAllBytes(filePath, ole2File);
        }

        private byte[] BuildBiff8WorkbookStream(ReconstructedTable table)
        {
            // Determine active columns and rows
            List<int> activeCols = new List<int>();
            for (int c = 0; c < table.SelectedColumns.Count; c++)
            {
                if (table.SelectedColumns[c])
                    activeCols.Add(c);
            }

            List<GridRowData> activeRows = new List<GridRowData>();
            for (int r = 0; r < table.Rows.Count; r++)
            {
                if (table.Rows[r].Selected)
                    activeRows.Add(table.Rows[r]);
            }

            // Collect Shared String Table (SST) entries and track column widths
            List<string> sstStrings = new List<string>();
            Dictionary<string, int> sstIndexMap = new Dictionary<string, int>();
            int[] maxColChars = new int[Math.Max(1, activeCols.Count)];
            for (int c = 0; c < maxColChars.Length; c++)
                maxColChars[c] = 10;

            for (int r = 0; r < activeRows.Count; r++)
            {
                GridRowData row = activeRows[r];
                for (int cIdx = 0; cIdx < activeCols.Count; cIdx++)
                {
                    int srcCol = activeCols[cIdx];
                    if (srcCol >= row.Cells.Length) continue;
                    string txt = row.Cells[srcCol].Text;
                    if (string.IsNullOrEmpty(txt)) continue;

                    if (txt.Length + 4 > maxColChars[cIdx])
                        maxColChars[cIdx] = Math.Min(65, txt.Length + 4);

                    double dummyNum;
                    if (!AutoParseNumbers || !TryParseNumericCell(txt, out dummyNum))
                    {
                        if (!sstIndexMap.ContainsKey(txt))
                        {
                            sstIndexMap[txt] = sstStrings.Count;
                            sstStrings.Add(txt);
                        }
                    }
                }
            }

            // First build the Worksheet substream so we know the exact byte offset for BOUNDSHEET
            byte[] sheetBytes = BuildWorksheetSubstream(activeRows, activeCols, sstIndexMap, maxColChars);

            // Now build the Workbook Globals substream
            MemoryStream globalsMs = new MemoryStream();
            BinaryWriter gw = new BinaryWriter(globalsMs);

            // 1. BOF (Workbook Globals = 0x0005)
            WriteBof(gw, 0x0005);

            // 2. CODEPAGE (1200 = UTF-16 Unicode)
            WriteRecordHeader(gw, REC_CODEPAGE, 2);
            gw.Write((ushort)0x04B0);

            // 3. WINDOW1
            WriteRecordHeader(gw, REC_WINDOW1, 18);
            gw.Write((ushort)0x00F0); // xWn
            gw.Write((ushort)0x0078); // yWn
            gw.Write((ushort)0x3800); // dxWn
            gw.Write((ushort)0x1C00); // dyWn
            gw.Write((ushort)0x0038); // grbit (visible, hscroll, vscroll, tabs)
            gw.Write((ushort)0x0000); // itabCur
            gw.Write((ushort)0x0000); // itabFirst
            gw.Write((ushort)0x0001); // ctabSel
            gw.Write((ushort)0x0258); // wTabRatio

            // 4. FONT table (BIFF8 uses 0, 1, 2, 3, and skips 4! Then 5 is bold)
            // Font 0: Tahoma 10pt Regular (Supports Arabic & Latin cleanly on XP/Win7/Win10)
            WriteFontRecord(gw, 200, false, "Tahoma");
            // Font 1: Tahoma 10pt Regular
            WriteFontRecord(gw, 200, false, "Tahoma");
            // Font 2: Tahoma 10pt Regular
            WriteFontRecord(gw, 200, false, "Tahoma");
            // Font 3: Tahoma 10pt Regular
            WriteFontRecord(gw, 200, false, "Tahoma");
            // Font 5 (index 5 in BIFF8): Tahoma 10pt Bold
            WriteFontRecord(gw, 200, true, "Tahoma");

            // 5. XF records (0..14 are Style XFs, 15 is default Cell XF, 16..18 custom Cell XFs)
            for (int i = 0; i < 15; i++)
            {
                WriteXfRecord(gw, 0, 0, true, false, false);
            }
            // XF 15: Default Cell XF (Font 0, General, LTR)
            WriteXfRecord(gw, 0, 0, false, false, false);
            // XF 16: Bold Cell XF (Font 5, General, LTR)
            WriteXfRecord(gw, 5, 0, false, true, false);
            // XF 17: RTL Cell XF (Font 0, General, RTL reading order + right align)
            WriteXfRecord(gw, 0, 0, false, false, true);
            // XF 18: Bold RTL Cell XF (Font 5, General, RTL reading order + right align)
            WriteXfRecord(gw, 5, 0, false, true, true);

            // 6. STYLE record (Built-in Normal style referencing XF 0)
            WriteRecordHeader(gw, REC_STYLE, 4);
            gw.Write((ushort)0x8000); // Built-in style flag
            gw.Write((byte)0x00);     // Normal style
            gw.Write((byte)0xFF);     // Outline level

            // 7. Prepare BOUNDSHEET record (we need to patch the 4-byte stream offset of the worksheet BOF)
            byte[] boundSheetPayload = BuildBoundSheetPayload(string.IsNullOrEmpty(SheetName) ? "Sheet1" : SheetName);
            WriteRecordHeader(gw, REC_BOUNDSHEET, (ushort)boundSheetPayload.Length);
            long boundSheetOffsetPos = globalsMs.Position; // First 4 bytes of payload is lbPlyPos
            gw.Write(boundSheetPayload);

            // 8. SST (Shared String Table)
            WriteSstRecords(gw, sstStrings);

            // 9. EOF (End of Workbook Globals)
            WriteRecordHeader(gw, REC_EOF, 0);

            // Patch BOUNDSHEET lbPlyPos with current length of Workbook Globals stream
            uint sheetStartOffset = (uint)globalsMs.Length;
            globalsMs.Position = boundSheetOffsetPos;
            gw.Write(sheetStartOffset);
            globalsMs.Position = globalsMs.Length;

            // Append Worksheet substream
            gw.Write(sheetBytes);
            gw.Flush();

            return globalsMs.ToArray();
        }

        private byte[] BuildWorksheetSubstream(
            List<GridRowData> activeRows,
            List<int> activeCols,
            Dictionary<string, int> sstIndexMap,
            int[] maxColChars)
        {
            MemoryStream ms = new MemoryStream();
            BinaryWriter w = new BinaryWriter(ms);

            // 1. BOF (Worksheet = 0x0010)
            WriteBof(w, 0x0010);

            // 2. COLINFO records for column widths (in 1/256th of a character width)
            for (int c = 0; c < activeCols.Count; c++)
            {
                WriteRecordHeader(w, REC_COLINFO, 12);
                w.Write((ushort)c); // colFirst
                w.Write((ushort)c); // colLast
                ushort colWidth = (ushort)(Math.Max(10, maxColChars[c]) * 256);
                w.Write(colWidth);  // coldx
                w.Write(XF_DEFAULT);// ixfe
                w.Write((ushort)0); // grbit
                w.Write((ushort)0); // reserved
            }

            // 3. DIMENSIONS
            WriteRecordHeader(w, REC_DIMENSIONS, 14);
            w.Write((uint)0); // rwMic
            w.Write((uint)activeRows.Count); // rwMac
            w.Write((ushort)0); // colMic
            w.Write((ushort)activeCols.Count); // colMac
            w.Write((ushort)0); // reserved

            // 4. Cell records (NUMBER and LABELSST)
            for (int r = 0; r < activeRows.Count; r++)
            {
                GridRowData row = activeRows[r];
                for (int cIdx = 0; cIdx < activeCols.Count; cIdx++)
                {
                    int srcCol = activeCols[cIdx];
                    if (srcCol >= row.Cells.Length) continue;
                    GridCellData cell = row.Cells[srcCol];
                    if (string.IsNullOrEmpty(cell.Text)) continue;

                    ushort xfIndex = XF_DEFAULT;
                    if (cell.IsBold && cell.IsRtl) xfIndex = XF_BOLD_RTL;
                    else if (cell.IsBold) xfIndex = XF_BOLD;
                    else if (cell.IsRtl) xfIndex = XF_RTL;

                    double numVal;
                    if (AutoParseNumbers && TryParseNumericCell(cell.Text, out numVal))
                    {
                        WriteRecordHeader(w, REC_NUMBER, 14);
                        w.Write((ushort)r);
                        w.Write((ushort)cIdx);
                        w.Write(xfIndex);
                        w.Write(numVal);
                    }
                    else
                    {
                        int sstIdx;
                        if (sstIndexMap.TryGetValue(cell.Text, out sstIdx))
                        {
                            WriteRecordHeader(w, REC_LABELSST, 10);
                            w.Write((ushort)r);
                            w.Write((ushort)cIdx);
                            w.Write(xfIndex);
                            w.Write((uint)sstIdx);
                        }
                    }
                }
            }

            // 5. WINDOW2 (Sheet view options, including Arabic Right-to-Left mode bit 0x0040)
            WriteRecordHeader(w, REC_WINDOW2, 18);
            ushort grbit = 0x06B6; // Default gridlines, headers, zero values, selected
            if (RightToLeftSheet)
                grbit |= 0x0040; // fRightToLeft (Arabic RTL sheet view)
            w.Write(grbit);
            w.Write((ushort)0); // rwTop
            w.Write((ushort)0); // colLeft
            w.Write((uint)0x00000040); // icvHdr
            w.Write((ushort)0); // wScaleSLV
            w.Write((ushort)0); // wScaleNormal
            w.Write((uint)0);   // reserved

            // 6. EOF
            WriteRecordHeader(w, REC_EOF, 0);

            w.Flush();
            return ms.ToArray();
        }

        private static bool TryParseNumericCell(string text, out double value)
        {
            value = 0;
            if (string.IsNullOrEmpty(text)) return false;

            string trimmed = text.Trim();
            // Do not convert strings with leading zeros (like phone numbers, account codes "00123")
            if (trimmed.Length > 1 && trimmed[0] == '0' && trimmed[1] != '.')
                return false;

            return double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out value);
        }

        private static void WriteRecordHeader(BinaryWriter w, ushort recId, ushort length)
        {
            w.Write(recId);
            w.Write(length);
        }

        private static void WriteBof(BinaryWriter w, ushort substreamType)
        {
            WriteRecordHeader(w, REC_BOF, 16);
            w.Write((ushort)0x0600); // BIFF8 version
            w.Write(substreamType);  // 0x0005 = Workbook Globals, 0x0010 = Worksheet
            w.Write((ushort)0x0DBB); // rupBuild
            w.Write((ushort)0x07CC); // rupYear (1996)
            w.Write((uint)0x00000000); // bfh
            w.Write((uint)0x00000006); // sfo
        }

        private static void WriteFontRecord(BinaryWriter w, ushort heightTwips, bool bold, string fontName)
        {
            byte[] nameBytes = Encoding.Unicode.GetBytes(fontName);
            ushort recLen = (ushort)(14 + 2 + nameBytes.Length);
            WriteRecordHeader(w, REC_FONT, recLen);
            w.Write(heightTwips); // dyHeight (200 = 10pt)
            w.Write((ushort)0);   // grbit
            w.Write((ushort)0x7FFF); // icv (automatic color)
            w.Write((ushort)(bold ? 0x02BC : 0x0190)); // bls (700 = Bold, 400 = Normal)
            w.Write((ushort)0);   // sss
            w.Write((byte)0);     // uls
            w.Write((byte)0);     // bFamily
            w.Write((byte)1);     // bCharSet (DEFAULT_CHARSET = 1, supports Unicode/Arabic)
            w.Write((byte)0);     // reserved
            w.Write((byte)fontName.Length); // cch
            w.Write((byte)0x01);  // fHighByte = 1 (UTF-16LE)
            w.Write(nameBytes);
        }

        private static void WriteXfRecord(BinaryWriter w, ushort fontIndex, ushort numFmtIndex, bool isStyleXf, bool isBold, bool isRtl)
        {
            WriteRecordHeader(w, REC_XF, 20);
            w.Write(fontIndex);
            w.Write(numFmtIndex);
            w.Write((ushort)(isStyleXf ? 0xFFF5 : 0x0001)); // fLocked + parent style 0

            // Alignment byte:
            // Bits 0-2: Horizontal alignment (0 = General, 3 = Right)
            // Bits 4-6: Vertical alignment (2 = Bottom)
            byte align = (byte)(isRtl ? 0x23 : 0x20);
            w.Write(align);

            // Text rotation
            w.Write((byte)0);

            // Indentation & Reading Order byte:
            // Bits 6-7: Reading order (0 = Context, 1 = LTR, 2 = RTL)
            // Bit 10..15 in uint (used attributes flags in next byte)
            byte indentAndDir = (byte)(isRtl ? 0x80 : 0x00); // 2 << 6 = 0x80 (Right-to-Left reading order)
            byte usedAttribs = 0xFC; // Apply font/alignment/border flags
            w.Write(indentAndDir);
            w.Write(usedAttribs);

            w.Write((uint)0); // Borders 1
            w.Write((uint)0); // Borders 2
            w.Write((ushort)0x20C0); // Palette color default
        }

        private static byte[] BuildBoundSheetPayload(string sheetName)
        {
            if (sheetName.Length > 31)
                sheetName = sheetName.Substring(0, 31);
            byte[] nameBytes = Encoding.Unicode.GetBytes(sheetName);
            MemoryStream ms = new MemoryStream();
            BinaryWriter w = new BinaryWriter(ms);
            w.Write((uint)0);     // lbPlyPos (patched later)
            w.Write((byte)0);     // hsState (visible)
            w.Write((byte)0);     // dt (worksheet)
            w.Write((byte)sheetName.Length); // cch
            w.Write((byte)0x01);  // UTF-16LE flag
            w.Write(nameBytes);
            w.Flush();
            return ms.ToArray();
        }

        /// <summary>
        /// Writes BIFF8 SST (0x00FC) and CONTINUE (0x003C) records when the string table exceeds 8224 bytes,
        /// ensuring large multi-page Arabic/English reports never corrupt the BIFF8 record boundary.
        /// </summary>
        private static void WriteSstRecords(BinaryWriter w, List<string> strings)
        {
            const int MAX_RECORD_DATA = 8220;
            const ushort REC_CONTINUE = 0x003C;

            List<byte[]> recordChunks = new List<byte[]>();
            MemoryStream curMs = new MemoryStream();
            BinaryWriter cw = new BinaryWriter(curMs);

            // SST header: cstTotal (4 bytes) + cstUnique (4 bytes)
            cw.Write((uint)strings.Count);
            cw.Write((uint)strings.Count);

            for (int i = 0; i < strings.Count; i++)
            {
                string s = strings[i] ?? string.Empty;
                if (s.Length > 32767)
                    s = s.Substring(0, 32767);

                byte[] strBytes = Encoding.Unicode.GetBytes(s);
                int neededForHeaderAndFirstChar = 3 + Math.Min(2, strBytes.Length);

                // If header + at least 1 char doesn't fit in current record, flush to CONTINUE
                if (curMs.Length + neededForHeaderAndFirstChar > MAX_RECORD_DATA)
                {
                    cw.Flush();
                    recordChunks.Add(curMs.ToArray());
                    curMs = new MemoryStream();
                    cw = new BinaryWriter(curMs);
                }

                // Write string header: cch (2 bytes) + grbit (1 byte: 0x01 = UTF-16LE)
                cw.Write((ushort)s.Length);
                cw.Write((byte)0x01);

                int offset = 0;
                while (offset < strBytes.Length)
                {
                    int spaceLeft = MAX_RECORD_DATA - (int)curMs.Length;
                    // Align to 2-byte UTF-16 character boundary
                    spaceLeft &= ~1;

                    if (spaceLeft <= 0)
                    {
                        cw.Flush();
                        recordChunks.Add(curMs.ToArray());
                        curMs = new MemoryStream();
                        cw = new BinaryWriter(curMs);
                        // First byte of a CONTINUE record continuing a string is the option flag (0x01 = UTF-16LE)
                        cw.Write((byte)0x01);
                        spaceLeft = (MAX_RECORD_DATA - 1) & ~1;
                    }

                    int bytesToCopy = Math.Min(spaceLeft, strBytes.Length - offset);
                    cw.Write(strBytes, offset, bytesToCopy);
                    offset += bytesToCopy;
                }
            }

            cw.Flush();
            recordChunks.Add(curMs.ToArray());

            for (int r = 0; r < recordChunks.Count; r++)
            {
                ushort recType = (r == 0) ? REC_SST : REC_CONTINUE;
                byte[] chunk = recordChunks[r];
                WriteRecordHeader(w, recType, (ushort)chunk.Length);
                w.Write(chunk);
            }
        }

        /// <summary>
        /// Wraps the BIFF8 "Workbook" stream inside a standard Microsoft OLE2 Compound File Binary (CFB) container
        /// (512-byte sectors) so Excel 97-2003, Excel 2007-2024, LibreOffice, and legacy ERP imports open it natively.
        /// </summary>
        private static byte[] WrapInOle2CompoundDocument(byte[] workbookBytes)
        {
            const int SECTOR_SIZE = 512;
            const uint ENDOFCHAIN = 0xFFFFFFFE;
            const uint FREESECT = 0xFFFFFFFF;
            const uint FATSECT = 0xFFFFFFFD;

            // Ensure stream is at least 4096 bytes so OLE2 stores it in standard 512-byte sectors
            // rather than the MiniFAT / Short-Sector container (simplifies OLE2 structure and maximizes compatibility)
            int paddedWorkbookLen = Math.Max(4096, ((workbookBytes.Length + SECTOR_SIZE - 1) / SECTOR_SIZE) * SECTOR_SIZE);
            byte[] paddedWorkbook = new byte[paddedWorkbookLen];
            Buffer.BlockCopy(workbookBytes, 0, paddedWorkbook, 0, workbookBytes.Length);
            // Fill any padding after EOF with 0
            int numWorkbookSectors = paddedWorkbookLen / SECTOR_SIZE;

            // Sector Layout:
            // Sector 0: Root Directory Entry + "Workbook" Directory Entry
            // Sector 1 .. numWorkbookSectors: Workbook Stream Data
            // Sector (1 + numWorkbookSectors) .. : FAT Sectors
            int totalDataAndDirSectors = 1 + numWorkbookSectors;
            int entriesPerFatSector = SECTOR_SIZE / 4; // 128
            int numFatSectors = (totalDataAndDirSectors + entriesPerFatSector) / entriesPerFatSector;
            while ((totalDataAndDirSectors + numFatSectors + entriesPerFatSector - 1) / entriesPerFatSector > numFatSectors)
            {
                numFatSectors++;
            }

            int totalSectors = totalDataAndDirSectors + numFatSectors;

            MemoryStream ms = new MemoryStream((1 + totalSectors) * SECTOR_SIZE);
            BinaryWriter w = new BinaryWriter(ms);

            // ================= 1. OLE2 512-byte Header =================
            w.Write(new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 }); // Signature
            w.Write(new byte[16]); // CLSID (zeros)
            w.Write((ushort)0x003E); // Minor version
            w.Write((ushort)0x0003); // Major version (3 = 512 byte sectors)
            w.Write((ushort)0xFFFE); // Byte order (little-endian)
            w.Write((ushort)9);      // Sector shift (2^9 = 512)
            w.Write((ushort)6);      // Mini-sector shift (2^6 = 64)
            w.Write(new byte[6]);    // Reserved
            w.Write((uint)0);        // Number of directory sectors (0 for v3)
            w.Write((uint)numFatSectors); // Number of FAT sectors
            w.Write((uint)0);        // First Directory Sector Location (Sector 0)
            w.Write((uint)0);        // Transaction signature
            w.Write((uint)0x00001000); // Mini stream cutoff size (4096)
            w.Write(ENDOFCHAIN);     // First MiniFAT sector
            w.Write((uint)0);        // Number of MiniFAT sectors
            w.Write(ENDOFCHAIN);     // First DIFAT sector
            w.Write((uint)0);        // Number of DIFAT sectors

            // DIFAT array in header (109 entries)
            for (int i = 0; i < 109; i++)
            {
                if (i < numFatSectors)
                    w.Write((uint)(totalDataAndDirSectors + i));
                else
                    w.Write(FREESECT);
            }

            // ================= 2. Sector 0: Directory Entries (4 x 128 bytes = 512 bytes) =================
            // Entry 0: "Root Entry"
            WriteDirectoryEntry(w, "Root Entry", 5, 1, 0xFFFFFFFF, 0xFFFFFFFF, 1, ENDOFCHAIN, 0);
            // Entry 1: "Workbook"
            WriteDirectoryEntry(w, "Workbook", 2, 1, 0xFFFFFFFF, 0xFFFFFFFF, 0xFFFFFFFF, 1, (uint)paddedWorkbookLen);
            // Entry 2: Empty
            WriteDirectoryEntry(w, "", 0, 0, 0xFFFFFFFF, 0xFFFFFFFF, 0xFFFFFFFF, 0, 0);
            // Entry 3: Empty
            WriteDirectoryEntry(w, "", 0, 0, 0xFFFFFFFF, 0xFFFFFFFF, 0xFFFFFFFF, 0, 0);

            // ================= 3. Sectors 1..N: Workbook BIFF8 Stream =================
            w.Write(paddedWorkbook);

            // ================= 4. FAT Sectors =================
            int fatEntriesWritten = 0;
            // Sector 0 is Directory (end of chain)
            w.Write(ENDOFCHAIN);
            fatEntriesWritten++;

            // Sectors 1..numWorkbookSectors are Workbook chain
            for (int i = 1; i <= numWorkbookSectors; i++)
            {
                if (i < numWorkbookSectors)
                    w.Write((uint)(i + 1));
                else
                    w.Write(ENDOFCHAIN);
                fatEntriesWritten++;
            }

            // FAT sectors themselves
            for (int i = 0; i < numFatSectors; i++)
            {
                w.Write(FATSECT);
                fatEntriesWritten++;
            }

            // Pad remaining FAT entries with FREESECT
            int totalFatCapacity = numFatSectors * entriesPerFatSector;
            while (fatEntriesWritten < totalFatCapacity)
            {
                w.Write(FREESECT);
                fatEntriesWritten++;
            }

            w.Flush();
            return ms.ToArray();
        }

        private static void WriteDirectoryEntry(
            BinaryWriter w, string name, byte objectType, byte colorFlag,
            uint leftSibling, uint rightSibling, uint childId,
            uint startSector, uint streamSize)
        {
            byte[] nameBuf = new byte[64];
            ushort nameLenBytes = 0;
            if (!string.IsNullOrEmpty(name))
            {
                byte[] utf16 = Encoding.Unicode.GetBytes(name);
                Buffer.BlockCopy(utf16, 0, nameBuf, 0, Math.Min(62, utf16.Length));
                nameLenBytes = (ushort)(utf16.Length + 2); // Includes null terminator
            }

            w.Write(nameBuf);        // 64 bytes: Entry Name
            w.Write(nameLenBytes);   // 2 bytes: Name length
            w.Write(objectType);     // 1 byte: 5=Root, 2=Stream, 0=Empty
            w.Write(colorFlag);      // 1 byte: 1=Black, 0=Red
            w.Write(leftSibling);    // 4 bytes
            w.Write(rightSibling);   // 4 bytes
            w.Write(childId);        // 4 bytes
            w.Write(new byte[16]);   // 16 bytes: CLSID
            w.Write((uint)0);        // 4 bytes: State bits
            w.Write((ulong)0);       // 8 bytes: Creation time
            w.Write((ulong)0);       // 8 bytes: Modified time
            w.Write(startSector);    // 4 bytes: Starting sector
            w.Write(streamSize);     // 4 bytes: Stream size low
            w.Write((uint)0);        // 4 bytes: Stream size high
        }
    }
}
