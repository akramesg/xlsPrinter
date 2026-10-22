using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace XlsPrinter.Core
{
    /// <summary>
    /// Pure-managed .XLSX (OpenXML SpreadsheetML) and .CSV (UTF-8 with BOM for Arabic Excel support)
    /// exporter compatible with .NET 2.0+ (uses a self-contained ZIP store/deflate writer so it has
    /// zero external library dependencies and runs on Windows XP SP2+ / Win7 / Win10).
    /// </summary>
    public class XlsxAndCsvWriter
    {
        public bool RightToLeftSheet { get; set; }
        public bool AutoParseNumbers { get; set; }
        public string SheetName { get; set; }

        public XlsxAndCsvWriter()
        {
            RightToLeftSheet = false;
            AutoParseNumbers = true;
            SheetName = "PrintOutput";
        }

        public void SaveToCsv(ReconstructedTable table, string filePath)
        {
            List<int> activeCols = GetActiveColumns(table);
            List<GridRowData> activeRows = GetActiveRows(table);

            using (FileStream fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                // Write UTF-8 BOM (EF BB BF) so Excel on Windows automatically opens Arabic characters properly
                byte[] bom = new byte[] { 0xEF, 0xBB, 0xBF };
                fs.Write(bom, 0, bom.Length);

                using (StreamWriter sw = new StreamWriter(fs, new UTF8Encoding(false)))
                {
                    for (int r = 0; r < activeRows.Count; r++)
                    {
                        GridRowData row = activeRows[r];
                        StringBuilder line = new StringBuilder();
                        for (int cIdx = 0; cIdx < activeCols.Count; cIdx++)
                        {
                            if (cIdx > 0) line.Append(',');
                            int srcCol = activeCols[cIdx];
                            string val = (srcCol < row.Cells.Length) ? (row.Cells[srcCol].Text ?? "") : "";
                            if (val.IndexOfAny(new char[] { ',', '"', '\r', '\n' }) >= 0)
                            {
                                val = "\"" + val.Replace("\"", "\"\"") + "\"";
                            }
                            line.Append(val);
                        }
                        sw.WriteLine(line.ToString());
                    }
                }
            }
        }

        public void SaveToXlsx(ReconstructedTable table, string filePath)
        {
            List<int> activeCols = GetActiveColumns(table);
            List<GridRowData> activeRows = GetActiveRows(table);

            Dictionary<string, byte[]> zipEntries = new Dictionary<string, byte[]>();

            zipEntries["[Content_Types].xml"] = Encoding.UTF8.GetBytes(
@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
  <Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>
  <Default Extension=""xml"" ContentType=""application/xml""/>
  <Override PartName=""/xl/workbook.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml""/>
  <Override PartName=""/xl/worksheets/sheet1.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>
  <Override PartName=""/xl/styles.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml""/>
</Types>");

            zipEntries["_rels/.rels"] = Encoding.UTF8.GetBytes(
@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
  <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""xl/workbook.xml""/>
</Relationships>");

            string safeSheetName = EscapeXml(string.IsNullOrEmpty(SheetName) ? "PrintOutput" : SheetName);
            zipEntries["xl/workbook.xml"] = Encoding.UTF8.GetBytes(
@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
  <sheets>
    <sheet name=""" + safeSheetName + @""" sheetId=""1"" r:id=""rId1""/>
  </sheets>
</workbook>");

            zipEntries["xl/_rels/workbook.xml.rels"] = Encoding.UTF8.GetBytes(
@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
  <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/sheet1.xml""/>
  <Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""styles.xml""/>
</Relationships>");

            zipEntries["xl/styles.xml"] = Encoding.UTF8.GetBytes(
@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<styleSheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
  <fonts count=""2"">
    <font><sz val=""10""/><name val=""Tahoma""/></font>
    <font><b/><sz val=""10""/><name val=""Tahoma""/></font>
  </fonts>
  <fills count=""2"">
    <fill><patternFill patternType=""none""/></fill>
    <fill><patternFill patternType=""gray125""/></fill>
  </fills>
  <borders count=""1"">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count=""1"">
    <xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0""/>
  </cellStyleXfs>
  <cellXfs count=""4"">
    <xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0"" xfId=""0""/>
    <xf numFmtId=""0"" fontId=""1"" fillId=""0"" borderId=""0"" xfId=""0"" applyFont=""1""/>
    <xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0"" xfId=""0"" applyAlignment=""1""><alignment horizontal=""right"" readingOrder=""2""/></xf>
    <xf numFmtId=""0"" fontId=""1"" fillId=""0"" borderId=""0"" xfId=""0"" applyFont=""1"" applyAlignment=""1""><alignment horizontal=""right"" readingOrder=""2""/></xf>
  </cellXfs>
</styleSheet>");

            // Build worksheet XML
            StringBuilder sb = new StringBuilder();
            sb.Append(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>");
            sb.Append(@"<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">");
            sb.Append(@"<sheetViews><sheetView workbookViewId=""0""");
            if (RightToLeftSheet)
                sb.Append(@" rightToLeft=""1""");
            sb.Append(@"/></sheetViews>");

            // Column widths
            if (activeCols.Count > 0)
            {
                sb.Append("<cols>");
                for (int cIdx = 0; cIdx < activeCols.Count; cIdx++)
                {
                    int maxLen = 12;
                    int srcCol = activeCols[cIdx];
                    for (int r = 0; r < activeRows.Count; r++)
                    {
                        if (srcCol < activeRows[r].Cells.Length && activeRows[r].Cells[srcCol].Text != null)
                        {
                            int len = activeRows[r].Cells[srcCol].Text.Length + 4;
                            if (len > maxLen) maxLen = Math.Min(60, len);
                        }
                    }
                    int colNum = cIdx + 1;
                    sb.AppendFormat(CultureInfo.InvariantCulture, @"<col min=""{0}"" max=""{0}"" width=""{1}"" customWidth=""1""/>", colNum, maxLen);
                }
                sb.Append("</cols>");
            }

            sb.Append("<sheetData>");
            for (int r = 0; r < activeRows.Count; r++)
            {
                GridRowData row = activeRows[r];
                int rowNum = r + 1;
                sb.AppendFormat(@"<row r=""{0}"">", rowNum);

                for (int cIdx = 0; cIdx < activeCols.Count; cIdx++)
                {
                    int srcCol = activeCols[cIdx];
                    if (srcCol >= row.Cells.Length) continue;
                    GridCellData cell = row.Cells[srcCol];
                    if (string.IsNullOrEmpty(cell.Text)) continue;

                    string cellRef = TableLayoutEngine.GetExcelColumnName(cIdx) + rowNum.ToString(CultureInfo.InvariantCulture);
                    int styleIdx = 0;
                    if (cell.IsBold && cell.IsRtl) styleIdx = 3;
                    else if (cell.IsBold) styleIdx = 1;
                    else if (cell.IsRtl) styleIdx = 2;

                    double numVal;
                    if (AutoParseNumbers && TryParseNumeric(cell.Text, out numVal))
                    {
                        sb.AppendFormat(CultureInfo.InvariantCulture, @"<c r=""{0}"" s=""{1}""><v>{2}</v></c>", cellRef, styleIdx, numVal);
                    }
                    else
                    {
                        sb.AppendFormat(@"<c r=""{0}"" s=""{1}"" t=""inlineStr""><is><t>{2}</t></is></c>", cellRef, styleIdx, EscapeXml(cell.Text));
                    }
                }
                sb.Append("</row>");
            }
            sb.Append("</sheetData></worksheet>");

            zipEntries["xl/worksheets/sheet1.xml"] = Encoding.UTF8.GetBytes(sb.ToString());

            WriteZipFile(filePath, zipEntries);
        }

        private static List<int> GetActiveColumns(ReconstructedTable table)
        {
            List<int> cols = new List<int>();
            for (int i = 0; i < table.SelectedColumns.Count; i++)
            {
                if (table.SelectedColumns[i]) cols.Add(i);
            }
            return cols;
        }

        private static List<GridRowData> GetActiveRows(ReconstructedTable table)
        {
            List<GridRowData> rows = new List<GridRowData>();
            for (int i = 0; i < table.Rows.Count; i++)
            {
                if (table.Rows[i].Selected) rows.Add(table.Rows[i]);
            }
            return rows;
        }

        private static bool TryParseNumeric(string text, out double val)
        {
            val = 0;
            if (string.IsNullOrEmpty(text)) return false;
            string t = text.Trim();
            if (t.Length > 1 && t[0] == '0' && t[1] != '.') return false;
            return double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, out val);
        }

        private static string EscapeXml(string s)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;
            return s.Replace("&", "&amp;")
                    .Replace("<", "&lt;")
                    .Replace(">", "&gt;")
                    .Replace("\"", "&quot;")
                    .Replace("'", "&apos;");
        }

        /// <summary>
        /// Minimal ZIP archive writer compatible with .NET 2.0 (using DeflateStream + CRC32).
        /// </summary>
        private static void WriteZipFile(string zipPath, Dictionary<string, byte[]> files)
        {
            using (FileStream fs = new FileStream(zipPath, FileMode.Create, FileAccess.Write))
            using (BinaryWriter w = new BinaryWriter(fs))
            {
                List<ZipDirEntry> dirEntries = new List<ZipDirEntry>();

                foreach (KeyValuePair<string, byte[]> kvp in files)
                {
                    byte[] nameBytes = Encoding.UTF8.GetBytes(kvp.Key);
                    byte[] rawData = kvp.Value;
                    uint crc = ComputeCrc32(rawData);

                    byte[] compressedData;
                    using (MemoryStream cms = new MemoryStream())
                    {
                        using (DeflateStream ds = new DeflateStream(cms, CompressionMode.Compress, true))
                        {
                            ds.Write(rawData, 0, rawData.Length);
                        }
                        compressedData = cms.ToArray();
                    }

                    uint localHeaderOffset = (uint)fs.Position;

                    // Local file header signature (0x04034b50)
                    w.Write((uint)0x04034B50);
                    w.Write((ushort)20); // Version needed (2.0)
                    w.Write((ushort)0);  // General purpose bit flag
                    w.Write((ushort)8);  // Compression method (8 = Deflate)
                    w.Write((ushort)0);  // Last mod file time
                    w.Write((ushort)0x2100); // Last mod file date
                    w.Write(crc);
                    w.Write((uint)compressedData.Length);
                    w.Write((uint)rawData.Length);
                    w.Write((ushort)nameBytes.Length);
                    w.Write((ushort)0); // Extra field length
                    w.Write(nameBytes);
                    w.Write(compressedData);

                    ZipDirEntry de = new ZipDirEntry();
                    de.NameBytes = nameBytes;
                    de.Crc = crc;
                    de.CompressedSize = (uint)compressedData.Length;
                    de.UncompressedSize = (uint)rawData.Length;
                    de.LocalHeaderOffset = localHeaderOffset;
                    dirEntries.Add(de);
                }

                uint centralDirOffset = (uint)fs.Position;
                for (int i = 0; i < dirEntries.Count; i++)
                {
                    ZipDirEntry de = dirEntries[i];
                    w.Write((uint)0x02014B50); // Central directory file header signature
                    w.Write((ushort)20); // Version made by
                    w.Write((ushort)20); // Version needed
                    w.Write((ushort)0);  // Flags
                    w.Write((ushort)8);  // Compression method (Deflate)
                    w.Write((ushort)0);  // File time
                    w.Write((ushort)0x2100); // File date
                    w.Write(de.Crc);
                    w.Write(de.CompressedSize);
                    w.Write(de.UncompressedSize);
                    w.Write((ushort)de.NameBytes.Length);
                    w.Write((ushort)0); // Extra field length
                    w.Write((ushort)0); // File comment length
                    w.Write((ushort)0); // Disk number start
                    w.Write((ushort)0); // Internal file attributes
                    w.Write((uint)0);   // External file attributes
                    w.Write(de.LocalHeaderOffset);
                    w.Write(de.NameBytes);
                }

                uint centralDirSize = (uint)fs.Position - centralDirOffset;

                // End of central directory record (0x06054b50)
                w.Write((uint)0x06054B50);
                w.Write((ushort)0); // Number of this disk
                w.Write((ushort)0); // Disk where central directory starts
                w.Write((ushort)dirEntries.Count);
                w.Write((ushort)dirEntries.Count);
                w.Write(centralDirSize);
                w.Write(centralDirOffset);
                w.Write((ushort)0); // Comment length
            }
        }

        private class ZipDirEntry
        {
            public byte[] NameBytes;
            public uint Crc;
            public uint CompressedSize;
            public uint UncompressedSize;
            public uint LocalHeaderOffset;
        }

        private static readonly uint[] CrcTable = BuildCrc32Table();

        private static uint[] BuildCrc32Table()
        {
            uint[] table = new uint[256];
            for (uint i = 0; i < 256; i++)
            {
                uint c = i;
                for (int k = 0; k < 8; k++)
                {
                    if ((c & 1) != 0)
                        c = 0xEDB88320 ^ (c >> 1);
                    else
                        c >>= 1;
                }
                table[i] = c;
            }
            return table;
        }

        private static uint ComputeCrc32(byte[] data)
        {
            uint crc = 0xFFFFFFFF;
            for (int i = 0; i < data.Length; i++)
            {
                byte index = (byte)((crc ^ data[i]) & 0xFF);
                crc = (crc >> 8) ^ CrcTable[index];
            }
            return ~crc;
        }
    }
}
