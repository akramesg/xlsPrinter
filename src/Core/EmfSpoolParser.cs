using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace XlsPrinter.Core
{
    /// <summary>
    /// Parses Windows Spooler .SPL files (NT EMF 1.00x format) and standalone .EMF files,
    /// extracting all GDI text records (EMR_EXTTEXTOUTW, EMR_EXTTEXTOUTA, EMR_SMALLTEXTOUT,
    /// EMR_POLYTEXTOUTW/A) along with world/viewport coordinate transforms and font properties.
    /// </summary>
    public class EmfSpoolParser
    {
        // EMF Record Types (wingdi.h)
        private const uint EMR_HEADER = 1;
        private const uint EMR_EOF = 14;
        private const uint EMR_SETMAPMODE = 17;
        private const uint EMR_SETTEXTALIGN = 22;
        private const uint EMR_SAVEDC = 33;
        private const uint EMR_RESTOREDC = 34;
        private const uint EMR_SETWORLDTRANSFORM = 35;
        private const uint EMR_MODIFYWORLDTRANSFORM = 36;
        private const uint EMR_SELECTOBJECT = 37;
        private const uint EMR_DELETEOBJECT = 40;
        private const uint EMR_CREATEFONTINDIRECTW = 82;
        private const uint EMR_EXTTEXTOUTA = 83;
        private const uint EMR_EXTTEXTOUTW = 84;
        private const uint EMR_POLYTEXTOUTA = 96;
        private const uint EMR_POLYTEXTOUTW = 97;
        private const uint EMR_SMALLTEXTOUT = 108;
        private const uint EMR_SETVIEWPORTEXTEX = 11;
        private const uint EMR_SETVIEWPORTORGEX = 12;
        private const uint EMR_SETWINDOWEXTEX = 9;
        private const uint EMR_SETWINDOWORGEX = 10;

        // ETO flags
        private const uint ETO_GLYPH_INDEX = 0x0010;
        private const uint ETO_RTLREADING = 0x0080;
        private const uint ETO_NO_RECT = 0x0100;
        private const uint ETO_SMALL_CHARS = 0x0200;

        // TextAlign flags
        private const uint TA_RIGHT = 2;
        private const uint TA_CENTER = 6;
        private const uint TA_RTLREADING = 256;

        // ModifyWorldTransform modes
        private const uint MWT_IDENTITY = 1;
        private const uint MWT_LEFTMULTIPLY = 2;
        private const uint MWT_RIGHTMULTIPLY = 3;

        [DllImport("gdi32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr CreateFont(int nHeight, int nWidth, int nEscapement, int nOrientation, int fnWeight, uint fdwItalic, uint fdwUnderline, uint fdwStrikeOut, uint fdwCharSet, uint fdwOutputPrecision, uint fdwClipPrecision, uint fdwQuality, uint fdwPitchAndFamily, string lpszFace);

        [DllImport("gdi32.dll")]
        private static extern IntPtr SelectObject(IntPtr hdc, IntPtr hgdiobj);

        [DllImport("gdi32.dll")]
        private static extern bool DeleteObject(IntPtr hObject);

        [DllImport("gdi32.dll", SetLastError = true)]
        private static extern uint GetFontData(IntPtr hdc, uint dwTable, uint dwOffset, [Out] byte[] lpvBuffer, uint cbData);

        private static readonly Dictionary<string, Dictionary<ushort, char>> _fontCmapCache = new Dictionary<string, Dictionary<ushort, char>>(StringComparer.OrdinalIgnoreCase);

        internal static Dictionary<ushort, char> GetFontGlyphMapForTest(string fontName, byte charSet)
        {
            return GetFontGlyphMap(fontName, charSet);
        }

        private static Dictionary<ushort, char> GetFontGlyphMap(string fontName, byte charSet)
        {
            string cacheKey = (fontName ?? "Arial") + "_" + charSet;
            lock (_fontCmapCache)
            {
                Dictionary<ushort, char> map;
                if (_fontCmapCache.TryGetValue(cacheKey, out map))
                    return map;

                map = BuildGlyphToUnicodeMap(fontName, charSet);
                _fontCmapCache[cacheKey] = map;
                return map;
            }
        }

        private static Dictionary<ushort, char> BuildGlyphToUnicodeMap(string fontName, byte charSet)
        {
            Dictionary<ushort, char> map = new Dictionary<ushort, char>();
            string cleanFontName = fontName ?? "Arial";
            if (cleanFontName.StartsWith("@"))
                cleanFontName = cleanFontName.Substring(1);
            if (cleanFontName.EndsWith(" (Arabic)", StringComparison.OrdinalIgnoreCase))
                cleanFontName = cleanFontName.Substring(0, cleanFontName.Length - 9).Trim();

            // Handle Arial Unicode MS fallback
            if (cleanFontName.IndexOf("Arial Unicode", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                // Exact mapping for Arabic Presentation Forms in Arial Unicode MS
                map[0x0003] = '\u0020'; // space
                map[0x04C1] = '\u0642'; // ق
                map[0x04C2] = '\u0643'; // ك
                map[0x04B5] = '\u0631'; // ر
                map[0x04AB] = '\u0627'; // ا
                map[0x196E] = '\uFE8E'; // ﺎ
                map[0x1972] = '\uFE92'; // ﺒ
                map[0x1974] = '\uFE94'; // ﺔ
                map[0x1978] = '\uFE98'; // ﺘ
                map[0x1984] = '\uFEA4'; // ﺤ
                map[0x198E] = '\uFEAE'; // ﺮ
                map[0x1998] = '\uFEB8'; // ﺸ
                map[0x199C] = '\uFEBC'; // ﺼ
                map[0x19BB] = '\uFEDB'; // ﻛ
                map[0x19BE] = '\uFEDE'; // ﻞ
                map[0x19BF] = '\uFEDF'; // ﻟ
                map[0x19C4] = '\uFEE4'; // ﻤ
                map[0x19CE] = '\uFEEE'; // ﻮ
                map[0x19D3] = '\uFEF3'; // ﻳ

                string winDir = Environment.GetEnvironmentVariable("WINDIR") ?? @"C:\Windows";
                string loc = System.Reflection.Assembly.GetExecutingAssembly().Location;
                string asmDir = !string.IsNullOrEmpty(loc) ? Path.GetDirectoryName(loc) : null;
                string currDir = Directory.GetCurrentDirectory();

                List<string> candidatesList = new List<string>();
                if (!string.IsNullOrEmpty(AppDomain.CurrentDomain.BaseDirectory))
                {
                    candidatesList.Add(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "arialuni.ttf"));
                    candidatesList.Add(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"bin\arialuni.ttf"));
                }
                if (!string.IsNullOrEmpty(asmDir))
                {
                    candidatesList.Add(Path.Combine(asmDir, "arialuni.ttf"));
                    candidatesList.Add(Path.Combine(asmDir, @"bin\arialuni.ttf"));
                }
                if (!string.IsNullOrEmpty(currDir))
                {
                    candidatesList.Add(Path.Combine(currDir, "arialuni.ttf"));
                    candidatesList.Add(Path.Combine(currDir, @"bin\arialuni.ttf"));
                }
                candidatesList.Add(Path.Combine(winDir, @"Fonts\arialuni.ttf"));
                candidatesList.Add(Path.Combine(winDir, @"Fonts\ARIALUNI.TTF"));

                foreach (string candidate in candidatesList)
                {
                    if (!string.IsNullOrEmpty(candidate) && File.Exists(candidate))
                    {
                        try
                        {
                            byte[] ttfData = File.ReadAllBytes(candidate);
                            ParseTtfCmapData(ttfData, map);
                            break;
                        }
                        catch { }
                    }
                }
                return map;
            }

            try
            {
                using (Bitmap bmp = new Bitmap(1, 1))
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    IntPtr hdc = g.GetHdc();
                    try
                    {
                        IntPtr hFont = CreateFont(24, 0, 0, 0, 400, 0, 0, 0, charSet != 0 ? charSet : (uint)178, 0, 0, 0, 0, cleanFontName);
                        IntPtr oldFont = SelectObject(hdc, hFont);

                        // Table tag 'cmap' = 0x70616D63 in little endian
                        uint CMAP_TAG = 0x70616D63;
                        uint size = GetFontData(hdc, CMAP_TAG, 0, null, 0);
                        if (size != 0xFFFFFFFF && size > 0)
                        {
                            byte[] cmapData = new byte[size];
                            GetFontData(hdc, CMAP_TAG, 0, cmapData, size);
                            ushort numTables = (ushort)((cmapData[2] << 8) | cmapData[3]);
                            for (int t = 0; t < numTables; t++)
                            {
                                ushort platformId = (ushort)((cmapData[4 + t * 8] << 8) | cmapData[5 + t * 8]);
                                ushort encodingId = (ushort)((cmapData[6 + t * 8] << 8) | cmapData[7 + t * 8]);
                                uint subtableOffset = (uint)((cmapData[8 + t * 8] << 24) | (cmapData[9 + t * 8] << 16) | (cmapData[10 + t * 8] << 8) | cmapData[11 + t * 8]);
                                if ((platformId == 3 && encodingId == 1) || (platformId == 0))
                                {
                                    ushort format = (ushort)((cmapData[subtableOffset] << 8) | cmapData[subtableOffset + 1]);
                                    if (format == 4)
                                    {
                                        ushort segCountX2 = (ushort)((cmapData[subtableOffset + 6] << 8) | cmapData[subtableOffset + 7]);
                                        int segCount = segCountX2 / 2;
                                        int endCodesOffset = (int)subtableOffset + 14;
                                        int startCodesOffset = endCodesOffset + segCountX2 + 2;
                                        int idDeltaOffset = startCodesOffset + segCountX2;
                                        int idRangeOffset = idDeltaOffset + segCountX2;
                                        for (int s = 0; s < segCount; s++)
                                        {
                                            ushort startCode = (ushort)((cmapData[startCodesOffset + s * 2] << 8) | cmapData[startCodesOffset + s * 2 + 1]);
                                            ushort endCode = (ushort)((cmapData[endCodesOffset + s * 2] << 8) | cmapData[endCodesOffset + s * 2 + 1]);
                                            short idDelta = (short)((cmapData[idDeltaOffset + s * 2] << 8) | cmapData[idDeltaOffset + s * 2 + 1]);
                                            ushort rangeOffset = (ushort)((cmapData[idRangeOffset + s * 2] << 8) | cmapData[idRangeOffset + s * 2 + 1]);
                                            if (startCode == 0xFFFF) break;
                                            for (int c = startCode; c <= endCode; c++)
                                            {
                                                ushort glyphIndex = 0;
                                                if (rangeOffset == 0)
                                                {
                                                    glyphIndex = (ushort)((c + idDelta) & 0xFFFF);
                                                }
                                                else
                                                {
                                                    int glyphOffset = idRangeOffset + s * 2 + rangeOffset + (c - startCode) * 2;
                                                    glyphIndex = (ushort)((cmapData[glyphOffset] << 8) | cmapData[glyphOffset + 1]);
                                                    if (glyphIndex != 0) glyphIndex = (ushort)((glyphIndex + idDelta) & 0xFFFF);
                                                }
                                                if (glyphIndex > 0 && !map.ContainsKey(glyphIndex))
                                                {
                                                    map[glyphIndex] = (char)c;
                                                }
                                            }
                                        }
                                        break;
                                    }
                                }
                            }
                        }
                        SelectObject(hdc, oldFont);
                        DeleteObject(hFont);
                    }
                    finally
                    {
                        g.ReleaseHdc(hdc);
                    }
                }
            }
            catch
            {
                // Fallback
            }
            return map;
        }

        private static void ParseTtfCmapData(byte[] data, Dictionary<ushort, char> map)
        {
            if (data == null || data.Length < 12) return;
            ushort numTables = (ushort)((data[4] << 8) | data[5]);
            int cmapOffset = 0;
            for (int i = 0; i < numTables; i++)
            {
                int off = 12 + i * 16;
                if (off + 16 > data.Length) break;
                string tag = Encoding.ASCII.GetString(data, off, 4);
                if (tag == "cmap")
                {
                    cmapOffset = (int)((data[off + 8] << 24) | (data[off + 9] << 16) | (data[off + 10] << 8) | data[off + 11]);
                    break;
                }
            }
            if (cmapOffset == 0 || cmapOffset + 4 > data.Length) return;
            ushort subTables = (ushort)((data[cmapOffset + 2] << 8) | data[cmapOffset + 3]);
            for (int t = 0; t < subTables; t++)
            {
                int toff = cmapOffset + 4 + t * 8;
                if (toff + 8 > data.Length) break;
                ushort platformId = (ushort)((data[toff] << 8) | data[toff + 1]);
                ushort encodingId = (ushort)((data[toff + 2] << 8) | data[toff + 3]);
                uint subOffset = (uint)((data[toff + 4] << 24) | (data[toff + 5] << 16) | (data[toff + 6] << 8) | data[toff + 7]);
                int absSub = cmapOffset + (int)subOffset;
                if (absSub + 4 > data.Length) continue;
                ushort format = (ushort)((data[absSub] << 8) | data[absSub + 1]);
                if (format == 4 && ((platformId == 3 && encodingId == 1) || platformId == 0))
                {
                    ushort segCountX2 = (ushort)((data[absSub + 6] << 8) | data[absSub + 7]);
                    int segCount = segCountX2 / 2;
                    int endCodes = absSub + 14;
                    int startCodes = endCodes + segCountX2 + 2;
                    int idDeltas = startCodes + segCountX2;
                    int idRangeOffsets = idDeltas + segCountX2;
                    if (idRangeOffsets + segCountX2 > data.Length) break;
                    for (int s = 0; s < segCount; s++)
                    {
                        ushort start = (ushort)((data[startCodes + s * 2] << 8) | data[startCodes + s * 2 + 1]);
                        ushort end = (ushort)((data[endCodes + s * 2] << 8) | data[endCodes + s * 2 + 1]);
                        short delta = (short)((data[idDeltas + s * 2] << 8) | data[idDeltas + s * 2 + 1]);
                        ushort ro = (ushort)((data[idRangeOffsets + s * 2] << 8) | data[idRangeOffsets + s * 2 + 1]);
                        if (start == 0xFFFF) break;
                        for (int c = start; c <= end; c++)
                        {
                            ushort gid = 0;
                            if (ro == 0) gid = (ushort)((c + delta) & 0xFFFF);
                            else
                            {
                                int go = idRangeOffsets + s * 2 + ro + (c - start) * 2;
                                if (go + 2 <= data.Length)
                                {
                                    gid = (ushort)((data[go] << 8) | data[go + 1]);
                                    if (gid != 0) gid = (ushort)((gid + delta) & 0xFFFF);
                                }
                            }
                            if (gid > 0 && !map.ContainsKey(gid)) map[gid] = (char)c;
                        }
                    }
                    break;
                }
            }
        }

        private class FontEntry
        {
            public string FaceName = "Arial";
            public float Height = 10f;
            public bool Bold = false;
            public byte CharSet = 1; // DEFAULT_CHARSET
        }

        private class DcState
        {
            public float M11 = 1f, M12 = 0f, M21 = 0f, M22 = 1f, Dx = 0f, Dy = 0f;
            public int WinOrgX = 0, WinOrgY = 0, WinExtX = 1, WinExtY = 1;
            public int VpOrgX = 0, VpOrgY = 0, VpExtX = 1, VpExtY = 1;
            public uint TextAlign = 0;
            public uint SelectedFontHandle = 0;

            public DcState Clone()
            {
                return (DcState)this.MemberwiseClone();
            }
        }

        public bool NormalizePresentationForms { get; set; }
        public bool ReverseVisualArabic { get; set; }

        public EmfSpoolParser()
        {
            NormalizePresentationForms = true;
            ReverseVisualArabic = true;
        }

        /// <summary>
        /// Parses either a Windows Spooler .SPL file (containing 1 or more EMF pages),
        /// a single .EMF file, or a raw text/PRN file fallback.
        /// </summary>
        public PrintJobDocument ParseFile(string filePath)
        {
            byte[] data = File.ReadAllBytes(filePath);
            return ParseBytes(data, Path.GetFileNameWithoutExtension(filePath));
        }

        public PrintJobDocument ParseBytes(byte[] data, string documentTitle)
        {
            PrintJobDocument doc = new PrintJobDocument();
            doc.DocumentTitle = string.IsNullOrEmpty(documentTitle) ? "Captured Print Job" : documentTitle;

            if (data == null || data.Length < 16)
                return doc;

            // 1. Check if the print job is an OpenXPS or XPS package (ZIP archive containing [Content_Types].xml)
            if (data.Length >= 4 && data[0] == 0x50 && data[1] == 0x4B && data[2] == 0x03 && data[3] == 0x04)
            {
                List<ExtractedPage> xpsPages = ParseXpsPackage(data);
                if (xpsPages.Count > 0)
                {
                    for (int i = 0; i < xpsPages.Count; i++)
                    {
                        doc.Pages.Add(xpsPages[i]);
                    }
                    return doc;
                }
            }

            // 2. Scan for all embedded EMF streams inside the SPL or EMF buffer.
            // Every valid EMF stream starts with EMR_HEADER (type == 1) and contains signature 0x464D4520 (" EMF") at offset 40.
            List<byte[]> emfStreams = ExtractEmfStreams(data, doc);

            if (emfStreams.Count > 0)
            {
                for (int i = 0; i < emfStreams.Count; i++)
                {
                    ExtractedPage page = ParseSingleEmfPage(emfStreams[i], i + 1);
                    doc.Pages.Add(page);
                }
            }
            else
            {
                // Fallback: if the print job is raw text / line printer stream
                ExtractedPage textPage = ParseRawTextFallback(data, 1);
                if (textPage.Elements.Count > 0)
                {
                    doc.Pages.Add(textPage);
                }
            }

            return doc;
        }

        /// <summary>
        /// Parses an XPS or OpenXPS print stream (standard spool format on Windows 10/11 v4 print drivers),
        /// extracting FixedPage Glyphs, coordinates (OriginX, OriginY, FontRenderingEmSize), and Unicode text.
        /// </summary>
        private List<ExtractedPage> ParseXpsPackage(byte[] zipData)
        {
            List<ExtractedPage> pages = new List<ExtractedPage>();
            try
            {
                Dictionary<string, byte[]> zipEntries = ExtractZipEntries(zipData);
                // Look for FixedPage parts: Documents/1/Pages/1.fpage, 2.fpage, etc.
                List<string> pageKeys = new List<string>();
                foreach (string key in zipEntries.Keys)
                {
                    if (key.IndexOf(".fpage", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        key.IndexOf("FixedDocument", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        key.IndexOf("/Pages/", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        pageKeys.Add(key);
                    }
                }

                // Natural sort page files (1.fpage, 2.fpage, etc.)
                pageKeys.Sort(StringComparer.OrdinalIgnoreCase);

                for (int p = 0; p < pageKeys.Count; p++)
                {
                    byte[] pageXmlBytes = zipEntries[pageKeys[p]];
                    ExtractedPage page = ParseSingleXpsPage(pageXmlBytes, p + 1);
                    if (page.Elements.Count > 0)
                    {
                        pages.Add(page);
                    }
                }
            }
            catch
            {
                // Ignore and let fallback handle
            }

            return pages;
        }

        private ExtractedPage ParseSingleXpsPage(byte[] xmlBytes, int pageNumber)
        {
            ExtractedPage page = new ExtractedPage(pageNumber);

            try
            {
                string xmlStr = Encoding.UTF8.GetString(xmlBytes);
                System.Xml.XmlDocument xdoc = new System.Xml.XmlDocument();
                xdoc.LoadXml(xmlStr);

                // Check page dimensions from <FixedPage Width="816" Height="1056" ...>
                // XPS coordinates are in 1/96 inch DIPs (1 inch = 25.4 mm => 1 DIP = 0.264583 mm)
                const float DIP_TO_MM = 25.4f / 96.0f;

                System.Xml.XmlElement root = xdoc.DocumentElement;
                if (root != null)
                {
                    float wDip = 816f, hDip = 1056f;
                    float.TryParse(root.GetAttribute("Width"), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out wDip);
                    float.TryParse(root.GetAttribute("Height"), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out hDip);
                    page.Width = wDip * DIP_TO_MM;
                    page.Height = hDip * DIP_TO_MM;
                }

                System.Xml.XmlNodeList glyphs = xdoc.SelectNodes("//*[local-name()='Glyphs']");
                if (glyphs != null)
                {
                    foreach (System.Xml.XmlNode gNode in glyphs)
                    {
                        if (gNode.Attributes == null) continue;

                        System.Xml.XmlAttribute attrText = gNode.Attributes["UnicodeString"];
                        if (attrText == null || string.IsNullOrEmpty(attrText.Value))
                            continue;

                        string rawText = attrText.Value;
                        string text = ArabicTextHelper.NormalizeArabicText(rawText, NormalizePresentationForms, ReverseVisualArabic);
                        if (string.IsNullOrEmpty(text.Trim())) continue;

                        float originX = 0f, originY = 0f, emSize = 12f;
                        if (gNode.Attributes["OriginX"] != null)
                            float.TryParse(gNode.Attributes["OriginX"].Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out originX);
                        if (gNode.Attributes["OriginY"] != null)
                            float.TryParse(gNode.Attributes["OriginY"].Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out originY);
                        if (gNode.Attributes["FontRenderingEmSize"] != null)
                            float.TryParse(gNode.Attributes["FontRenderingEmSize"].Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out emSize);

                        float xMm = originX * DIP_TO_MM;
                        float yMm = (originY - emSize * 0.8f) * DIP_TO_MM; // OriginY is baseline in XPS
                        float heightMm = Math.Max(2.5f, emSize * DIP_TO_MM);
                        float widthMm = Math.Max(3.0f, text.Length * heightMm * 0.55f);

                        ExtractedTextElement elem = new ExtractedTextElement();
                        elem.PageNumber = pageNumber;
                        elem.Text = text.Trim();
                        elem.X = xMm;
                        elem.Y = yMm;
                        elem.Width = widthMm;
                        elem.Height = heightMm;
                        elem.FontSize = emSize * 0.75f; // Convert 1/96 to pt (72/96)
                        elem.FontName = "Arial";
                        elem.IsBold = false;
                        elem.IsRtl = ArabicTextHelper.ContainsArabic(text);
                        elem.Selected = true;

                        page.Elements.Add(elem);
                    }
                }

                MergeFragmentedTextElements(page);
            }
            catch
            {
                // Fallback
            }

            return page;
        }

        private static Dictionary<string, byte[]> ExtractZipEntries(byte[] zipData)
        {
            Dictionary<string, byte[]> entries = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
            int pos = 0;
            while (pos + 30 <= zipData.Length)
            {
                uint sig = BitConverter.ToUInt32(zipData, pos);
                if (sig != 0x04034B50) break; // Local file header signature

                ushort method = BitConverter.ToUInt16(zipData, pos + 8);
                uint compSize = BitConverter.ToUInt32(zipData, pos + 18);
                ushort nameLen = BitConverter.ToUInt16(zipData, pos + 26);
                ushort extraLen = BitConverter.ToUInt16(zipData, pos + 28);
                int namePos = pos + 30;
                if (namePos + nameLen > zipData.Length) break;

                string name = Encoding.UTF8.GetString(zipData, namePos, nameLen);
                int dataPos = namePos + nameLen + extraLen;
                if (dataPos + compSize > zipData.Length) break;

                byte[] uncompBytes = null;
                if (method == 0) // Stored
                {
                    uncompBytes = new byte[compSize];
                    Buffer.BlockCopy(zipData, dataPos, uncompBytes, 0, (int)compSize);
                }
                else if (method == 8) // Deflate
                {
                    using (MemoryStream inMs = new MemoryStream(zipData, dataPos, (int)compSize))
                    using (System.IO.Compression.DeflateStream ds = new System.IO.Compression.DeflateStream(inMs, System.IO.Compression.CompressionMode.Decompress))
                    using (MemoryStream outMs = new MemoryStream())
                    {
                        byte[] buf = new byte[4096];
                        int r;
                        while ((r = ds.Read(buf, 0, buf.Length)) > 0)
                        {
                            outMs.Write(buf, 0, r);
                        }
                        uncompBytes = outMs.ToArray();
                    }
                }

                if (uncompBytes != null)
                {
                    entries[name] = uncompBytes;
                }

                pos = dataPos + (int)compSize;
            }

            return entries;
        }

        /// <summary>
        /// Locates all EMF streams inside an NT EMF Spool (.SPL) file or standalone .EMF buffer.
        /// </summary>
        private List<byte[]> ExtractEmfStreams(byte[] data, PrintJobDocument doc)
        {
            List<byte[]> streams = new List<byte[]>();

            // Also try to read document title from NT SPL header if present:
            // SPL header starts with 0x00010000 (NT EMF 1.003) or 0x00010002 (NT EMF 1.008)
            uint firstDword = BitConverter.ToUInt32(data, 0);
            if (firstDword == 0x00010000 || firstDword == 0x00010002 || firstDword == 0x00010001)
            {
                uint headerSize = BitConverter.ToUInt32(data, 4);
                if (headerSize >= 20 && headerSize <= data.Length)
                {
                    uint docNameOffset = BitConverter.ToUInt32(data, 12);
                    if (docNameOffset >= 20 && docNameOffset < headerSize)
                    {
                        string splDocName = ReadNullTerminatedUnicode(data, (int)docNameOffset, (int)(headerSize - docNameOffset));
                        if (!string.IsNullOrEmpty(splDocName))
                            doc.DocumentTitle = splDocName;
                    }
                }
            }

            // Scan for EMF header signature:
            // Offset +0: iType == 1 (EMR_HEADER)
            // Offset +4: nSize >= 88 and multiple of 4
            // Offset +40: dSignature == 0x464D4520 (" EMF")
            // Offset +48: nBytes (total size of this EMF page stream)
            int pos = 0;
            while (pos <= data.Length - 88)
            {
                if (BitConverter.ToUInt32(data, pos) == EMR_HEADER)
                {
                    uint hdrSize = BitConverter.ToUInt32(data, pos + 4);
                    uint sig = BitConverter.ToUInt32(data, pos + 40);
                    if (hdrSize >= 88 && hdrSize <= 1024 && sig == 0x464D4520)
                    {
                        uint emfTotalBytes = BitConverter.ToUInt32(data, pos + 48);
                        if (emfTotalBytes >= hdrSize && pos + emfTotalBytes <= data.Length)
                        {
                            byte[] emfBytes = new byte[emfTotalBytes];
                            Buffer.BlockCopy(data, pos, emfBytes, 0, (int)emfTotalBytes);
                            streams.Add(emfBytes);
                            pos += (int)emfTotalBytes;
                            continue;
                        }
                        else if (pos + hdrSize <= data.Length)
                        {
                            // In case nBytes in header is slightly off, walk records until EMR_EOF
                            int endPos = WalkUntilEmfEof(data, pos);
                            if (endPos > pos)
                            {
                                int len = endPos - pos;
                                byte[] emfBytes = new byte[len];
                                Buffer.BlockCopy(data, pos, emfBytes, 0, len);
                                streams.Add(emfBytes);
                                pos = endPos;
                                continue;
                            }
                        }
                    }
                }
                pos += 4;
            }

            return streams;
        }

        private int WalkUntilEmfEof(byte[] data, int startPos)
        {
            int pos = startPos;
            while (pos <= data.Length - 8)
            {
                uint recType = BitConverter.ToUInt32(data, pos);
                uint recSize = BitConverter.ToUInt32(data, pos + 4);
                if (recSize < 8 || pos + recSize > data.Length)
                    break;
                pos += (int)recSize;
                if (recType == EMR_EOF)
                    return pos;
            }
            return -1;
        }

        /// <summary>
        /// Parses a single EMF page stream and extracts all text records with page coordinates.
        /// </summary>
        private ExtractedPage ParseSingleEmfPage(byte[] emfData, int pageNumber)
        {
            ExtractedPage page = new ExtractedPage(pageNumber);
            page.RawEmfBytes = emfData;

            Dictionary<uint, FontEntry> fonts = new Dictionary<uint, FontEntry>();
            Stack<DcState> dcStack = new Stack<DcState>();
            DcState dc = new DcState();

            // Bounds from EMF Header
            int frameLeft = BitConverter.ToInt32(emfData, 24);
            int frameTop = BitConverter.ToInt32(emfData, 28);
            int frameRight = BitConverter.ToInt32(emfData, 32);
            int frameBottom = BitConverter.ToInt32(emfData, 36);

            // Frame is in 0.01 millimeter units
            if (frameRight > frameLeft && frameBottom > frameTop)
            {
                page.Width = (frameRight - frameLeft) / 100.0f;
                page.Height = (frameBottom - frameTop) / 100.0f;
            }

            int boundsLeft = BitConverter.ToInt32(emfData, 8);
            int boundsTop = BitConverter.ToInt32(emfData, 12);
            int boundsRight = BitConverter.ToInt32(emfData, 16);
            int boundsBottom = BitConverter.ToInt32(emfData, 20);

            // Reference device DPI/resolution (ENHMETAHEADER szlDevice at +72, szlMillimeters at +80)
            int devCxPx = BitConverter.ToInt32(emfData, 72);
            int devCyPx = BitConverter.ToInt32(emfData, 76);
            int devCxMm = BitConverter.ToInt32(emfData, 80);
            int devCyMm = BitConverter.ToInt32(emfData, 84);

            float scaleXToMm = (devCxPx > 0 && devCxMm > 0) ? ((float)devCxMm / devCxPx) : 0.254f; // default ~100 DPI
            float scaleYToMm = (devCyPx > 0 && devCyMm > 0) ? ((float)devCyMm / devCyPx) : 0.254f;

            int pos = 0;
            while (pos <= emfData.Length - 8)
            {
                uint iType = BitConverter.ToUInt32(emfData, pos);
                uint nSize = BitConverter.ToUInt32(emfData, pos + 4);

                if (nSize < 8 || pos + nSize > emfData.Length)
                    break;

                switch (iType)
                {
                    case EMR_SAVEDC:
                        dcStack.Push(dc.Clone());
                        break;

                    case EMR_RESTOREDC:
                        if (nSize >= 12)
                        {
                            int relative = BitConverter.ToInt32(emfData, pos + 8);
                            int count = Math.Abs(relative);
                            for (int i = 0; i < count && dcStack.Count > 0; i++)
                            {
                                dc = dcStack.Pop();
                            }
                        }
                        break;

                    case EMR_SETWORLDTRANSFORM:
                        if (nSize >= 32)
                        {
                            dc.M11 = BitConverter.ToSingle(emfData, pos + 8);
                            dc.M12 = BitConverter.ToSingle(emfData, pos + 12);
                            dc.M21 = BitConverter.ToSingle(emfData, pos + 16);
                            dc.M22 = BitConverter.ToSingle(emfData, pos + 20);
                            dc.Dx = BitConverter.ToSingle(emfData, pos + 24);
                            dc.Dy = BitConverter.ToSingle(emfData, pos + 28);
                        }
                        break;

                    case EMR_MODIFYWORLDTRANSFORM:
                        if (nSize >= 36)
                        {
                            uint mode = BitConverter.ToUInt32(emfData, pos + 32);
                            if (mode == MWT_IDENTITY)
                            {
                                dc.M11 = 1f; dc.M12 = 0f; dc.M21 = 0f; dc.M22 = 1f; dc.Dx = 0f; dc.Dy = 0f;
                            }
                            else if (mode == MWT_LEFTMULTIPLY || mode == MWT_RIGHTMULTIPLY)
                            {
                                float m11 = BitConverter.ToSingle(emfData, pos + 8);
                                float m12 = BitConverter.ToSingle(emfData, pos + 12);
                                float m21 = BitConverter.ToSingle(emfData, pos + 16);
                                float m22 = BitConverter.ToSingle(emfData, pos + 20);
                                float dx = BitConverter.ToSingle(emfData, pos + 24);
                                float dy = BitConverter.ToSingle(emfData, pos + 28);

                                if (mode == MWT_LEFTMULTIPLY)
                                {
                                    float nM11 = m11 * dc.M11 + m12 * dc.M21;
                                    float nM12 = m11 * dc.M12 + m12 * dc.M22;
                                    float nM21 = m21 * dc.M11 + m22 * dc.M21;
                                    float nM22 = m21 * dc.M12 + m22 * dc.M22;
                                    float nDx = dx * dc.M11 + dy * dc.M21 + dc.Dx;
                                    float nDy = dx * dc.M12 + dy * dc.M22 + dc.Dy;
                                    dc.M11 = nM11; dc.M12 = nM12; dc.M21 = nM21; dc.M22 = nM22; dc.Dx = nDx; dc.Dy = nDy;
                                }
                                else
                                {
                                    float nM11 = dc.M11 * m11 + dc.M12 * m21;
                                    float nM12 = dc.M11 * m12 + dc.M12 * m22;
                                    float nM21 = dc.M21 * m11 + dc.M22 * m21;
                                    float nM22 = dc.M21 * m12 + dc.M22 * m22;
                                    float nDx = dc.Dx * m11 + dc.Dy * m21 + dx;
                                    float nDy = dc.Dx * m12 + dc.Dy * m22 + dy;
                                    dc.M11 = nM11; dc.M12 = nM12; dc.M21 = nM21; dc.M22 = nM22; dc.Dx = nDx; dc.Dy = nDy;
                                }
                            }
                        }
                        break;

                    case EMR_SETWINDOWORGEX:
                        if (nSize >= 16)
                        {
                            dc.WinOrgX = BitConverter.ToInt32(emfData, pos + 8);
                            dc.WinOrgY = BitConverter.ToInt32(emfData, pos + 12);
                        }
                        break;

                    case EMR_SETWINDOWEXTEX:
                        if (nSize >= 16)
                        {
                            dc.WinExtX = BitConverter.ToInt32(emfData, pos + 8);
                            dc.WinExtY = BitConverter.ToInt32(emfData, pos + 12);
                        }
                        break;

                    case EMR_SETVIEWPORTORGEX:
                        if (nSize >= 16)
                        {
                            dc.VpOrgX = BitConverter.ToInt32(emfData, pos + 8);
                            dc.VpOrgY = BitConverter.ToInt32(emfData, pos + 12);
                        }
                        break;

                    case EMR_SETVIEWPORTEXTEX:
                        if (nSize >= 16)
                        {
                            dc.VpExtX = BitConverter.ToInt32(emfData, pos + 8);
                            dc.VpExtY = BitConverter.ToInt32(emfData, pos + 12);
                        }
                        break;

                    case EMR_SETTEXTALIGN:
                        if (nSize >= 12)
                        {
                            dc.TextAlign = BitConverter.ToUInt32(emfData, pos + 8);
                        }
                        break;

                    case EMR_CREATEFONTINDIRECTW:
                        if (nSize >= 104)
                        {
                            uint ihFont = BitConverter.ToUInt32(emfData, pos + 8);
                            int lfHeight = BitConverter.ToInt32(emfData, pos + 12);
                            int lfWeight = BitConverter.ToInt32(emfData, pos + 28);
                            byte lfCharSet = emfData[pos + 35];
                            string faceName = ReadNullTerminatedUnicode(emfData, pos + 40, 64);

                            FontEntry fe = new FontEntry();
                            fe.FaceName = string.IsNullOrEmpty(faceName) ? "Arial" : faceName;
                            fe.Height = Math.Abs(lfHeight) * scaleYToMm * 2.8346f; // Approximate points
                            if (fe.Height < 6f || fe.Height > 72f) fe.Height = 10f;
                            fe.Bold = lfWeight >= 700;
                            fe.CharSet = lfCharSet;
                            fonts[ihFont] = fe;
                        }
                        break;

                    case EMR_SELECTOBJECT:
                        if (nSize >= 12)
                        {
                            uint ihObject = BitConverter.ToUInt32(emfData, pos + 8);
                            if (fonts.ContainsKey(ihObject))
                            {
                                dc.SelectedFontHandle = ihObject;
                            }
                        }
                        break;

                    case EMR_EXTTEXTOUTW:
                    case EMR_EXTTEXTOUTA:
                        ParseExtTextOut(emfData, pos, (int)nSize, iType == EMR_EXTTEXTOUTW, dc, fonts, scaleXToMm, scaleYToMm, page);
                        break;

                    case EMR_SMALLTEXTOUT:
                        ParseSmallTextOut(emfData, pos, (int)nSize, dc, fonts, scaleXToMm, scaleYToMm, page);
                        break;
                }

                if (iType == EMR_EOF)
                    break;

                pos += (int)nSize;
            }

            // Merge adjacent character fragments if the application emitted character-by-character textouts
            MergeFragmentedTextElements(page);

            return page;
        }

        private void ParseExtTextOut(
            byte[] data, int recStart, int recSize, bool isUnicode,
            DcState dc, Dictionary<uint, FontEntry> fonts,
            float scaleXToMm, float scaleYToMm, ExtractedPage page)
        {
            // EMR_EXTTEXTOUTA/W structure:
            // +0: iType (4)
            // +4: nSize (4)
            // +8: rclBounds (16: Left, Top, Right, Bottom in device units)
            // +24: iGraphicsMode (4)
            // +28: exScale (4 float)
            // +32: eyScale (4 float)
            // +36: emrtext:
            //      +36: ptlReference (8: X, Y)
            //      +44: nChars (4)
            //      +48: offString (4)
            //      +52: fOptions (4)
            //      +56: rcl (16)
            //      +72: offDx (4)
            if (recSize < 76) return;

            int boundsLeft = BitConverter.ToInt32(data, recStart + 8);
            int boundsTop = BitConverter.ToInt32(data, recStart + 12);
            int boundsRight = BitConverter.ToInt32(data, recStart + 16);
            int boundsBottom = BitConverter.ToInt32(data, recStart + 20);

            int refX = BitConverter.ToInt32(data, recStart + 36);
            int refY = BitConverter.ToInt32(data, recStart + 40);
            uint nChars = BitConverter.ToUInt32(data, recStart + 44);
            uint offString = BitConverter.ToUInt32(data, recStart + 48);
            uint fOptions = BitConverter.ToUInt32(data, recStart + 52);

            if (nChars == 0 || offString < 36 || offString >= recSize)
                return;

            FontEntry font = GetCurrentFont(dc, fonts);
            string rawText = string.Empty;
            int strOffset = recStart + (int)offString;

            // Handle ETO_GLYPH_INDEX (0x0010): The text buffer contains TrueType font glyph indices!
            if ((fOptions & ETO_GLYPH_INDEX) != 0)
            {
                Dictionary<ushort, char> glyphMap = GetFontGlyphMap(font.FaceName, font.CharSet);
                StringBuilder sb = new StringBuilder((int)nChars);
                for (int i = 0; i < (int)nChars; i++)
                {
                    int off = strOffset + i * 2;
                    if (off + 2 <= recStart + recSize)
                    {
                        ushort gi = BitConverter.ToUInt16(data, off);
                        if (gi == 3) // Standard TrueType space glyph index
                        {
                            sb.Append(' ');
                        }
                        else if (glyphMap.ContainsKey(gi))
                        {
                            sb.Append(glyphMap[gi]);
                        }
                        else if (gi < 128)
                        {
                            sb.Append((char)gi);
                        }
                    }
                }
                rawText = sb.ToString();
            }
            else if (isUnicode && (fOptions & ETO_SMALL_CHARS) == 0)
            {
                int byteLen = (int)nChars * 2;
                if (strOffset + byteLen <= recStart + recSize)
                {
                    rawText = Encoding.Unicode.GetString(data, strOffset, byteLen);
                }
            }
            else
            {
                int byteLen = (int)nChars;
                if (strOffset + byteLen <= recStart + recSize)
                {
                    rawText = ArabicTextHelper.DecodeAnsiBytes(data, strOffset, byteLen, font.CharSet);
                }
            }

            AddTextElement(page, rawText, boundsLeft, boundsTop, boundsRight, boundsBottom, refX, refY, fOptions, dc, font, scaleXToMm, scaleYToMm);
        }

        private void ParseSmallTextOut(
            byte[] data, int recStart, int recSize,
            DcState dc, Dictionary<uint, FontEntry> fonts,
            float scaleXToMm, float scaleYToMm, ExtractedPage page)
        {
            if (recSize < 36) return;

            int refX = BitConverter.ToInt32(data, recStart + 8);
            int refY = BitConverter.ToInt32(data, recStart + 12);
            uint nChars = BitConverter.ToUInt32(data, recStart + 16);
            uint fuOptions = BitConverter.ToUInt32(data, recStart + 20);

            if (nChars == 0)
                return;

            int textOffset = recStart + 36;
            int boundsLeft = 0, boundsTop = 0, boundsRight = -1, boundsBottom = -1;

            if ((fuOptions & ETO_NO_RECT) == 0)
            {
                if (recSize < 52) return;
                boundsLeft = BitConverter.ToInt32(data, recStart + 36);
                boundsTop = BitConverter.ToInt32(data, recStart + 40);
                boundsRight = BitConverter.ToInt32(data, recStart + 44);
                boundsBottom = BitConverter.ToInt32(data, recStart + 48);
                textOffset = recStart + 52;
            }

            FontEntry font = GetCurrentFont(dc, fonts);
            string rawText = string.Empty;

            if ((fuOptions & ETO_GLYPH_INDEX) != 0)
            {
                Dictionary<ushort, char> glyphMap = GetFontGlyphMap(font.FaceName, font.CharSet);
                StringBuilder sb = new StringBuilder((int)nChars);
                for (int i = 0; i < (int)nChars; i++)
                {
                    int off = textOffset + i * 2;
                    if (off + 2 <= recStart + recSize)
                    {
                        ushort gi = BitConverter.ToUInt16(data, off);
                        if (gi == 3)
                        {
                            sb.Append(' ');
                        }
                        else if (glyphMap.ContainsKey(gi))
                        {
                            sb.Append(glyphMap[gi]);
                        }
                        else
                        {
                            sb.Append((char)gi);
                        }
                    }
                }
                rawText = sb.ToString();
            }
            else if ((fuOptions & ETO_SMALL_CHARS) != 0)
            {
                int byteLen = (int)nChars;
                if (textOffset + byteLen <= recStart + recSize)
                {
                    rawText = ArabicTextHelper.DecodeAnsiBytes(data, textOffset, byteLen, font.CharSet);
                }
            }
            else
            {
                int byteLen = (int)nChars * 2;
                if (textOffset + byteLen <= recStart + recSize)
                {
                    rawText = Encoding.Unicode.GetString(data, textOffset, byteLen);
                }
            }

            AddTextElement(page, rawText, boundsLeft, boundsTop, boundsRight, boundsBottom, refX, refY, fuOptions, dc, font, scaleXToMm, scaleYToMm);
        }

        private void AddTextElement(
            ExtractedPage page, string rawText,
            int boundsLeft, int boundsTop, int boundsRight, int boundsBottom,
            int refX, int refY, uint fOptions,
            DcState dc, FontEntry font,
            float scaleXToMm, float scaleYToMm)
        {
            if (string.IsNullOrEmpty(rawText)) return;

            string text = ArabicTextHelper.NormalizeArabicText(rawText, NormalizePresentationForms, ReverseVisualArabic);
            if (string.IsNullOrEmpty(text.Trim())) return;

            float xMm, yMm, widthMm, heightMm;

            // Prefer device rclBounds when valid (non-empty rectangle)
            if (boundsRight >= boundsLeft && boundsBottom >= boundsTop && (boundsRight > 0 || boundsBottom > 0))
            {
                xMm = boundsLeft * scaleXToMm;
                yMm = boundsTop * scaleYToMm;
                widthMm = Math.Max(2.0f, (boundsRight - boundsLeft) * scaleXToMm);
                heightMm = Math.Max(2.5f, (boundsBottom - boundsTop) * scaleYToMm);
            }
            else
            {
                // Transform logical reference point through WorldTransform + Window/Viewport mapping
                float worldX = refX * dc.M11 + refY * dc.M21 + dc.Dx;
                float worldY = refX * dc.M12 + refY * dc.M22 + dc.Dy;

                float devX = worldX;
                float devY = worldY;
                if (dc.WinExtX != 0 && dc.VpExtX != 0)
                    devX = (worldX - dc.WinOrgX) * ((float)dc.VpExtX / dc.WinExtX) + dc.VpOrgX;
                if (dc.WinExtY != 0 && dc.VpExtY != 0)
                    devY = (worldY - dc.WinOrgY) * ((float)dc.VpExtY / dc.WinExtY) + dc.VpOrgY;

                xMm = devX * scaleXToMm;
                yMm = devY * scaleYToMm;
                heightMm = Math.Max(3.0f, font.Height * 0.3527f);
                widthMm = Math.Max(3.0f, text.Length * heightMm * 0.55f);

                if ((dc.TextAlign & TA_CENTER) == TA_CENTER)
                    xMm -= widthMm / 2.0f;
                else if ((dc.TextAlign & TA_RIGHT) == TA_RIGHT)
                    xMm -= widthMm;
            }

            bool isRtl = ((fOptions & ETO_RTLREADING) != 0) ||
                         ((dc.TextAlign & TA_RTLREADING) != 0) ||
                         ArabicTextHelper.ContainsArabic(text);

            ExtractedTextElement elem = new ExtractedTextElement();
            elem.PageNumber = page.PageNumber;
            elem.Text = text.Trim();
            elem.X = xMm;
            elem.Y = yMm;
            elem.Width = widthMm;
            elem.Height = heightMm;
            elem.FontName = font.FaceName;
            elem.FontSize = font.Height;
            elem.IsBold = font.Bold;
            elem.IsRtl = isRtl;
            elem.Selected = true;

            page.Elements.Add(elem);
        }

        private FontEntry GetCurrentFont(DcState dc, Dictionary<uint, FontEntry> fonts)
        {
            FontEntry font;
            if (dc.SelectedFontHandle != 0 && fonts.TryGetValue(dc.SelectedFontHandle, out font))
                return font;
            return new FontEntry();
        }

        /// <summary>
        /// Some GDI programs print strings one character or word fragment at a time.
        /// This merges fragments that sit on the same baseline and are immediately adjacent horizontally.
        /// When a fragment cluster contains Arabic, it merges them in Right-to-Left logical reading order.
        /// </summary>
        private void MergeFragmentedTextElements(ExtractedPage page)
        {
            if (page.Elements.Count <= 1) return;

            // Sort top-to-bottom, then left-to-right
            page.Elements.Sort(delegate(ExtractedTextElement a, ExtractedTextElement b)
            {
                if (Math.Abs(a.CenterY - b.CenterY) <= Math.Max(1.5f, Math.Max(a.Height, b.Height) * 0.55f))
                    return a.X.CompareTo(b.X);
                return a.CenterY.CompareTo(b.CenterY);
            });

            List<ExtractedTextElement> merged = new List<ExtractedTextElement>();
            List<ExtractedTextElement> currentCluster = new List<ExtractedTextElement>();
            currentCluster.Add(page.Elements[0]);

            for (int i = 1; i < page.Elements.Count; i++)
            {
                ExtractedTextElement next = page.Elements[i];
                ExtractedTextElement lastInCluster = currentCluster[currentCluster.Count - 1];

                bool sameLine = Math.Abs(lastInCluster.CenterY - next.CenterY) <= Math.Max(1.5f, Math.Max(lastInCluster.Height, next.Height) * 0.55f);
                float gap = next.X - lastInCluster.Right;

                // Merge if fragments sit on the same line and are touching or adjacent (gap between -4.0mm and 0.6mm)
                if (sameLine && gap >= -4.0f && gap <= 0.6f)
                {
                    currentCluster.Add(next);
                }
                else
                {
                    merged.Add(MergeCluster(currentCluster));
                    currentCluster.Clear();
                    currentCluster.Add(next);
                }
            }

            if (currentCluster.Count > 0)
            {
                merged.Add(MergeCluster(currentCluster));
            }

            page.Elements = merged;
        }

        private ExtractedTextElement MergeCluster(List<ExtractedTextElement> cluster)
        {
            if (cluster.Count == 1)
                return cluster[0];

            ExtractedTextElement first = cluster[0];
            ExtractedTextElement res = new ExtractedTextElement();
            res.PageNumber = first.PageNumber;
            res.FontName = first.FontName;
            res.FontSize = first.FontSize;
            res.IsBold = first.IsBold;

            float minX = first.X;
            float maxRight = first.Right;
            float maxHeight = first.Height;
            float avgY = 0f;

            bool hasArabic = false;
            for (int i = 0; i < cluster.Count; i++)
            {
                ExtractedTextElement el = cluster[i];
                if (el.X < minX) minX = el.X;
                if (el.Right > maxRight) maxRight = el.Right;
                if (el.Height > maxHeight) maxHeight = el.Height;
                avgY += el.Y;
                if (el.IsRtl || ArabicTextHelper.ContainsArabic(el.Text))
                    hasArabic = true;
            }
            res.X = minX;
            res.Width = maxRight - minX;
            res.Y = avgY / cluster.Count;
            res.Height = maxHeight;
            res.IsRtl = hasArabic;

            StringBuilder sb = new StringBuilder();
            if (hasArabic)
            {
                // In Right-to-Left Arabic text, the rightmost fragment was printed first in logical reading order
                for (int i = cluster.Count - 1; i >= 0; i--)
                {
                    string t = cluster[i].Text;
                    if (string.IsNullOrEmpty(t)) continue;

                    if (sb.Length > 0)
                    {
                        ExtractedTextElement prevRight = cluster[i + 1];
                        ExtractedTextElement currLeft = cluster[i];
                        float gap = prevRight.X - currLeft.Right;
                        bool isNumericContinuation = (char.IsDigit(currLeft.Text[currLeft.Text.Length - 1]) || currLeft.Text.EndsWith(",") || currLeft.Text.EndsWith(".") || currLeft.Text.EndsWith("/") || currLeft.Text.EndsWith("-")) &&
                                                    (char.IsDigit(prevRight.Text[0]) || prevRight.Text.StartsWith(",") || prevRight.Text.StartsWith(".") || prevRight.Text.StartsWith("/") || prevRight.Text.StartsWith("-"));
                        if (gap > 0.5f && !isNumericContinuation)
                            sb.Append(" ");
                    }
                    sb.Append(t);
                }
            }
            else
            {
                // Standard LTR reading order
                for (int i = 0; i < cluster.Count; i++)
                {
                    string t = cluster[i].Text;
                    if (string.IsNullOrEmpty(t)) continue;

                    if (sb.Length > 0)
                    {
                        ExtractedTextElement prev = cluster[i - 1];
                        ExtractedTextElement curr = cluster[i];
                        float gap = curr.X - prev.Right;
                        bool isNumericContinuation = (char.IsDigit(prev.Text[prev.Text.Length - 1]) || prev.Text.EndsWith(",") || prev.Text.EndsWith(".") || prev.Text.EndsWith("/") || prev.Text.EndsWith("-")) &&
                                                    (char.IsDigit(curr.Text[0]) || curr.Text.StartsWith(",") || curr.Text.StartsWith(".") || curr.Text.StartsWith("/") || curr.Text.StartsWith("-"));
                        if (gap > 0.5f && !isNumericContinuation)
                            sb.Append(" ");
                    }
                    sb.Append(t);
                }
            }

            res.Text = ArabicTextHelper.NormalizeArabicText(sb.ToString(), NormalizePresentationForms, ReverseVisualArabic);
            return res;
        }

        /// <summary>
        /// Fallback parser in case a classic app sends raw text / line-printer data instead of EMF.
        /// Supports UTF-8, UTF-16, and Windows-1256 (Arabic ANSI).
        /// </summary>
        private ExtractedPage ParseRawTextFallback(byte[] data, int pageNumber)
        {
            ExtractedPage page = new ExtractedPage(pageNumber);
            string content;
            if (data.Length >= 2 && data[0] == 0xFF && data[1] == 0xFE)
                content = Encoding.Unicode.GetString(data, 2, data.Length - 2);
            else
                content = ArabicTextHelper.DecodeAnsiBytes(data, 0, data.Length, 178); // 1256 Arabic

            string[] lines = content.Replace("\r\n", "\n").Split('\n');
            float y = 10f;
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i];
                if (string.IsNullOrEmpty(line.Trim()))
                {
                    y += 5f;
                    continue;
                }

                // Split columns by 2+ spaces or tabs
                int colStart = 0;
                while (colStart < line.Length)
                {
                    while (colStart < line.Length && (line[colStart] == ' ' || line[colStart] == '\t'))
                        colStart++;
                    if (colStart >= line.Length) break;

                    int colEnd = colStart;
                    while (colEnd < line.Length)
                    {
                        if (line[colEnd] == '\t') break;
                        if (colEnd + 1 < line.Length && line[colEnd] == ' ' && line[colEnd + 1] == ' ') break;
                        colEnd++;
                    }

                    string cellText = line.Substring(colStart, colEnd - colStart).Trim();
                    if (!string.IsNullOrEmpty(cellText))
                    {
                        ExtractedTextElement elem = new ExtractedTextElement();
                        elem.PageNumber = pageNumber;
                        elem.Text = ArabicTextHelper.NormalizeArabicText(cellText, NormalizePresentationForms, ReverseVisualArabic);
                        elem.X = 10f + colStart * 2.2f;
                        elem.Y = y;
                        elem.Width = cellText.Length * 2.2f;
                        elem.Height = 4.5f;
                        elem.IsRtl = ArabicTextHelper.ContainsArabic(elem.Text);
                        page.Elements.Add(elem);
                    }
                    colStart = colEnd;
                }
                y += 5f;
            }

            return page;
        }

        private static string ReadNullTerminatedUnicode(byte[] data, int offset, int maxBytes)
        {
            int end = offset;
            int limit = Math.Min(data.Length - 1, offset + maxBytes);
            while (end < limit)
            {
                if (data[end] == 0 && data[end + 1] == 0)
                    break;
                end += 2;
            }
            if (end <= offset) return string.Empty;
            return Encoding.Unicode.GetString(data, offset, end - offset);
        }
    }
}
