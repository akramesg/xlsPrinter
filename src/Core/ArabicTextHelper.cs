using System;
using System.Collections.Generic;
using System.Text;

namespace XlsPrinter.Core
{
    /// <summary>
    /// Handles Arabic character encoding (Windows-1256, UTF-16LE), Arabic Presentation Forms
    /// (U+FE70..U+FEFF) normalization to standard Unicode Arabic (U+0600..U+06FF),
    /// and visual-order vs. logical-order Right-to-Left (RTL) text reconstruction.
    /// </summary>
    public static class ArabicTextHelper
    {
        // Mapping from Arabic Presentation Forms-B (U+FE70 - U+FEFC) and Forms-A (common ligatures)
        // back to standard canonical Arabic characters (U+0600 - U+06FF) so Excel can search,
        // sort, and calculate formulas properly.
        private static readonly Dictionary<char, string> PresentationFormMap = BuildPresentationFormMap();

        private static Dictionary<char, string> BuildPresentationFormMap()
        {
            Dictionary<char, string> map = new Dictionary<char, string>();

            // Standard Arabic Presentation Forms-B (FE80 - FEFC)
            AddMapping(map, '\u0621', '\uFE80'); // Hamza
            AddMapping(map, '\u0622', '\uFE81', '\uFE82'); // Alef with Madda
            AddMapping(map, '\u0623', '\uFE83', '\uFE84'); // Alef with Hamza Above
            AddMapping(map, '\u0624', '\uFE85', '\uFE86'); // Waw with Hamza
            AddMapping(map, '\u0625', '\uFE87', '\uFE88'); // Alef with Hamza Below
            AddMapping(map, '\u0626', '\uFE89', '\uFE8A', '\uFE8B', '\uFE8C'); // Yeh with Hamza
            AddMapping(map, '\u0627', '\uFE8D', '\uFE8E'); // Alef
            AddMapping(map, '\u0628', '\uFE8F', '\uFE90', '\uFE91', '\uFE92'); // Beh
            AddMapping(map, '\u0629', '\uFE93', '\uFE94'); // Teh Marbuta
            AddMapping(map, '\u062A', '\uFE95', '\uFE96', '\uFE97', '\uFE98'); // Teh
            AddMapping(map, '\u062B', '\uFE99', '\uFE9A', '\uFE9B', '\uFE9C'); // Theh
            AddMapping(map, '\u062C', '\uFE9D', '\uFE9E', '\uFE9F', '\uFEA0'); // Jeem
            AddMapping(map, '\u062D', '\uFEA1', '\uFEA2', '\uFEA3', '\uFEA4'); // Hah
            AddMapping(map, '\u062E', '\uFEA5', '\uFEA6', '\uFEA7', '\uFEA8'); // Khah
            AddMapping(map, '\u062F', '\uFEA9', '\uFEAA'); // Dal
            AddMapping(map, '\u0630', '\uFEAB', '\uFEAC'); // Thal
            AddMapping(map, '\u0631', '\uFEAD', '\uFEAE'); // Reh
            AddMapping(map, '\u0632', '\uFEAF', '\uFEB0'); // Zain
            AddMapping(map, '\u0633', '\uFEB1', '\uFEB2', '\uFEB3', '\uFEB4'); // Seen
            AddMapping(map, '\u0634', '\uFEB5', '\uFEB6', '\uFEB7', '\uFEB8'); // Sheen
            AddMapping(map, '\u0635', '\uFEB9', '\uFEBA', '\uFEBB', '\uFEBC'); // Sad
            AddMapping(map, '\u0636', '\uFEBD', '\uFEBE', '\uFEBF', '\uFEC0'); // Dad
            AddMapping(map, '\u0637', '\uFEC1', '\uFEC2', '\uFEC3', '\uFEC4'); // Tah
            AddMapping(map, '\u0638', '\uFEC5', '\uFEC6', '\uFEC7', '\uFEC8'); // Zah
            AddMapping(map, '\u0639', '\uFEC9', '\uFECA', '\uFECB', '\uFECC'); // Ain
            AddMapping(map, '\u063A', '\uFECD', '\uFECE', '\uFECF', '\uFED0'); // Ghain
            AddMapping(map, '\u0641', '\uFED1', '\uFED2', '\uFED3', '\uFED4'); // Feh
            AddMapping(map, '\u0642', '\uFED5', '\uFED6', '\uFED7', '\uFED8'); // Qaf
            AddMapping(map, '\u0643', '\uFED9', '\uFEDA', '\uFEDB', '\uFEDC'); // Kaf
            AddMapping(map, '\u0644', '\uFEDD', '\uFEDE', '\uFEDF', '\uFEE0'); // Lam
            AddMapping(map, '\u0645', '\uFEE1', '\uFEE2', '\uFEE3', '\uFEE4'); // Meem
            AddMapping(map, '\u0646', '\uFEE5', '\uFEE6', '\uFEE7', '\uFEE8'); // Noon
            AddMapping(map, '\u0647', '\uFEE9', '\uFEEA', '\uFEEB', '\uFEEC'); // Heh
            AddMapping(map, '\u0648', '\uFEED', '\uFEEE'); // Waw
            AddMapping(map, '\u0649', '\uFEEF', '\uFEF0'); // Alef Maksura
            AddMapping(map, '\u064A', '\uFEF1', '\uFEF2', '\uFEF3', '\uFEF4'); // Yeh

            // Persian / Urdu / Extended Arabic Presentation Forms-A common chars
            AddMapping(map, '\u067E', '\uFB56', '\uFB57', '\uFB58', '\uFB59'); // Peh
            AddMapping(map, '\u0686', '\uFB7A', '\uFB7B', '\uFB7C', '\uFB7D'); // Tcheh
            AddMapping(map, '\u0698', '\uFB8A', '\uFB8B'); // Jeh
            AddMapping(map, '\u06AF', '\uFB92', '\uFB93', '\uFB94', '\uFB95'); // Gaf
            AddMapping(map, '\u06A9', '\uFB8E', '\uFB8F', '\uFB90', '\uFB91'); // Keheh
            AddMapping(map, '\u06CC', '\uFBFD', '\uFBFE', '\uFBFF'); // Farsi Yeh

            // Lam-Alef Ligatures (FEF5 - FEFC)
            map['\uFEF5'] = "\u0644\u0622"; // Lam + Alef Madda Isolated
            map['\uFEF6'] = "\u0644\u0622"; // Lam + Alef Madda Final
            map['\uFEF7'] = "\u0644\u0623"; // Lam + Alef Hamza Above Isolated
            map['\uFEF8'] = "\u0644\u0623"; // Lam + Alef Hamza Above Final
            map['\uFEF9'] = "\u0644\u0625"; // Lam + Alef Hamza Below Isolated
            map['\uFEFA'] = "\u0644\u0625"; // Lam + Alef Hamza Below Final
            map['\uFEFB'] = "\u0644\u0627"; // Lam + Alef Isolated
            map['\uFEFC'] = "\u0644\u0627"; // Lam + Alef Final

            // Allah ligature
            map['\uFDF2'] = "\u0627\u0644\u0644\u0647";

            return map;
        }

        private static void AddMapping(Dictionary<char, string> map, char canonical, params char[] forms)
        {
            string target = canonical.ToString();
            for (int i = 0; i < forms.Length; i++)
            {
                map[forms[i]] = target;
            }
        }

        /// <summary>
        /// Decodes an 8-bit ANSI byte array from EMR_EXTTEXTOUTA using Windows-1256 (Arabic)
        /// or the specified GDI lfCharSet.
        /// </summary>
        public static string DecodeAnsiBytes(byte[] bytes, int offset, int count, byte lfCharSet)
        {
            if (bytes == null || count <= 0 || offset + count > bytes.Length)
                return string.Empty;

            int codePage = 1256; // Default to Arabic codepage 1256 for high-bit bytes
            // GDI Charset constants:
            // ARABIC_CHARSET = 178
            // HEBREW_CHARSET = 177
            // GREEK_CHARSET = 161
            // TURKISH_CHARSET = 162
            // VIETNAMESE_CHARSET = 163
            // BALTIC_CHARSET = 186
            //EASTEUROPE_CHARSET = 238
            // RUSSIAN_CHARSET = 204
            if (lfCharSet == 161) codePage = 1253;
            else if (lfCharSet == 162) codePage = 1254;
            else if (lfCharSet == 177) codePage = 1255;
            else if (lfCharSet == 178) codePage = 1256;
            else if (lfCharSet == 186) codePage = 1257;
            else if (lfCharSet == 204) codePage = 1251;
            else if (lfCharSet == 238) codePage = 1250;

            try
            {
                Encoding enc = Encoding.GetEncoding(codePage);
                return enc.GetString(bytes, offset, count);
            }
            catch
            {
                return Encoding.Default.GetString(bytes, offset, count);
            }
        }

        /// <summary>
        /// Returns true if the character is in any Arabic Unicode block.
        /// </summary>
        public static bool IsArabicChar(char c)
        {
            return (c >= '\u0600' && c <= '\u06FF') ||
                   (c >= '\u0750' && c <= '\u077F') ||
                   (c >= '\u08A0' && c <= '\u08FF') ||
                   (c >= '\uFB50' && c <= '\uFDFF') ||
                   (c >= '\uFE70' && c <= '\uFEFF');
        }

        /// <summary>
        /// Returns true if the string contains any Arabic character.
        /// </summary>
        public static bool ContainsArabic(string text)
        {
            if (string.IsNullOrEmpty(text)) return false;
            for (int i = 0; i < text.Length; i++)
            {
                if (IsArabicChar(text[i])) return true;
            }
            return false;
        }

        /// <summary>
        /// Returns true if the character is in Arabic Presentation Forms A or B (shaped glyphs).
        /// Classic GDI apps often emit shaped glyphs in visual (left-to-right) order when printing.
        /// </summary>
        public static bool ContainsPresentationForms(string text)
        {
            if (string.IsNullOrEmpty(text)) return false;
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if ((c >= '\uFB50' && c <= '\uFDFF') || (c >= '\uFE70' && c <= '\uFEFF'))
                    return true;
            }
            return false;
        }

        /// <summary>
        /// Returns true if the Arabic text is detected to be in visual (inverted/reversed) order.
        /// Detects both presentation-form reversed glyph sequences and canonical Unicode inverted words
        /// (e.g. words starting with Teh Marbuta, words ending in 'لا', or reversed common prepositions).
        /// </summary>
        public static bool IsVisualArabic(string text)
        {
            if (string.IsNullOrEmpty(text) || !ContainsArabic(text))
                return false;

            int visualScore = 0;
            int logicalScore = 0;

            // 1. Check Presentation Forms characteristics
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                // Final forms
                bool isFinal = (c >= '\uFE8E' && c <= '\uFEF0' && ((c - 0xFE80) % 2 == 0)) || c == '\uFE94';
                // Initial forms
                bool isInitial = (c >= '\uFE8F' && c <= '\uFEF3' && ((c - 0xFE80) % 2 == 1));

                if (isFinal && i + 1 < text.Length)
                {
                    char next = text[i + 1];
                    bool nextIsInitial = (next >= '\uFE8F' && next <= '\uFEF3' && ((next - 0xFE80) % 2 == 1));
                    if (nextIsInitial) visualScore += 2; // Final before Initial is strong indicator of visual LTR
                }
                if (isInitial && i + 1 < text.Length)
                {
                    char next = text[i + 1];
                    bool nextIsFinal = (next >= '\uFE8E' && next <= '\uFEF0' && ((next - 0xFE80) % 2 == 0));
                    if (nextIsFinal) logicalScore += 2;
                }
            }

            // 2. Token-level analysis (handles canonical Unicode as well as presentation forms)
            string[] words = text.Split(new char[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            for (int w = 0; w < words.Length; w++)
            {
                string word = words[w].Trim();
                if (word.Length < 2) continue;

                // Teh Marbuta ('ة' or '\uFE93' / '\uFE94') NEVER begins a word in Arabic.
                // If a word starts with Teh Marbuta, it is 100% reversed!
                if (word[0] == '\u0629' || word[0] == '\uFE93' || word[0] == '\uFE94')
                {
                    visualScore += 5;
                }
                if (word[word.Length - 1] == '\u0629' || word[word.Length - 1] == '\uFE93' || word[word.Length - 1] == '\uFE94')
                {
                    logicalScore += 3;
                }

                // Definite article 'ال' (U+0627 U+0644) always begins words in Arabic.
                // In visual order, it appears as 'لا' (U+0644 U+0627) at the END of words.
                if (word.Length >= 3 && word.EndsWith("\u0644\u0627"))
                {
                    visualScore += 3;
                }
                if (word.Length >= 3 && word.StartsWith("\u0627\u0644"))
                {
                    logicalScore += 3;
                }

                // Common reversed Arabic keywords vs logical keywords
                if (word == "نم" || word == "ىتح" || word == "ىلا" || word == "ىلع" || word == "وه" ||
                    word == "ةنس" || word == "رهش" || word == "زاغ" || word == "خيرات" || word == "رعس" ||
                    word == "طسوتم" || word == "تابوحسم" || word == "كالهتسا")
                {
                    visualScore += 4;
                }
                if (word == "من" || word == "حتى" || word == "الى" || word == "إلى" || word == "على" ||
                    word == "سنة" || word == "شهر" || word == "غاز" || word == "تاريخ" || word == "سعر" ||
                    word == "متوسط" || word == "مسحوبات" || word == "استهلاك")
                {
                    logicalScore += 4;
                }
            }

            return visualScore > logicalScore;
        }

        /// <summary>
        /// Returns true if the string is clearly in standard logical Right-to-Left order.
        /// </summary>
        public static bool IsStrictlyLogicalArabic(string text)
        {
            if (string.IsNullOrEmpty(text) || !ContainsArabic(text))
                return false;

            string[] words = text.Split(new char[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            int logicalHits = 0;
            int visualHits = 0;

            for (int w = 0; w < words.Length; w++)
            {
                string word = words[w].Trim();
                if (word.Length < 2) continue;

                if (word[0] == '\u0629' || word[0] == '\uFE93' || word[0] == '\uFE94' || word.EndsWith("\u0644\u0627"))
                    visualHits++;

                if (word.StartsWith("\u0627\u0644") || word.EndsWith("\u0629") ||
                    word == "من" || word == "حتى" || word == "الى" || word == "إلى" || word == "على" ||
                    word == "سنة" || word == "شهر" || word == "غاز" || word == "تاريخ" || word == "استهلاك")
                    logicalHits++;
            }

            return logicalHits > visualHits && visualHits == 0;
        }

        /// <summary>
        /// Normalizes Arabic text:
        /// 1. Converts Arabic Presentation Forms (FE70-FEFF) to standard canonical Unicode (0600-06FF).
        /// 2. If reverseVisualArabic is enabled (or automatically when input was stored in visual LTR order),
        ///    reverses the characters while keeping numbers, dates, and Latin words in proper order.
        /// 3. Cleans control characters and null terminators.
        /// </summary>
        public static string NormalizeArabicText(string input, bool normalizePresentationForms, bool reverseVisualArabic)
        {
            if (string.IsNullOrEmpty(input))
                return string.Empty;

            // Strip null chars or non-printable control chars (except tab/newline)
            StringBuilder cleanSb = new StringBuilder(input.Length);
            for (int i = 0; i < input.Length; i++)
            {
                char c = input[i];
                if (c == '\0') continue;
                if (char.IsControl(c) && c != '\t' && c != '\r' && c != '\n') continue;
                cleanSb.Append(c);
            }
            string cleaned = cleanSb.ToString();

            // Reverse if explicitly requested (unless strictly already logical) or automatically if visual Arabic detected
            if (ContainsArabic(cleaned))
            {
                bool shouldReverse = reverseVisualArabic ? !IsStrictlyLogicalArabic(cleaned) : IsVisualArabic(cleaned);
                if (shouldReverse)
                {
                    cleaned = ReverseArabicRuns(cleaned);
                }
            }

            if (normalizePresentationForms && ContainsPresentationForms(cleaned))
            {
                StringBuilder normSb = new StringBuilder(cleaned.Length);
                for (int i = 0; i < cleaned.Length; i++)
                {
                    char c = cleaned[i];
                    string mapped;
                    if (PresentationFormMap.TryGetValue(c, out mapped))
                    {
                        normSb.Append(mapped);
                    }
                    else
                    {
                        normSb.Append(c);
                    }
                }
                cleaned = normSb.ToString();
            }

            // If reverseVisualArabic was requested and after presentation forms normalization
            // the canonical Arabic text is still detected as visual order, reverse it
            if (reverseVisualArabic && ContainsArabic(cleaned) && IsVisualArabic(cleaned))
            {
                cleaned = ReverseArabicRuns(cleaned);
            }

            return cleaned;
        }

        public static string ReverseArabicRuns(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;

            // Break string into directional tokens (LTR words/numbers/dates vs Arabic text words vs spaces)
            List<string> tokens = new List<string>();
            int i = 0;
            while (i < text.Length)
            {
                if (IsLtrRunChar(text[i]))
                {
                    int start = i;
                    while (i < text.Length && (IsLtrRunChar(text[i]) || ((text[i] == '.' || text[i] == ',' || text[i] == '/' || text[i] == '-' || text[i] == ':') && i + 1 < text.Length && IsLtrRunChar(text[i + 1]))))
                    {
                        i++;
                    }
                    tokens.Add(text.Substring(start, i - start));
                }
                else if (text[i] == ' ' || text[i] == '\t')
                {
                    tokens.Add(text[i].ToString());
                    i++;
                }
                else
                {
                    int start = i;
                    while (i < text.Length && !IsLtrRunChar(text[i]) && text[i] != ' ' && text[i] != '\t')
                    {
                        i++;
                    }
                    // Reverse letter order within the Arabic token
                    char[] w = text.Substring(start, i - start).ToCharArray();
                    Array.Reverse(w);
                    tokens.Add(new string(w));
                }
            }

            // Reverse token list so reading order becomes Right-to-Left logical
            tokens.Reverse();
            return string.Join("", tokens.ToArray());
        }

        private static bool IsLtrRunChar(char c)
        {
            return (c >= '0' && c <= '9') ||
                   (c >= '\u0660' && c <= '\u0669') || // Arabic-Indic digits
                   (c >= 'A' && c <= 'Z') ||
                   (c >= 'a' && c <= 'z');
        }
    }
}
