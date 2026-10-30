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

        private static readonly HashSet<string> LogicalKeywords = new HashSet<string>(StringComparer.Ordinal)
        {
            "من", "حتى", "الى", "إلى", "على", "هو", "في", "عن", "مع", "أو", "ثم",
            "سنة", "شهر", "يوم", "تاريخ", "سعر", "رقم", "حساب", "فاتورة", "كمية",
            "إجمالي", "اجمالي", "مجموع", "عميل", "زبون", "ريال", "دينار", "جنيه", "جنية", "دولار",
            "خصم", "ضريبة", "صافي", "مدين", "دائن", "رصيد", "سند", "قبض", "صرف",
            "فرع", "مركز", "اسم", "بيان", "صنف", "وحدة", "كود", "صندوق", "كشف",
            "وحدات", "تخزين", "احتياطي", "كابلات", "معدات", "ربط", "أجهزة", "طابعات",
            "شاشات", "خدمات", "مؤسسة", "شركة", "تقرير", "فترة", "نقد", "شيك", "بنك", "دفعة",
            "ورق", "خام", "عدد", "كجم", "اكرم", "إكرم", "أكرم", "طباعة", "جالسين", "ميكرون",
            "دبل", "كرتون", "ابيض", "مواسير", "كيماويات", "تصوير", "احبار", "واحد", "متر",
            "بصمة", "سم", "مللي", "ايباك", "سادة", "ساده", "داخلي", "داخلى", "علبة", "اكلاشيهات",
            "مجهز", "معالج", "مسلكن", "شريط", "لاصق", "شفاف", "اوفست", "ميكرو", "افلام", "معالجة"
        };

        private static readonly HashSet<string> VisualKeywords = new HashSet<string>(StringComparer.Ordinal)
        {
            "نم", "ىتح", "ىلا", "ىلإ", "ىلع", "وه", "يف", "نع",
            "ةنس", "رهش", "موي", "خيرات", "رعس", "مقر", "باسح", "ةروتاف", "ةيمك",
            "يلامجا", "يلامجإ", "عومجم", "ليمع", "نوبز", "لاير", "رانيد", "هينج", "ةينج", "رالود",
            "مصح", "ةبيرض", "يفاص", "نيدم", "نئاد", "ديصر", "دنس", "ضبق", "فرص",
            "عرف", "زكرم", "مسا", "نايب", "فنص", "ةدحو", "دوك", "قودنص", "فشك",
            "تادحو", "نيزخت", "يطايتحا", "تالباك", "تادعم", "طبر", "ةزهجا", "تاعابط",
            "تاشاش", "تامدخ", "ةسسوم", "ةكرش", "ريرقت", "ةرتف", "دقن", "كيش", "كنب", "ةعفد",
            "قرو", "ماخ", "ددع", "مجك", "مركا", "مركإ", "مركأ", "ةعابط", "نیسلاج", "نوركیم",
            "لبد", "نوترك", "ضيبا", "ريساوم", "تايواميك", "ريوصت", "رابحا", "دحاو", "رتم",
            "همصب", "مس", "يللم", "كابيا", "هداس", "يلخاد", "ةبلع", "فافش", "قصلا", "طيرش"
        };

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
            AddMapping(map, '\u0643', '\uFB8E', '\uFB8F', '\uFB90', '\uFB91'); // Keheh -> standard Arabic Kaf
            AddMapping(map, '\u064A', '\uFBAE', '\uFBAF', '\uFBFD', '\uFBFE', '\uFBFF'); // Yeh forms -> standard Arabic Yeh
            AddMapping(map, '\u0647', '\uFBA6', '\uFBA7', '\uFBA8', '\uFBA9', '\uFBAA', '\uFBAC', '\uFBAD'); // Heh Goal / Doachashmee
            map['\uFBAB'] = "\u0629"; // Teh Marbuta isolated (widely used in Arabic Windows ERP printing)

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
        /// Detects presentation forms, canonical Unicode inverted words
        /// (e.g. words starting with Teh Marbuta or Alef Maksura, words ending in 'لا', or reversed common keywords).
        /// </summary>
        public static bool IsVisualArabic(string text)
        {
            if (string.IsNullOrEmpty(text) || !ContainsArabic(text))
                return false;

            // 1. Any string containing presentation forms was pre-shaped for visual LTR rendering
            if (ContainsPresentationForms(text))
                return true;

            int visualScore = 0;
            int logicalScore = 0;

            // 2. Token-level analysis (handles canonical Unicode as well as presentation forms)
            string[] words = text.Split(new char[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            for (int w = 0; w < words.Length; w++)
            {
                string raw = words[w].Trim();
                string word = raw.Trim(' ', '\t', '\r', '\n', ':', '-', '/', '\\', '.', ',', '(', ')', '[', ']', '{', '}', '%', '#', '@', '*', '"', '\'', '؟', '،', '؛');
                if (word.Length < 2) continue;

                // Teh Marbuta or Alef Maksura NEVER begins a word in Arabic.
                // If a word starts with Teh Marbuta or Alef Maksura, it is 100% reversed!
                if (word[0] == '\u0629' || word[0] == '\uFE93' || word[0] == '\uFE94' || word[0] == '\uFBAB' ||
                    word[0] == '\u0649' || word[0] == '\uFEEF' || word[0] == '\uFEF0')
                {
                    visualScore += 5;
                }

                // Teh Marbuta or Alef Maksura at end of word -> strong indicator of logical order
                if (word[word.Length - 1] == '\u0629' || word[word.Length - 1] == '\uFE93' || word[word.Length - 1] == '\uFE94' || word[word.Length - 1] == '\uFBAB')
                {
                    logicalScore += 3;
                }
                else if (word[word.Length - 1] == '\u0649' || word[word.Length - 1] == '\uFEEF' || word[word.Length - 1] == '\uFEF0')
                {
                    logicalScore += 2;
                }

                // Definite article prefixes (ال, لل, وال, بال, فال, كال) always begin words in logical Arabic.
                if (word.Length >= 3 && (word.StartsWith("\u0627\u0644", StringComparison.Ordinal) ||
                                         word.StartsWith("\u0644\u0644", StringComparison.Ordinal) ||
                                         word.StartsWith("\u0648\u0627\u0644", StringComparison.Ordinal) ||
                                         word.StartsWith("\u0628\u0627\u0644", StringComparison.Ordinal) ||
                                         word.StartsWith("\u0641\u0627\u0644", StringComparison.Ordinal) ||
                                         word.StartsWith("\u0643\u0627\u0644", StringComparison.Ordinal)))
                {
                    logicalScore += 3;
                }

                // In visual order, the definite article appears as 'لا' (U+0644 U+0627) at the END of words.
                if (word.Length >= 3 && (word.EndsWith("\u0644\u0627", StringComparison.Ordinal) || word.EndsWith("\uFEFB", StringComparison.Ordinal)) &&
                    word != "إلا" && word != "كلا" && word != "لولا" && word != "هلا")
                {
                    visualScore += 4;
                }

                // Common reversed Arabic keywords vs logical keywords (including single-letter prefix stripping)
                string rootWord = word;
                if (word.Length >= 4 && (word[0] == 'و' || word[0] == 'ف' || word[0] == 'ب' || word[0] == 'ل' || word[0] == 'ك'))
                {
                    rootWord = word.Substring(1);
                }

                if (VisualKeywords.Contains(word) || VisualKeywords.Contains(rootWord))
                {
                    visualScore += 4;
                }
                if (LogicalKeywords.Contains(word) || LogicalKeywords.Contains(rootWord))
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

            int logicalHits = 0;
            int visualHits = 0;

            string[] words = text.Split(new char[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            for (int w = 0; w < words.Length; w++)
            {
                string raw = words[w].Trim();
                string word = raw.Trim(' ', '\t', '\r', '\n', ':', '-', '/', '\\', '.', ',', '(', ')', '[', ']', '{', '}', '%', '#', '@', '*', '"', '\'', '؟', '،', '؛');
                if (word.Length < 2) continue;

                if (word[0] == '\u0629' || word[0] == '\uFE93' || word[0] == '\uFE94' || word[0] == '\uFBAB' ||
                    word[0] == '\u0649' || word[0] == '\uFEEF' || word[0] == '\uFEF0')
                    visualHits++;

                if (word.Length >= 3 && (word.EndsWith("\u0644\u0627", StringComparison.Ordinal) || word.EndsWith("\uFEFB", StringComparison.Ordinal)) &&
                    word != "إلا" && word != "كلا" && word != "لولا" && word != "هلا")
                    visualHits++;

                if (word[word.Length - 1] == '\u0629' || word[word.Length - 1] == '\uFE93' || word[word.Length - 1] == '\uFE94' || word[word.Length - 1] == '\uFBAB' ||
                    word[word.Length - 1] == '\u0649' || word[word.Length - 1] == '\uFEEF' || word[word.Length - 1] == '\uFEF0')
                    logicalHits++;

                if (word.Length >= 3 && (word.StartsWith("\u0627\u0644", StringComparison.Ordinal) ||
                                         word.StartsWith("\u0644\u0644", StringComparison.Ordinal) ||
                                         word.StartsWith("\u0648\u0627\u0644", StringComparison.Ordinal) ||
                                         word.StartsWith("\u0628\u0627\u0644", StringComparison.Ordinal) ||
                                         word.StartsWith("\u0641\u0627\u0644", StringComparison.Ordinal) ||
                                         word.StartsWith("\u0643\u0627\u0644", StringComparison.Ordinal)))
                    logicalHits++;

                string rootWord = word;
                if (word.Length >= 4 && (word[0] == 'و' || word[0] == 'ف' || word[0] == 'ب' || word[0] == 'ل' || word[0] == 'ك'))
                {
                    rootWord = word.Substring(1);
                }

                if (VisualKeywords.Contains(word) || VisualKeywords.Contains(rootWord))
                    visualHits++;

                if (LogicalKeywords.Contains(word) || LogicalKeywords.Contains(rootWord))
                    logicalHits++;
            }

            return logicalHits > visualHits && visualHits == 0 && logicalHits >= 1;
        }

        /// <summary>
        /// Normalizes Arabic text:
        /// 1. Converts Arabic Presentation Forms (FE70-FEFF, FB50-FDFF) to standard canonical Unicode (0600-06FF).
        /// 2. If reverseVisualArabic is enabled AND the input exhibits visual order,
        ///    reverses the characters while keeping numbers, dates, and Latin words in proper order.
        ///    Normal logical Arabic text (including text with numbers, dates, or symbols) is NEVER reversed.
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
                if (c == '\u06CC') c = '\u064A'; // Map Farsi Yeh to standard Arabic Yeh
                if (c == '\u06BE') c = '\u0647'; // Map Urdu Heh Doachashmee to standard Arabic Heh
                cleanSb.Append(c);
            }
            string cleaned = cleanSb.ToString();

            if (!ContainsArabic(cleaned))
                return cleaned;

            bool hadPresentationForms = ContainsPresentationForms(cleaned);

            // Reversal check:
            // Reverse if requested and either has presentation forms OR positive visual indicators,
            // provided the text is not confirmed to be strictly logical Arabic.
            bool isVisual = hadPresentationForms || IsVisualArabic(cleaned);
            bool isStrictlyLogical = IsStrictlyLogicalArabic(cleaned);

            if (reverseVisualArabic && isVisual && !isStrictlyLogical)
            {
                cleaned = ReverseArabicRuns(cleaned);
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

            return PostProcessArabicString(cleaned);
        }

        private static string PostProcessArabicString(string text)
        {
            if (string.IsNullOrEmpty(text) || !ContainsArabic(text))
                return text;

            // 1. Relocate leading Latin/model codes (e.g. "4a" -> "a4", "N") printed at position 0
            if (text.Length > 2 && !IsArabicChar(text[0]))
            {
                int prefixLen = 0;
                bool hasLetter = false;
                while (prefixLen < text.Length && !IsArabicChar(text[prefixLen]) && text[prefixLen] != ' ' && text[prefixLen] != '\t')
                {
                    char c = text[prefixLen];
                    if ((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z'))
                        hasLetter = true;
                    prefixLen++;
                }

                if (hasLetter && prefixLen >= 1 && prefixLen <= 5 && prefixLen < text.Length && IsArabicChar(text[prefixLen]))
                {
                    string rawCode = text.Substring(0, prefixLen);
                    if (char.IsDigit(rawCode[0]) && rawCode.Length >= 2)
                    {
                        char[] ca = rawCode.ToCharArray();
                        Array.Reverse(ca);
                        rawCode = new string(ca);
                    }
                    string rest = text.Substring(prefixLen).Trim();
                    text = rest + " " + rawCode;
                }
            }

            // 2. Insert spacing between Arabic letters and adjacent numbers
            StringBuilder sb = new StringBuilder(text.Length + 4);
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                sb.Append(c);

                if (i + 1 < text.Length)
                {
                    char next = text[i + 1];

                    // Arabic letter followed by digit: insert space unless it's "م2" or "م3"
                    if (IsArabicLetter(c) && char.IsDigit(next))
                    {
                        bool isSquareOrCubicUnit = (c == 'م' || c == 'M') && (next == '2' || next == '3') &&
                            (i + 2 >= text.Length || !char.IsDigit(text[i + 2]));
                        if (!isSquareOrCubicUnit)
                        {
                            sb.Append(' ');
                        }
                    }
                    // Arabic unit 'م' preceded by digit and followed by digit: "2000م30" -> "2000م 30"
                    else if (c == 'م' && i > 0 && char.IsDigit(text[i - 1]) && char.IsDigit(next))
                    {
                        sb.Append(' ');
                    }
                    // Digit followed by unit word: "3مللي" -> "3 مللي", "30ميكرون" -> "30 ميكرون"
                    else if (char.IsDigit(c) && (text.Substring(i + 1).StartsWith("مللي") || text.Substring(i + 1).StartsWith("مللى") || text.Substring(i + 1).StartsWith("ميكرون")))
                    {
                        sb.Append(' ');
                    }
                }
            }
            text = sb.ToString();

            // 3. Normalize "مللى" -> "مللي"
            if (text.Contains("مللى"))
            {
                text = text.Replace("مللى", "مللي");
            }

            return text;
        }

        private static bool IsArabicLetter(char c)
        {
            return IsArabicChar(c) && !char.IsDigit(c) && !(c >= '\u0660' && c <= '\u0669') &&
                   c != '،' && c != '؛' && c != '؟' && c != '٪' && c != 'ـ';
        }

        public static string ReverseArabicRuns(string text)
        {
            if (string.IsNullOrEmpty(text) || !ContainsArabic(text))
                return text;

            // In visual LTR rendering, trailing Latin/alphanumeric suffixes (e.g. A4, 4a, N)
            // were printed at the far left (position 0). Detect if string starts with a Latin code
            // immediately followed by visual Arabic, and relocate it to the logical suffix position.
            string trailingSuffix = null;
            if (text.Length > 2 && !IsArabicChar(text[0]))
            {
                int prefixEnd = 0;
                bool hasLetter = false;
                while (prefixEnd < text.Length && !IsArabicChar(text[prefixEnd]) && text[prefixEnd] != ' ' && text[prefixEnd] != '\t')
                {
                    char c = text[prefixEnd];
                    if ((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z'))
                        hasLetter = true;
                    prefixEnd++;
                }

                int arabicStart = prefixEnd;
                while (arabicStart < text.Length && (text[arabicStart] == ' ' || text[arabicStart] == '\t'))
                    arabicStart++;

                if (hasLetter && prefixEnd <= 6 && arabicStart < text.Length && IsArabicChar(text[arabicStart]))
                {
                    string rawCode = text.Substring(0, prefixEnd);
                    if (char.IsDigit(rawCode[0]) && rawCode.Length >= 2)
                    {
                        char[] ca = rawCode.ToCharArray();
                        Array.Reverse(ca);
                        rawCode = new string(ca);
                    }
                    trailingSuffix = rawCode;
                    text = text.Substring(arabicStart);
                }
            }

            StringBuilder result = new StringBuilder(text.Length);
            int i = 0;
            while (i < text.Length)
            {
                if (IsArabicChar(text[i]))
                {
                    // Collect contiguous Arabic run (Arabic characters and spaces between Arabic words)
                    int start = i;
                    while (i < text.Length)
                    {
                        if (IsArabicChar(text[i]))
                        {
                            i++;
                        }
                        else if (text[i] == ' ' || text[i] == '\t')
                        {
                            // Space is part of Arabic run only if followed by another Arabic character
                            int lookAhead = i;
                            while (lookAhead < text.Length && (text[lookAhead] == ' ' || text[lookAhead] == '\t'))
                                lookAhead++;

                            if (lookAhead < text.Length && IsArabicChar(text[lookAhead]))
                            {
                                i = lookAhead;
                            }
                            else
                            {
                                break;
                            }
                        }
                        else
                        {
                            break;
                        }
                    }

                    string arabicSegment = text.Substring(start, i - start);
                    string reversedSegment = ReverseArabicSegment(arabicSegment);

                    // If previous non-Arabic segment ended with a digit and this segment is a full word (>2 chars), add space
                    if (result.Length > 0)
                    {
                        char prev = result[result.Length - 1];
                        if (char.IsDigit(prev) && reversedSegment.Length > 2)
                        {
                            result.Append(" ");
                        }
                    }

                    result.Append(reversedSegment);
                }
                else
                {
                    int start = i;
                    while (i < text.Length && !IsArabicChar(text[i]))
                    {
                        i++;
                    }
                    string nonArabic = text.Substring(start, i - start);

                    // If previous segment ended with an Arabic word (>2 chars) and this non-Arabic segment starts with a digit, add space
                    if (result.Length > 0 && nonArabic.Length > 0)
                    {
                        char prev = result[result.Length - 1];
                        if (IsArabicChar(prev) && char.IsDigit(nonArabic[0]))
                        {
                            result.Append(" ");
                        }
                    }

                    result.Append(nonArabic);
                }
            }

            if (!string.IsNullOrEmpty(trailingSuffix))
            {
                string r = result.ToString().TrimEnd();
                return r + " " + trailingSuffix;
            }

            return result.ToString();
        }

        private static string ReverseArabicSegment(string segment)
        {
            if (string.IsNullOrEmpty(segment)) return segment;

            List<string> tokens = new List<string>();
            int i = 0;
            while (i < segment.Length)
            {
                if (segment[i] == ' ' || segment[i] == '\t')
                {
                    int start = i;
                    while (i < segment.Length && (segment[i] == ' ' || segment[i] == '\t'))
                        i++;
                    tokens.Add(segment.Substring(start, i - start));
                }
                else
                {
                    int start = i;
                    while (i < segment.Length && segment[i] != ' ' && segment[i] != '\t')
                        i++;

                    char[] chars = segment.Substring(start, i - start).ToCharArray();
                    Array.Reverse(chars);
                    tokens.Add(new string(chars));
                }
            }

            // Reverse word order within this Arabic segment so reading order is RTL
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
