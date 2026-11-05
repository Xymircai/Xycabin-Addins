using System;
using System.Text;

namespace Excel_Iterate
{
    [Flags]
    internal enum CellCharacterKinds { None = 0, Spaces = 1, English = 2, Chinese = 4, Digits = 8, Other = 16, All = 31 }

    internal static class CellCharacterClassifier
    {
        // Keep the established OCR/name-extraction Han ranges, including supplementary characters.
        internal static bool IsChinese(int code)
        { return code >= 0x3400 && code <= 0x4dbf || code >= 0x4e00 && code <= 0x9fff || code >= 0xf900 && code <= 0xfaff || code >= 0x20000 && code <= 0x323af; }

        internal static CellCharacterKinds Kind(string text, int at)
        {
            if (char.IsWhiteSpace(text, at)) return CellCharacterKinds.Spaces;
            int code = char.IsHighSurrogate(text[at]) && at + 1 < text.Length && char.IsLowSurrogate(text[at + 1]) ? char.ConvertToUtf32(text[at], text[at + 1]) : text[at];
            if (IsChinese(code)) return CellCharacterKinds.Chinese;
            if (code >= '0' && code <= '9' || code >= 0xff10 && code <= 0xff19) return CellCharacterKinds.Digits;
            if (code >= 'a' && code <= 'z' || code >= 'A' && code <= 'Z' || code >= 0xff21 && code <= 0xff3a || code >= 0xff41 && code <= 0xff5a) return CellCharacterKinds.English;
            return CellCharacterKinds.Other;
        }

        internal static string Filter(string text, CellCharacterKinds kinds, bool extract)
        {
            if (kinds == CellCharacterKinds.None || (kinds & ~CellCharacterKinds.All) != 0) throw new ArgumentException("请至少选择一种有效字符类型。");
            var result = new StringBuilder(); text = text ?? "";
            for (int at = 0; at < text.Length; at++)
            {
                int count = char.IsHighSurrogate(text[at]) && at + 1 < text.Length && char.IsLowSurrogate(text[at + 1]) ? 2 : 1;
                if (((Kind(text, at) & kinds) != 0) == extract) result.Append(text, at, count);
                at += count - 1;
            }
            return result.ToString();
        }
    }
}
