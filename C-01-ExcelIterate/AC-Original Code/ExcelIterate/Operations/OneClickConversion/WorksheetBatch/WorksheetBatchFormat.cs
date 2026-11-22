using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.CSharp.RuntimeBinder;

namespace Excel_Iterate
{
    internal sealed class WorksheetBatchFormat : IWorksheetChange
    {
        private const int NoColor = -4142, Automatic = -4105, Solid = 1;
        private const int UnderlineNone = -4142, UnderlineSingle = 2;
        private const int WpsUndefined = 9999999;

        internal sealed class ColorState
        {
            internal int Rgb, Index, Theme, Pattern;
            internal double Tint;

            internal static ColorState Read(dynamic style, bool fill)
            {
                int pattern = fill ? ReadInt((object)style.Pattern, "填充图案") : Solid;
                if (pattern == NoColor) return new ColorState { Pattern = pattern };
                if (pattern != Solid) throw new InvalidOperationException("图案或渐变填充需保留原样，请缩小格式选区。");
                int index = ReadInt((object)style.ColorIndex, "颜色索引");
                int theme;
                try { theme = ReadInt((object)style.ThemeColor, "主题颜色"); }
                catch (COMException error)
                {
                    if (error.ErrorCode != unchecked((int)0x800A03EC)) throw;
                    theme = NoColor;
                }
                catch (ArgumentException error)
                {
                    if (error.HResult != unchecked((int)0x80070057)) throw;
                    theme = NoColor;
                }
                return new ColorState { Pattern = pattern, Index = index, Theme = theme > 0 ? theme : 0,
                    Rgb = ReadInt((object)style.Color, "颜色", allowWpsUndefined: true),
                    Tint = Convert.ToDouble((object)style.TintAndShade, CultureInfo.InvariantCulture) };
            }

            internal bool Matches(ColorState other)
            {
                if (Pattern != other.Pattern) return false;
                if (Pattern == NoColor) return true;
                return (Index == Automatic) == (other.Index == Automatic) && Theme == other.Theme &&
                    Rgb == other.Rgb && Math.Abs(Tint - other.Tint) < 0.00001;
            }

            internal bool MatchesRgb(int rgb)
            { return Pattern == Solid && Index != Automatic && Theme == 0 && Rgb == rgb && Math.Abs(Tint) < 0.00001; }

            internal static void WriteRgb(dynamic style, bool fill, int rgb)
            {
                if (fill) style.Pattern = Solid;
                style.TintAndShade = 0.0;
                style.Color = rgb;
            }

            internal void Restore(dynamic style, bool fill)
            {
                if (Pattern == NoColor) { style.Pattern = NoColor; return; }
                if (fill) style.Pattern = Solid;
                if (Theme > 0) style.ThemeColor = Theme;
                else if (Index == Automatic) style.ColorIndex = Automatic;
                else style.Color = Rgb;
                style.TintAndShade = Tint;
            }

            public override string ToString()
            { return string.Format(CultureInfo.InvariantCulture, "pattern={0}, index={1}, theme={2}, color={3}, tint={4}", Pattern, Index, Theme, Rgb, Tint); }
        }

        internal sealed class TypefaceState
        {
            internal readonly string Name;
            private readonly int? theme;

            private TypefaceState(string name, int? theme)
            { Name = name; this.theme = theme; }

            internal static TypefaceState Read(dynamic font)
            {
                string name = ReadText((object)font.Name, "字体");
                object theme = ReadOptional(() => (object)font.ThemeFont);
                return new TypefaceState(name, theme == null ? (int?)null : ReadInt(theme, "主题字体"));
            }

            internal bool Matches(TypefaceState other)
            { return theme == other.theme && SameFontName(Name, other.Name); }

            internal void Restore(dynamic font)
            {
                if (theme > 0) font.ThemeFont = theme.Value;
                else
                {
                    if (theme == 0 && (object)font.ThemeFont != null &&
                        ReadInt((object)font.ThemeFont, "主题字体") != 0) font.ThemeFont = 0;
                    font.Name = Name;
                }
            }

            public override string ToString() { return Name + (theme.HasValue ? " (主题 " + theme.Value + ")" : ""); }
        }

        internal sealed class ClearDefaults
        {
            internal readonly TypefaceState Typeface;
            internal readonly ColorState Color;
            private ClearDefaults(TypefaceState typeface, ColorState color)
            { Typeface = typeface; Color = color; }

            internal static ClearDefaults Read(object workbook)
            {
                using (var bag = new ComBag())
                {
                    dynamic styles = bag.Own(((dynamic)workbook).Styles);
                    object normal;
                    try { normal = styles.Item("Normal"); }
                    catch (Exception englishError)
                    {
                        try { normal = styles.Item("常规"); }
                        catch (Exception localError)
                        { throw new InvalidOperationException("无法读取本工作簿的常规样式。", new AggregateException(englishError, localError)); }
                    }
                    dynamic style = bag.Own(normal);
                    // WPS may report BuiltIn=false for the workbook's named Normal style.
                    dynamic font = bag.Own(style.Font);
                    return new ClearDefaults(TypefaceState.Read(font), ColorState.Read(font, false));
                }
            }
        }

        private sealed class CellLayout
        {
            private readonly bool wrap, shrink;
            private readonly int indent;
            private readonly bool? autoIndent;

            private CellLayout(bool wrapText, bool shrinkToFit, int indentLevel, bool? addIndent)
            { wrap = wrapText; shrink = shrinkToFit; indent = indentLevel; autoIndent = addIndent; }

            internal static CellLayout Read(dynamic cell)
            {
                object addIndent = ReadOptional(() => (object)cell.AddIndent);
                return new CellLayout(ReadBool((object)cell.WrapText, "自动换行"), ReadBool((object)cell.ShrinkToFit, "缩小字体填充"),
                    ReadInt((object)cell.IndentLevel, "缩进量"), addIndent == null ? (bool?)null : ReadBool(addIndent, "自动缩进"));
            }

            internal bool Matches(CellLayout other)
            { return wrap == other.wrap && shrink == other.shrink && indent == other.indent && autoIndent == other.autoIndent; }

            internal void Restore(dynamic cell)
            {
                var current = Read(cell);
                if (current.wrap != wrap) cell.WrapText = wrap;
                if (current.shrink != shrink) cell.ShrinkToFit = shrink;
                if (autoIndent.HasValue && current.autoIndent != autoIndent) cell.AddIndent = autoIndent.Value;
                if (current.indent != indent) cell.IndentLevel = indent;
            }

            public override string ToString()
            { return "换行=" + wrap + "，缩小字体填充=" + shrink + "，缩进量=" + indent + "，自动缩进=" + autoIndent; }
        }

        private sealed class PropertyEdit<T>
        {
            private readonly T before, after;
            private PropertyEdit(T original, T requested) { before = original; after = requested; }
            internal static PropertyEdit<T> Plan(T before, T after)
            { return EqualityComparer<T>.Default.Equals(before, after) ? null : new PropertyEdit<T>(before, after); }
            internal T Value(bool applied) { return applied ? after : before; }
            internal void Check(T actual, bool applied, string label)
            { Require(EqualityComparer<T>.Default.Equals(Value(applied), actual), label); }
        }

        private PropertyEdit<bool> bold, italic, strikethrough;
        private PropertyEdit<int> underline, horizontal, vertical;
        private PropertyEdit<string> number;
        private TypefaceState typefaceBefore, defaultTypeface, typefaceAfter;
        private ColorState fontBefore, defaultFontColor, fillBefore, fontAfter, fillAfter;
        private string fontName;
        private int? fontRgb, fillRgb;
        private CellLayout layoutBefore, layoutAfter;

        internal bool AffectsRowLayout { get { return layoutBefore != null; } }
        public string Address { get; private set; }
        private WorksheetBatchFormat(string address) { Address = address; }

        internal static string CanonicalFontName(string name)
        {
            name = name.Trim();
            if (name == "宋体") return "SimSun";
            if (name == "楷体") return "KaiTi";
            return name;
        }

        internal static bool SameFontName(string a, string b)
        { return string.Equals(CanonicalFontName(a), CanonicalFontName(b), StringComparison.OrdinalIgnoreCase); }

        internal static object ReadOptional(Func<object> read)
        {
            try
            {
                object value = read();
                // WPS can return this sentinel for optional properties even on a single cell.
                return value == DBNull.Value || value == Type.Missing || Equals(value, WpsUndefined) ? null : value;
            }
            catch (RuntimeBinderException) { return null; }
            catch (MissingMemberException) { return null; }
            catch (COMException error)
            {
                if (error.ErrorCode != unchecked((int)0x80020003) && error.ErrorCode != unchecked((int)0x80020006) &&
                    error.ErrorCode != unchecked((int)0x80004001)) throw;
                return null;
            }
        }

        private static string ReadText(object value, string label)
        {
            string text = value as string;
            if (string.IsNullOrWhiteSpace(text)) throw new InvalidOperationException("无法可靠读取" + label + "，可能含混合格式。");
            return text;
        }

        private static bool ReadBool(object value, string label)
        {
            if (value is bool) return (bool)value;
            int number = ReadInt(value, label);
            if (number == 0 || number == 1 || number == -1) return number != 0;
            throw new InvalidOperationException("无法可靠读取" + label + "，可能含混合格式。");
        }

        private static int ReadInt(object value, string label, bool allowWpsUndefined = false)
        {
            if (value is int || value is short || value is long || value is byte || value is uint ||
                value is double || value is float || value is decimal)
            {
                double number = Convert.ToDouble(value, CultureInfo.InvariantCulture);
                if ((allowWpsUndefined || number != WpsUndefined) && number >= int.MinValue &&
                    number <= int.MaxValue && number == Math.Truncate(number)) return (int)number;
            }
            throw new InvalidOperationException("无法可靠读取" + label + "，可能含混合格式。");
        }

        internal static WorksheetBatchFormat Plan(object range, string address, WorksheetBatchRequest request, ClearDefaults defaults)
        {
            var edit = new WorksheetBatchFormat(address);
            bool clear = request.ClearFormatting;
            bool clearFill = clear || request.ClearFill;
            int? horizontal = clear ? (int)WorksheetBatchHorizontal.General : (int?)request.Horizontal;
            int? vertical = clear ? (int)WorksheetBatchVertical.Bottom : (int?)request.Vertical;
            string number = clear ? "General" : request.NumberFormat;
            using (var bag = new ComBag())
            {
                dynamic cell = range;
                if (clear || request.TextStyle != WorksheetBatchTextStyle.Keep || request.FontName != null || request.FontColor.HasValue)
                    edit.PlanFont(bag.Own(cell.Font), request, defaults);
                if (clearFill || request.FillColor.HasValue)
                {
                    ColorState before = ColorState.Read((dynamic)bag.Own(cell.Interior), true);
                    if (clearFill ? before.Pattern != NoColor : !before.MatchesRgb(request.FillColor.Value))
                    { edit.fillBefore = before; edit.fillRgb = request.FillColor; }
                }
                if (number != null)
                    edit.number = PropertyEdit<string>.Plan(ReadText((object)cell.NumberFormat, "数字格式"), number);
                if (horizontal.HasValue)
                    edit.horizontal = PropertyEdit<int>.Plan(ReadInt((object)cell.HorizontalAlignment, "水平对齐方式"), horizontal.Value);
                if (vertical.HasValue)
                    edit.vertical = PropertyEdit<int>.Plan(ReadInt((object)cell.VerticalAlignment, "垂直对齐方式"), vertical.Value);
                if (edit.typefaceBefore != null || edit.bold != null || edit.italic != null || edit.underline != null ||
                    edit.strikethrough != null || edit.horizontal != null || edit.vertical != null || (clear && edit.number != null))
                    edit.layoutBefore = CellLayout.Read(cell);
            }
            return edit.layoutBefore != null || edit.fontBefore != null || edit.fillBefore != null || edit.number != null ? edit : null;
        }

        private void PlanFont(dynamic font, WorksheetBatchRequest request, ClearDefaults defaults)
        {
            bool clear = request.ClearFormatting;
            if (clear || request.TextStyle == WorksheetBatchTextStyle.Bold)
                bold = PropertyEdit<bool>.Plan(ReadBool((object)font.Bold, "加粗状态"), !clear);
            if (clear || request.TextStyle == WorksheetBatchTextStyle.Italic)
                italic = PropertyEdit<bool>.Plan(ReadBool((object)font.Italic, "斜体状态"), !clear);
            if (clear || request.TextStyle == WorksheetBatchTextStyle.Underline)
                underline = PropertyEdit<int>.Plan(ReadInt((object)font.Underline, "下划线类型"), clear ? UnderlineNone : UnderlineSingle);
            if (clear || request.TextStyle == WorksheetBatchTextStyle.Strikethrough)
                strikethrough = PropertyEdit<bool>.Plan(ReadBool((object)font.Strikethrough, "删除线状态"), !clear);
            if (clear || request.FontName != null)
            {
                var current = TypefaceState.Read(font);
                if (!(clear ? defaults.Typeface.Matches(current) : SameFontName(current.Name, request.FontName)))
                {
                    typefaceBefore = current;
                    if (clear) defaultTypeface = defaults.Typeface;
                    else fontName = request.FontName;
                }
            }
            if (clear || request.FontColor.HasValue)
            {
                ColorState current = ColorState.Read(font, false);
                if (!(clear ? defaults.Color.Matches(current) : current.MatchesRgb(request.FontColor.Value)))
                {
                    fontBefore = current;
                    if (clear) defaultFontColor = defaults.Color;
                    else fontRgb = request.FontColor;
                }
            }
        }

        private static void Require(bool condition, string property)
        { if (!condition) throw new InvalidOperationException(property + "已变化，或宿主返回了不同的格式。"); }

        private bool ChangesFont
        { get { return typefaceBefore != null || bold != null || italic != null || underline != null || strikethrough != null || fontBefore != null; } }

        public void Check(object sheet, bool after)
        {
            using (var bag = new ComBag())
            {
                dynamic cell = bag.Own(((dynamic)sheet).Range[Address]);
                Require(!(bool)((dynamic)sheet).ProtectContents && !(bool)cell.MergeCells, "工作表保护或单元格合并状态");
                if (ChangesFont) CheckFont(bag.Own(cell.Font), after);
                if (fillBefore != null)
                {
                    ColorState current = ColorState.Read((dynamic)bag.Own(cell.Interior), true);
                    bool matches = !after ? fillBefore.Matches(current) : fillAfter != null ? fillAfter.Matches(current)
                        : fillRgb.HasValue ? current.MatchesRgb(fillRgb.Value) : current.Pattern == NoColor;
                    Require(matches, "填充颜色");
                    if (after && fillAfter == null) fillAfter = current;
                }
                if (number != null) number.Check(ReadText((object)cell.NumberFormat, "数字格式"), after, "数字格式");
                if (horizontal != null) horizontal.Check(ReadInt((object)cell.HorizontalAlignment, "水平对齐方式"), after, "水平对齐方式");
                if (vertical != null) vertical.Check(ReadInt((object)cell.VerticalAlignment, "垂直对齐方式"), after, "垂直对齐方式");
                if (layoutBefore != null)
                {
                    CellLayout current = CellLayout.Read(cell);
                    if (after && layoutAfter == null) layoutAfter = current;
                    Require((after ? layoutAfter : layoutBefore).Matches(current), "换行或缩进");
                }
            }
        }

        private void CheckFont(dynamic font, bool after)
        {
            if (typefaceBefore != null)
            {
                TypefaceState current = TypefaceState.Read(font);
                TypefaceState expected = after ? typefaceAfter ?? defaultTypeface : typefaceBefore;
                Require(expected != null ? expected.Matches(current) : SameFontName(fontName, current.Name), "字体");
                if (after && typefaceAfter == null) typefaceAfter = current;
            }
            if (bold != null) bold.Check(ReadBool((object)font.Bold, "加粗状态"), after, "加粗状态");
            if (italic != null) italic.Check(ReadBool((object)font.Italic, "斜体状态"), after, "斜体状态");
            if (underline != null) underline.Check(ReadInt((object)font.Underline, "下划线类型"), after, "下划线类型");
            if (strikethrough != null) strikethrough.Check(ReadBool((object)font.Strikethrough, "删除线状态"), after, "删除线状态");
            if (fontBefore != null)
            {
                ColorState current = ColorState.Read(font, false);
                ColorState expected = after ? fontAfter ?? defaultFontColor : fontBefore;
                Require(expected != null ? expected.Matches(current) : current.MatchesRgb(fontRgb.Value), "字体颜色");
                if (after && fontAfter == null) fontAfter = current;
            }
        }

        public void Write(object sheet, bool after)
        {
            using (var bag = new ComBag())
            {
                dynamic cell = bag.Own(((dynamic)sheet).Range[Address]);
                if (ChangesFont) WriteFont(bag.Own(cell.Font), after);
                if (fillBefore != null)
                {
                    dynamic fill = bag.Own(cell.Interior);
                    if (!after) fillBefore.Restore(fill, true);
                    else if (fillRgb.HasValue) ColorState.WriteRgb(fill, true, fillRgb.Value);
                    else fill.Pattern = NoColor;
                }
                if (number != null) cell.NumberFormat = number.Value(after);
                if (horizontal != null) cell.HorizontalAlignment = horizontal.Value(after);
                if (vertical != null) cell.VerticalAlignment = vertical.Value(after);
                CellLayout layout = after ? layoutAfter : layoutBefore;
                if (layout != null) layout.Restore(cell);
            }
        }

        private void WriteFont(dynamic font, bool after)
        {
            if (typefaceBefore != null)
            {
                TypefaceState value = after ? typefaceAfter ?? defaultTypeface : typefaceBefore;
                if (value != null) value.Restore(font);
                else font.Name = fontName;
            }
            if (bold != null) font.Bold = bold.Value(after);
            if (italic != null) font.Italic = italic.Value(after);
            if (underline != null) font.Underline = underline.Value(after);
            if (strikethrough != null) font.Strikethrough = strikethrough.Value(after);
            if (fontBefore != null)
            {
                ColorState value = after ? fontAfter ?? defaultFontColor : fontBefore;
                if (value != null) value.Restore(font, false);
                else ColorState.WriteRgb(font, false, fontRgb.Value);
            }
        }

        private static string ValueText<T>(PropertyEdit<T> edit, bool after)
        { return edit == null ? "" : Convert.ToString(edit.Value(after), CultureInfo.InvariantCulture); }

        private string Describe(bool after)
        {
            string typeface = after ? defaultTypeface == null ? fontName : defaultTypeface.ToString()
                : typefaceBefore == null ? null : typefaceBefore.ToString();
            string fontColor = after ? defaultFontColor == null ? fontRgb.ToString() : defaultFontColor.ToString()
                : fontBefore == null ? null : fontBefore.ToString();
            return "字体=" + typeface + "，加粗=" + ValueText(bold, after) + "，斜体=" + ValueText(italic, after) +
                "，下划线=" + ValueText(underline, after) + "，删除线=" + ValueText(strikethrough, after) +
                "，字体颜色=" + fontColor + "，填充=" + (after ? fillRgb.ToString() : (object)fillBefore) +
                "，数字格式=" + ValueText(number, after) + "，水平对齐=" + ValueText(horizontal, after) +
                "，垂直对齐=" + ValueText(vertical, after) + (after || layoutBefore == null ? "" : "，" + layoutBefore);
        }

        public string DescribeBefore() { return Describe(false); }
        internal string DescribePlan() { return Describe(true); }
    }
}
