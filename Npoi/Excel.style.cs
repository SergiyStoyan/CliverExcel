//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text.RegularExpressions;
using System.Drawing;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.SS.Formula.PTG;
using NPOI.SS.Formula;
using Newtonsoft.Json.Serialization;
using System.Reflection;
using Newtonsoft.Json;
using System.Xml.Linq;
using NPOI.HSSF.Util;
using NPOI.XSSF.Streaming;

namespace Cliver
{
    public partial class Excel : IDisposable
    {
        public class Color
        {
            public readonly byte R;
            public readonly byte G;
            public readonly byte B;
            readonly public byte[] RGB = new byte[3];

            public Color(byte r, byte g, byte b)
            {
                R = r;
                G = g;
                B = b;
                RGB[0] = R;
                RGB[1] = G;
                RGB[2] = B;
            }

            public Color(int aRGB) : this((byte)((aRGB >> 16) & 0xFF), (byte)((aRGB >> 8) & 0xFF), (byte)(aRGB & 0xFF))
            {
            }

            public Color(byte[] RGB) : this(RGB[0], RGB[1], RGB[2])
            {
            }

            public Color(IColor c) : this(c.RGB[0], c.RGB[1], c.RGB[2])
            {
            }

            public Color(System.Drawing.Color color) : this(color.ToArgb())
            {
            }
        }

        //public void Unhighlight(Color color)
        //{
        //    if (Workbook is XSSFWorkbook)
        //    {
        //        if (color == null)
        //        {
        //            foreach (XSSFCellStyle s in GetStyles())
        //                s.SetFillForegroundColor(null);
        //            return;
        //        }
        //        XSSFColor c = new XSSFColor(color.RGB);
        //        foreach (XSSFCellStyle s in GetStyles())
        //        {
        //            if (AreColorsEqual(s.FillForegroundColorColor, c))
        //                s.SetFillForegroundColor(null);
        //        }
        //    }
        //    else if (Workbook is HSSFWorkbook hw)
        //    {
        //        if (color == null)
        //        {
        //            foreach (HSSFCellStyle s in GetStyles())
        //                s.FillForegroundColor = 0;
        //            return;
        //        }
        //        HSSFPalette palette = hw.GetCustomPalette();
        //        HSSFColor c = palette.FindColor(color.RGB[0], color.RGB[1], color.RGB[2]);
        //        foreach (HSSFCellStyle s in GetStyles())
        //        {
        //            if (AreColorsEqual(color, c))
        //                s.FillForegroundColor = 0;
        //        }
        //    }
        //    else
        //        throw new Exception("Unsupported workbook type: " + Workbook.GetType().FullName);
        //}

        public ICellStyle Highlight(ICellStyle style, Color color, FillPattern fillPattern = FillPattern.SolidForeground, bool createOnlyUniqueStyle = true)
        {
            return highlight(this, style, createOnlyUniqueStyle, color, fillPattern);
        }

        /// <summary>
        /// Intended for either adding or removing backgound color.
        /// (!)When createUniqueStyleOnly, it is slow.
        /// </summary>
        /// <param name="excel"></param>
        /// <param name="style"></param>
        /// <param name="createUniqueStyleOnly"></param>
        /// <param name="color"></param>
        /// <param name="fillPattern"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        static internal ICellStyle highlight(Excel excel, ICellStyle style, bool createOnlyUniqueStyle, Color color, FillPattern fillPattern = FillPattern.SolidForeground)
        {
            if (excel.Workbook is XSSFWorkbook)
            {
                XSSFCellStyle cs;
                if (color == null)
                {
                    if (style == null)
                        return null;
                    cs = (XSSFCellStyle)style;
                    cs.SetFillForegroundColor(null);
                    cs.FillPattern = FillPattern.NoFill;
                    return cs;
                }
                if (createOnlyUniqueStyle)
                {
                    cs = style == null ? (XSSFCellStyle)excel.CreateUnregisteredStyle() : (XSSFCellStyle)excel.CloneUnregisteredStyle(style);
                    cs.SetFillForegroundColor(new XSSFColor(color.RGB));
                    cs.FillPattern = fillPattern;
                    return excel.GetRegisteredStyle(cs);
                }
                cs = style == null ? (XSSFCellStyle)excel.Workbook.CreateCellStyle() : (XSSFCellStyle)style;
                cs.SetFillForegroundColor(new XSSFColor(color.RGB));
                cs.FillPattern = fillPattern;
                return cs;
            }
            if (excel.Workbook is HSSFWorkbook)
            {
                if (color == null)
                {
                    if (style == null)
                        return null;
                    style.FillForegroundColor = 0;
                    style.FillPattern = FillPattern.NoFill;
                    return style;
                }
                HSSFPalette palette = ((HSSFWorkbook)excel.Workbook).GetCustomPalette();
                HSSFColor hssfColor = palette.FindColor(color.R, color.G, color.B);
                if (hssfColor == null)
                {
                    hssfColor = getRegisteredHSSFColor((HSSFWorkbook)excel.Workbook, color);
                    HSSFCellStyle hcs = style == null ? (HSSFCellStyle)excel.Workbook.CreateCellStyle() : (HSSFCellStyle)style;
                    hcs.FillForegroundColor = hssfColor.Indexed;
                    hcs.FillPattern = fillPattern;
                    return hcs;
                }
                ICellStyle cs;
                if (createOnlyUniqueStyle)
                {
                    if (style == null)
                        cs = excel.CreateUnregisteredStyle();
                    else
                        cs = excel.CloneUnregisteredStyle(style);
                    cs.FillForegroundColor = hssfColor.Indexed;
                    cs.FillPattern = fillPattern;
                    return excel.GetRegisteredStyle(cs);
                }
                cs = style == null ? (HSSFCellStyle)excel.Workbook.CreateCellStyle() : (HSSFCellStyle)style;
                cs.FillForegroundColor = hssfColor.Indexed;
                cs.FillPattern = fillPattern;
                return cs;
            }
            throw new Exception("Unsupported workbook type: " + excel.Workbook.GetType().FullName);
        }

        static HSSFColor getRegisteredHSSFColor(HSSFWorkbook workbook, Color color)
        {
            HSSFPalette palette = workbook.GetCustomPalette();
            HSSFColor hssfColor = palette.FindColor(color.R, color.G, color.B);
            if (hssfColor != null)
                return hssfColor;
            try
            {
                hssfColor = palette.AddColor(color.R, color.G, color.B);
            }
            catch (Exception e)
            {//pallete is full
                short? findUnusedColorIndex()
                {
                    for (short j = 0x8; j <= 0x40; j++)//the first color in the palette has the index 0x8, the second has the index 0x9, etc. through 0x40
                    {
                        int i = 0;
                        for (; i < workbook.NumCellStyles; i++)
                        {
                            var s = workbook.GetCellStyleAt(i);
                            if (s.BorderDiagonalColor == j
                                || s.BottomBorderColor == j
                                || s.FillBackgroundColor == j
                                || s.FillForegroundColor == j
                                || s.LeftBorderColor == j
                                || s.RightBorderColor == j
                                || s.TopBorderColor == j
                                )
                                break;
                        }
                        if (i >= workbook.NumCellStyles)
                            return j;
                    }
                    return null;
                }
                short? ci = findUnusedColorIndex();
                if (ci == null)
                    ci = palette.FindSimilarColor(color.R, color.G, color.B).Indexed;
                palette.SetColorAtIndex(ci.Value, color.R, color.G, color.B);
                hssfColor = palette.GetColor(ci.Value);
            }
            return hssfColor;
        }

        /// <summary>
        /// Looks for an equal style in the workbook and, if it does not exists, creates a new one.
        /// (!)Incidentally, there is a somewhat analogous method NPOI.SS.Util.CellUtil.SetCellStyleProperties() which is not as handy in use though.
        /// </summary>
        /// <param name="style">it is a style created by CreateUnregisteredStyle() and then modified as needed. But it can be a registered style, too.</param>
        /// <param name="unregisteredStyleWorkbook"></param>
        /// <returns></returns>
        public ICellStyle GetRegisteredStyle(ICellStyle unregisteredStyle, IWorkbook unregisteredStyleWorkbook = null)
        {
            if (unregisteredStyleWorkbook != null && unregisteredStyleWorkbook.GetType() != Workbook.GetType())
                throw new Exception("Registering a style in a different type workbook is not supported: " + Workbook.GetType().FullName);

            ICellStyle style = FindEqualStyles(unregisteredStyle, unregisteredStyleWorkbook).FirstOrDefault();
            if (style != null)
                return style;
            style = Workbook.CreateCellStyle();
            return CopyStyle(unregisteredStyle, style);
        }

        public IEnumerable<ICellStyle> FindEqualStyles(ICellStyle style, IWorkbook styleWorkbook = null)
        {
            if (styleWorkbook != null && styleWorkbook.GetType() != Workbook.GetType())
                throw new Exception("Comparing a style from a different type workbook is not supported: " + Workbook.GetType().FullName);

            HSSFColor hSSFForegroundColor = null;
            HSSFColor hSSFBackgroundColor = null;
            HSSFColor hSSFBorderDiagonalColor = null;
            HSSFColor hSSFBottomBorderColor = null;
            HSSFColor hSSFLeftBorderColor = null;
            HSSFColor hSSFRightBorderColor = null;
            HSSFColor hSSFTopBorderColor = null;
            if (Workbook is HSSFWorkbook hw)
            {
                HSSFPalette palette = hw.GetCustomPalette();
                HSSFColor findColor(IColor c)
                {
                    return c == null ? null : palette.FindColor(c.RGB[0], c.RGB[1], c.RGB[2]);
                }
                hSSFForegroundColor = findColor(style.FillForegroundColorColor);
                if (hSSFForegroundColor == null)
                    yield break;
                hSSFBackgroundColor = findColor(style.FillBackgroundColorColor);
                if (hSSFBackgroundColor == null)
                    yield break;
                HSSFPalette uPalette = ((HSSFWorkbook)styleWorkbook).GetCustomPalette();
                hSSFBorderDiagonalColor = findColor(uPalette.GetColor(style.BorderDiagonalColor));
                if (hSSFBorderDiagonalColor == null)
                    yield break;
                hSSFBottomBorderColor = findColor(uPalette.GetColor(style.BottomBorderColor));
                if (hSSFBottomBorderColor == null)
                    yield break;
                hSSFLeftBorderColor = findColor(uPalette.GetColor(style.LeftBorderColor));
                if (hSSFLeftBorderColor == null)
                    yield break;
                hSSFRightBorderColor = findColor(uPalette.GetColor(style.RightBorderColor));
                if (hSSFRightBorderColor == null)
                    yield break;
                hSSFTopBorderColor = findColor(uPalette.GetColor(style.TopBorderColor));
                if (hSSFTopBorderColor == null)
                    yield break;
            }

            string unregisteredStyleDataFormatString = null;
            IDataFormat sDataFormat = null;
            IFont unregisteredStyleFont = null;
            if (styleWorkbook != null && styleWorkbook != Workbook)
            {
                unregisteredStyleDataFormatString = styleWorkbook.CreateDataFormat().GetFormat(style.DataFormat);
                sDataFormat = Workbook.CreateDataFormat();
                unregisteredStyleFont = style.GetFont(styleWorkbook);
            }

            foreach (ICellStyle s in GetStyles())
            {
                if (styleWorkbook == null || styleWorkbook == Workbook)
                    if (s.Index == style.Index)
                        continue;
                if (style.Alignment != s.Alignment
                || style.BorderBottom != s.BorderBottom
                || style.BorderDiagonal != s.BorderDiagonal
                || style.BorderDiagonalLineStyle != s.BorderDiagonalLineStyle
                || style.BorderLeft != s.BorderLeft
                || style.BorderRight != s.BorderRight
                || style.BorderTop != s.BorderTop
                || style.FillPattern != s.FillPattern
                || style.Indention != s.Indention
                || style.IsHidden != s.IsHidden
                || style.IsLocked != s.IsLocked
                || style.Rotation != s.Rotation
                || style.ShrinkToFit != s.ShrinkToFit
                || style.VerticalAlignment != s.VerticalAlignment
                || style.WrapText != s.WrapText
                //|| style.BorderDiagonalColor != s.BorderDiagonalColor
                //|| style.BottomBorderColor != s.BottomBorderColor
                //|| style.LeftBorderColor != s.LeftBorderColor
                //|| style.RightBorderColor != s.RightBorderColor
                //|| style.TopBorderColor != s.TopBorderColor
                )
                    continue;

                if (style is XSSFCellStyle uxcs)
                {
                    XSSFCellStyle xcs = s as XSSFCellStyle;
                    if (!AreColorsEqual(uxcs.FillForegroundXSSFColor, xcs.FillForegroundXSSFColor)
                        || !AreColorsEqual(uxcs.FillBackgroundXSSFColor, xcs.FillBackgroundXSSFColor)
                        || !AreColorsEqual(uxcs.DiagonalBorderXSSFColor, xcs.DiagonalBorderXSSFColor)
                        || !AreColorsEqual(uxcs.BottomBorderXSSFColor, xcs.BottomBorderXSSFColor)
                        || !AreColorsEqual(uxcs.LeftBorderXSSFColor, xcs.LeftBorderXSSFColor)
                        || !AreColorsEqual(uxcs.RightBorderXSSFColor, xcs.RightBorderXSSFColor)
                        || !AreColorsEqual(uxcs.TopBorderXSSFColor, xcs.TopBorderXSSFColor)
                        )
                        continue;
                }
                else if (style is HSSFCellStyle)
                {
                    if (hSSFForegroundColor.Indexed != s.FillForegroundColor
                         || hSSFBackgroundColor.Indexed != s.FillBackgroundColor
                         || hSSFBorderDiagonalColor.Indexed != s.BorderDiagonalColor
                         || hSSFBottomBorderColor.Indexed != s.BottomBorderColor
                         || hSSFLeftBorderColor.Indexed != s.LeftBorderColor
                         || hSSFRightBorderColor.Indexed != s.RightBorderColor
                         || hSSFTopBorderColor.Indexed != s.TopBorderColor
                         )
                        continue;
                }
                else
                    throw new Exception("Unsupported style type: " + style.GetType().FullName);

                if (styleWorkbook == null)
                {
                    if (style.DataFormat != s.DataFormat
                       || style.FontIndex != s.FontIndex
                       )
                        continue;
                }
                else
                {
                    if (unregisteredStyleDataFormatString != sDataFormat.GetFormat(s.DataFormat))
                        continue;

                    IFont sFont = s.GetFont(Workbook);
                    if (unregisteredStyleFont.Charset != sFont.Charset
                        || unregisteredStyleFont.Color != sFont.Color
                        || unregisteredStyleFont.FontHeight != sFont.FontHeight
                        || unregisteredStyleFont.FontName != sFont.FontName
                        || unregisteredStyleFont.IsBold != sFont.IsBold
                        || unregisteredStyleFont.IsItalic != sFont.IsItalic
                        || unregisteredStyleFont.IsStrikeout != sFont.IsStrikeout
                        || unregisteredStyleFont.TypeOffset != sFont.TypeOffset
                        || unregisteredStyleFont.Underline != sFont.Underline
                        )
                        continue;
                }
                yield return s;
            }
        }

        /// <summary>
        /// Both styles can be unregistered. (!)However, font and format used by them must be registered in the respective workbooks.
        /// </summary>
        /// <param name="fromStyle"></param>
        /// <param name="toStyle"></param>
        /// <param name="toStyleWorkbook"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public ICellStyle CopyStyle(ICellStyle fromStyle, ICellStyle toStyle, IWorkbook toStyleWorkbook = null)
        {
            if (toStyleWorkbook != null && toStyleWorkbook.GetType() != Workbook.GetType())
                throw new Exception("Copying a style in a different type workbook is not supported: " + toStyleWorkbook.GetType().FullName);
            toStyle.Alignment = fromStyle.Alignment;
            toStyle.BorderBottom = fromStyle.BorderBottom;
            toStyle.BorderDiagonal = fromStyle.BorderDiagonal;
            toStyle.BorderDiagonalColor = fromStyle.BorderDiagonalColor;
            toStyle.BorderDiagonalLineStyle = fromStyle.BorderDiagonalLineStyle;
            toStyle.BorderLeft = fromStyle.BorderLeft;
            toStyle.BorderRight = fromStyle.BorderRight;
            toStyle.BorderTop = fromStyle.BorderTop;
            toStyle.BottomBorderColor = fromStyle.BottomBorderColor;
            if (toStyleWorkbook == null)
                toStyle.DataFormat = fromStyle.DataFormat;
            else
            {
                var dataFormat1 = Workbook.CreateDataFormat();
                var dataFormat2 = toStyleWorkbook.CreateDataFormat();
                string sDataFormat;
                try
                {
                    sDataFormat = dataFormat1.GetFormat(fromStyle.DataFormat);
                }
                catch (Exception e)
                {
                    throw new Exception("Style fromStyle has DataFormat=" + fromStyle.DataFormat + " that does not exists in the workbook.", e);
                }
                toStyle.DataFormat = dataFormat2.GetFormat(sDataFormat);
            }
            if (fromStyle is XSSFCellStyle xcs)
            {
                XSSFCellStyle toXcs = toStyle as XSSFCellStyle;
                if (toXcs == null)
                    throw new Exception("Copying style to a different type is not supported: " + toStyle.GetType().FullName);
                toXcs.FillForegroundColorColor = fromStyle.FillForegroundColorColor;
                toXcs.FillBackgroundColorColor = fromStyle.FillBackgroundColorColor;
                toXcs.SetDiagonalBorderColor(xcs.DiagonalBorderXSSFColor);
                toXcs.SetBottomBorderColor(xcs.BottomBorderXSSFColor);
                toXcs.SetLeftBorderColor(xcs.LeftBorderXSSFColor);
                toXcs.SetRightBorderColor(xcs.RightBorderXSSFColor);
                toXcs.SetTopBorderColor(xcs.TopBorderXSSFColor);
            }
            else if (fromStyle is HSSFCellStyle)
            {
                if (!(toStyle is HSSFCellStyle))
                    throw new Exception("Copying style to a different type is not supported: " + toStyle.GetType().FullName);
                if (toStyleWorkbook != null && toStyleWorkbook != Workbook)
                {
                    if (fromStyle.FillForegroundColor > 0)
                    {
                        HSSFColor c = getRegisteredHSSFColor((HSSFWorkbook)toStyleWorkbook, new Color(fromStyle.FillForegroundColorColor));
                        toStyle.FillForegroundColor = c.Indexed;//(!)might be not exactly same color
                    }
                    if (fromStyle.FillBackgroundColor > 0)
                    {
                        HSSFColor c = getRegisteredHSSFColor((HSSFWorkbook)toStyleWorkbook, new Color(fromStyle.FillBackgroundColorColor));
                        toStyle.FillBackgroundColor = c.Indexed;//(!)might be not exactly same color
                    }
                    HSSFPalette palette = ((HSSFWorkbook)Workbook).GetCustomPalette();
                    if (fromStyle.BorderDiagonalColor > 0)
                    {
                        HSSFColor c = getRegisteredHSSFColor((HSSFWorkbook)toStyleWorkbook, new Color(palette.GetColor(fromStyle.BorderDiagonalColor)));
                        toStyle.BorderDiagonalColor = c.Indexed;//(!)might be not exactly same color
                    }
                    if (fromStyle.BottomBorderColor > 0)
                    {
                        HSSFColor c = getRegisteredHSSFColor((HSSFWorkbook)toStyleWorkbook, new Color(palette.GetColor(fromStyle.BottomBorderColor)));
                        toStyle.BottomBorderColor = c.Indexed;//(!)might be not exactly same color
                    }
                    if (fromStyle.LeftBorderColor > 0)
                    {
                        HSSFColor c = getRegisteredHSSFColor((HSSFWorkbook)toStyleWorkbook, new Color(palette.GetColor(fromStyle.LeftBorderColor)));
                        toStyle.LeftBorderColor = c.Indexed;//(!)might be not exactly same color
                    }
                    if (fromStyle.RightBorderColor > 0)
                    {
                        HSSFColor c = getRegisteredHSSFColor((HSSFWorkbook)toStyleWorkbook, new Color(palette.GetColor(fromStyle.RightBorderColor)));
                        toStyle.RightBorderColor = c.Indexed;//(!)might be not exactly same color
                    }
                    if (fromStyle.TopBorderColor > 0)
                    {
                        HSSFColor c = getRegisteredHSSFColor((HSSFWorkbook)toStyleWorkbook, new Color(palette.GetColor(fromStyle.TopBorderColor)));
                        toStyle.TopBorderColor = c.Indexed;//(!)might be not exactly same color
                    }
                }
                else
                {
                    toStyle.FillForegroundColor = fromStyle.FillForegroundColor;
                    toStyle.FillBackgroundColor = fromStyle.FillBackgroundColor;
                    toStyle.BorderDiagonalColor = fromStyle.BorderDiagonalColor;
                    toStyle.BottomBorderColor = fromStyle.BottomBorderColor;
                    toStyle.LeftBorderColor = fromStyle.LeftBorderColor;
                    toStyle.RightBorderColor = fromStyle.RightBorderColor;
                    toStyle.TopBorderColor = fromStyle.TopBorderColor;
                }
            }
            else
                throw new Exception("Unsupported style type: " + fromStyle.GetType().FullName);
            toStyle.FillPattern = fromStyle.FillPattern;
            toStyle.Indention = fromStyle.Indention;
            toStyle.IsHidden = fromStyle.IsHidden;
            toStyle.IsLocked = fromStyle.IsLocked;
            toStyle.Rotation = fromStyle.Rotation;
            toStyle.ShrinkToFit = fromStyle.ShrinkToFit;
            toStyle.VerticalAlignment = fromStyle.VerticalAlignment;
            toStyle.WrapText = fromStyle.WrapText;
            IFont f1;
            try
            {
                f1 = Workbook.GetFontAt(fromStyle.FontIndex);
            }
            catch (Exception e)
            {
                throw new Exception("Style fromStyle has font[@index=" + fromStyle.FontIndex + "] that does not exists in the workbook.", e);
            }
            if (toStyleWorkbook == null)
                toStyle.SetFont(f1);
            else
            {
                IFont f2 = GetRegisteredFont(f1.IsBold, (IndexedColors)Enum.ToObject(typeof(IndexedColors), f1.Color), (short)f1.FontHeight, f1.FontName, f1.IsItalic, f1.IsStrikeout, f1.TypeOffset, f1.Underline);
                toStyle.SetFont(f2);
            }
            return toStyle;
        }

        public ICellStyle CreateUnregisteredStyle()
        {
            if (Workbook is XSSFWorkbook)
            {
                XSSFWorkbook w = new XSSFWorkbook();
                ICellStyle s = new XSSFCellStyle(w.GetStylesSource());
                IFont f = Workbook.NumberOfFonts > 0 ? Workbook.GetFontAt(0) : w.CreateFont();
                s.SetFont(f);//otherwise it throws an exception on accessing font
                return s;
            }
            if (Workbook is HSSFWorkbook)
            {
                HSSFWorkbook w = new HSSFWorkbook();
                ICellStyle s = new HSSFCellStyle(0, new NPOI.HSSF.Record.ExtendedFormatRecord(), w);
                IFont f = Workbook.NumberOfFonts > 0 ? Workbook.GetFontAt(0) : w.CreateFont();
                s.SetFont(f);//set default font
                return s;
            }
            throw new Exception("Unsupported workbook type: " + Workbook.GetType().FullName);
        }

        /// <summary>
        /// Creates an unregistered copy of a style.
        /// </summary>
        /// <param name="fromStyle"></param>
        /// <param name="cloneStyleWorkbook"></param>
        /// <returns></returns>
        public ICellStyle CloneUnregisteredStyle(ICellStyle fromStyle, IWorkbook cloneStyleWorkbook = null)
        {
            ICellStyle toStyle = CreateUnregisteredStyle();
            return CopyStyle(fromStyle, toStyle, cloneStyleWorkbook);
        }

        /// <summary>
        /// If the font does not exists, it is created.
        /// </summary>
        /// <param name="bold"></param>
        /// <param name="color"></param>
        /// <param name="fontHeight"></param>
        /// <param name="name"></param>
        /// <param name="italic"></param>
        /// <param name="strikeout"></param>
        /// <param name="fontSuperScript"></param>
        /// <param name="fontUnderlineType"></param>
        /// <returns></returns>
        public IFont GetRegisteredFont(bool bold, IndexedColors color, short fontHeightInPoints, string name, bool italic = false, bool strikeout = false, FontSuperScript typeOffset = FontSuperScript.None, FontUnderlineType underline = FontUnderlineType.None)
        {
            short fontHeight = (short)(20 * fontHeightInPoints);
            IFont f = Workbook.FindFont(bold, color.Index, fontHeight, name, italic, strikeout, typeOffset, underline);
            if (f == null)
            {
                f = Workbook.CreateFont();
                f.IsBold = bold;
                f.Color = color.Index;
                f.FontHeight = fontHeight;
                f.FontName = name;
                f.IsItalic = italic;
                f.IsStrikeout = strikeout;
                f.TypeOffset = typeOffset;
                f.Underline = underline;
            }
            return f;
        }

        public IEnumerable<ICellStyle> GetStyles()
        {
            for (int i = 0; i < Workbook.NumCellStyles; i++)
            {
                yield return Workbook.GetCellStyleAt(i);
            }
        }

        public IEnumerable<ICellStyle> GetUnusedStyles(params short[] ignoredStyleIds)
        {
            bool usedBySheet(ISheet sheet, ICellStyle style)
            {
                for (int r = 0; r <= sheet.LastRowNum; r++)
                {
                    IRow row = GetRow(r, false);
                    if (row == null)
                        continue;
                    if (row.RowStyle?.Index == style.Index)
                        return true;
                    foreach (ICell c in row.Cells)
                    {
                        if (c.CellStyle?.Index == style.Index)
                            return true;
                    }
                }
                return false;
            }
            bool used(ICellStyle style)
            {
                for (int s = 0; s < Workbook.NumberOfSheets; s++)
                {
                    ISheet sheet = Workbook.GetSheetAt(s);
                    if (usedBySheet(sheet, style))
                        return true;
                }
                return false;
            }
            for (int i = 0; i < Workbook.NumCellStyles; i++)
            {
                var style = Workbook.GetCellStyleAt(i);
                if (ignoredStyleIds.Contains(style.Index))
                    continue;
                if (!used(style))
                    yield return style;
            }
            yield break;
        }
        /// <summary>
        /// Call GetUnusedStyles() after this method to re-use styles.
        /// </summary>
        /// <exception cref="Exception"></exception>
        public void OptimiseStyles()
        {
            if (Workbook is XSSFWorkbook xw)
            {
                //NPOI.XSSF.Model.StylesTable st = xSSFWorkbook.GetStylesSource();
                for (int i = 0; i < Workbook.NumCellStyles; i++)
                {
                    var style = Workbook.GetCellStyleAt(i);
                    foreach (var s in findEqualStyles(style))
                        ReplaceStyle(s, style);
                }
            }
            else if (Workbook is HSSFWorkbook hSSFWorkbook)
            {
                HSSFOptimiser.OptimiseFonts(hSSFWorkbook);
                HSSFOptimiser.OptimiseCellStyles(hSSFWorkbook);
            }
            else
                throw new Exception("Unsupported workbook type: " + Workbook.GetType().FullName);
        }

        IEnumerable<ICellStyle> findEqualStyles(ICellStyle style)
        {
            for (int i = style.Index + 1; i < Workbook.NumCellStyles; i++)
            {
                ICellStyle s = Workbook.GetCellStyleAt(i);
                if (style.Alignment != s.Alignment
                || style.BorderBottom != s.BorderBottom
                || style.BorderDiagonal != s.BorderDiagonal
                || style.BorderDiagonalLineStyle != s.BorderDiagonalLineStyle
                || style.BorderLeft != s.BorderLeft
                || style.BorderRight != s.BorderRight
                || style.BorderTop != s.BorderTop
                || style.FillPattern != s.FillPattern
                || style.Indention != s.Indention
                || style.IsHidden != s.IsHidden
                || style.IsLocked != s.IsLocked
                || style.Rotation != s.Rotation
                || style.ShrinkToFit != s.ShrinkToFit
                || style.VerticalAlignment != s.VerticalAlignment
                || style.WrapText != s.WrapText
                //|| style.BorderDiagonalColor != s.BorderDiagonalColor
                //|| style.BottomBorderColor != s.BottomBorderColor
                //|| style.LeftBorderColor != s.LeftBorderColor
                //|| style.RightBorderColor != s.RightBorderColor
                //|| style.TopBorderColor != s.TopBorderColor
                )
                    continue;

                if (style is XSSFCellStyle uxcs)
                {
                    XSSFCellStyle xcs = s as XSSFCellStyle;
                    if (!AreColorsEqual(uxcs.FillForegroundXSSFColor, xcs.FillForegroundXSSFColor)
                        || !AreColorsEqual(uxcs.FillBackgroundXSSFColor, xcs.FillBackgroundXSSFColor)
                        || !AreColorsEqual(uxcs.DiagonalBorderXSSFColor, xcs.DiagonalBorderXSSFColor)
                        || !AreColorsEqual(uxcs.BottomBorderXSSFColor, xcs.BottomBorderXSSFColor)
                        || !AreColorsEqual(uxcs.LeftBorderXSSFColor, xcs.LeftBorderXSSFColor)
                        || !AreColorsEqual(uxcs.RightBorderXSSFColor, xcs.RightBorderXSSFColor)
                        || !AreColorsEqual(uxcs.TopBorderXSSFColor, xcs.TopBorderXSSFColor)
                        )
                        continue;
                }
                else if (style is HSSFCellStyle hcs)
                {
                    if (hcs.FillForegroundColor != s.FillForegroundColor
                         || hcs.FillBackgroundColor != s.FillBackgroundColor
                         || hcs.BorderDiagonalColor != s.BorderDiagonalColor
                         || hcs.BottomBorderColor != s.BottomBorderColor
                         || hcs.LeftBorderColor != s.LeftBorderColor
                         || hcs.RightBorderColor != s.RightBorderColor
                         || hcs.TopBorderColor != s.TopBorderColor
                         )
                        continue;
                }
                else
                    throw new Exception("Unsupported style type: " + style.GetType().FullName);

                if (style.DataFormat != s.DataFormat
                   || style.FontIndex != s.FontIndex
                   )
                    continue;
                yield return s;
            }
        }

        public void ReplaceStyle(ICellStyle style1, ICellStyle style2)
        {
            ReplaceStyle(null, style1, style2);
        }

        public void SetStyle(ICellStyle style, bool createCells)
        {
            SetStyle(null, style, createCells);
        }

        public void UnsetStyle(ICellStyle style)
        {
            UnsetStyle(null, style);
        }
    }
}