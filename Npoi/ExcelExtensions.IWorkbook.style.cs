//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace Cliver
{
    static public partial class ExcelExtensions
    {
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

        /// <summary>
        /// Intended for either adding or removing backgound color.
        /// The style can be unregistered but on HSSFWorkbook the color will be added to the palette.
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="style"></param>
        /// <param name="color"></param>
        /// <param name="fillPattern"></param>
        /// <returns></returns>
        static public void _Highlight(this IWorkbook workbook, ICellStyle style, Excel.Color color, FillPattern fillPattern = FillPattern.SolidForeground)
        {
            workbook._highlight(style, color, fillPattern);
        }

        //public enum HighlightOption
        //{
        //    GetRegisteredUniqueStyle,
        //    GetRegisteredStyle,
        //    GetUnregisteredStyle
        //}

        /// <summary>
        /// Intended for either adding or removing backgound color.
        /// The style can be unregistered but on HSSFWorkbook the color will be added to the palette.
        /// </summary>
        static internal void _highlight(this IWorkbook workbook, ICellStyle style, Excel.Color color, FillPattern fillPattern = FillPattern.SolidForeground)
        {
            if (style == null)
                return;
            if (workbook is XSSFWorkbook)
            {
                XSSFCellStyle cs = (XSSFCellStyle)style;
                if (color == null)
                {
                    cs.SetFillForegroundColor(null);
                    cs.FillPattern = FillPattern.NoFill;
                    return;
                }
                cs.SetFillForegroundColor(new XSSFColor(color.RGB));
                cs.FillPattern = fillPattern;
                return;
            }
            if (workbook is HSSFWorkbook)
            {
                HSSFCellStyle cs = (HSSFCellStyle)style;
                if (color == null)
                {
                    cs.FillForegroundColor = 0;
                    cs.FillPattern = FillPattern.NoFill;
                    return;
                }
                HSSFColor hssfColor = getRegisteredHSSFColor((HSSFWorkbook)workbook, color);
                cs.FillForegroundColor = hssfColor.Indexed;
                cs.FillPattern = fillPattern;
                return;
            }
            throw new Exception("Unsupported workbook type: " + workbook.GetType().FullName);
        }

        static HSSFColor getRegisteredHSSFColor(HSSFWorkbook workbook, Excel.Color color)
        {
            HSSFPalette palette = workbook.GetCustomPalette();
            HSSFColor hssfColor = palette.FindColor(color.R, color.G, color.B);
            if (hssfColor != null)
                return hssfColor;
            try
            {
                hssfColor = palette.AddColor(color.R, color.G, color.B);
            }
            catch
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
        static public ICellStyle _GetRegisteredStyle(this IWorkbook workbook, ICellStyle unregisteredStyle, IWorkbook unregisteredStyleWorkbook = null)
        {
            if (unregisteredStyleWorkbook != null && unregisteredStyleWorkbook.GetType() != workbook.GetType())
                throw new Exception("Registering a style in a different type workbook is not supported: " + workbook.GetType().FullName);

            ICellStyle style = workbook._FindEqualStyles(unregisteredStyle, unregisteredStyleWorkbook).FirstOrDefault();
            if (style != null)
                return style;
            style = workbook.CreateCellStyle();
            return workbook._CopyStyle(unregisteredStyle, style);
        }

        /// <summary>
        /// Style can be unregistered.
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="style"></param>
        /// <param name="styleWorkbook"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        static public IEnumerable<ICellStyle> _FindEqualStyles(this IWorkbook workbook, ICellStyle style, IWorkbook styleWorkbook = null)
        {
            if (styleWorkbook != null && styleWorkbook.GetType() != workbook.GetType())
                throw new Exception("Comparing a style from a different type workbook is not supported: " + workbook.GetType().FullName);

            HSSFColor hSSFForegroundColor = null;
            HSSFColor hSSFBackgroundColor = null;
            HSSFColor hSSFBorderDiagonalColor = null;
            HSSFColor hSSFBottomBorderColor = null;
            HSSFColor hSSFLeftBorderColor = null;
            HSSFColor hSSFRightBorderColor = null;
            HSSFColor hSSFTopBorderColor = null;
            if (workbook is HSSFWorkbook hw)
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
            if (styleWorkbook != null && styleWorkbook != workbook)
            {
                unregisteredStyleDataFormatString = styleWorkbook.CreateDataFormat().GetFormat(style.DataFormat);
                sDataFormat = workbook.CreateDataFormat();
                unregisteredStyleFont = style.GetFont(styleWorkbook);
            }

            foreach (ICellStyle s in workbook._GetStyles())
            {
                if (styleWorkbook == null || styleWorkbook == workbook)
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
                    if (!Excel.AreColorsEqual(uxcs.FillForegroundXSSFColor, xcs.FillForegroundXSSFColor)
                        || !Excel.AreColorsEqual(uxcs.FillBackgroundXSSFColor, xcs.FillBackgroundXSSFColor)
                        || !Excel.AreColorsEqual(uxcs.DiagonalBorderXSSFColor, xcs.DiagonalBorderXSSFColor)
                        || !Excel.AreColorsEqual(uxcs.BottomBorderXSSFColor, xcs.BottomBorderXSSFColor)
                        || !Excel.AreColorsEqual(uxcs.LeftBorderXSSFColor, xcs.LeftBorderXSSFColor)
                        || !Excel.AreColorsEqual(uxcs.RightBorderXSSFColor, xcs.RightBorderXSSFColor)
                        || !Excel.AreColorsEqual(uxcs.TopBorderXSSFColor, xcs.TopBorderXSSFColor)
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

                    IFont sFont = s.GetFont(workbook);
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
        /// Both styles can be unregistered. (!)However, font and format used by them must exist in the source workbook.
        /// Font and format, if do not exist in the destination workbook, will be created there.
        /// </summary>
        /// <param name="fromStyle"></param>
        /// <param name="toStyle"></param>
        /// <param name="toStyleWorkbook"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        static public ICellStyle _CopyStyle(this IWorkbook workbook, ICellStyle fromStyle, ICellStyle toStyle, IWorkbook toStyleWorkbook = null)
        {
            if (toStyleWorkbook != null && toStyleWorkbook.GetType() != workbook.GetType())
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
                var dataFormat1 = workbook.CreateDataFormat();
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
                if (toStyleWorkbook != null && toStyleWorkbook != workbook)
                {
                    if (fromStyle.FillForegroundColor > 0)
                    {
                        HSSFColor c = getRegisteredHSSFColor((HSSFWorkbook)toStyleWorkbook, new Excel.Color(fromStyle.FillForegroundColorColor));
                        toStyle.FillForegroundColor = c.Indexed;//(!)might be not exactly same color
                    }
                    if (fromStyle.FillBackgroundColor > 0)
                    {
                        HSSFColor c = getRegisteredHSSFColor((HSSFWorkbook)toStyleWorkbook, new Excel.Color(fromStyle.FillBackgroundColorColor));
                        toStyle.FillBackgroundColor = c.Indexed;//(!)might be not exactly same color
                    }
                    HSSFPalette palette = ((HSSFWorkbook)workbook).GetCustomPalette();
                    if (fromStyle.BorderDiagonalColor > 0)
                    {
                        HSSFColor c = getRegisteredHSSFColor((HSSFWorkbook)toStyleWorkbook, new Excel.Color(palette.GetColor(fromStyle.BorderDiagonalColor)));
                        toStyle.BorderDiagonalColor = c.Indexed;//(!)might be not exactly same color
                    }
                    if (fromStyle.BottomBorderColor > 0)
                    {
                        HSSFColor c = getRegisteredHSSFColor((HSSFWorkbook)toStyleWorkbook, new Excel.Color(palette.GetColor(fromStyle.BottomBorderColor)));
                        toStyle.BottomBorderColor = c.Indexed;//(!)might be not exactly same color
                    }
                    if (fromStyle.LeftBorderColor > 0)
                    {
                        HSSFColor c = getRegisteredHSSFColor((HSSFWorkbook)toStyleWorkbook, new Excel.Color(palette.GetColor(fromStyle.LeftBorderColor)));
                        toStyle.LeftBorderColor = c.Indexed;//(!)might be not exactly same color
                    }
                    if (fromStyle.RightBorderColor > 0)
                    {
                        HSSFColor c = getRegisteredHSSFColor((HSSFWorkbook)toStyleWorkbook, new Excel.Color(palette.GetColor(fromStyle.RightBorderColor)));
                        toStyle.RightBorderColor = c.Indexed;//(!)might be not exactly same color
                    }
                    if (fromStyle.TopBorderColor > 0)
                    {
                        HSSFColor c = getRegisteredHSSFColor((HSSFWorkbook)toStyleWorkbook, new Excel.Color(palette.GetColor(fromStyle.TopBorderColor)));
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
                f1 = workbook.GetFontAt(fromStyle.FontIndex);
            }
            catch (Exception e)
            {
                throw new Exception("Style fromStyle has font[@index=" + fromStyle.FontIndex + "] that does not exists in the workbook.", e);
            }
            if (toStyleWorkbook == null)
                toStyle.SetFont(f1);
            else
            {
                IFont f2 = toStyleWorkbook._GetRegisteredFont(f1);
                toStyle.SetFont(f2);
            }
            return toStyle;
        }

        static public ICellStyle _CreateUnregisteredStyle(this IWorkbook workbook)
        {
            IFont f = workbook.NumberOfFonts > 0 ? workbook.GetFontAt(0) : workbook.CreateFont();
            //IFont f = workbook._CreateUnregisteredFont();
            if (workbook is XSSFWorkbook)
            {
                XSSFWorkbook w = new XSSFWorkbook();
                ICellStyle s = new XSSFCellStyle(w.GetStylesSource());
                s.SetFont(f);//otherwise it throws an exception on accessing font
                return s;
            }
            if (workbook is HSSFWorkbook)
            {
                HSSFWorkbook w = new HSSFWorkbook();
                ICellStyle s = new HSSFCellStyle(0, new NPOI.HSSF.Record.ExtendedFormatRecord(), w);
                s.SetFont(f);//set default font
                return s;
            }
            throw new Exception("Unsupported workbook type: " + workbook.GetType().FullName);
        }

        static public IFont _CreateUnregisteredFont(this IWorkbook workbook)
        {
            if (workbook is XSSFWorkbook)
                return new XSSFFont();
            if (workbook is HSSFWorkbook)
                return new HSSFFont(0, new NPOI.HSSF.Record.FontRecord());
            throw new Exception("Unsupported workbook type: " + workbook.GetType().FullName);
        }

        /// <summary>
        /// Creates an unregistered copy of a style.
        /// </summary>
        /// <param name="fromStyle"></param>
        /// <param name="cloneStyleWorkbook"></param>
        /// <returns></returns>
        static public ICellStyle _CloneUnregisteredStyle(this IWorkbook workbook, ICellStyle fromStyle, IWorkbook cloneStyleWorkbook = null)
        {
            ICellStyle toStyle = workbook._CreateUnregisteredStyle();
            return workbook._CopyStyle(fromStyle, toStyle, cloneStyleWorkbook);
        }

        /// <summary>
        /// Creates an unregistered copy of a font.
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="font"></param>
        /// <returns></returns>
        static public IFont _CloneUnregisteredFont(this IWorkbook workbook, IFont font)
        {
            IFont f = workbook._CreateUnregisteredFont();
            f.IsBold = font.IsBold;
            f.Color = font.Color;
            f.FontHeight = font.FontHeight;
            f.FontName = font.FontName;
            f.IsItalic = font.IsItalic;
            f.IsStrikeout = font.IsStrikeout;
            f.TypeOffset = font.TypeOffset;
            f.Underline = font.Underline;
            return f;
        }

        /// <summary>
        /// If the font does not exists, it is created.
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="font"></param>
        /// <returns></returns>
        static public IFont _GetRegisteredFont(this IWorkbook workbook, IFont font)
        {
            IFont f = workbook.FindFont(font.IsBold, font.Color, (short)font.FontHeight, font.FontName, font.IsItalic, font.IsStrikeout, font.TypeOffset, font.Underline);
            if (f == null)
            {
                f = workbook.CreateFont();
                f.IsBold = font.IsBold;
                f.Color = font.Color;
                f.FontHeight = font.FontHeight;
                f.FontName = font.FontName;
                f.IsItalic = font.IsItalic;
                f.IsStrikeout = font.IsStrikeout;
                f.TypeOffset = font.TypeOffset;
                f.Underline = font.Underline;
            }
            return f;
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
        static public IFont _GetRegisteredFont(this IWorkbook workbook, bool bold, short color, short fontHeightInPoints, string name, bool italic = false, bool strikeout = false, FontSuperScript typeOffset = FontSuperScript.None, FontUnderlineType underline = FontUnderlineType.None)
        {
            short fontHeight = (short)(20 * fontHeightInPoints);
            IFont f = workbook.FindFont(bold, color, fontHeight, name, italic, strikeout, typeOffset, underline);
            if (f == null)
            {
                f = workbook.CreateFont();
                f.IsBold = bold;
                f.Color = color;
                f.FontHeight = fontHeight;
                f.FontName = name;
                f.IsItalic = italic;
                f.IsStrikeout = strikeout;
                f.TypeOffset = typeOffset;
                f.Underline = underline;
            }
            return f;
        }

        static public IEnumerable<ICellStyle> _GetStyles(this IWorkbook workbook)
        {
            for (int i = 0; i < workbook.NumCellStyles; i++)
            {
                yield return workbook.GetCellStyleAt(i);
            }
        }

        static public IEnumerable<ICellStyle> _GetUnusedStyles(this IWorkbook workbook, params short[] ignoredStyleIds)
        {
            bool usedBySheet(ISheet sheet, ICellStyle style)
            {
                for (int r = 0; r <= sheet.LastRowNum; r++)
                {
                    IRow row = sheet._GetRow(r, false);
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
                for (int s = 0; s < workbook.NumberOfSheets; s++)
                {
                    ISheet sheet = workbook.GetSheetAt(s);
                    if (usedBySheet(sheet, style))
                        return true;
                }
                return false;
            }
            for (int i = 0; i < workbook.NumCellStyles; i++)
            {
                var style = workbook.GetCellStyleAt(i);
                if (ignoredStyleIds.Contains(style.Index))
                    continue;
                if (!used(style))
                    yield return style;
            }
        }

        /// <summary>
        /// Makes all the duplicated styles unused. Call GetUnusedStyles() after this method to re-use styles.
        /// (!)Tends to be slow on big sheets.
        /// </summary>
        /// <exception cref="Exception"></exception>
        static public void _OptimiseStyles(this IWorkbook workbook)
        {
            //if (workbook is XSSFWorkbook xw)
            //{
            for (short i = 0; i < workbook.NumberOfFonts; i++)
            {
                var font = workbook.GetFontAt(i);
                for (short j = (short)(i + 1); j < workbook.NumberOfFonts; j++)
                {
                    var f = workbook.GetFontAt(j);
                    if (font.IsBold == f.IsBold
                        && font.Color == f.Color
                        && font.FontHeight == f.FontHeight
                        && font.FontName == f.FontName
                        && font.IsItalic == f.IsItalic
                        && font.IsStrikeout == f.IsStrikeout
                        && font.TypeOffset == f.TypeOffset
                        && font.Underline == f.Underline
                        )
                        for (int s = 0; s < workbook.NumCellStyles; s++)
                        {
                            var style = workbook.GetCellStyleAt(s);
                            if (style.FontIndex == f.Index)
                                style.SetFont(font);
                        }
                }
            }

            //NPOI.XSSF.Model.StylesTable st = xSSFWorkbook.GetStylesSource();
            for (int i = 0; i < workbook.NumCellStyles; i++)
            {
                var style = workbook.GetCellStyleAt(i);
                foreach (var s in workbook.findEqualStyles(style))
                    foreach (var sheet in workbook._GetSheets())
                        sheet._ReplaceStyle(s, style);
            }
            //}
            //else if (workbook is HSSFWorkbook hSSFWorkbook)
            //{
            //    HSSFOptimiser.OptimiseFonts(hSSFWorkbook);
            //    HSSFOptimiser.OptimiseCellStyles(hSSFWorkbook);
            //}
            //else
            //    throw new Exception("Unsupported workbook type: " + workbook.GetType().FullName);
        }

        static IEnumerable<ICellStyle> findEqualStyles(this IWorkbook workbook, ICellStyle style)
        {
            for (int i = style.Index + 1; i < workbook.NumCellStyles; i++)
            {
                ICellStyle s = workbook.GetCellStyleAt(i);
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
                    if (!Excel.AreColorsEqual(uxcs.FillForegroundXSSFColor, xcs.FillForegroundXSSFColor)
                        || !Excel.AreColorsEqual(uxcs.FillBackgroundXSSFColor, xcs.FillBackgroundXSSFColor)
                        || !Excel.AreColorsEqual(uxcs.DiagonalBorderXSSFColor, xcs.DiagonalBorderXSSFColor)
                        || !Excel.AreColorsEqual(uxcs.BottomBorderXSSFColor, xcs.BottomBorderXSSFColor)
                        || !Excel.AreColorsEqual(uxcs.LeftBorderXSSFColor, xcs.LeftBorderXSSFColor)
                        || !Excel.AreColorsEqual(uxcs.RightBorderXSSFColor, xcs.RightBorderXSSFColor)
                        || !Excel.AreColorsEqual(uxcs.TopBorderXSSFColor, xcs.TopBorderXSSFColor)
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
    }
}
