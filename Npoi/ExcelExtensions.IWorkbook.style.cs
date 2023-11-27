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
        /// <summary>
        /// Works for unregistered styles too.
        /// (!)Font must be registered in the workbook though.
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="style"></param>
        /// <exception cref="Exception"></exception>
        public static IFont _GetFont(this IWorkbook workbook, ICellStyle style)
        {
            try
            {
                return workbook.GetFontAt(style.FontIndex);
            }
            catch (Exception e)
            {
                throw new Exception("Could not get font[ID=" + style.FontIndex + " for style[ID=" + style.Index + "]. The font is not registered in the workbook.", e);
            }
        }

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
        /// Comparison is performed by actual parameters. Therefore:
        /// - styles with different indexes and font indexes can be equal;
        /// - styles can be unregistered;
        /// - styles can be of different types;
        /// (!)Unregistered styles must have their fonts registerd in the workbooks.
        /// (!)Unregistered HSSF styles must have their colors registerd in the workbook palette.
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="style"></param>
        /// <param name="searchWorkbook"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        static public IEnumerable<ICellStyle> _FindEqualStyles(this IWorkbook workbook, ICellStyle style, IWorkbook searchWorkbook = null)
        {
            if (searchWorkbook == null)
                searchWorkbook = workbook;
            return _FindEqualStyles(workbook, style, searchWorkbook._GetStyles(), searchWorkbook);
        }

        /// <summary>
        /// Comparison is performed by actual parameters. Therefore:
        /// - styles with different indexes and font indexes can be equal;
        /// - styles can be unregistered;
        /// - styles can be of different types;
        /// (!)Unregistered styles must have their fonts registerd in the workbooks.
        /// (!)Unregistered HSSF styles must have their colors registerd in the workbook palette.
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="style"></param>
        /// <param name="style2"></param>
        /// <param name="workbook2"></param>
        /// <returns></returns>
        static public bool _AreStylesEqual(this IWorkbook workbook, ICellStyle style, ICellStyle style2, IWorkbook workbook2 = null)
        {
            if (workbook2 == null)
                workbook2 = workbook;
            return _FindEqualStyles(workbook, style, new ICellStyle[] { style2 }, workbook2).FirstOrDefault() != null;
        }

        /// <summary>
        /// Comparison is performed by actual parameters. Therefore:
        /// - styles with different indexes and font indexes can be equal;
        /// - styles can be unregistered;
        /// - styles can be of different types;
        /// (!)Unregistered styles must have their fonts registerd in the workbooks.
        /// (!)Unregistered HSSF styles must have their colors registerd in the workbook palette.
        /// </summary>
        /// <param name="workbook">the workbook which the style belongs to</param>
        /// <param name="style">the style to search for</param>
        /// <param name="style2s">the styles to compare with</param>
        /// <param name="workbook2">the workbook which the style2s belong to</param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        static public IEnumerable<ICellStyle> _FindEqualStyles(this IWorkbook workbook, ICellStyle style, IEnumerable<ICellStyle> style2s, IWorkbook workbook2 = null)
        {
            if (workbook2 == null)
                workbook2 = workbook;

            HSSFPalette palette = null;
            HSSFPalette palette2 = null;

            //[System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
            HSSFColor getHSSFColor(HSSFPalette p, short c)
            {
                try
                {
                    return p.GetColor(c);
                }
                catch (Exception e)
                {
                    throw new Exception("Could not get HSSF color[ID=" + c + "]. Most likely the color is not registered.", e);
                }
            }
            bool areHSSFHSSFColorsEqual(short c1, short c2)
            {
                return Excel.AreColorsEqual(getHSSFColor(palette, c1), getHSSFColor(palette2, c2));
            }
            bool areXSSFHSSFColorsEqual(XSSFColor c1, short c2)
            {
                return Excel.AreColorsEqual(c1, getHSSFColor(palette2, c2));
            }
            bool areXSSFXSSFStyleColorsEqual(ICellStyle s1_, ICellStyle s2_)
            {
                XSSFCellStyle s1 = (XSSFCellStyle)s1_;
                XSSFCellStyle s2 = (XSSFCellStyle)s2_;
                return Excel.AreColorsEqual(s1.FillForegroundXSSFColor, s2.FillForegroundXSSFColor)
                    && Excel.AreColorsEqual(s1.FillBackgroundXSSFColor, s2.FillBackgroundXSSFColor)
                    && Excel.AreColorsEqual(s1.DiagonalBorderXSSFColor, s2.DiagonalBorderXSSFColor)
                    && Excel.AreColorsEqual(s1.BottomBorderXSSFColor, s2.BottomBorderXSSFColor)
                    && Excel.AreColorsEqual(s1.LeftBorderXSSFColor, s2.LeftBorderXSSFColor)
                    && Excel.AreColorsEqual(s1.RightBorderXSSFColor, s2.RightBorderXSSFColor)
                    && Excel.AreColorsEqual(s1.TopBorderXSSFColor, s2.TopBorderXSSFColor);
            }
            bool areXSSFHSSFStyleColorsEqual(ICellStyle s1_, ICellStyle s2_)
            {
                XSSFCellStyle s1 = (XSSFCellStyle)s1_;
                HSSFCellStyle s2 = (HSSFCellStyle)s2_;
                return areXSSFHSSFColorsEqual(s1.FillForegroundXSSFColor, s2.FillForegroundColor)
                     && areXSSFHSSFColorsEqual(s1.FillBackgroundXSSFColor, s2.FillBackgroundColor)
                     && areXSSFHSSFColorsEqual(s1.DiagonalBorderXSSFColor, s2.BorderDiagonalColor)
                     && areXSSFHSSFColorsEqual(s1.BottomBorderXSSFColor, s2.BottomBorderColor)
                     && areXSSFHSSFColorsEqual(s1.LeftBorderXSSFColor, s2.LeftBorderColor)
                     && areXSSFHSSFColorsEqual(s1.RightBorderXSSFColor, s2.RightBorderColor)
                     && areXSSFHSSFColorsEqual(s1.TopBorderXSSFColor, s2.TopBorderColor);
            }
            bool areHSSFXSSFStyleColorsEqual(ICellStyle s1_, ICellStyle s2_)
            {
                return areXSSFHSSFStyleColorsEqual(s2_, s1_);
            }
            bool areHSSFHSSFStyleColorsEqualByIndex(ICellStyle s1_, ICellStyle s2_)
            {
                HSSFCellStyle s1 = (HSSFCellStyle)s1_;
                HSSFCellStyle s2 = (HSSFCellStyle)s2_;
                return style.FillForegroundColor == s2.FillForegroundColor
                 && style.FillBackgroundColor == s2.FillBackgroundColor
                 && style.BorderDiagonalColor == s2.BorderDiagonalColor
                 && style.BottomBorderColor == s2.BottomBorderColor
                 && style.LeftBorderColor == s2.LeftBorderColor
                 && style.RightBorderColor == s2.RightBorderColor
                 && style.TopBorderColor == s2.TopBorderColor;
            }
            bool areHSSFHSSFStyleColorsEqualByValue(ICellStyle s1_, ICellStyle s2_)
            {
                return areHSSFHSSFColorsEqual(s1_.FillForegroundColor, s2_.FillForegroundColor)
                && areHSSFHSSFColorsEqual(s1_.FillBackgroundColor, s2_.FillBackgroundColor)
                && areHSSFHSSFColorsEqual(s1_.BorderDiagonalColor, s2_.BorderDiagonalColor)
                && areHSSFHSSFColorsEqual(s1_.BottomBorderColor, s2_.BottomBorderColor)
                && areHSSFHSSFColorsEqual(s1_.LeftBorderColor, s2_.LeftBorderColor)
                && areHSSFHSSFColorsEqual(s1_.RightBorderColor, s2_.RightBorderColor)
                && areHSSFHSSFColorsEqual(s1_.TopBorderColor, s2_.TopBorderColor);
            }

            Func<ICellStyle, ICellStyle, bool> areStyleColorsEqual;

            if (workbook is XSSFWorkbook)
            {
                if (workbook2 is XSSFWorkbook)
                    areStyleColorsEqual = areXSSFXSSFStyleColorsEqual;
                else if (workbook2 is HSSFWorkbook)
                    areStyleColorsEqual = areXSSFHSSFStyleColorsEqual;
                else
                    throw new Exception("Unsupported workbook type: " + workbook2.GetType().FullName);
            }
            else if (workbook is HSSFWorkbook)
            {
                if (workbook2 is XSSFWorkbook)
                    areStyleColorsEqual = areHSSFXSSFStyleColorsEqual;
                else if (workbook2 is HSSFWorkbook)
                {
                    if (workbook2 == workbook)
                        areStyleColorsEqual = areHSSFHSSFStyleColorsEqualByIndex;
                    else
                    {
                        palette = ((HSSFWorkbook)workbook).GetCustomPalette();
                        palette2 = ((HSSFWorkbook)workbook2).GetCustomPalette();
                        areStyleColorsEqual = areHSSFHSSFStyleColorsEqualByValue;
                    }
                }
                else
                    throw new Exception("Unsupported workbook type: " + workbook2.GetType().FullName);
            }
            else
                throw new Exception("Unsupported workbook type: " + workbook.GetType().FullName);

            string dataFormatString = workbook.CreateDataFormat().GetFormat(style.DataFormat);
            IDataFormat dataFormat2 = workbook2.CreateDataFormat();
            IFont font = workbook._GetFont(style);

            foreach (ICellStyle style2 in style2s)
            {
                if (style.Alignment != style2.Alignment
                || style.BorderBottom != style2.BorderBottom
                || style.BorderDiagonal != style2.BorderDiagonal
                || style.BorderDiagonalLineStyle != style2.BorderDiagonalLineStyle
                || style.BorderLeft != style2.BorderLeft
                || style.BorderRight != style2.BorderRight
                || style.BorderTop != style2.BorderTop
                || style.FillPattern != style2.FillPattern
                || style.Indention != style2.Indention
                || style.IsHidden != style2.IsHidden
                || style.IsLocked != style2.IsLocked
                || style.Rotation != style2.Rotation
                || style.ShrinkToFit != style2.ShrinkToFit
                || style.VerticalAlignment != style2.VerticalAlignment
                || style.WrapText != style2.WrapText
                )
                    continue;

                if (!areStyleColorsEqual(style, style2))
                    continue;

                IFont font2 = workbook2._GetFont(style2);
                if (font.Charset != font2.Charset
                    || font.Color != font2.Color
                    || font.FontHeight != font2.FontHeight
                    || font.FontName != font2.FontName
                    || font.IsBold != font2.IsBold
                    || font.IsItalic != font2.IsItalic
                    || font.IsStrikeout != font2.IsStrikeout
                    || font.TypeOffset != font2.TypeOffset
                    || font.Underline != font2.Underline
                    )
                    continue;

                if (dataFormatString != dataFormat2.GetFormat(style2.DataFormat))
                    continue;

                yield return style2;
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
            IFont f1 = workbook._GetFont(fromStyle);
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
                return new XSSFFont(new NPOI.OpenXmlFormats.Spreadsheet.CT_Font(), -1);
            if (workbook is HSSFWorkbook)
                return new HSSFFont(-1, new NPOI.HSSF.Record.FontRecord());
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
