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
        /// Intended for either adding or removing backgound color.
        /// The style can be unregistered but on HSSFWorkbook the color will be added to the workbook's palette.
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="style"></param>
        /// <param name="color"></param>
        /// <param name="fillPattern"></param>
        static public void _Highlight(this IWorkbook workbook, ICellStyle style, Excel.Color color, FillPattern fillPattern = FillPattern.SolidForeground)
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
                HSSFColor hssfColor = Excel.GetRegisteredHSSFColor((HSSFWorkbook)workbook, color);
                cs.FillForegroundColor = hssfColor.Indexed;
                cs.FillPattern = fillPattern;
                return;
            }
            throw new Exception("Unsupported workbook type: " + workbook.GetType().FullName);
        }

        /// <summary>
        /// Looks for the equal style in the workbook and, if it does not exists, creates a new one.
        /// (!)Incidentally, there is a somewhat analogous method NPOI.SS.Util.CellUtil.SetCellStyleProperties() which is not as handy in use though.
        /// </summary>
        /// <param name="unregisteredStyle">it is a style created by CreateUnregisteredStyle() and then modified as needed. But it can be a registered style, too.</param>
        /// <param name="reuseUnusedStyle">(!)slows down performance. It makes sense ony when styles need optimization</param>
        /// <param name="unregisteredStyleWorkbook"></param>
        /// <returns></returns>
        static public ICellStyle _GetRegisteredStyle(this IWorkbook workbook, ICellStyle unregisteredStyle, bool reuseUnusedStyle = false, IWorkbook unregisteredStyleWorkbook = null)
        {
            ICellStyle style = workbook._FindEqualStyles(unregisteredStyle, unregisteredStyleWorkbook).FirstOrDefault();
            if (style != null)
                return style;
            if (reuseUnusedStyle)
            {
                style = workbook._GetUnusedStyles().FirstOrDefault();
                if (style == null)
                    style = workbook.CreateCellStyle();
            }
            else
                style = workbook.CreateCellStyle();
            return workbook._CopyStyle(unregisteredStyle, style);
        }

        /// <summary>
        /// Comparison is performed by actual parameters. Hence:
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
        /// <param name="workbook1"></param>
        /// <param name="style1"></param>
        /// <param name="style2"></param>
        /// <param name="workbook2"></param>
        /// <returns></returns>
        static public bool _AreStylesEqual(this IWorkbook workbook1, ICellStyle style1, ICellStyle style2, IWorkbook workbook2 = null)
        {
            if (workbook2 == null)
                workbook2 = workbook1;
            return _FindEqualStyles(workbook1, style1, new ICellStyle[] { style2 }, workbook2).FirstOrDefault() != null;
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
        /// <param name="searchStyles">the styles to compare with</param>
        /// <param name="searchWorkbook">the workbook which the style2s belong to</param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        static public IEnumerable<ICellStyle> _FindEqualStyles(this IWorkbook workbook, ICellStyle style, IEnumerable<ICellStyle> searchStyles, IWorkbook searchWorkbook = null)
        {
            if (searchWorkbook == null)
                searchWorkbook = workbook;

            HSSFPalette palette1 = null;
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
                return Excel.AreColorsEqual(getHSSFColor(palette1, c1), getHSSFColor(palette2, c2));
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
                if (searchWorkbook is XSSFWorkbook)
                    areStyleColorsEqual = areXSSFXSSFStyleColorsEqual;
                else if (searchWorkbook is HSSFWorkbook)
                    areStyleColorsEqual = areXSSFHSSFStyleColorsEqual;
                else
                    throw new Exception("Unsupported workbook type: " + searchWorkbook.GetType().FullName);
            }
            else if (workbook is HSSFWorkbook)
            {
                if (searchWorkbook is XSSFWorkbook)
                    areStyleColorsEqual = areHSSFXSSFStyleColorsEqual;
                else if (searchWorkbook is HSSFWorkbook)
                {
                    if (searchWorkbook == workbook)
                        areStyleColorsEqual = areHSSFHSSFStyleColorsEqualByIndex;
                    else
                    {
                        palette1 = ((HSSFWorkbook)workbook).GetCustomPalette();
                        palette2 = ((HSSFWorkbook)searchWorkbook).GetCustomPalette();
                        areStyleColorsEqual = areHSSFHSSFStyleColorsEqualByValue;
                    }
                }
                else
                    throw new Exception("Unsupported workbook type: " + searchWorkbook.GetType().FullName);
            }
            else
                throw new Exception("Unsupported workbook type: " + workbook.GetType().FullName);

            string dataFormat1String = workbook.CreateDataFormat().GetFormat(style.DataFormat);
            IDataFormat dataFormat2 = searchWorkbook.CreateDataFormat();
            IFont font1 = workbook._GetFont(style);

            foreach (ICellStyle style2 in searchStyles)
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

                IFont font2 = searchWorkbook._GetFont(style2);
                if (!Excel.AreFontsEqual(font1, font2))
                    continue;

                if (dataFormat1String != dataFormat2.GetFormat(style2.DataFormat))
                    continue;

                yield return style2;
            }
        }

        /// <summary>
        /// Both styles can be unregistered. (!)However, font, format and indexed colors used by them must exist in the source workbook.
        /// Font, format and indexed colors, if do not exist in the destination workbook, will be created there.
        /// Allows copying between styles of different types.
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="style"></param>
        /// <param name="style2"></param>
        /// <param name="workbook2"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        static public ICellStyle _CopyStyle(this IWorkbook workbook, ICellStyle style, ICellStyle style2, IWorkbook workbook2 = null)
        {
            if (workbook2 == null)
                workbook2 = workbook;

            style2.Alignment = style.Alignment;
            style2.BorderBottom = style.BorderBottom;
            style2.BorderDiagonal = style.BorderDiagonal;
            style2.BorderDiagonalLineStyle = style.BorderDiagonalLineStyle;
            style2.BorderLeft = style.BorderLeft;
            style2.BorderRight = style.BorderRight;
            style2.BorderTop = style.BorderTop;
            if (workbook2 == workbook)
                style2.DataFormat = style.DataFormat;
            else
            {
                var dataFormat1 = workbook.CreateDataFormat();
                var dataFormat2 = workbook2.CreateDataFormat();
                string sDataFormat;
                try
                {
                    sDataFormat = dataFormat1.GetFormat(style.DataFormat);
                }
                catch (Exception e)
                {
                    throw new Exception("Style fromStyle has DataFormat=" + style.DataFormat + " that does not exists in the workbook.", e);
                }
                style2.DataFormat = dataFormat2.GetFormat(sDataFormat);
            }

            if (style is XSSFCellStyle xcs)
            {
                if (style2 is XSSFCellStyle xcs2)
                {
                    xcs2.FillForegroundColorColor = style.FillForegroundColorColor;
                    xcs2.FillBackgroundColorColor = style.FillBackgroundColorColor;
                    xcs2.SetDiagonalBorderColor(xcs.DiagonalBorderXSSFColor);
                    xcs2.SetBottomBorderColor(xcs.BottomBorderXSSFColor);
                    xcs2.SetLeftBorderColor(xcs.LeftBorderXSSFColor);
                    xcs2.SetRightBorderColor(xcs.RightBorderXSSFColor);
                    xcs2.SetTopBorderColor(xcs.TopBorderXSSFColor);
                }
                else if (style2 is HSSFCellStyle)
                {
                    short getXSSFHSSFColor(XSSFColor color)
                    {
                        if (color == null)
                            return 0;
                        HSSFColor c = Excel.GetRegisteredHSSFColor((HSSFWorkbook)workbook2, new Excel.Color(color));
                        return c.Indexed;//(!)might be not exactly same color
                    }
                    style2.FillForegroundColor = getXSSFHSSFColor(xcs.FillForegroundXSSFColor);
                    style2.FillBackgroundColor = getXSSFHSSFColor(xcs.FillBackgroundXSSFColor);
                    style2.BorderDiagonalColor = getXSSFHSSFColor(xcs.DiagonalBorderXSSFColor);
                    style2.BottomBorderColor = getXSSFHSSFColor(xcs.BottomBorderXSSFColor);
                    style2.LeftBorderColor = getXSSFHSSFColor(xcs.LeftBorderXSSFColor);
                    style2.RightBorderColor = getXSSFHSSFColor(xcs.RightBorderXSSFColor);
                    style2.TopBorderColor = getXSSFHSSFColor(xcs.TopBorderXSSFColor);
                }
                else
                    throw new Exception("Unsupported workbook type: " + workbook2.GetType().FullName);
            }
            else if (style is HSSFCellStyle)
            {
                if (style2 is XSSFCellStyle xcs2)
                {
                    HSSFPalette palette = ((HSSFWorkbook)workbook).GetCustomPalette();
                    XSSFColor getHSSFXSSFColor(short color)
                    {
                        if (color == 0)
                            return null;
                        return new XSSFColor(new Excel.Color(palette.GetColor(color)).RGB);
                    }
                    xcs2.FillForegroundXSSFColor = getHSSFXSSFColor(style.FillForegroundColor);
                    xcs2.FillBackgroundXSSFColor = getHSSFXSSFColor(style.FillBackgroundColor);
                    xcs2.SetDiagonalBorderColor(getHSSFXSSFColor(style.BorderDiagonalColor));
                    xcs2.SetBottomBorderColor(getHSSFXSSFColor(style.BottomBorderColor));
                    xcs2.SetLeftBorderColor(getHSSFXSSFColor(style.LeftBorderColor));
                    xcs2.SetRightBorderColor(getHSSFXSSFColor(style.RightBorderColor));
                    xcs2.SetTopBorderColor(getHSSFXSSFColor(style.TopBorderColor));
                }
                else if (style2 is HSSFCellStyle)
                {
                    if (workbook2 != workbook)
                    {
                        HSSFPalette palette = ((HSSFWorkbook)workbook).GetCustomPalette();
                        short getHSSFHSSFColor(short color)
                        {
                            if (color == 0)
                                return 0;
                            HSSFColor c = Excel.GetRegisteredHSSFColor((HSSFWorkbook)workbook2, new Excel.Color(palette.GetColor(color)));
                            return c.Indexed;//(!)might be not exactly same color
                        }
                        style2.FillForegroundColor = getHSSFHSSFColor(style.FillForegroundColor);
                        style2.FillBackgroundColor = getHSSFHSSFColor(style.FillBackgroundColor);
                        style2.BorderDiagonalColor = getHSSFHSSFColor(style.BorderDiagonalColor);
                        style2.BottomBorderColor = getHSSFHSSFColor(style.BottomBorderColor);
                        style2.LeftBorderColor = getHSSFHSSFColor(style.LeftBorderColor);
                        style2.RightBorderColor = getHSSFHSSFColor(style.RightBorderColor);
                        style2.TopBorderColor = getHSSFHSSFColor(style.TopBorderColor);
                    }
                    else
                    {
                        style2.FillForegroundColor = style.FillForegroundColor;
                        style2.FillBackgroundColor = style.FillBackgroundColor;
                        style2.BorderDiagonalColor = style.BorderDiagonalColor;
                        style2.BottomBorderColor = style.BottomBorderColor;
                        style2.LeftBorderColor = style.LeftBorderColor;
                        style2.RightBorderColor = style.RightBorderColor;
                        style2.TopBorderColor = style.TopBorderColor;
                    }
                }
                else
                    throw new Exception("Unsupported workbook type: " + workbook2.GetType().FullName);
            }
            else
                throw new Exception("Unsupported workbook type: " + workbook.GetType().FullName);

            style2.FillPattern = style.FillPattern;
            style2.Indention = style.Indention;
            style2.IsHidden = style.IsHidden;
            style2.IsLocked = style.IsLocked;
            style2.Rotation = style.Rotation;
            style2.ShrinkToFit = style.ShrinkToFit;
            style2.VerticalAlignment = style.VerticalAlignment;
            style2.WrapText = style.WrapText;
            IFont f1 = workbook._GetFont(style);
            if (workbook2 == workbook)
                style2.SetFont(f1);
            else
            {
                IFont f2 = workbook2._GetRegisteredFont(f1);
                style2.SetFont(f2);
            }
            return style2;
        }

        /// <summary>
        /// (!)Experimental. Copies listes properties from maskStyle to style. Both styles can be unregistered.
        /// </summary>
        /// <param name="maskWorkbook"></param>
        /// <param name="stylePropertieNames"></param>
        /// <param name="maskStyle"></param>
        /// <param name="style2"></param>
        /// <param name="workbook2"></param>
        /// <exception cref="Exception"></exception>
        static public void _BlendStyle(this IWorkbook maskWorkbook, IEnumerable<string> stylePropertieNames, ICellStyle maskStyle, ICellStyle style2, IWorkbook workbook2 = null)
        {
            if (workbook2 == null)
                workbook2 = maskWorkbook;

            HashSet<string> spns = new HashSet<string>(stylePropertieNames);

            if (spns.Contains("Alignment"))
                style2.Alignment = maskStyle.Alignment;
            if (spns.Contains("BorderBottom"))
                style2.BorderBottom = maskStyle.BorderBottom;
            if (spns.Contains("BorderDiagonal"))
                style2.BorderDiagonal = maskStyle.BorderDiagonal;
            if (spns.Contains("BorderDiagonalLineStyle"))
                style2.BorderDiagonalLineStyle = maskStyle.BorderDiagonalLineStyle;
            if (spns.Contains("BorderLeft"))
                style2.BorderLeft = maskStyle.BorderLeft;
            if (spns.Contains("BorderRight"))
                style2.BorderRight = maskStyle.BorderRight;
            if (spns.Contains("BorderTop"))
                style2.BorderTop = maskStyle.BorderTop;

            if (maskWorkbook == workbook2)
                style2.DataFormat = maskStyle.DataFormat;
            else
            {
                var dataFormat1 = maskWorkbook.CreateDataFormat();
                var dataFormat2 = workbook2.CreateDataFormat();
                string sDataFormat;
                try
                {
                    sDataFormat = dataFormat1.GetFormat(maskStyle.DataFormat);
                }
                catch (Exception e)
                {
                    throw new Exception("Style maskStyle has DataFormat=" + maskStyle.DataFormat + " that does not exists in the maskWorkbook.", e);
                }
                style2.DataFormat = dataFormat2.GetFormat(sDataFormat);
            }

            if (maskStyle is XSSFCellStyle xcs)
            {
                if (style2 is XSSFCellStyle xcs2)
                {
                    if (spns.Contains("FillForegroundColorColor"))
                        xcs2.FillForegroundColorColor = maskStyle.FillForegroundColorColor;
                    if (spns.Contains("FillBackgroundColorColor"))
                        xcs2.FillBackgroundColorColor = maskStyle.FillBackgroundColorColor;
                    if (spns.Contains("DiagonalBorderXSSFColor"))
                        xcs2.SetDiagonalBorderColor(xcs.DiagonalBorderXSSFColor);
                    if (spns.Contains("BottomBorderXSSFColor"))
                        xcs2.SetBottomBorderColor(xcs.BottomBorderXSSFColor);
                    if (spns.Contains("LeftBorderXSSFColor"))
                        xcs2.SetLeftBorderColor(xcs.LeftBorderXSSFColor);
                    if (spns.Contains("RightBorderXSSFColor"))
                        xcs2.SetRightBorderColor(xcs.RightBorderXSSFColor);
                    if (spns.Contains("TopBorderXSSFColor"))
                        xcs2.SetTopBorderColor(xcs.TopBorderXSSFColor);
                }
                else if (style2 is HSSFCellStyle)
                {
                    short getXSSFHSSFColor(XSSFColor color)
                    {
                        if (color == null)
                            return 0;
                        NPOI.HSSF.Util.HSSFColor c = Excel.GetRegisteredHSSFColor((HSSFWorkbook)workbook2, new Excel.Color(color));
                        return c.Indexed;//(!)might be not exactly same color
                    }
                    if (spns.Contains("FillForegroundXSSFColor"))
                        style2.FillForegroundColor = getXSSFHSSFColor(xcs.FillForegroundXSSFColor);
                    if (spns.Contains("FillBackgroundXSSFColor"))
                        style2.FillBackgroundColor = getXSSFHSSFColor(xcs.FillBackgroundXSSFColor);
                    if (spns.Contains("DiagonalBorderXSSFColor"))
                        style2.BorderDiagonalColor = getXSSFHSSFColor(xcs.DiagonalBorderXSSFColor);
                    if (spns.Contains("BottomBorderXSSFColor"))
                        style2.BottomBorderColor = getXSSFHSSFColor(xcs.BottomBorderXSSFColor);
                    if (spns.Contains("LeftBorderXSSFColor"))
                        style2.LeftBorderColor = getXSSFHSSFColor(xcs.LeftBorderXSSFColor);
                    if (spns.Contains("RightBorderXSSFColor"))
                        style2.RightBorderColor = getXSSFHSSFColor(xcs.RightBorderXSSFColor);
                    if (spns.Contains("TopBorderXSSFColor"))
                        style2.TopBorderColor = getXSSFHSSFColor(xcs.TopBorderXSSFColor);
                }
                else
                    throw new Exception("Unsupported workbook2 type: " + workbook2.GetType().FullName);
            }
            else if (maskStyle is HSSFCellStyle)
            {
                if (style2 is XSSFCellStyle xcs2)
                {
                    HSSFPalette palette = ((HSSFWorkbook)maskWorkbook).GetCustomPalette();
                    XSSFColor getHSSFXSSFColor(short color)
                    {
                        if (color == 0)
                            return null;
                        return new XSSFColor(new Excel.Color(palette.GetColor(color)).RGB);
                    }
                    if (spns.Contains("FillForegroundColor"))
                        xcs2.FillForegroundXSSFColor = getHSSFXSSFColor(maskStyle.FillForegroundColor);
                    if (spns.Contains("FillBackgroundColor"))
                        xcs2.FillBackgroundXSSFColor = getHSSFXSSFColor(maskStyle.FillBackgroundColor);
                    if (spns.Contains("BorderDiagonalColor"))
                        xcs2.SetDiagonalBorderColor(getHSSFXSSFColor(maskStyle.BorderDiagonalColor));
                    if (spns.Contains("BottomBorderColor"))
                        xcs2.SetBottomBorderColor(getHSSFXSSFColor(maskStyle.BottomBorderColor));
                    if (spns.Contains("LeftBorderColor"))
                        xcs2.SetLeftBorderColor(getHSSFXSSFColor(maskStyle.LeftBorderColor));
                    if (spns.Contains("RightBorderColor"))
                        xcs2.SetRightBorderColor(getHSSFXSSFColor(maskStyle.RightBorderColor));
                    if (spns.Contains("TopBorderColor"))
                        xcs2.SetTopBorderColor(getHSSFXSSFColor(maskStyle.TopBorderColor));
                }
                else if (style2 is HSSFCellStyle)
                {
                    if (workbook2 != maskWorkbook)
                    {
                        HSSFPalette palette = ((HSSFWorkbook)maskWorkbook).GetCustomPalette();
                        short getHSSFHSSFColor(short color)
                        {
                            if (color == 0)
                                return 0;
                            NPOI.HSSF.Util.HSSFColor c = Excel.GetRegisteredHSSFColor((HSSFWorkbook)workbook2, new Excel.Color(palette.GetColor(color)));
                            return c.Indexed;//(!)might be not exactly same color
                        }
                        if (spns.Contains("FillForegroundColor"))
                            style2.FillForegroundColor = getHSSFHSSFColor(maskStyle.FillForegroundColor);
                        if (spns.Contains("FillBackgroundColor"))
                            style2.FillBackgroundColor = getHSSFHSSFColor(maskStyle.FillBackgroundColor);
                        if (spns.Contains("BorderDiagonalColor"))
                            style2.BorderDiagonalColor = getHSSFHSSFColor(maskStyle.BorderDiagonalColor);
                        if (spns.Contains("BottomBorderColor"))
                            style2.BottomBorderColor = getHSSFHSSFColor(maskStyle.BottomBorderColor);
                        if (spns.Contains("LeftBorderColor"))
                            style2.LeftBorderColor = getHSSFHSSFColor(maskStyle.LeftBorderColor);
                        if (spns.Contains("RightBorderColor"))
                            style2.RightBorderColor = getHSSFHSSFColor(maskStyle.RightBorderColor);
                        if (spns.Contains("TopBorderColor"))
                            style2.TopBorderColor = getHSSFHSSFColor(maskStyle.TopBorderColor);
                    }
                    else
                    {
                        if (spns.Contains("FillForegroundColor"))
                            style2.FillForegroundColor = maskStyle.FillForegroundColor;
                        if (spns.Contains("FillBackgroundColor"))
                            style2.FillBackgroundColor = maskStyle.FillBackgroundColor;
                        if (spns.Contains("BorderDiagonalColor"))
                            style2.BorderDiagonalColor = maskStyle.BorderDiagonalColor;
                        if (spns.Contains("BottomBorderColor"))
                            style2.BottomBorderColor = maskStyle.BottomBorderColor;
                        if (spns.Contains("LeftBorderColor"))
                            style2.LeftBorderColor = maskStyle.LeftBorderColor;
                        if (spns.Contains("RightBorderColor"))
                            style2.RightBorderColor = maskStyle.RightBorderColor;
                        if (spns.Contains("TopBorderColor"))
                            style2.TopBorderColor = maskStyle.TopBorderColor;
                    }
                }
                else
                    throw new Exception("Unsupported workbook2 type: " + workbook2.GetType().FullName);
            }
            else
                throw new Exception("Unsupported maskWorkbook type: " + maskWorkbook.GetType().FullName);

            if (spns.Contains("FillPattern"))
                style2.FillPattern = maskStyle.FillPattern;
            if (spns.Contains("Indention"))
                style2.Indention = maskStyle.Indention;
            if (spns.Contains("IsHidden"))
                style2.IsHidden = maskStyle.IsHidden;
            if (spns.Contains("IsLocked"))
                style2.IsLocked = maskStyle.IsLocked;
            if (spns.Contains("Rotation"))
                style2.Rotation = maskStyle.Rotation;
            if (spns.Contains("ShrinkToFit"))
                style2.ShrinkToFit = maskStyle.ShrinkToFit;
            if (spns.Contains("VerticalAlignment"))
                style2.VerticalAlignment = maskStyle.VerticalAlignment;
            if (spns.Contains("WrapText"))
                style2.WrapText = maskStyle.WrapText;
            if (spns.Contains("FontIndex"))
            {
                IFont f1 = maskWorkbook._GetFont(maskStyle);
                if (workbook2 == maskWorkbook)
                    style2.SetFont(f1);
                else
                {
                    IFont f2 = workbook2._GetRegisteredFont(f1);
                    style2.SetFont(f2);
                }
            }
        }

        /// <summary>
        /// Unregistered style's index = -1
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        static public ICellStyle _CreateUnregisteredStyle(this IWorkbook workbook)
        {
            IFont f = workbook.NumberOfFonts > 0 ? workbook.GetFontAt(0) : workbook.CreateFont();
            if (workbook is XSSFWorkbook)
            {
                XSSFWorkbook w = new XSSFWorkbook();
                ICellStyle s = w.GetStylesSource().CreateCellStyle();
                if (XSSFCellStyle_cellXfId_FI == null)
                    XSSFCellStyle_cellXfId_FI = s.GetType().GetField("_cellXfId", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                XSSFCellStyle_cellXfId_FI.SetValue(s, -1);
                s.SetFont(f);//otherwise it throws an exception on accessing font
                return s;
            }
            if (workbook is HSSFWorkbook)
            {
                HSSFWorkbook w = new HSSFWorkbook();
                ICellStyle s = new HSSFCellStyle(-1, new NPOI.HSSF.Record.ExtendedFormatRecord(), w);
                s.SetFont(f);//set default font
                return s;
            }
            throw new Exception("Unsupported workbook type: " + workbook.GetType().FullName);
        }
        static System.Reflection.FieldInfo XSSFCellStyle_cellXfId_FI = null;

        /// Unregistered font's index = -1
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
        /// <param name="style">can be unregistered</param>
        /// <param name="clonedStyleWorkbook"></param>
        /// <returns></returns>
        static public ICellStyle _CloneUnregisteredStyle(this IWorkbook workbook, ICellStyle style, IWorkbook clonedStyleWorkbook = null)
        {
            if (clonedStyleWorkbook == null)
                clonedStyleWorkbook = workbook;
            ICellStyle toStyle = clonedStyleWorkbook._CreateUnregisteredStyle();
            return workbook._CopyStyle(style, toStyle, clonedStyleWorkbook);
        }

        static public IEnumerable<ICellStyle> _GetStyles(this IWorkbook workbook)
        {
            for (int i = 0; i < workbook.NumCellStyles; i++)
                yield return workbook.GetCellStyleAt(i);
        }

        /// <summary>
        /// Finds styles in the workbook that are not used and hence can be used as new.
        /// (!)To make it efficient, run _OptimizeStyles() once and then time to time call this to get unused styles until it return nothing.
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="ignoredStyleIds"></param>
        /// <returns></returns>
        static public IEnumerable<ICellStyle> _GetUnusedStyles(this IWorkbook workbook, params short[] ignoredStyleIds)
        {
            bool isUsed(ICellStyle style)
            {
                foreach (var sheet in workbook._GetSheets())
                {
                    int maxY = sheet.LastRowNum + 1;
                    for (int y = 1; y <= maxY; y++)
                    {
                        IRow row = sheet._GetRow(y, false);
                        if (row == null)
                            continue;
                        if (row.RowStyle?.Index == style.Index)
                            return true;
                        int maxX = row.LastCellNum;
                        for (int x = 1; x <= maxX; x++)
                        {
                            ICell c = row._GetCell(x, false);
                            if (c?.CellStyle.Index == style.Index)
                                return true;
                        }
                    }
                }
                return false;
            }
            foreach (var style in workbook._GetStyles().Where(a => !ignoredStyleIds.Contains(a.Index)).OrderByDescending(a => a.Index))
                if (!isUsed(style))
                    yield return style;
        }

        static public void _OptimizeStylesAndFonts(this IWorkbook workbook, out List<ICellStyle> unusedStyles, out List<IFont> unusedFonts)
        {
            workbook._OptimizeFonts(out unusedFonts);
            workbook._OptimizeStyles(out unusedStyles);
        }

        /// <summary>
        /// Makes all the duplicated styles unused so they can be used as new.
        /// (!)Tends to be slow on large sheets.
        /// </summary>
        static public void _OptimizeStyles(this IWorkbook workbook, out List<ICellStyle> unusedStyles)
        {

            unusedStyles = new List<ICellStyle>();
            var styles = workbook._GetStyles().ToList();
            while (styles.Count > 0)
            {
                var style = styles[0];
                styles.RemoveAt(0);
                List<ICellStyle> style2s = workbook._FindEqualStyles(style, styles).ToList();

                foreach (var sheet in workbook._GetSheets())
                {
                    int maxY = sheet.LastRowNum + 1;
                    for (int y = 1; y <= maxY; y++)
                    {
                        IRow row = sheet._GetRow(y, false);
                        if (row == null)
                            continue;
                        if (row.RowStyle != null && style2s.Contains(row.RowStyle))
                            row.RowStyle = style;
                        int maxX = row.LastCellNum;
                        for (int x = 1; x <= maxX; x++)
                        {
                            ICell c = row._GetCell(x, false);
                            if (c != null && style2s.Contains(c.CellStyle))
                                c.CellStyle = style;
                        }
                    }
                }

                styles = styles.Except(style2s).ToList();
                unusedStyles.AddRange(style2s);
            }
        }
    }
}
