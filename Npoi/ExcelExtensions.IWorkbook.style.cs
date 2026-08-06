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
                cs.SetFillForegroundColor(new XSSFColor(color.RGB, null));
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
        /// Looks for an equal style in the workbook and, if it does not exists, creates a new one.
        /// (!)Incidentally, there is a somewhat analogous method NPOI.SS.Util.CellUtil.SetCellStyleProperties() which is not as handy in use though.
        /// </summary>
        /// <param name="workbook1">workbook which style1 belongs to (even if it is not registered)</param>
        /// <param name="style1">it can be either a unregistered style created by CreateUnregisteredStyle() and modified as needed, or a registered style.</param>
        ///// <param name="reuseUnusedStyle">(!)slows down performance. It makes sense ony when styles need optimization</param>
        /// <param name="workbook2">workbook where an equivalent of style1 to be registered</param>
        /// <returns></returns>
        static public ICellStyle _GetRegisteredStyle(this IWorkbook workbook1, ICellStyle style1/*, bool reuseUnusedStyle = false*/, IWorkbook workbook2 = null)
        {
            if (workbook2 == null)
                workbook2 = workbook1;
            ICellStyle style2 = workbook1._FindEqualStyles(style1, workbook2).FirstOrDefault();
            if (style2 != null)
                return style2;
            //if (reuseUnusedStyle)
            //{
            //    style2 = workbook1._GetUnusedStyles().FirstOrDefault();
            //    if (style2 == null)
            //        style2 = workbook1.CreateCellStyle();
            //}
            //else
            style2 = workbook2.CreateCellStyle();
            return workbook1._CopyStyle(style1, style2, workbook2);
        }

        /// <summary>
        /// Comparison is performed by actual parameters. Hence:
        /// - styles with different indexes and font indexes can be equal;
        /// - styles can be unregistered;
        /// - styles can be of different types;
        /// (!)Unregistered styles must have their fonts registerd in the workbooks.
        /// (!)Unregistered HSSF styles must have their colors registerd in the workbook palette.
        /// </summary>
        /// <param name="workbook1">workbook which style1 belongs to</param>
        /// <param name="style1">it can be either a unregistered style created by CreateUnregisteredStyle() and modified as needed, or a registered style.</param>
        /// <param name="workbook2">workbook where equivalents of style1 to be searched</param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        static public IEnumerable<ICellStyle> _FindEqualStyles(this IWorkbook workbook1, ICellStyle style1, IWorkbook workbook2 = null)
        {
            if (workbook2 == null)
                workbook2 = workbook1;
            return _FindEqualStyles(workbook1, style1, workbook2._GetStyles(), workbook2);
        }

        /// <summary>
        /// Comparison is performed by actual parameters. Therefore:
        /// - styles with different indexes and font indexes can be equal;
        /// - styles can be unregistered;
        /// - styles can be of different types;
        /// (!)Unregistered styles must have their fonts registerd in the workbooks.
        /// (!)Unregistered HSSF styles must have their colors registerd in the workbook palette.
        /// </summary>
        /// <param name="workbook1">workbook which style1 belongs to</param>
        /// <param name="style1">it can be either a unregistered style created by CreateUnregisteredStyle() and modified as needed, or a registered style.</param>
        /// <param name="style2">it can be either a unregistered style created by CreateUnregisteredStyle() and modified as needed, or a registered style.</param>
        /// <param name="workbook2">workbook which style2 belongs to</param>
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
        /// <param name="workbook1">the workbook which style1 belongs to</param>
        /// <param name="style1">the style to search for</param>
        /// <param name="style2s">the styles to compare with</param>
        /// <param name="workbook2">the workbook which the style2s belong to</param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        static public IEnumerable<ICellStyle> _FindEqualStyles(this IWorkbook workbook1, ICellStyle style1, IEnumerable<ICellStyle> style2s, IWorkbook workbook2 = null)
        {
            if (workbook2 == null)
                workbook2 = workbook1;

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
                return style1.FillForegroundColor == s2.FillForegroundColor
                 && style1.FillBackgroundColor == s2.FillBackgroundColor
                 && style1.BorderDiagonalColor == s2.BorderDiagonalColor
                 && style1.BottomBorderColor == s2.BottomBorderColor
                 && style1.LeftBorderColor == s2.LeftBorderColor
                 && style1.RightBorderColor == s2.RightBorderColor
                 && style1.TopBorderColor == s2.TopBorderColor;
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

            if (workbook1 is XSSFWorkbook)
            {
                if (workbook2 is XSSFWorkbook)
                    areStyleColorsEqual = areXSSFXSSFStyleColorsEqual;
                else if (workbook2 is HSSFWorkbook)
                    areStyleColorsEqual = areXSSFHSSFStyleColorsEqual;
                else
                    throw new Exception("Unsupported workbook type: " + workbook2.GetType().FullName);
            }
            else if (workbook1 is HSSFWorkbook)
            {
                if (workbook2 is XSSFWorkbook)
                    areStyleColorsEqual = areHSSFXSSFStyleColorsEqual;
                else if (workbook2 is HSSFWorkbook)
                {
                    if (workbook2 == workbook1)
                        areStyleColorsEqual = areHSSFHSSFStyleColorsEqualByIndex;
                    else
                    {
                        palette1 = ((HSSFWorkbook)workbook1).GetCustomPalette();
                        palette2 = ((HSSFWorkbook)workbook2).GetCustomPalette();
                        areStyleColorsEqual = areHSSFHSSFStyleColorsEqualByValue;
                    }
                }
                else
                    throw new Exception("Unsupported workbook type: " + workbook2.GetType().FullName);
            }
            else
                throw new Exception("Unsupported workbook type: " + workbook1.GetType().FullName);

            string dataFormat1String = workbook1.CreateDataFormat().GetFormat(style1.DataFormat);
            IDataFormat dataFormat2 = workbook2.CreateDataFormat();
            IFont font1 = workbook1._GetFont(style1);

            foreach (ICellStyle style2 in style2s)
            {
                if (style1.Alignment != style2.Alignment
                || style1.BorderBottom != style2.BorderBottom
                || style1.BorderDiagonal != style2.BorderDiagonal
                || style1.BorderDiagonalLineStyle != style2.BorderDiagonalLineStyle
                || style1.BorderLeft != style2.BorderLeft
                || style1.BorderRight != style2.BorderRight
                || style1.BorderTop != style2.BorderTop
                || style1.FillPattern != style2.FillPattern
                || style1.Indention != style2.Indention
                || style1.IsHidden != style2.IsHidden
                || style1.IsLocked != style2.IsLocked
                || style1.Rotation != style2.Rotation
                || style1.ShrinkToFit != style2.ShrinkToFit
                || style1.VerticalAlignment != style2.VerticalAlignment
                || style1.WrapText != style2.WrapText
                )
                    continue;

                if (!areStyleColorsEqual(style1, style2))
                    continue;

                IFont font2 = workbook2._GetFont(style2);
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
        /// <param name="workbook1">the workbook which style1 belongs to</param>
        /// <param name="style1">the style to be copied</param>
        /// <param name="style2">the style to copy to</param>
        /// <param name="workbook2">the workbook which the style2 belongs to</param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        static public ICellStyle _CopyStyle(this IWorkbook workbook1, ICellStyle style1, ICellStyle style2, IWorkbook workbook2 = null)
        {
            if (workbook2 == null)
                workbook2 = workbook1;

            style2.Alignment = style1.Alignment;
            style2.BorderBottom = style1.BorderBottom;
            style2.BorderDiagonal = style1.BorderDiagonal;
            style2.BorderDiagonalLineStyle = style1.BorderDiagonalLineStyle;
            style2.BorderLeft = style1.BorderLeft;
            style2.BorderRight = style1.BorderRight;
            style2.BorderTop = style1.BorderTop;
            if (workbook2 == workbook1)
                style2.DataFormat = style1.DataFormat;
            else
            {
                var dataFormat1 = workbook1.CreateDataFormat();
                var dataFormat2 = workbook2.CreateDataFormat();
                string sDataFormat;
                try
                {
                    sDataFormat = dataFormat1.GetFormat(style1.DataFormat);
                }
                catch (Exception e)
                {
                    throw new Exception("Style fromStyle has DataFormat=" + style1.DataFormat + " that does not exists in the workbook.", e);
                }
                style2.DataFormat = dataFormat2.GetFormat(sDataFormat);
            }

            if (style1 is XSSFCellStyle xcs)
            {
                if (style2 is XSSFCellStyle xcs2)
                {
                    xcs2.FillForegroundColorColor = style1.FillForegroundColorColor;
                    xcs2.FillBackgroundColorColor = style1.FillBackgroundColorColor;
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
            else if (style1 is HSSFCellStyle)
            {
                if (style2 is XSSFCellStyle xcs2)
                {
                    HSSFPalette palette = ((HSSFWorkbook)workbook1).GetCustomPalette();
                    XSSFColor getHSSFXSSFColor(short color)
                    {
                        if (color == 0)
                            return null;
                        return new XSSFColor(new Excel.Color(palette.GetColor(color)).RGB, null);
                    }
                    xcs2.FillForegroundXSSFColor = getHSSFXSSFColor(style1.FillForegroundColor);
                    xcs2.FillBackgroundXSSFColor = getHSSFXSSFColor(style1.FillBackgroundColor);
                    xcs2.SetDiagonalBorderColor(getHSSFXSSFColor(style1.BorderDiagonalColor));
                    xcs2.SetBottomBorderColor(getHSSFXSSFColor(style1.BottomBorderColor));
                    xcs2.SetLeftBorderColor(getHSSFXSSFColor(style1.LeftBorderColor));
                    xcs2.SetRightBorderColor(getHSSFXSSFColor(style1.RightBorderColor));
                    xcs2.SetTopBorderColor(getHSSFXSSFColor(style1.TopBorderColor));
                }
                else if (style2 is HSSFCellStyle)
                {
                    if (workbook2 != workbook1)
                    {
                        HSSFPalette palette = ((HSSFWorkbook)workbook1).GetCustomPalette();
                        short getHSSFHSSFColor(short color)
                        {
                            if (color == 0)
                                return 0;
                            HSSFColor c = Excel.GetRegisteredHSSFColor((HSSFWorkbook)workbook2, new Excel.Color(palette.GetColor(color)));
                            return c.Indexed;//(!)might be not exactly same color
                        }
                        style2.FillForegroundColor = getHSSFHSSFColor(style1.FillForegroundColor);
                        style2.FillBackgroundColor = getHSSFHSSFColor(style1.FillBackgroundColor);
                        style2.BorderDiagonalColor = getHSSFHSSFColor(style1.BorderDiagonalColor);
                        style2.BottomBorderColor = getHSSFHSSFColor(style1.BottomBorderColor);
                        style2.LeftBorderColor = getHSSFHSSFColor(style1.LeftBorderColor);
                        style2.RightBorderColor = getHSSFHSSFColor(style1.RightBorderColor);
                        style2.TopBorderColor = getHSSFHSSFColor(style1.TopBorderColor);
                    }
                    else
                    {
                        style2.FillForegroundColor = style1.FillForegroundColor;
                        style2.FillBackgroundColor = style1.FillBackgroundColor;
                        style2.BorderDiagonalColor = style1.BorderDiagonalColor;
                        style2.BottomBorderColor = style1.BottomBorderColor;
                        style2.LeftBorderColor = style1.LeftBorderColor;
                        style2.RightBorderColor = style1.RightBorderColor;
                        style2.TopBorderColor = style1.TopBorderColor;
                    }
                }
                else
                    throw new Exception("Unsupported workbook type: " + workbook2.GetType().FullName);
            }
            else
                throw new Exception("Unsupported workbook type: " + workbook1.GetType().FullName);

            style2.FillPattern = style1.FillPattern;
            style2.Indention = style1.Indention;
            style2.IsHidden = style1.IsHidden;
            style2.IsLocked = style1.IsLocked;
            style2.Rotation = style1.Rotation;
            style2.ShrinkToFit = style1.ShrinkToFit;
            style2.VerticalAlignment = style1.VerticalAlignment;
            style2.WrapText = style1.WrapText;
            IFont f1 = workbook1._GetFont(style1);
            if (workbook2 == workbook1)
                style2.SetFont(f1);
            else
            {
                IFont f2 = workbook2._GetRegisteredFont(f1);
                style2.SetFont(f2);
            }
            return style2;
        }

        /// <summary>
        /// (!)Experimental. Copies listes properties from style1 to style2. Both styles can be unregistered.
        /// </summary>
        /// <param name="workbook1">the workbook which style1 belongs to</param>
        /// <param name="stylePropertieNames">properties to be copied</param>
        /// <param name="style1">the style to be copied</param>
        /// <param name="style2">the style to copy into</param>
        /// <param name="workbook2">the workbook which the style2 belongs to</param>
        /// <exception cref="Exception"></exception>
        static public void _BlendStyle(this IWorkbook workbook1, IEnumerable<string> stylePropertieNames, ICellStyle style1, ICellStyle style2, IWorkbook workbook2 = null)
        {
            if (workbook2 == null)
                workbook2 = workbook1;

            HashSet<string> spns = new HashSet<string>(stylePropertieNames);

            if (spns.Contains("Alignment"))
                style2.Alignment = style1.Alignment;
            if (spns.Contains("BorderBottom"))
                style2.BorderBottom = style1.BorderBottom;
            if (spns.Contains("BorderDiagonal"))
                style2.BorderDiagonal = style1.BorderDiagonal;
            if (spns.Contains("BorderDiagonalLineStyle"))
                style2.BorderDiagonalLineStyle = style1.BorderDiagonalLineStyle;
            if (spns.Contains("BorderLeft"))
                style2.BorderLeft = style1.BorderLeft;
            if (spns.Contains("BorderRight"))
                style2.BorderRight = style1.BorderRight;
            if (spns.Contains("BorderTop"))
                style2.BorderTop = style1.BorderTop;

            if (workbook1 == workbook2)
                style2.DataFormat = style1.DataFormat;
            else
            {
                var dataFormat1 = workbook1.CreateDataFormat();
                var dataFormat2 = workbook2.CreateDataFormat();
                string sDataFormat;
                try
                {
                    sDataFormat = dataFormat1.GetFormat(style1.DataFormat);
                }
                catch (Exception e)
                {
                    throw new Exception("Style style1 has DataFormat=" + style1.DataFormat + " that does not exists in the workbook1.", e);
                }
                style2.DataFormat = dataFormat2.GetFormat(sDataFormat);
            }

            if (style1 is XSSFCellStyle xcs)
            {
                if (style2 is XSSFCellStyle xcs2)
                {
                    if (spns.Contains("FillForegroundColorColor"))
                        xcs2.FillForegroundColorColor = style1.FillForegroundColorColor;
                    if (spns.Contains("FillBackgroundColorColor"))
                        xcs2.FillBackgroundColorColor = style1.FillBackgroundColorColor;
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
            else if (style1 is HSSFCellStyle)
            {
                if (style2 is XSSFCellStyle xcs2)
                {
                    HSSFPalette palette = ((HSSFWorkbook)workbook1).GetCustomPalette();
                    XSSFColor getHSSFXSSFColor(short color)
                    {
                        if (color == 0)
                            return null;
                        return new XSSFColor(new Excel.Color(palette.GetColor(color)).RGB, null);
                    }
                    if (spns.Contains("FillForegroundColor"))
                        xcs2.FillForegroundXSSFColor = getHSSFXSSFColor(style1.FillForegroundColor);
                    if (spns.Contains("FillBackgroundColor"))
                        xcs2.FillBackgroundXSSFColor = getHSSFXSSFColor(style1.FillBackgroundColor);
                    if (spns.Contains("BorderDiagonalColor"))
                        xcs2.SetDiagonalBorderColor(getHSSFXSSFColor(style1.BorderDiagonalColor));
                    if (spns.Contains("BottomBorderColor"))
                        xcs2.SetBottomBorderColor(getHSSFXSSFColor(style1.BottomBorderColor));
                    if (spns.Contains("LeftBorderColor"))
                        xcs2.SetLeftBorderColor(getHSSFXSSFColor(style1.LeftBorderColor));
                    if (spns.Contains("RightBorderColor"))
                        xcs2.SetRightBorderColor(getHSSFXSSFColor(style1.RightBorderColor));
                    if (spns.Contains("TopBorderColor"))
                        xcs2.SetTopBorderColor(getHSSFXSSFColor(style1.TopBorderColor));
                }
                else if (style2 is HSSFCellStyle)
                {
                    if (workbook2 != workbook1)
                    {
                        HSSFPalette palette = ((HSSFWorkbook)workbook1).GetCustomPalette();
                        short getHSSFHSSFColor(short color)
                        {
                            if (color == 0)
                                return 0;
                            NPOI.HSSF.Util.HSSFColor c = Excel.GetRegisteredHSSFColor((HSSFWorkbook)workbook2, new Excel.Color(palette.GetColor(color)));
                            return c.Indexed;//(!)might be not exactly same color
                        }
                        if (spns.Contains("FillForegroundColor"))
                            style2.FillForegroundColor = getHSSFHSSFColor(style1.FillForegroundColor);
                        if (spns.Contains("FillBackgroundColor"))
                            style2.FillBackgroundColor = getHSSFHSSFColor(style1.FillBackgroundColor);
                        if (spns.Contains("BorderDiagonalColor"))
                            style2.BorderDiagonalColor = getHSSFHSSFColor(style1.BorderDiagonalColor);
                        if (spns.Contains("BottomBorderColor"))
                            style2.BottomBorderColor = getHSSFHSSFColor(style1.BottomBorderColor);
                        if (spns.Contains("LeftBorderColor"))
                            style2.LeftBorderColor = getHSSFHSSFColor(style1.LeftBorderColor);
                        if (spns.Contains("RightBorderColor"))
                            style2.RightBorderColor = getHSSFHSSFColor(style1.RightBorderColor);
                        if (spns.Contains("TopBorderColor"))
                            style2.TopBorderColor = getHSSFHSSFColor(style1.TopBorderColor);
                    }
                    else
                    {
                        if (spns.Contains("FillForegroundColor"))
                            style2.FillForegroundColor = style1.FillForegroundColor;
                        if (spns.Contains("FillBackgroundColor"))
                            style2.FillBackgroundColor = style1.FillBackgroundColor;
                        if (spns.Contains("BorderDiagonalColor"))
                            style2.BorderDiagonalColor = style1.BorderDiagonalColor;
                        if (spns.Contains("BottomBorderColor"))
                            style2.BottomBorderColor = style1.BottomBorderColor;
                        if (spns.Contains("LeftBorderColor"))
                            style2.LeftBorderColor = style1.LeftBorderColor;
                        if (spns.Contains("RightBorderColor"))
                            style2.RightBorderColor = style1.RightBorderColor;
                        if (spns.Contains("TopBorderColor"))
                            style2.TopBorderColor = style1.TopBorderColor;
                    }
                }
                else
                    throw new Exception("Unsupported workbook2 type: " + workbook2.GetType().FullName);
            }
            else
                throw new Exception("Unsupported workbook1 type: " + workbook1.GetType().FullName);

            if (spns.Contains("FillPattern"))
                style2.FillPattern = style1.FillPattern;
            if (spns.Contains("Indention"))
                style2.Indention = style1.Indention;
            if (spns.Contains("IsHidden"))
                style2.IsHidden = style1.IsHidden;
            if (spns.Contains("IsLocked"))
                style2.IsLocked = style1.IsLocked;
            if (spns.Contains("Rotation"))
                style2.Rotation = style1.Rotation;
            if (spns.Contains("ShrinkToFit"))
                style2.ShrinkToFit = style1.ShrinkToFit;
            if (spns.Contains("VerticalAlignment"))
                style2.VerticalAlignment = style1.VerticalAlignment;
            if (spns.Contains("WrapText"))
                style2.WrapText = style1.WrapText;
            if (spns.Contains("FontIndex"))
            {
                IFont f1 = workbook1._GetFont(style1);
                if (workbook2 == workbook1)
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
                return new XSSFFont(new NPOI.OpenXmlFormats.Spreadsheet.CT_Font(), -1, null);
            if (workbook is HSSFWorkbook)
                return new HSSFFont(-1, new NPOI.HSSF.Record.FontRecord());
            throw new Exception("Unsupported workbook type: " + workbook.GetType().FullName);
        }

        /// <summary>
        /// Creates an unregistered copy of a style which can be unregistered or registered. 
        /// (!)However, font, format and indexed colors used by it must exist in the source workbook.
        /// </summary>
        /// <param name="workbook1">workbook which style1 belongs to</param>
        /// <param name="style1">style to be copied</param>
        /// <param name="workbook2">workbook where the cloned style will belong to</param>
        /// <returns></returns>
        static public ICellStyle _CloneUnregisteredStyle(this IWorkbook workbook1, ICellStyle style1, IWorkbook workbook2 = null)
        {
            if (workbook2 == null)
                workbook2 = workbook1;
            ICellStyle style2 = workbook2._CreateUnregisteredStyle();
            return workbook1._CopyStyle(style1, style2, workbook2);
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
