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

            public Color(System.Drawing.Color color) : this(color.ToArgb())
            {
            }
        }

        public ICellStyle Highlight(ICellStyle style, Color color, FillPattern fillPattern = FillPattern.SolidForeground, bool createUniqueStyleOnly = true)
        {
            return highlight(this, style, createUniqueStyleOnly, color, fillPattern);
        }

        ///// <summary>
        ///// Is intended for either adding or removing backgound color.
        ///// </summary>
        ///// <exception cref="Exception"></exception>
        //static internal ICellStyle highlight(IWorkbook workbook, ICellStyle style, Color color, FillPattern fillPattern = FillPattern.SolidForeground)
        //{
        //    if (workbook is XSSFWorkbook)
        //    {
        //        XSSFCellStyle cs = style == null ? (XSSFCellStyle)workbook.CreateCellStyle() : (XSSFCellStyle)style;
        //        if (color == null)
        //        {
        //            cs.SetFillForegroundColor(null);
        //            cs.FillPattern = FillPattern.NoFill;
        //        }
        //        else
        //        {
        //            cs.SetFillForegroundColor(new XSSFColor(color.RGB));
        //            cs.FillPattern = fillPattern;
        //        }
        //        return cs;
        //    }
        //    if (workbook is HSSFWorkbook)
        //    {
        //        HSSFCellStyle cs = style == null ? (HSSFCellStyle)workbook.CreateCellStyle() : (HSSFCellStyle)style;
        //        if (color == null)
        //        {
        //            cs.FillForegroundColor = 0;
        //            cs.FillPattern = FillPattern.NoFill;
        //        }
        //        else
        //        {
        //            HSSFPalette palette = ((HSSFWorkbook)workbook).GetCustomPalette();
        //            NPOI.HSSF.Util.HSSFColor hssfColor = palette.FindColor(color.R, color.G, color.B);
        //            if (hssfColor == null)
        //            {
        //                try
        //                {
        //                    hssfColor = palette.AddColor(color.R, color.G, color.B);
        //                }
        //                catch (Exception e)
        //                {//pallete is full
        //                    short? findUnusedColorIndex()
        //                    {
        //                        for (short j = 0x8; j <= 0x40; j++)//the first color in the palette has the index 0x8, the second has the index 0x9, etc. through 0x40
        //                        {
        //                            int i = 0;
        //                            for (; i < workbook.NumCellStyles; i++)
        //                            {
        //                                var s = workbook.GetCellStyleAt(i);
        //                                if (s.BorderDiagonalColor == j
        //                                    || s.BottomBorderColor == j
        //                                    || s.FillBackgroundColor == j
        //                                    || s.FillForegroundColor == j
        //                                    || s.LeftBorderColor == j
        //                                    || s.RightBorderColor == j
        //                                    || s.TopBorderColor == j
        //                                    )
        //                                    break;
        //                            }
        //                            if (i >= workbook.NumCellStyles)
        //                                return j;
        //                        }
        //                        return null;
        //                    }
        //                    short? ci = findUnusedColorIndex();
        //                    if (ci == null)
        //                        ci = palette.FindSimilarColor(color.R, color.G, color.B).Indexed;
        //                    palette.SetColorAtIndex(ci.Value, color.R, color.G, color.B);
        //                    hssfColor = palette.GetColor(ci.Value);
        //                }
        //            }
        //            cs.FillForegroundColor = hssfColor.Indexed;
        //            cs.FillPattern = fillPattern;
        //        }
        //        return cs;
        //    }
        //    throw new Exception("Unsupported workbook type: " + workbook.GetType().FullName);
        //}

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
        static internal ICellStyle highlight(Excel excel, ICellStyle style, bool createUniqueStyleOnly, Color color, FillPattern fillPattern = FillPattern.SolidForeground)
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
                if (createUniqueStyleOnly)
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
                if (createUniqueStyleOnly)
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

            HSSFColor hSSFForegroundColor = null;
            HSSFColor hSSFBackgroundColor = null;
            if (Workbook is HSSFWorkbook hw)
            {
                HSSFPalette palette = hw.GetCustomPalette();
                hSSFForegroundColor = palette.FindColor(unregisteredStyle.FillForegroundColorColor.RGB[0], unregisteredStyle.FillForegroundColorColor.RGB[1], unregisteredStyle.FillForegroundColorColor.RGB[2]);
                if (hSSFForegroundColor == null)
                    goto CREATE_STYLE;
                hSSFBackgroundColor = palette.FindColor(unregisteredStyle.FillBackgroundColorColor.RGB[0], unregisteredStyle.FillBackgroundColorColor.RGB[1], unregisteredStyle.FillBackgroundColorColor.RGB[2]);
                if (hSSFBackgroundColor == null)
                    goto CREATE_STYLE;
            }

            string unregisteredStyleDataFormatString = null;
            IDataFormat sDataFormat = null;
            IFont unregisteredStyleFont = null;
            if (unregisteredStyleWorkbook != null && unregisteredStyleWorkbook != Workbook)
            {
                unregisteredStyleDataFormatString = unregisteredStyleWorkbook.CreateDataFormat().GetFormat(unregisteredStyle.DataFormat);
                sDataFormat = Workbook.CreateDataFormat();
                unregisteredStyleFont = unregisteredStyle.GetFont(unregisteredStyleWorkbook);
            }

            foreach (ICellStyle s in GetStyles())
            {
                if (unregisteredStyle.Alignment != s.Alignment
                    || unregisteredStyle.BorderBottom != s.BorderBottom
                    || unregisteredStyle.BorderDiagonal != s.BorderDiagonal
                    || unregisteredStyle.BorderDiagonalColor != s.BorderDiagonalColor
                    || unregisteredStyle.BorderDiagonalLineStyle != s.BorderDiagonalLineStyle
                    || unregisteredStyle.BorderLeft != s.BorderLeft
                    || unregisteredStyle.BorderRight != s.BorderRight
                    || unregisteredStyle.BorderTop != s.BorderTop
                    || unregisteredStyle.BottomBorderColor != s.BottomBorderColor
                    || unregisteredStyle.FillPattern != s.FillPattern
                    || unregisteredStyle.Indention != s.Indention
                    || unregisteredStyle.IsHidden != s.IsHidden
                    || unregisteredStyle.IsLocked != s.IsLocked
                    || unregisteredStyle.LeftBorderColor != s.LeftBorderColor
                    || unregisteredStyle.RightBorderColor != s.RightBorderColor
                    || unregisteredStyle.Rotation != s.Rotation
                    || unregisteredStyle.ShrinkToFit != s.ShrinkToFit
                    || unregisteredStyle.TopBorderColor != s.TopBorderColor
                    || unregisteredStyle.VerticalAlignment != s.VerticalAlignment
                    || unregisteredStyle.WrapText != s.WrapText
                    )
                    continue;

                if (unregisteredStyle is XSSFCellStyle xcs)
                {
                    if (!Serialization.Json.IsEqual(xcs.FillForegroundColorColor?.RGB, s.FillForegroundColorColor?.RGB)
                        || !Serialization.Json.IsEqual(xcs.FillBackgroundColorColor?.RGB, s.FillBackgroundColorColor?.RGB)
                        )
                        continue;
                }
                else if (unregisteredStyle is HSSFCellStyle hcs)
                {
                    if (hSSFForegroundColor.Indexed != s.FillForegroundColor
                         || hSSFBackgroundColor.Indexed != s.FillBackgroundColor
                         )
                        continue;
                }
                else
                    throw new Exception("Unsupported style type: " + unregisteredStyle.GetType().FullName);

                if (unregisteredStyleWorkbook == null)
                {
                    if (unregisteredStyle.DataFormat != s.DataFormat
                       || unregisteredStyle.FontIndex != s.FontIndex
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
                return s;
            }
        CREATE_STYLE:
            ICellStyle style = Workbook.CreateCellStyle();
            return CopyStyle(unregisteredStyle, style);
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
            }
            else if (fromStyle is HSSFCellStyle hcs)
            {
                if (!(toStyle is HSSFCellStyle))
                    throw new Exception("Copying style to a different type is not supported: " + toStyle.GetType().FullName);
                if ((fromStyle.FillForegroundColor > 0 || fromStyle.FillBackgroundColor > 0)
                    && toStyleWorkbook != null && toStyleWorkbook != Workbook
                    )
                {
                    HSSFColor hSSFForegroundColor = getRegisteredHSSFColor((HSSFWorkbook)toStyleWorkbook, new Color(fromStyle.FillForegroundColorColor.RGB[0], fromStyle.FillForegroundColorColor.RGB[1], fromStyle.FillForegroundColorColor.RGB[2]));
                    HSSFColor hSSFBackgroundColor = getRegisteredHSSFColor((HSSFWorkbook)toStyleWorkbook, new Color(fromStyle.FillBackgroundColorColor.RGB[0], fromStyle.FillBackgroundColorColor.RGB[1], fromStyle.FillBackgroundColorColor.RGB[2]));
                    toStyle.FillForegroundColor = hSSFForegroundColor.Indexed;//(!)might be not exactly same color
                    toStyle.FillBackgroundColor = hSSFBackgroundColor.Indexed;//(!)might be not exactly same color
                }
                else
                {
                    toStyle.FillForegroundColor = fromStyle.FillForegroundColor;
                    toStyle.FillBackgroundColor = fromStyle.FillBackgroundColor;
                }
            }
            else
                throw new Exception("Unsupported style type: " + fromStyle.GetType().FullName);
            toStyle.FillPattern = fromStyle.FillPattern;
            toStyle.Indention = fromStyle.Indention;
            toStyle.IsHidden = fromStyle.IsHidden;
            toStyle.IsLocked = fromStyle.IsLocked;
            toStyle.LeftBorderColor = fromStyle.LeftBorderColor;
            toStyle.RightBorderColor = fromStyle.RightBorderColor;
            toStyle.Rotation = fromStyle.Rotation;
            toStyle.ShrinkToFit = fromStyle.ShrinkToFit;
            toStyle.TopBorderColor = fromStyle.TopBorderColor;
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

        //public void OptimiseStyles()
        //{
        //    if (Workbook is XSSFWorkbook xSSFWorkbook)
        //    {
        //        NPOI.XSSF.Model.StylesTable st = xSSFWorkbook.GetStylesSource();
        //        st.GetTableStyle()..re().(0).;
        //        new NPOI.XSSF.Model.StylesTable(Workbook.pa;
        //    }
        //    else if (Workbook is HSSFWorkbook hSSFWorkbook)
        //    {
        //        HSSFOptimiser.OptimiseCellStyles(hSSFWorkbook);
        //    }
        //    else
        //        throw new Exception("Unsupported workbook type: " + Workbook.GetType().FullName);
        //}

        //public void OptimiseFonts()
        //{
        //    if (Workbook is XSSFWorkbook xSSFWorkbook)
        //    {
        //    }
        //    else if (Workbook is HSSFWorkbook hSSFWorkbook)
        //    {
        //        HSSFOptimiser.OptimiseFonts(hSSFWorkbook);
        //    }
        //    else
        //        throw new Exception("Unsupported workbook type: " + Workbook.GetType().FullName);
        //}

        public void ReplaceStyle(ICellStyle style1, ICellStyle style2)
        {
            ReplaceStyle(null, style1, style2);
        }

        public void SetStyle(ICellStyle style, bool createCells)
        {
            SetStyle(null, style, createCells);
        }

        public void ClearStyle(ICellStyle style)
        {
            ClearStyle(null, style);
        }
    }
}