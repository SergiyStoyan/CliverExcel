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

        public ICellStyle Highlight(ICellStyle style, Color color)
        {
            return highlight(Workbook, style, color);
        }

        /// <summary>
        /// Removes highlighting from all the styles.
        /// </summary>
        public void ClearStyle(ICellStyle style)
        {
            SetStyle(new Range(this), style, false);
        }

        /// <summary>
        /// Is intended for either adding or removing backgound color.
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="style"></param>
        /// <param name="color"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        static internal ICellStyle highlight(IWorkbook workbook, ICellStyle style, Color color)
        {
            if (workbook is XSSFWorkbook)
            {
                XSSFCellStyle cs = style == null ? (XSSFCellStyle)workbook.CreateCellStyle() : (XSSFCellStyle)style;
                if (color == null)
                {
                    cs.SetFillForegroundColor(null);
                    cs.FillPattern = FillPattern.NoFill;
                }
                else
                {
                    cs.SetFillForegroundColor(new XSSFColor(color.RGB));
                    cs.FillPattern = FillPattern.SolidForeground;
                }
                return cs;
            }
            if (workbook is HSSFWorkbook)
            {
                HSSFCellStyle cs = style == null ? (HSSFCellStyle)workbook.CreateCellStyle() : (HSSFCellStyle)style;
                if (color == null)
                {
                    cs.FillForegroundColor = 0;
                    cs.FillPattern = FillPattern.NoFill;
                }
                else
                {
                    HSSFPalette palette = ((HSSFWorkbook)workbook).GetCustomPalette();
                    NPOI.HSSF.Util.HSSFColor hssfColor = palette.FindColor(color.R, color.G, color.B);
                    if (hssfColor == null)
                    {
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
                    }
                    cs.FillForegroundColor = hssfColor.Indexed;
                    cs.FillPattern = FillPattern.SolidForeground;
                }
                return cs;
            }
            throw new Exception("Unsupported workbook type: " + workbook.GetType().FullName);
        }

        /// <summary>
        /// Looks for an equal style in the workbook and, if it does not exists, creates a new one.
        /// (!)Incidentally, there is a somewhat analogous method NPOI.SS.Util.CellUtil.SetCellStyleProperties() which is not as handy in use though.
        /// </summary>
        /// <param name="style">it must be a style created by CreateUnregisteredStyle() and then modified as needed</param>
        /// <param name="unregisteredStyleWorkbook"></param>
        /// <returns></returns>
        public ICellStyle GetRegisteredStyle(ICellStyle unregisteredStyle, IWorkbook unregisteredStyleWorkbook = null)
        {
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
                    //|| unregisteredStyle.FillBackgroundColor != s.FillBackgroundColor //(NPOI bug?) FillBackgroundColor cannot be set and remains 64. It is not used, so can be ignored. 
                    || unregisteredStyle.FillForegroundColor != s.FillForegroundColor
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

                if (unregisteredStyleWorkbook == null)
                {
                    if (unregisteredStyle.DataFormat != s.DataFormat
                       || unregisteredStyle.FontIndex != s.FontIndex
                       )
                        continue;
                }
                else
                {
                    var unregisteredStyleDataFormat = unregisteredStyleWorkbook.CreateDataFormat();
                    var sDataFormat = Workbook.CreateDataFormat();
                    if (unregisteredStyleDataFormat.GetFormat(unregisteredStyle.DataFormat) != sDataFormat.GetFormat(s.DataFormat))
                        continue;

                    IFont unregisteredStyleFont = unregisteredStyle.GetFont(unregisteredStyleWorkbook);
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
            toStyle.FillBackgroundColor = fromStyle.FillBackgroundColor;
            //if (toStyle.FillBackgroundColor != fromStyle.FillBackgroundColor)//it happens when FillBackgroundColor = 0 (bug?)
            //    throw new Exception("FillBackgroundColor could not be copied: " + fromStyle.FillBackgroundColor + " -> " + toStyle.FillBackgroundColor);
            toStyle.FillForegroundColor = fromStyle.FillForegroundColor;
            if (toStyle.FillForegroundColor != fromStyle.FillForegroundColor)
                throw new Exception("FillForegroundColor could not be copied: " + fromStyle.FillForegroundColor + " -> " + toStyle.FillForegroundColor);
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

        public void OptimiseStyles()
        {
            if (Workbook is XSSFWorkbook xSSFWorkbook)
            {
                //xSSFWorkbook.GetStylesSource().RemoveNumberFormat().
                //NPOI.XSSF.UserModel.XSSFBuiltinTableStyle h;
                //h...sty.OptimiseCellStyles();
            }
            else if (Workbook is HSSFWorkbook hSSFWorkbook)
            {
                HSSFOptimiser.OptimiseCellStyles(hSSFWorkbook);
            }
            else
                throw new Exception("Unsupported workbook type: " + Workbook.GetType().FullName);
        }
    }
}