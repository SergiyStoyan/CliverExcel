///********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************/
//using NPOI.HSSF.UserModel;
//using NPOI.HSSF.Util;
//using NPOI.SS.UserModel;
//using NPOI.XSSF.UserModel;
//using System;
//using System.Collections.Generic;
//using System.Linq;

//namespace Cliver
//{
//    public partial class Excel
//    {
//        public partial class Style
//        {
//            public Style(ICellStyle style, IWorkbook workbook)
//            {
//                _ = style;
//                Workbook = workbook;
//            }

//            public ICellStyle _ { get; private set; }
//            public IWorkbook Workbook { get; private set; }

//            static public Style CreateUnregisteredStyle(IWorkbook workbook)
//            {
//                if (workbook is XSSFWorkbook)
//                {
//                    XSSFWorkbook w = new XSSFWorkbook();
//                    ICellStyle s = new XSSFCellStyle(w.GetStylesSource());
//                    IFont f = workbook.NumberOfFonts > 0 ? workbook.GetFontAt(0) : w.CreateFont();
//                    s.SetFont(f);//otherwise it throws an exception on accessing font
//                    return new Style(s, workbook);
//                }
//                if (workbook is HSSFWorkbook)
//                {
//                    HSSFWorkbook w = new HSSFWorkbook();
//                    ICellStyle s = new HSSFCellStyle(0, new NPOI.HSSF.Record.ExtendedFormatRecord(), w);
//                    IFont f = workbook.NumberOfFonts > 0 ? workbook.GetFontAt(0) : w.CreateFont();
//                    s.SetFont(f);//set default font
//                    return new Style(s, workbook);
//                }
//                throw new Exception("Unsupported workbook type: " + workbook.GetType().FullName);
//            }

//            static public Style CreateRegisteredStyle(IWorkbook workbook)
//            {
//                return new Style(workbook.CreateCellStyle(), workbook);
//            }

//            public Style CloneUnregistered(IWorkbook cloneWorkbook = null)
//            {
//                Style toStyle = CreateUnregisteredStyle(Workbook);
//                _ = copyStyle(Workbook, _, toStyle._, cloneWorkbook);
//                return toStyle;
//            }

//            static ICellStyle copyStyle(IWorkbook fromWorkbook, ICellStyle fromStyle, ICellStyle toStyle, IWorkbook toStyleWorkbook = null)
//            {
//                if (toStyleWorkbook != null && toStyleWorkbook.GetType() != fromWorkbook.GetType())
//                    throw new Exception("Copying a style in a different type workbook is not supported: " + toStyleWorkbook.GetType().FullName);
//                toStyle.Alignment = fromStyle.Alignment;
//                toStyle.BorderBottom = fromStyle.BorderBottom;
//                toStyle.BorderDiagonal = fromStyle.BorderDiagonal;
//                toStyle.BorderDiagonalColor = fromStyle.BorderDiagonalColor;
//                toStyle.BorderDiagonalLineStyle = fromStyle.BorderDiagonalLineStyle;
//                toStyle.BorderLeft = fromStyle.BorderLeft;
//                toStyle.BorderRight = fromStyle.BorderRight;
//                toStyle.BorderTop = fromStyle.BorderTop;
//                toStyle.BottomBorderColor = fromStyle.BottomBorderColor;
//                if (toStyleWorkbook == null)
//                    toStyle.DataFormat = fromStyle.DataFormat;
//                else
//                {
//                    var dataFormat1 = fromWorkbook.CreateDataFormat();
//                    var dataFormat2 = toStyleWorkbook.CreateDataFormat();
//                    string sDataFormat;
//                    try
//                    {
//                        sDataFormat = dataFormat1.GetFormat(fromStyle.DataFormat);
//                    }
//                    catch (Exception e)
//                    {
//                        throw new Exception("Style fromStyle has DataFormat=" + fromStyle.DataFormat + " that does not exists in the workbook.", e);
//                    }
//                    toStyle.DataFormat = dataFormat2.GetFormat(sDataFormat);
//                }
//                if (fromStyle is XSSFCellStyle xcs)
//                {
//                    XSSFCellStyle toXcs = toStyle as XSSFCellStyle;
//                    if (toXcs == null)
//                        throw new Exception("Copying style to a different type is not supported: " + toStyle.GetType().FullName);
//                    toXcs.FillForegroundColorColor = fromStyle.FillForegroundColorColor;
//                    toXcs.FillBackgroundColorColor = fromStyle.FillBackgroundColorColor;
//                    toXcs.SetDiagonalBorderColor(xcs.DiagonalBorderXSSFColor);
//                    toXcs.SetBottomBorderColor(xcs.BottomBorderXSSFColor);
//                    toXcs.SetLeftBorderColor(xcs.LeftBorderXSSFColor);
//                    toXcs.SetRightBorderColor(xcs.RightBorderXSSFColor);
//                    toXcs.SetTopBorderColor(xcs.TopBorderXSSFColor);
//                }
//                else if (fromStyle is HSSFCellStyle)
//                {
//                    if (!(toStyle is HSSFCellStyle))
//                        throw new Exception("Copying style to a different type is not supported: " + toStyle.GetType().FullName);
//                    if (toStyleWorkbook != null && toStyleWorkbook != fromWorkbook)
//                    {
//                        if (fromStyle.FillForegroundColor > 0)
//                        {
//                            HSSFColor c = getRegisteredHSSFColor((HSSFWorkbook)toStyleWorkbook, new Excel.Color(fromStyle.FillForegroundColorColor));
//                            toStyle.FillForegroundColor = c.Indexed;//(!)might be not exactly same color
//                        }
//                        if (fromStyle.FillBackgroundColor > 0)
//                        {
//                            HSSFColor c = getRegisteredHSSFColor((HSSFWorkbook)toStyleWorkbook, new Excel.Color(fromStyle.FillBackgroundColorColor));
//                            toStyle.FillBackgroundColor = c.Indexed;//(!)might be not exactly same color
//                        }
//                        HSSFPalette palette = ((HSSFWorkbook)fromWorkbook).GetCustomPalette();
//                        if (fromStyle.BorderDiagonalColor > 0)
//                        {
//                            HSSFColor c = getRegisteredHSSFColor((HSSFWorkbook)toStyleWorkbook, new Excel.Color(palette.GetColor(fromStyle.BorderDiagonalColor)));
//                            toStyle.BorderDiagonalColor = c.Indexed;//(!)might be not exactly same color
//                        }
//                        if (fromStyle.BottomBorderColor > 0)
//                        {
//                            HSSFColor c = getRegisteredHSSFColor((HSSFWorkbook)toStyleWorkbook, new Excel.Color(palette.GetColor(fromStyle.BottomBorderColor)));
//                            toStyle.BottomBorderColor = c.Indexed;//(!)might be not exactly same color
//                        }
//                        if (fromStyle.LeftBorderColor > 0)
//                        {
//                            HSSFColor c = getRegisteredHSSFColor((HSSFWorkbook)toStyleWorkbook, new Excel.Color(palette.GetColor(fromStyle.LeftBorderColor)));
//                            toStyle.LeftBorderColor = c.Indexed;//(!)might be not exactly same color
//                        }
//                        if (fromStyle.RightBorderColor > 0)
//                        {
//                            HSSFColor c = getRegisteredHSSFColor((HSSFWorkbook)toStyleWorkbook, new Excel.Color(palette.GetColor(fromStyle.RightBorderColor)));
//                            toStyle.RightBorderColor = c.Indexed;//(!)might be not exactly same color
//                        }
//                        if (fromStyle.TopBorderColor > 0)
//                        {
//                            HSSFColor c = getRegisteredHSSFColor((HSSFWorkbook)toStyleWorkbook, new Excel.Color(palette.GetColor(fromStyle.TopBorderColor)));
//                            toStyle.TopBorderColor = c.Indexed;//(!)might be not exactly same color
//                        }
//                    }
//                    else
//                    {
//                        toStyle.FillForegroundColor = fromStyle.FillForegroundColor;
//                        toStyle.FillBackgroundColor = fromStyle.FillBackgroundColor;
//                        toStyle.BorderDiagonalColor = fromStyle.BorderDiagonalColor;
//                        toStyle.BottomBorderColor = fromStyle.BottomBorderColor;
//                        toStyle.LeftBorderColor = fromStyle.LeftBorderColor;
//                        toStyle.RightBorderColor = fromStyle.RightBorderColor;
//                        toStyle.TopBorderColor = fromStyle.TopBorderColor;
//                    }
//                }
//                else
//                    throw new Exception("Unsupported style type: " + fromStyle.GetType().FullName);
//                toStyle.FillPattern = fromStyle.FillPattern;
//                toStyle.Indention = fromStyle.Indention;
//                toStyle.IsHidden = fromStyle.IsHidden;
//                toStyle.IsLocked = fromStyle.IsLocked;
//                toStyle.Rotation = fromStyle.Rotation;
//                toStyle.ShrinkToFit = fromStyle.ShrinkToFit;
//                toStyle.VerticalAlignment = fromStyle.VerticalAlignment;
//                toStyle.WrapText = fromStyle.WrapText;
//                IFont f1;
//                try
//                {
//                    f1 = fromWorkbook.GetFontAt(fromStyle.FontIndex);
//                }
//                catch (Exception e)
//                {
//                    throw new Exception("Style fromStyle has font[@index=" + fromStyle.FontIndex + "] that does not exists in the workbook.", e);
//                }
//                if (toStyleWorkbook == null)
//                    toStyle.SetFont(f1);
//                else
//                {
//                    IFont f2 = fromWorkbook._GetRegisteredFont(f1.IsBold, (IndexedColors)Enum.ToObject(typeof(IndexedColors), f1.Color), (short)f1.FontHeight, f1.FontName, f1.IsItalic, f1.IsStrikeout, f1.TypeOffset, f1.Underline);
//                    toStyle.SetFont(f2);
//                }
//                return toStyle;
//            }

//            static HSSFColor getRegisteredHSSFColor(HSSFWorkbook workbook, Excel.Color color)
//            {
//                HSSFPalette palette = workbook.GetCustomPalette();
//                HSSFColor hssfColor = palette.FindColor(color.R, color.G, color.B);
//                if (hssfColor != null)
//                    return hssfColor;
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
//                return hssfColor;
//            }

//            public Style CloneRegistered(IWorkbook workbook = null)
//            {
//                Style toStyle = CreateRegisteredStyle(Workbook);
//                _ = copyStyle(Workbook, _, toStyle._, workbook);
//                return toStyle;
//            }

//            public bool Register()
//            {
//                Style style = FindEquivalents().FirstOrDefault();
//                if (style != null)
//                    return false;
//                style = CreateRegisteredStyle(Workbook);
//                _ = copyStyle(Workbook, _, style._, Workbook);
//                return true;
//            }

//            public IEnumerable<Style> FindEquivalents()
//            {
//            }

//            public void Highlight(Color color, FillPattern fillPattern = FillPattern.SolidForeground)
//            {
//            }

//            public IFont GetRegisteredFont(bool bold, IndexedColors color, short fontHeightInPoints, string name, bool italic = false, bool strikeout = false, FontSuperScript typeOffset = FontSuperScript.None, FontUnderlineType underline = FontUnderlineType.None)
//            {
//            }

//            public IEnumerable<Style> GetStyles()
//            {
//                for (int i = 0; i < Workbook.NumCellStyles; i++)
//                    yield return new Style(Workbook.GetCellStyleAt(i), Workbook);
//            }

//            public IEnumerable<Style> GetUnused(params short[] ignoredStyleIds)
//            {
//            }

//            public void Optimise()
//            {
//            }
//        }
//    }
//}