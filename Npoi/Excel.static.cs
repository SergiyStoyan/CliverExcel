﻿//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.Util;
using System;
using System.Linq;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;

namespace Cliver
{
    public partial class Excel
    {
        static internal Excel Get(IWorkbook workbook)
        {
            return workbooks2Excel.GetValue(workbook, (IWorkbook w) =>
            {
                return new Excel(null);
            });
        }

        /// <summary>
        /// Set it to make Excel keep links absolute
        /// </summary>
        public const string AbsoluteLinksHyperlinkBase = "x";

        static public string GetSafeSheetName(string name)
        {
            name = Regex.Replace(name, @"\:", "-");//npoi does not accept :
            return WorkbookUtil.CreateSafeSheetName(name);
        }

        static public string GetColumnName(int x)
        {
            return CellReference.ConvertNumToColString(x - 1);
        }

        static public int GetX(string columnName)
        {
            return CellReference.ConvertColStringToIndex(columnName) + 1;
        }

        static public (int Y, int X) GetCoordinates(string address)
        {
            var a = ParseAddress(address);
            return (a.Y, GetX(a.ColumnName));
        }

        static public (int Y, string ColumnName) ParseAddress(string address)
        {
            Match m = Regex.Match(address, @"^\s*([a-z]+)(\d+)\s*$", RegexOptions.IgnoreCase);
            if (!m.Success)
                throw new Exception("Address is not parsable: " + address);
            return (int.Parse(m.Groups[2].Value), m.Groups[1].Value);
        }

        static public bool AreColorsEqual(IColor c1, IColor c2)
        {
            if (c1?.RGB == null)
                return c2?.RGB == null;
            if (c2?.RGB == null)
                return false;
            return c1.RGB[0] == c2.RGB[0] && c1.RGB[1] == c2.RGB[1] && c1.RGB[2] == c2.RGB[2];
        }

        static public bool AreColorsEqual(Color c1, IColor c2)
        {
            if (c1 == null)
                return c2?.RGB == null;
            if (c2?.RGB == null)
                return false;
            return c1.R == c2.RGB[0] && c1.G == c2.RGB[1] && c1.B == c2.RGB[2];
        }

        static public bool AreColorsEqual(Color c1, Color c2)
        {
            if (c1 == null)
                return c2 == null;
            if (c2 == null)
                return false;
            return c1.R == c2.R && c1.G == c2.G && c1.B == c2.B;
        }

        static public void PasteRange(ICell[][] rangeCells, int y2, int x2, CopyCellMode copyCellMode, ISheet sheet2 = null, StyleMap styleMap2 = null)
        {
            if (sheet2 == null)
                throw new Exception("sheet2 must not be NULL.");
            for (int yi = rangeCells.Length - 1; yi >= 0; yi--)
            {
                ICell[] rowCells = rangeCells[yi];
                for (int xi = rowCells.Length - 1; xi >= 0; xi--)
                {
                    var c = rowCells[xi];
                    if (c != null)
                        c._Copy(y2 + yi, x2 + xi, copyCellMode, sheet2, styleMap2);
                    else
                        sheet2._RemoveCell(y2 + yi, x2 + xi, copyCellMode?.CopyComment == true);
                }
            }
        }

        static public bool AreFontsEqual(IFont font1, IFont font2)
        {
            return font1.Charset == font2.Charset
                && font1.Color == font2.Color
                && font1.FontHeight == font2.FontHeight
                && font1.FontName == font2.FontName
                && font1.IsBold == font2.IsBold
                && font1.IsItalic == font2.IsItalic
                && font1.IsStrikeout == font2.IsStrikeout
                && font1.TypeOffset == font2.TypeOffset
                && font1.Underline == font2.Underline;
        }

        static public IEnumerable<IFont> FindEqualFonts(IFont font, IEnumerable<IFont> searchFonts)
        {
            foreach (IFont font2 in searchFonts)
                if (AreFontsEqual(font, font2))
                    yield return font2;
        }

        /// <summary>
        /// Find or resgister the color. (!)It may be not exact match.
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="color"></param>
        /// <returns></returns>
        public static HSSFColor GetRegisteredHSSFColor(HSSFWorkbook workbook, Excel.Color color, bool allowSimilarColor = true)
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
                bool isIndexedColorUsed(short c/*, List<ICellStyle> unusedStyles, List<IFont> unusedFonts*/)
                {
                    for (int i = workbook.NumCellStyles - 1; i >= 0; i--)
                    {
                        ICellStyle s = workbook.GetCellStyleAt(i);
                        if (s.BorderDiagonalColor == c
                            || s.BottomBorderColor == c
                            || s.FillBackgroundColor == c
                            || s.FillForegroundColor == c
                            || s.LeftBorderColor == c
                            || s.RightBorderColor == c
                            || s.TopBorderColor == c
                            )
                            //if (!unusedStyles.Contains(s))
                            return true;
                    }
                    for (short i = (short)(workbook.NumberOfFonts - 1); i >= 0; i--)
                    {
                        IFont f = workbook.GetFontAt(i);
                        if (f.Color == c)
                            //if (!unusedFonts.Contains(f))
                            return true;
                    }
                    return false;
                }
                short? findUnusedColorIndex()
                {
                    //workbook._OptimizeStylesAndFonts(out List<ICellStyle> unusedStyles, out List<IFont> unusedFonts);
                    for (short j = 0x8; j <= 0x40; j++)//the first color in the palette has the index 0x8, the second has the index 0x9, etc. through 0x40
                    {
                        if (!isIndexedColorUsed(j))
                            return j;
                    }
                    return null;
                }
                short? ci = findUnusedColorIndex();
                if (ci == null)
                {
                    if (!allowSimilarColor)
                        throw new Exception("The palette of indexed colors is full and all the colors are in use. Consider optimizing styles and fonts.");
                    ci = palette.FindSimilarColor(color.R, color.G, color.B).Indexed;
                }
                palette.SetColorAtIndex(ci.Value, color.R, color.G, color.B);
                hssfColor = palette.GetColor(ci.Value);
            }
            return hssfColor;
        }
    }
}