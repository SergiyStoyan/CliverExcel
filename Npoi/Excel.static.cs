//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.Util;
using System;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;

namespace Cliver
{
    public partial class Excel
    {
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

        static public void PasteRange(ICell[][] rangeCells, int toY, int toX, OnFormulaCellMoved onFormulaCellMoved = null, ISheet toSheet = null, StyleMap toStyleMap = null)
        {
            for (int yi = rangeCells.Length - 1; yi >= 0; yi--)
            {
                ICell[] rowCells = rangeCells[yi];
                for (int xi = rowCells.Length - 1; xi >= 0; xi--)
                {
                    var c = rowCells[xi];
                    if (c != null)
                        c._Copy(toY + yi, toX + xi, onFormulaCellMoved, toSheet, toStyleMap);
                    else
                        toSheet._RemoveCell(toY + yi, toX + xi);
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
        public static HSSFColor GetRegisteredHSSFColor(HSSFWorkbook workbook, Excel.Color color)
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
    }
}