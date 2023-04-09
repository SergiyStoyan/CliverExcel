//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Text.RegularExpressions;

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
            if (c1 == null)
                return c2 == null;
            if (c2 == null)
                return false;
            return c1.RGB[0] == c2.RGB[0] && c1.RGB[1] == c2.RGB[1] && c1.RGB[2] == c2.RGB[2];
        }

        static public bool AreColorsEqual(Color c1, IColor c2)
        {
            if (c1 == null)
                return c2 == null;
            if (c2 == null)
                return false;
            return c1.RGB[0] == c2.RGB[0] && c1.RGB[1] == c2.RGB[1] && c1.RGB[2] == c2.RGB[2];
        }

        static public bool AreColorsEqual(Color c1, Color c2)
        {
            if (c1 == null)
                return c2 == null;
            if (c2 == null)
                return false;
            return c1.RGB[0] == c2.RGB[0] && c1.RGB[1] == c2.RGB[1] && c1.RGB[2] == c2.RGB[2];
        }
<<<<<<< Updated upstream
=======

        static public void PasteRange(Cell[][] rangeCells, int toY, int toX, Sheet toSheet = null)
        {
            for (int yi = rangeCells.Length - 1; yi >= 0; yi--)
            {
                Cell[] rowCells = rangeCells[yi];
                for (int xi = rowCells.Length - 1; xi >= 0; xi--)
                {
                    var c = rowCells[xi];
                    if (c != null)
                        c.Copy(toY + yi, toX + xi, toSheet);
                    else
                        toSheet.RemoveCell(toY + yi, toX + xi);
                }
            }
        }
>>>>>>> Stashed changes
    }
}