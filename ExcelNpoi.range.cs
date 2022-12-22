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

namespace Cliver
{
    public partial class Excel
    {
        /// <summary>
        /// (!) 1-based
        /// </summary>
        public class Range
        {
            public int X;
            public int LastX = 0;
            public int Y;
            public int LastY = 0;

            public Range(int y, int lastY, int x, int lastX)
            {
                Y = y;
                LastY = lastY;
                X = x;
                LastX = lastX;
            }

            public ICell GetMainCell(Excel excel, bool create)
            {
                return excel.GetCell(Y, X, create);
            }

            public string GetStringAddress()
            {
                return CellReference.ConvertNumToColString(X - 1) + Y + ":" + CellReference.ConvertNumToColString(LastX - 1) + LastY;
            }

            /// <summary>
            /// 
            /// </summary>
            /// <returns>(!) 0-based</returns>
            public CellRangeAddress GetCellRangeAddress()
            {
                return new CellRangeAddress(Y - 1, LastY - 1, X - 1, LastX - 1);
            }
        }

        public void Highlight(Range range, Color color)
        {
            for (int y = range.Y; y <= range.LastY; y++)
            {
                IRow row = GetRow(y, color != null);
                if (row == null)
                    continue;
                for (int x = range.X; x <= row.LastCellNum && x <= range.LastX; x++)
                {
                    ICell c = row.GetCell(x, true);
                    c.CellStyle = highlight(Workbook, c.CellStyle, color);
                }
            }
        }

        public void SetStyle(Range range, ICellStyle style)
        {
            for (int y = range.Y; y <= range.LastY; y++)
            {
                IRow row = GetRow(y, true);
                for (int x = range.X; x <= row.LastCellNum && x <= range.LastX; x++)
                {
                    ICell c = row.GetCell(x, true);
                    c.CellStyle = style;
                }
            }
        }

        public void ClearMerging(Range range)
        {
            CellRangeAddress cra = range.GetCellRangeAddress();
            for (int i = Sheet.MergedRegions.Count - 1; i >= 0; i--)
                if (Sheet.MergedRegions[i].Intersects(cra))
                    Sheet.RemoveMergedRegion(i);
        }

        public void Merge(Range range, bool clearOldMerging = false)
        {
            if (clearOldMerging)
                ClearMerging(range);
            Sheet.AddMergedRegion(range.GetCellRangeAddress());
        }

        public Range GetMergedRange(int y, int x)
        {
            return getMergedRange(Sheet, y, x);
        }

        static internal Range getMergedRange(ISheet sheet, int y, int x)
        {
            foreach (var mr in sheet.MergedRegions)
                if (mr.IsInRange(y - 1, x - 1))
                    return new Range(mr.FirstRow + 1, mr.LastRow + 1, mr.FirstColumn + 1, mr.LastColumn + 1);
            return null;
        }

        /// <summary>
        /// !!!test
        /// </summary>
        /// <param name="rangeCells"></param>
        /// <param name="y"></param>
        /// <param name="x"></param>
        /// <returns></returns>
        public void CopyRange(CellRangeAddress range, ISheet sourceSheet, ISheet destinationSheet)
        {
            for (int y = range.FirstRow; y <= range.LastRow; y++)
            {
                IRow sourceRow = sourceSheet.GetRow(y);
                if (sourceRow == null)
                    continue;
                IRow destinationRow = destinationSheet.GetRow(y);
                if (destinationRow == null)
                    destinationRow = destinationSheet.CreateRow(y);
                for (int x = range.FirstColumn; x < sourceRow.LastCellNum && x <= range.LastColumn; x++)
                {
                    ICell sourceCell = sourceRow.GetCell(x);
                    ICell destinationCell = destinationRow.GetCell(x);
                    if (sourceCell == null)
                    {
                        if (destinationCell == null)
                            continue;
                        destinationRow.RemoveCell(destinationCell);
                    }
                    else
                    {
                        destinationCell = destinationRow.CreateCell(x);
                        CopyCell(sourceCell, destinationCell);
                    }
                }
            }
        }

        public ICell[,] CutRange(Range range)
        {
            if (range.LastY <= 0)
                range.LastY = Sheet.LastRowNum + 1;
            if (range.LastX <= 0)
                for (int y0 = range.Y - 1; y0 < range.LastY; y0++)
                {
                    IRow row = Sheet.GetRow(y0);
                    if (range.LastX < row?.LastCellNum)
                        range.LastX = row.LastCellNum;
                }

            ICell[,] rangeCells = new ICell[range.LastY - range.Y + 1, range.LastX - range.X + 1];
            for (int y = range.Y; y <= range.LastY; y++)
            {
                IRow row = Sheet.GetRow(y - 1);
                if (row == null)
                    continue;
                for (int x = range.X; x <= row.LastCellNum && x <= range.LastX; x++)
                {
                    ICell cell = row.GetCell(x - 1);
                    rangeCells[y - range.Y, x - range.X] = cell;
                    row.RemoveCell(cell);
                }
            }
            return rangeCells;
        }

        public void PasteRange(ICell[,] rangeCells, int y, int x)
        {
            int height = rangeCells.GetLength(0);
            int width = rangeCells.GetLength(1);
            for (int yi = 0; yi < height; yi++)
                for (int xi = 0; xi < width; xi++)
                    CopyCell(rangeCells[yi, xi], y + yi, x + xi);
        }

        public void MoveRange(Range sourceRange, int y, int x)
        {
            PasteRange(CutRange(sourceRange), y, x);
        }
    }
}