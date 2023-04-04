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
            public int X1 = 1;
            public int X2 = 0;
            public int Y1 = 1;
            public int Y2 = 0;

            public Range(int y1, int y2, int x1, int x2)
            {
                Y1 = y1;
                Y2 = y2;
                X1 = x1;
                X2 = x2;
            }

            /// <summary>
            /// The entire sheet.
            /// </summary>
            /// <param name="xls"></param>
            public Range(Excel xls)
            {
                Y1 = 1;
                Y2 = xls.GetLastNotEmptyColumnInRowRange(1, null, true);
                X1 = 1;
                X2 = xls.GetLastNotEmptyRow(true);
            }

            public ICell GetMainCell(Excel excel, bool create)
            {
                return excel.GetCell(Y1, X1, create);
            }

            public string GetStringAddress()
            {
                return CellReference.ConvertNumToColString(X1 - 1) + Y1 + ":" + CellReference.ConvertNumToColString(X2 - 1) + Y2;
            }

            /// <summary>
            /// 
            /// </summary>
            /// <returns>(!) 0-based</returns>
            public CellRangeAddress GetCellRangeAddress()
            {
                return new CellRangeAddress(Y1 - 1, Y2 - 1, X1 - 1, X2 - 1);
            }
        }

        /// <summary>
        /// (!)Each call registers a new style for non-styled cells. If many calls, consider rather highlighting styles.
        /// </summary>
        /// <param name="range"></param>
        /// <param name="color"></param>
        public void Highlight(Range range, Color color)
        {
            ICellStyle newStyle = null;
            for (int y = range.Y1; y <= range.Y2; y++)
            {
                IRow row = GetRow(y, color != null);
                if (row == null)
                    continue;
                int maxX = Math.Min(row.LastCellNum, range.X2);
                for (int x = range.X1; x <= maxX; x++)
                {
                    ICell c = row.GetCell(x, true);
                    if (c.CellStyle == null)
                    {
                        if (newStyle == null)
                            newStyle = highlight(Workbook, null, color);
                        c.CellStyle = newStyle;
                    }
                    c.CellStyle = highlight(Workbook, c.CellStyle, color);
                }
            }
        }

        public void SetStyle(Range range, ICellStyle style, bool createCells)
        {
            for (int y = range.Y1; y <= range.Y2; y++)
            {
                IRow row = GetRow(y, createCells);
                if (row == null)
                    continue;
                int maxX = Math.Min(row.LastCellNum, range.X2);
                for (int x = range.X1; x <= maxX; x++)
                {
                    ICell c = row.GetCell(x, createCells);
                    if (c != null)
                        c.CellStyle = style;
                }
            }
        }

        public void ReplaceStyle(Range range, ICellStyle style1, ICellStyle style2)
        {
            for (int y = range.Y1; y <= range.Y2; y++)
            {
                IRow row = GetRow(y, false);
                if (row == null)
                    continue;
                int maxX = Math.Min(row.LastCellNum, range.X2);
                for (int x = range.X1; x <= maxX; x++)
                {
                    ICell c = row.GetCell(x, false);
                    if (c != null && c.CellStyle.Index == style1.Index)
                        c.CellStyle = style2;
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
        public void CopyRange(Range range, ISheet destinationSheet)
        {
            int maxY = Math.Min(range.Y2, Sheet.LastRowNum + 1);
            for (int y = range.Y1; y <= maxY; y++)
            {
                IRow sourceRow = Sheet.GetRow(y);
                if (sourceRow == null)
                    continue;
                IRow destinationRow = destinationSheet.GetRow(y);
                if (destinationRow == null)
                    destinationRow = destinationSheet.CreateRow(y);
                int maxX = Math.Min(range.X2, sourceRow.LastCellNum);
                for (int x = range.X1; x < maxX; x++)
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
            ICell[,] rangeCells = new ICell[range.Y2 - range.Y1 + 1, range.X2 - range.X1 + 1];
            int maxY = Math.Min(range.Y2, Sheet.LastRowNum + 1);
            for (int y = range.Y1; y <= maxY; y++)
            {
                IRow row = Sheet.GetRow(y - 1);
                if (row == null)
                    continue;
                int maxX = Math.Min(range.X2, row.LastCellNum);
                for (int x = range.X1; x <= maxX; x++)
                {
                    ICell cell = row.GetCell(x - 1);
                    rangeCells[y - range.Y1, x - range.X1] = cell;
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