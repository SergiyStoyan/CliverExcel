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
        public void ShiftCellsDown(int cellsY, int firstCellX, int lastCellX, int rowCount, Action<ICell> updateFormula = null)
        {
            for (int x = firstCellX; x <= lastCellX; x++)
            {
                for (int y = GetLastUsedRowInColumn(x); y >= cellsY; y--)
                {
                    CopyCell(y, x, y + rowCount, x);
                    if (updateFormula == null)
                        continue;
                    ICell formulaCell = GetCell(y + rowCount, x, false);
                    if (formulaCell?.CellType != CellType.Formula)
                        continue;
                    updateFormula(formulaCell);
                }
                GetCell(cellsY, x, false)?.SetBlank();
            }
        }

        public void ShiftColumns(int x, int shift, Action<ICell> onFormulaCellMoved = null)
        {
            Dictionary<int, int> columnXs2width = new Dictionary<int, int>();
            int lastColumnX = x;
            var rows = Sheet.GetRowEnumerator();
            while (rows.MoveNext())
            {
                IRow row = (IRow)rows.Current;
                int columnX = row.GetLastUsedColumnInRow(true);
                if (lastColumnX < columnX)
                {
                    for (int i = columnX; i > lastColumnX; i--)
                        columnXs2width[i] = Sheet.GetColumnWidth(i);
                    lastColumnX = columnX;
                }
                for (int i = columnX; i >= x; i--)
                    MoveCell(row.RowNum + 1, i, row.RowNum + 1, i + shift, onFormulaCellMoved);
            }
            foreach (int columnX in columnXs2width.Keys.OrderByDescending(a => a))
                Sheet.SetColumnWidth(columnX + shift - 1, columnXs2width[columnX]);
        }

        public void ShiftColumns(IRow row, int x, int shift, Action<ICell> onFormulaCellMoved = null)
        {
            for (int i = row.GetLastUsedColumnInRow(true); i >= x; i--)
                MoveCell(row.RowNum + 1, i, row.RowNum + 1, i + shift, onFormulaCellMoved);
        }

        public void AutosizeRowsInRange(int y1 = 1, int? y2 = null)
        {
            var rows = Sheet.GetRowEnumerator();
            while (rows.MoveNext())
            {
                IRow row = (IRow)rows.Current;
                if (row.RowNum + 1 < y1)
                    continue;
                if (row.RowNum >= y2)
                    return;
                row.Height = -1;
            }
        }

        public void AutosizeRows()
        {
            AutosizeRowsInRange();
        }

        public void AutosizeColumns(IEnumerable<int> columnIs, int padding = 0)
        {
            foreach (int i in columnIs)
            {
                Sheet.AutoSizeColumn(i - 1);
                if (padding > 0)
                    Sheet.SetColumnWidth(i - 1, Sheet.GetColumnWidth(i - 1) + padding);
            }
        }

        public void AutosizeColumnsInRange(int x1 = 1, int? x2 = null, int padding = 0)
        {
            if (x2 == null)
                x2 = GetLastUsedColumnInRowRange(x1, null, true);
            for (int x0 = x1 - 1; x0 <= x2; x0++)
            {
                Sheet.AutoSizeColumn(x0);
                if (padding > 0)
                    Sheet.SetColumnWidth(x0, Sheet.GetColumnWidth(x0) + padding);
            }
        }

        public void AutosizeColumns(int padding = 0)
        {
            AutosizeColumnsInRange(1, null, padding);
        }

        public ICell[,] CutRange(Range range)
        {
            if (range.LastY <= 0)
                range.LastY = Sheet.LastRowNum + 1;
            if (range.LastX <= 0)
                for (int y = range.Y - 1; y < range.LastY; y++)
                {
                    IRow row = Sheet.GetRow(y);
                    if (range.LastX < row?.LastCellNum)
                        range.LastX = row.LastCellNum;
                }

            ICell[,] rangeCells = new ICell[range.LastY - range.Y + 1, range.LastX - range.X + 1];
            for (int y = range.Y - 1; y < range.LastY; y++)
            {
                IRow row = Sheet.GetRow(y);
                if (row == null || row.PhysicalNumberOfCells < 1)
                    continue;
                for (int x = range.X - 1; x <= row.LastCellNum && x < range.LastX; x++)
                {
                    ICell cell = row.GetCell(x);
                    rangeCells[y - range.Y + 1, x - range.X + 1] = cell;
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

            public string GetStringAddress()
            {
                return CellReference.ConvertNumToColString(X - 1) + Y + ":" + CellReference.ConvertNumToColString(LastX - 1) + LastY;
            }

            public CellRangeAddress GetCellRangeAddress()
            {
                return new CellRangeAddress(Y - 1, LastY - 1, X - 1, LastX - 1);
            }
        }

        public void Highlight(Range range, Color color)
        {
            CellRangeAddress cra = range.GetCellRangeAddress();
            for (int y0 = cra.FirstRow; y0 <= cra.LastRow; y0++)
            {
                IRow row = GetRow(y0 + 1, true);
                for (int x0 = cra.FirstColumn; x0 < row.LastCellNum && x0 <= cra.LastColumn; x0++)
                {
                    ICell c = row.GetCell(x0 + 1, true);
                    c.CellStyle = highlight(Workbook, c.CellStyle, color);
                }
            }
        }

        public void SetStyle(Range range, ICellStyle style)
        {
            CellRangeAddress cra = range.GetCellRangeAddress();
            for (int y0 = cra.FirstRow; y0 <= cra.LastRow; y0++)
            {
                IRow row = GetRow(y0 + 1, true);
                for (int x0 = cra.FirstColumn; x0 < row.LastCellNum && x0 <= cra.LastColumn; x0++)
                {
                    ICell c = row.GetCell(x0 + 1, true);
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
    }
}