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

//works  
namespace Cliver
{
    public partial class Excel
    {
        public void ShiftCellsDown(int cellsY, int firstCellX, int lastCellX, int rowCount, Action<ICell> updateFormula = null)
        {
            for (int x = firstCellX; x <= lastCellX; x++)
            {
                for (int y = GetLastUsedRowInColumns(x); y >= cellsY; y--)
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
                int columnX = row.GetLastUsedColumnInRow();
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
            for (int i = row.GetLastUsedColumnInRow(); i >= x; i--)
                MoveCell(row.RowNum + 1, i, row.RowNum + 1, i + shift, onFormulaCellMoved);
        }

        public void AutosizeRows(int y1 = 1, int? y2 = null)
        {
            var rows = Sheet.GetRowEnumerator();
            while (rows.MoveNext())
            {
                IRow row = (IRow)rows.Current;
                if (row.RowNum + 1 < y1)
                    continue;
                if (row.RowNum + 1 >= y2)
                    return;
                row.Height = -1;
            }
            //if (y2 == null)
            //    y2 = GetLastUsedRow();
            //for (int y = y1; y <= y2; y++)
            //{
            //    var r = Sheet.GetRow(y - 1);
            //    if (r != null)
            //        r.Height = -1;
            //}
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
        }

        public void Highlight(Range range, Color color)
        {
            Highlight(new CellRangeAddress(range.Y, range.LastY, range.X, range.LastX), color);
        }

        public void Highlight(CellRangeAddress range, Color color)
        {
            for (int y0 = range.FirstRow; y0 <= range.LastRow; y0++)
            {
                IRow row = GetRow(y0 + 1, true);
                for (int x0 = range.FirstColumn; x0 < row.LastCellNum && x0 <= range.LastColumn; x0++)
                {
                    ICell c = row.GetCell(x0 + 1, true);
                    c.CellStyle = highlight(Workbook, c.CellStyle, color);
                }
            }
        }

        public void SetStyle(CellRangeAddress range, ICellStyle style)
        {
            for (int y0 = range.FirstRow; y0 <= range.LastRow; y0++)
            {
                IRow row = GetRow(y0 + 1, true);
                for (int x0 = range.FirstColumn; x0 < row.LastCellNum && x0 <= range.LastColumn; x0++)
                {
                    ICell c = row.GetCell(x0 + 1, true);
                    c.CellStyle = style;
                }
            }
        }
    }
}