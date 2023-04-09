//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Cliver
{
    public partial class Excel
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="header"></param>
        /// <param name="headerY"></param>
        /// <returns>1-based, otherwise 0</returns>
        public int FindColumnByHeader(Regex header, int headerY = 1)
        {
<<<<<<< Updated upstream
            IRow row = GetRow(headerY, false);
=======
            return Sheet.GetLastNotEmptyRowInColumnRange(x1, x2, includeMerged);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="includeMerged"></param>
        /// <param name="xs"></param>
        /// <returns>1-based, otherwise 0</returns>
        public int GetLastNotEmptyRowInColumns(bool includeMerged, params int[] xs)
        {
            return Sheet.GetLastNotEmptyRowInColumns(includeMerged, xs);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="x"></param>
        /// <param name="includeMerged"></param>
        /// <returns>1-based, otherwise 0</returns>
        public int GetLastNotEmptyRowInColumn(int x, bool includeMerged = true)
        {
            return Sheet.GetLastNotEmptyRowInColumn(x, includeMerged);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="x"></param>
        /// <param name="includeMerged"></param>
        /// <returns>1-based, otherwise 0</returns>
        public int GetLastRowInColumn(int x, bool includeMerged = true)
        {
            return Sheet.GetColumn(x).GetLastRow(includeMerged);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cellValue"></param>
        /// <param name="cellY"></param>
        /// <returns>1-based, otherwise 0</returns>
        public int FindColumnByCellValue(Regex cellValue, int cellY = 1)
        {
            IRow row = GetRow(cellY, false);
>>>>>>> Stashed changes
            if (row == null)
                return 0;
            for (int x = 1; x <= row.Cells.Count; x++)
                if (header.IsMatch(GetValueAsString(headerY, x, false)))
                    return x;
            return 0;
        }

        public void ShiftColumns(int x, int shift, Action<ICell> onFormulaCellMoved = null)
        {
            Dictionary<int, int> columnXs2width = new Dictionary<int, int>();
<<<<<<< Updated upstream
            int lastColumnX = x;
            columnXs2width[lastColumnX] = Sheet.GetColumnWidth(lastColumnX - 1);
=======
            int lastColumnX = x1;
            columnXs2width[lastColumnX] = Sheet._.GetColumnWidth(lastColumnX - 1);
>>>>>>> Stashed changes
            //var rows = Sheet.GetRowEnumerator();//!!!buggy: sometimes misses added rows
            //while (rows.MoveNext())
            for (int y0 = Sheet._.LastRowNum; y0 >= 0; y0--)
            {
                IRow row = Sheet._.GetRow(y0);
                if (row == null)
                    continue;
                int columnX = row.GetLastColumn(true);
                if (lastColumnX < columnX)
                {
                    for (int i = lastColumnX; i < columnX; i++)
                        columnXs2width[i + 1] = Sheet._.GetColumnWidth(i);
                    lastColumnX = columnX;
                }
                for (int i = columnX; i >= x; i--)
                    MoveCell(row.Y(), i, row.Y(), i + shift, onFormulaCellMoved);
            }
            foreach (int columnX in columnXs2width.Keys.OrderByDescending(a => a))
                SetColumnWidth(columnX + shift, columnXs2width[columnX]);
        }

        public void ShiftColumns(IRow row, int x, int shift, Action<ICell> onFormulaCellMoved = null)
        {
<<<<<<< Updated upstream
            for (int i = row.GetLastColumn(true); i >= x; i--)
                MoveCell(row.Y(), i, row.Y(), i + shift, onFormulaCellMoved);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="y"></param>
        /// <param name="includeMerged"></param>
        /// <returns>1-based, otherwise 0</returns>
        public int GetLastNotEmptyColumnInRow(int y, bool includeMerged = true)
        {
            IRow row = GetRow(y, false);
            if (row == null)
                return 0;
            return row.GetLastNotEmptyColumn(includeMerged);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="y"></param>
        /// <param name="includeMerged"></param>
        /// <returns>1-based, otherwise 0</returns>
        public int GetLastColumnInRow(int y, bool includeMerged = true)
        {
            IRow row = GetRow(y, false);
            if (row == null)
                return 0;
            return row.GetLastColumn(includeMerged);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="y1"></param>
        /// <param name="y2"></param>
        /// <param name="includeMerged"></param>
        /// <returns>1-based, otherwise 0</returns>
        public int GetLastNotEmptyColumnInRowRange(int y1 = 1, int? y2 = null, bool includeMerged = true)
        {
            if (y2 == null)
                y2 = Sheet.LastRowNum + 1;
            return GetRowsInRange(RowScope.OnlyExisting, y1, y2).Max(a => a.GetLastNotEmptyColumn(includeMerged));
=======
            Dictionary<int, int> columnXs2width = new Dictionary<int, int>();
            int lastColumnX = x1;
            columnXs2width[lastColumnX] = Sheet._.GetColumnWidth(lastColumnX - 1);
            //var rows = Sheet.GetRowEnumerator();//!!!buggy: sometimes misses added rows
            //while (rows.MoveNext())
            for (int y0 = Sheet._.LastRowNum; y0 >= 0; y0--)
            {
                IRow row = Sheet._.GetRow(y0);
                if (row == null)
                    continue;
                int columnX = row.GetLastColumn(true);
                if (lastColumnX < columnX)
                {
                    for (int i = lastColumnX; i < columnX; i++)
                        columnXs2width[i + 1] = Sheet._.GetColumnWidth(i);
                    lastColumnX = columnX;
                }
                for (int i = x1; i <= columnX; i++)
                    MoveCell(row.Y(), i, row.Y(), i - shift, onFormulaCellMoved);
            }
            foreach (int columnX in columnXs2width.Keys.OrderByDescending(a => a))
                SetColumnWidth(columnX - shift, columnXs2width[columnX]);
>>>>>>> Stashed changes
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="includeMerged"></param>
        /// <returns>1-based, otherwise 0</returns>
        public int GetLastNotEmptyColumn(bool includeMerged)
        {
            return GetLastNotEmptyColumnInRowRange(1, null, includeMerged);
        }

        public void CopyColumn(string columnName, ISheet destinationSheet)
        {
            int x = CellReference.ConvertColStringToIndex(columnName);
            CopyColumn(x, destinationSheet);
        }

        public void CopyColumn(int x, ISheet destinationSheet)
        {
<<<<<<< Updated upstream
            var range = new Range(1, x, null, x);
            CopyRange(range, destinationSheet);
=======
            Sheet.CopyColumn(fromX, toSheet, toX);
>>>>>>> Stashed changes
        }

        /// <summary>
        /// (!)Very slow on large data.
        /// </summary>
        /// <param name="columnIs"></param>
        /// <param name="padding">a character width</param>
        public void AutosizeColumns(IEnumerable<int> Xs, float padding = 0)
        {
            Sheet.AutosizeColumns(Xs, padding);
        }

        /// <summary>
        /// (!)Very slow on large data.
        /// </summary>
        /// <param name="x"></param>
        /// <param name="padding">a character width</param>
        public void AutosizeColumn(int x, float padding = 0)
        {
<<<<<<< Updated upstream
            Sheet.AutoSizeColumn(x - 1, false);

            //GetCellsInColumn(x).Max(a => a.GetValueAsString())
            //int width = ((int)(maxNumCharacters * 1.14388)) * 256;
            //sheet.setColumnWidth(i, width);

            if (padding > 0)
                SetColumnWidth(x, Sheet.GetColumnWidth(x - 1) + (int)(padding * 256));
=======
            Sheet.AutosizeColumn(x, padding);
>>>>>>> Stashed changes
        }

        public IEnumerable<ICell> GetCellsInColumn(int x)
        {
<<<<<<< Updated upstream
            return GetRows().Select(a => a.GetCell(x));
=======
            return Sheet.GetCellsInColumn(x);
>>>>>>> Stashed changes
        }

        /// <summary>
        /// Safe against the API's one
        /// </summary>
        /// <param name="x"></param>
        /// <param name="width">units of 1/256th of a character width</param>
        public void SetColumnWidth(int x, int width)
        {
<<<<<<< Updated upstream
            const int cellMaxWidth = 256 * 255;
            int w = MathRoutines.Truncate(width, cellMaxWidth);
            Sheet.SetColumnWidth(x - 1, w);
=======
            Sheet.SetColumnWidth(x, width);
>>>>>>> Stashed changes
        }

        /// <summary>
        /// Safe against the API's one
        /// </summary>
        /// <param name="x"></param>
        /// <param name="width">a character width</param>
        public void SetColumnWidth(int x, float width)
        {
<<<<<<< Updated upstream
            SetColumnWidth(x, (int)(width * 255));
=======
            Sheet.SetColumnWidth(x, width);
>>>>>>> Stashed changes
        }

        /// <summary>
        /// (!)Very slow on large data.
        /// </summary>
        /// <param name="x1"></param>
        /// <param name="x2"></param>
        /// <param name="padding">a character width</param>
        public void AutosizeColumnsInRange(int x1 = 1, int? x2 = null, float padding = 0)
        {
            Sheet.AutosizeColumnsInRange(x1, x2, padding);
        }

        public int GetLastColumnInRowRange(int y1 = 1, int? y2 = null, bool includeMerged = true)
        {
            return GetRowsInRange(RowScope.OnlyExisting, y1, y2).Max(a => a.GetLastColumn(includeMerged));
        }

        public int GetLastColumn(bool includeMerged = true)
        {
            return GetLastColumnInRowRange(1, null, includeMerged);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="padding">a character width</param>
        public void AutosizeColumns(float padding = 0)
        {
            Sheet.AutosizeColumns(padding);
        }

        public void ClearColumn(int x, bool clearMerging)
        {
<<<<<<< Updated upstream
            if (clearMerging)
                ClearMergingInColumn(x);
            //var rows = Sheet.GetRowEnumerator();//!!!buggy: sometimes misses added rows
            //while (rows.MoveNext())
            for (int y0 = Sheet.LastRowNum; y0 >= 0; y0--)
            {
                IRow row = Sheet.GetRow(y0);
                if (row == null)
                    continue;
                ICell c = row.GetCell(x);
                if (c != null)
                    row.RemoveCell(c);
            }
=======
            Sheet.ClearColumn(x, clearMerging);
>>>>>>> Stashed changes
        }

        public void ClearMergingInColumn(int x)
        {
<<<<<<< Updated upstream
            ClearMerging(new Range(1, x, null, x));
=======
            Sheet.ClearMergingInColumn(x);
>>>>>>> Stashed changes
        }

        public void SetStyleInColumn(ICellStyle style, bool createCells, int x)
        {
            Sheet.SetStyleInColumn(style, createCells, x);
        }

        public void SetStyleInColumnRange(ICellStyle style, bool createCells, int x1, int? x2 = null)
        {
<<<<<<< Updated upstream
            SetStyle(new Range(1, x1, null, x2), style, createCells);
=======
            Sheet.SetStyleInColumnRange(style, createCells, x1, x2);
>>>>>>> Stashed changes
        }

        public void ReplaceStyleInColumnRange(ICellStyle style1, ICellStyle style2, int x1, int? x2 = null)
        {
<<<<<<< Updated upstream
            ReplaceStyle(new Range(1, x1, null, x2), style1, style2);
=======
            Sheet.ReplaceStyleInColumnRange(style1, style2, x1, x2);
>>>>>>> Stashed changes
        }

        public void ClearStyleInColumnRange(ICellStyle style, int x1, int? x2 = null)
        {
            Sheet.ClearStyleInColumnRange(style, x1, x2);
        }
<<<<<<< Updated upstream
=======

        public Column GetColumn(int x)
        {
            return Sheet.GetColumn(x);
        }

        public IEnumerable<Column> GetColumnsInRange(int x1 = 1, int? x2 = null)
        {
            return Sheet.GetColumnsInRange(x1, x2);
        }

        public IEnumerable<Column> GetColumns()
        {
            return Sheet.GetColumns();
        }

        public void ShiftColumnCellsDown(int x, int y1, int shift, Action<ICell> onFormulaCellMoved = null)
        {
            Sheet.ShiftColumnCellsDown(x, y1, shift, onFormulaCellMoved);
        }

        public void ShiftColumnCellsUp(int x, int y1, int shift, Action<ICell> onFormulaCellMoved = null)
        {
            Sheet.ShiftColumnCellsUp(x, y1, shift, onFormulaCellMoved);
        }
>>>>>>> Stashed changes
    }
}