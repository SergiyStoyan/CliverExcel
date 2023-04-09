//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************

using System;
using System.Collections.Generic;
using NPOI.SS.UserModel;
using static Cliver.Excel;
using System.Linq;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System.Text.RegularExpressions;

namespace Cliver
{
    public partial class Sheet
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="includeMerged"></param>
        /// <param name="xs"></param>
        /// <returns>1-based, otherwise 0</returns>
        public int GetLastNotEmptyRowInColumns(bool includeMerged, params int[] xs)
        {
            for (int i = _.LastRowNum; i >= 0; i--)
            {
                IRow row = _.GetRow(i);
                if (row == null)
                    continue;
                var c = row.Cells.Find(a => xs.Contains(a.ColumnIndex + 1) && !string.IsNullOrEmpty(a.GetValueAsString()));
                if (string.IsNullOrEmpty(c?.GetValueAsString()))
                    continue;
                if (includeMerged)
                {
                    var r = c.GetMergedRange();
                    if (r != null)
                        return r.Y2.Value;
                }
                return c.RowIndex + 1;
            }
            return 0;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="x1"></param>
        /// <param name="x2"></param>
        /// <param name="includeMerged"></param>
        /// <returns>1-based, otherwise 0</returns>
        public int GetLastNotEmptyRowInColumnRange(int x1 = 1, int? x2 = null, bool includeMerged = true)
        {
            if (x2 == null)
                x2 = int.MaxValue;
            for (int i = _.LastRowNum; i >= 0; i--)
            {
                IRow row = _.GetRow(i);
                if (row == null)
                    continue;
                var c = row.Cells.Find(a => a.ColumnIndex + 1 >= x1 && a.ColumnIndex < x2 && !string.IsNullOrEmpty(a.GetValueAsString()));
                if (string.IsNullOrEmpty(c?.GetValueAsString()))
                    continue;
                if (includeMerged)
                {
                    var r = c.GetMergedRange();
                    if (r != null)
                        return r.Y2.Value;
                }
                return c.RowIndex + 1;
            }
            return 0;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="x"></param>
        /// <param name="includeMerged"></param>
        /// <returns>1-based, otherwise 0</returns>
        public int GetLastRowInColumn(int x, bool includeMerged = true)
        {
            return GetColumn(x).GetLastRow(includeMerged);
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
            if (row == null)
                return 0;
            for (int x = 1; x <= row.Cells.Count; x++)
                if (cellValue.IsMatch(GetValueAsString(cellY, x, false)))
                    return x;
            return 0;
        }

        public void ShiftColumnsRight(int x1, int shift, Action<ICell> onFormulaCellMoved = null)
        {
            Dictionary<int, int> columnXs2width = new Dictionary<int, int>();
            int lastColumnX = x1;
            columnXs2width[lastColumnX] = _.GetColumnWidth(lastColumnX - 1);
            //var rows = Sheet.GetRowEnumerator();//!!!buggy: sometimes misses added rows
            //while (rows.MoveNext())
            for (int y0 = _.LastRowNum; y0 >= 0; y0--)
            {
                IRow row = _.GetRow(y0);
                if (row == null)
                    continue;
                int columnX = row.GetLastColumn(true);
                if (lastColumnX < columnX)
                {
                    for (int i = lastColumnX; i < columnX; i++)
                        columnXs2width[i + 1] = _.GetColumnWidth(i);
                    lastColumnX = columnX;
                }
                for (int i = columnX; i >= x1; i--)
                    MoveCell(row.Y(), i, row.Y(), i + shift, onFormulaCellMoved);
            }
            foreach (int columnX in columnXs2width.Keys.OrderByDescending(a => a))
                SetColumnWidth(columnX + shift, columnXs2width[columnX]);
        }

        public void ShiftColumnsLeft(int x1, int shift, Action<ICell> onFormulaCellMoved = null)
        {
            Dictionary<int, int> columnXs2width = new Dictionary<int, int>();
            int lastColumnX = x1;
            columnXs2width[lastColumnX] = _.GetColumnWidth(lastColumnX - 1);
            //var rows = Sheet.GetRowEnumerator();//!!!buggy: sometimes misses added rows
            //while (rows.MoveNext())
            for (int y0 = _.LastRowNum; y0 >= 0; y0--)
            {
                IRow row = _.GetRow(y0);
                if (row == null)
                    continue;
                int columnX = row.GetLastColumn(true);
                if (lastColumnX < columnX)
                {
                    for (int i = lastColumnX; i < columnX; i++)
                        columnXs2width[i + 1] = _.GetColumnWidth(i);
                    lastColumnX = columnX;
                }
                for (int i = x1; i <= columnX; i++)
                    MoveCell(row.Y(), i, row.Y(), i - shift, onFormulaCellMoved);
            }
            foreach (int columnX in columnXs2width.Keys.OrderByDescending(a => a))
                SetColumnWidth(columnX - shift, columnXs2width[columnX]);
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

        public void CopyColumn(string fromColumnName, ISheet toSheet, string toColumnName = null)
        {
            GetColumn(fromColumnName).Copy(toSheet, toColumnName);
        }

        public void CopyColumn(int fromX, ISheet toSheet, int toX)
        {
            GetColumn(fromX).Copy(toSheet, toX);
        }

        public int GetLastNotEmptyRowInColumn(int x, bool includeMerged = true)
        {
            return GetColumn(x).GetLastNotEmptyRow(includeMerged);
        }
        public Column GetColumn(int x)
        {
            return new Column(_, x);
        }

        public Column GetColumn(string columnName)
        {
            return new Column(_, CellReference.ConvertColStringToIndex(columnName));
        }

        public IEnumerable<Column> GetColumns()
        {
            return GetColumnsInRange();
        }

        public IEnumerable<Column> GetColumnsInRange(int x1 = 1, int? x2 = null)
        {
            if (x2 == null)
                x2 = GetLastColumn(false);
            for (int x = x1; x <= x2; x++)
                yield return new Column(_, x);
        }

        public int GetLastColumn(bool includeMerged = true)
        {
            return GetLastColumnInRowRange(1, null, includeMerged);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="includeMerged"></param>
        /// <returns>1-based, otherwise 0</returns>
        public int GetLastNotEmptyRow(bool includeMerged = true)
        {
            return GetLastNotEmptyRowInColumnRange(1, null, includeMerged);
        }

        public void ShiftColumnCellsDown(int x, int y1, int shift, Action<ICell> onFormulaCellMoved = null)
        {
            GetColumn(x).ShiftCellsDown(y1, shift, onFormulaCellMoved);
        }

        public void ShiftColumnCellsUp(int x, int y1, int shift, Action<ICell> onFormulaCellMoved = null)
        {
            GetColumn(x).ShiftCellsUp(y1, shift, onFormulaCellMoved);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="padding">a character width</param>
        public void AutosizeColumns(float padding = 0)
        {
            AutosizeColumnsInRange(1, null, padding);
        }

        public void ClearColumn(int x, bool clearMerging)
        {
            GetColumn(x).Clear(clearMerging);
        }

        public void ClearMergingInColumn(int x)
        {
            NewRange(1, x, null, x).ClearMerging();
        }

        public void SetStyleInColumn(ICellStyle style, bool createCells, int x)
        {
            SetStyleInColumnRange(style, createCells, x, x);
        }

        public void SetStyleInColumnRange(ICellStyle style, bool createCells, int x1, int? x2 = null)
        {
            NewRange(1, x1, null, x2).SetStyle(style, createCells);
        }

        public void ReplaceStyleInColumnRange(ICellStyle style1, ICellStyle style2, int x1, int? x2 = null)
        {
            NewRange(1, x1, null, x2).ReplaceStyle(style1, style2);
        }

        public void ClearStyleInColumnRange(ICellStyle style, int x1, int? x2 = null)
        {
            ReplaceStyleInColumnRange(style, null, x1, x2);
        }

        /// <summary>
        /// (!)Very slow on large data.
        /// </summary>
        /// <param name="x1"></param>
        /// <param name="x2"></param>
        /// <param name="padding">a character width</param>
        public void AutosizeColumnsInRange(int x1 = 1, int? x2 = null, float padding = 0)
        {
            if (x2 == null)
                x2 = GetLastColumn();
            for (int x = x1; x <= x2; x++)
                AutosizeColumn(x, padding);
        }

        /// <summary>
        /// (!)Very slow on large data.
        /// </summary>
        /// <param name="columnIs"></param>
        /// <param name="padding">a character width</param>
        public void AutosizeColumns(IEnumerable<int> Xs, float padding = 0)
        {
            foreach (int y in Xs)
                AutosizeColumn(y, padding);
        }

        /// <summary>
        /// (!)Very slow on large data.
        /// </summary>
        /// <param name="x"></param>
        /// <param name="padding">a character width</param>
        public void AutosizeColumn(int x, float padding = 0)
        {
            GetColumn(x).Autosize(padding);
        }

        public IEnumerable<ICell> GetCellsInColumn(int x)
        {
            return GetColumn(x).GetCells();
        }

        /// <summary>
        /// Safe against the API's one
        /// </summary>
        /// <param name="x"></param>
        /// <param name="width">units of 1/256th of a character width</param>
        public void SetColumnWidth(int x, int width)
        {
            GetColumn(x).SetWidth(width);
        }

        /// <summary>
        /// Safe against the API's one
        /// </summary>
        /// <param name="x"></param>
        /// <param name="width">a character width</param>
        public void SetColumnWidth(int x, float width)
        {
            GetColumn(x).SetWidth(width);
        }
    }
}