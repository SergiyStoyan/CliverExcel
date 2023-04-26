//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.Util;
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
        /// <param name="includeMerged"></param>
        /// <param name="x1"></param>
        /// <param name="x2"></param>
        /// <returns>1-based, otherwise 0</returns>
        public int GetLastNotEmptyRowInColumnRange(bool includeMerged, int x1 = 1, int? x2 = null)
        {
            return Sheet._GetLastNotEmptyRowInColumnRange(includeMerged, x1, x2);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="includeMerged"></param>
        /// <param name="xs"></param>
        /// <returns>1-based, otherwise 0</returns>
        public int GetLastNotEmptyRowInColumns(bool includeMerged, params int[] xs)
        {
            return Sheet._GetLastNotEmptyRowInColumns(includeMerged, xs);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="x"></param>
        /// <param name="includeMerged"></param>
        /// <returns>1-based, otherwise 0</returns>
        public int GetLastNotEmptyRowInColumn(bool includeMerged, int x)
        {
            return Sheet._GetLastNotEmptyRowInColumn(includeMerged, x);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="includeMerged"></param>
        /// <param name="x"></param>
        /// <returns>1-based, otherwise 0</returns>
        public int GetLastRowInColumn(LastRowCondition lastRowCondition, bool includeMerged, int x)
        {
            return Sheet._GetColumn(x).GetLastRow(lastRowCondition, includeMerged);
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
            columnXs2width[lastColumnX] = Sheet.GetColumnWidth(lastColumnX - 1);
            //var rows = Sheet.GetRowEnumerator();//!!!buggy: sometimes misses added rows
            //while (rows.MoveNext())
            for (int y0 = Sheet.LastRowNum; y0 >= 0; y0--)
            {
                IRow row = Sheet.GetRow(y0);
                if (row == null)
                    continue;
                int columnX = row._GetLastColumn(true);
                if (lastColumnX < columnX)
                {
                    for (int i = lastColumnX; i < columnX; i++)
                        columnXs2width[i + 1] = Sheet.GetColumnWidth(i);
                    lastColumnX = columnX;
                }
                for (int i = columnX; i >= x1; i--)
                    MoveCell(row._Y(), i, row._Y(), i + shift, onFormulaCellMoved);
            }
            foreach (int columnX in columnXs2width.Keys.OrderByDescending(a => a))
                SetColumnWidth(columnX + shift, columnXs2width[columnX]);
        }

        public void ShiftColumnsLeft(int x1, int shift, Action<ICell> onFormulaCellMoved = null)
        {
            Dictionary<int, int> columnXs2width = new Dictionary<int, int>();
            int lastColumnX = x1;
            columnXs2width[lastColumnX] = Sheet.GetColumnWidth(lastColumnX - 1);
            //var rows = Sheet.GetRowEnumerator();//!!!buggy: sometimes misses added rows
            //while (rows.MoveNext())
            for (int y0 = Sheet.LastRowNum; y0 >= 0; y0--)
            {
                IRow row = Sheet.GetRow(y0);
                if (row == null)
                    continue;
                int columnX = row._GetLastColumn(true);
                if (lastColumnX < columnX)
                {
                    for (int i = lastColumnX; i < columnX; i++)
                        columnXs2width[i + 1] = Sheet.GetColumnWidth(i);
                    lastColumnX = columnX;
                }
                for (int i = x1; i <= columnX; i++)
                    MoveCell(row._Y(), i, row._Y(), i - shift, onFormulaCellMoved);
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
            return GetLastNotEmptyColumnInRowRange(includeMerged, 1, null);
        }

        public void CopyColumn(string fromColumnName, ISheet toSheet, string toColumnName = null)
        {
            Sheet._GetColumn(fromColumnName).Copy(toSheet, toColumnName);
        }

        public void CopyColumn(int fromX, ISheet toSheet, int toX)
        {
            Sheet._CopyColumn(fromX, toSheet, toX);
        }

        /// <summary>
        /// (!)Very slow on large data.
        /// </summary>
        /// <param name="columnIs"></param>
        /// <param name="padding">a character width</param>
        public void AutosizeColumns(IEnumerable<int> Xs, float padding = 0)
        {
            Sheet._AutosizeColumns(Xs, padding);
        }

        /// <summary>
        /// (!)Very slow on large data.
        /// </summary>
        /// <param name="x"></param>
        /// <param name="padding">a character width</param>
        public void AutosizeColumn(int x, float padding = 0)
        {
            Sheet._AutosizeColumn(x, padding);
        }

        public IEnumerable<ICell> GetCellsInColumn(int x, RowScope rowScope)
        {
            return Sheet._GetCellsInColumn(x, rowScope);
        }

        /// <summary>
        /// Safe against the API's one
        /// </summary>
        /// <param name="x"></param>
        /// <param name="width">units of 1/256th of a character width</param>
        public void SetColumnWidth(int x, int width)
        {
            Sheet._SetColumnWidth(x, width);
        }

        /// <summary>
        /// Safe against the API's one
        /// </summary>
        /// <param name="x"></param>
        /// <param name="width">a character width</param>
        public void SetColumnWidth(int x, float width)
        {
            Sheet._SetColumnWidth(x, width);
        }

        /// <summary>
        /// (!)Very slow on large data.
        /// </summary>
        /// <param name="x1"></param>
        /// <param name="x2"></param>
        /// <param name="padding">a character width</param>
        public void AutosizeColumnsInRange(int x1 = 1, int? x2 = null, float padding = 0)
        {
            Sheet._AutosizeColumnsInRange(x1, x2, padding);
        }

        public int GetLastColumn(bool includeMerged)
        {
            return Sheet._GetLastColumn(includeMerged);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="padding">a character width</param>
        public void AutosizeColumns(float padding = 0)
        {
            Sheet._AutosizeColumns(padding);
        }

        public void ClearColumn(int x, bool clearMerging)
        {
            Sheet._ClearColumn(x, clearMerging);
        }

        public void ClearMergingInColumn(int x)
        {
            Sheet._ClearMergingInColumn(x);
        }

        public void SetStyleInColumn(ICellStyle style, bool createCells, int x)
        {
            Sheet._SetStyleInColumn(style, createCells, x);
        }

        public void SetStyleInColumnRange(ICellStyle style, bool createCells, int x1, int? x2 = null)
        {
            Sheet._SetStyleInColumnRange(style, createCells, x1, x2);
        }

        public void ReplaceStyleInColumnRange(ICellStyle style1, ICellStyle style2, int x1, int? x2 = null)
        {
            Sheet._ReplaceStyleInColumnRange(style1, style2, x1, x2);
        }

        public void ClearStyleInColumnRange(ICellStyle style, int x1, int? x2 = null)
        {
            Sheet._ClearStyleInColumnRange(style, x1, x2);
        }

        public Column GetColumn(int x)
        {
            return Sheet._GetColumn(x);
        }

        public IEnumerable<Column> GetColumnsInRange(int x1 = 1, int? x2 = null)
        {
            return Sheet._GetColumnsInRange(x1, x2);
        }

        public IEnumerable<Column> GetColumns()
        {
            return Sheet._GetColumns();
        }

        public void ShiftColumnCellsDown(int x, int y1, int shift, Action<ICell> onFormulaCellMoved = null)
        {
            Sheet._ShiftColumnCellsDown(x, y1, shift, onFormulaCellMoved);
        }

        public void ShiftColumnCellsUp(int x, int y1, int shift, Action<ICell> onFormulaCellMoved = null)
        {
            Sheet._ShiftColumnCellsUp(x, y1, shift, onFormulaCellMoved);
        }
    }
}