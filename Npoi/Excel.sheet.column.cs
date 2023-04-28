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
        /// <param name="includeMerged"></param>
        /// <param name="x"></param>
        /// <returns>1-based, otherwise 0</returns>
        public int GetLastRowInColumn(LastRowCondition lastRowCondition, bool includeMerged, int x)
        {
            return Sheet._GetLastRowInColumn(lastRowCondition, includeMerged, x);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cellValue"></param>
        /// <param name="cellY"></param>
        /// <returns>1-based, otherwise 0</returns>
        public int FindColumnByCellValue(Regex cellValue, int cellY = 1)
        {
            return Sheet._FindColumnByCellValue(cellValue, cellY);
        }

        public void ShiftColumnsRight(int x1, int shift, Action<ICell> onFormulaCellMoved = null)
        {
            Sheet._ShiftColumnsRight(x1, shift, onFormulaCellMoved);
        }

        public void ShiftColumnsLeft(int x1, int shift, Action<ICell> onFormulaCellMoved = null)
        {
            Sheet._ShiftColumnsLeft(x1, shift, onFormulaCellMoved);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="includeMerged"></param>
        /// <returns>1-based, otherwise 0</returns>
        public int GetLastNotEmptyColumn(bool includeMerged)
        {
            return Sheet._GetLastNotEmptyColumn(includeMerged);
        }

        public void CopyColumn(string fromColumnName, ISheet toSheet, string toColumnName = null)
        {
            Sheet._CopyColumn(fromColumnName, toSheet, toColumnName);
        }

        public void CopyColumn(int fromX, ISheet toSheet, int toX)
        {
            Sheet._CopyColumn(fromX, toSheet, toX);
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

        public Column GetColumn(int x)
        {
            return Sheet._GetColumn(x);
        }

        public Column GetColumn(string columnName)
        {
            return Sheet._GetColumn(columnName);
        }

        public IEnumerable<Column> GetColumns()
        {
            return Sheet._GetColumns();
        }

        public IEnumerable<Column> GetColumnsInRange(int x1 = 1, int? x2 = null)
        {
            return Sheet._GetColumnsInRange(x1, x2);
        }

        public int GetLastColumn(bool includeMerged)
        {
            return Sheet._GetLastColumn(includeMerged);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="includeMerged"></param>
        /// <returns>1-based, otherwise 0</returns>
        public int GetLastNotEmptyRow(bool includeMerged)
        {
            return Sheet._GetLastNotEmptyRow(includeMerged);
        }

        public void ShiftColumnCellsDown(int x, int y1, int shift, Action<ICell> onFormulaCellMoved = null)
        {
            Sheet._ShiftColumnCellsDown(x, y1, shift, onFormulaCellMoved);
        }

        public void ShiftColumnCellsUp(int x, int y1, int shift, Action<ICell> onFormulaCellMoved = null)
        {
            Sheet._ShiftColumnCellsUp(x, y1, shift, onFormulaCellMoved);
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
    }
}