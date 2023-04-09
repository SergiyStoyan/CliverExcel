//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Cliver
{
    public partial class Excel : IDisposable
    {
        public int GetLastColumnInRowRange(int y1 = 1, int? y2 = null, bool includeMerged = true)
        {
            return Sheet.GetLastColumnInRowRange(y1, y2, includeMerged);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="y"></param>
        /// <param name="includeMerged"></param>
        /// <returns>1-based, otherwise 0</returns>
        public int GetLastNotEmptyColumnInRow(int y, bool includeMerged = true)
        {
            return Sheet.GetLastNotEmptyColumnInRow(y, includeMerged);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="y"></param>
        /// <param name="includeMerged"></param>
        /// <returns>1-based, otherwise 0</returns>
        public int GetLastColumnInRow(int y, bool includeMerged = true)
        {
            return Sheet.GetLastColumnInRow(y, includeMerged);
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
            return Sheet.GetLastNotEmptyColumnInRowRange(y1, y2, includeMerged);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="includeMerged"></param>
        /// <returns>1-based, otherwise 0</returns>
        public int GetLastNotEmptyRow(bool includeMerged = true)
        {
            return Sheet.GetLastNotEmptyRow(includeMerged);
        }

        public Row GetRow(int y, bool createRow)
        {
            return Sheet.GetRow(y, createRow);
        }

        public int GetLastRow(bool includeMerged = true)
        {
            return Sheet.GetLastRow(includeMerged);
        }

        //public void HighlightRow(int y, ICellStyle style, Color color)
        //{
        //    GetRow(y, true).Highlight(style, color);
        //}

        //public void Highlight(IRow row, ICellStyle style, Color color)
        //{
        //    row.Highlight(style, color);
        //}

        public void AutosizeRowsInRange(int y1 = 1, int? y2 = null)
        {
            Sheet.AutosizeRowsInRange(y1, y2);
        }

        public void AutosizeRows()
        {
            Sheet.AutosizeRows();
        }

        public void ClearRow(int y, bool clearMerging)
        {
            Sheet.ClearRow(y, clearMerging);
        }

        public void ClearMergingInRow(int y)
        {
            Sheet.ClearMergingInRow(y);
        }

        public enum RowScope
        {
            OnlyExisting,
            IncludeNull,
            CreateIfNull
        }
        public IEnumerable<Row> GetRowsInRange(RowScope rowScope = RowScope.IncludeNull, int y1 = 1, int? y2 = null)
        {
            return Sheet.GetRowsInRange(rowScope, y1, y2);
        }

        public void SetStyleInRow(ICellStyle style, bool createCells, int y)
        {
            Sheet.SetStyleInRow(style, createCells, y);
        }

        public void SetStyleInRowRange(ICellStyle style, bool createCells, int y1, int? y2 = null)
        {
            Sheet.SetStyleInRowRange(style, createCells, y1, y2);
        }

        public void ReplaceStyleInRowRange(ICellStyle style1, ICellStyle style2, int y1, int? y2 = null)
        {
            Sheet.ReplaceStyleInRowRange(style1, style2, y1, y2);
        }

        public void ClearStyleInRowRange(ICellStyle style, int y1, int? y2 = null)
        {
            Sheet.ClearStyleInRowRange(style, y1, y2);
        }

        public IEnumerable<Row> GetRows(RowScope rowScope = RowScope.IncludeNull)
        {
            return Sheet.GetRows(rowScope);
        }

        public Row AppendRow<T>(IEnumerable<T> values)
        {
            return Sheet.AppendRow(values);
        }

        public Row AppendRow<T>(params T[] values)
        {
            return Sheet.AppendRow(values);
        }

        public Row InsertRow<T>(int y, IEnumerable<T> values = null)
        {
            return Sheet.InsertRow(y, values);
        }

        public Row InsertRow<T>(int y, params T[] values)
        {
            return Sheet.InsertRow(y, values);
        }

        public Row WriteRow<T>(int y, IEnumerable<T> values)
        {
            return Sheet.WriteRow(y, values);
        }

        public Row WriteRow<T>(int y, params T[] values)
        {
            return Sheet.WriteRow(y, values);
        }

        public void ShiftRowCellsRight(int y, int x1, int shift, Action<Cell> onFormulaCellMoved = null)
        {
            Sheet.ShiftRowCellsRight(y, x1, shift, onFormulaCellMoved);
        }

        public void ShiftRowCellsLeft(int y, int x1, int shift, Action<Cell> onFormulaCellMoved = null)
        {
            Sheet.ShiftRowCellsLeft(y, x1, shift, onFormulaCellMoved);
        }
    }
}