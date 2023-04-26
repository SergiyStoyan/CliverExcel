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
        public int GetLastColumnInRowRange(bool includeMerged, int y1 = 1, int? y2 = null)
        {
            return GetRowsInRange(RowScope.NotNull, y1, y2).Max(a => a._GetLastColumn(includeMerged));
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="includeMerged"></param>
        /// <param name="y"></param>
        /// <returns>1-based, otherwise 0</returns>
        public int GetLastNotEmptyColumnInRow(bool includeMerged, int y)
        {
            IRow row = GetRow(y, false);
            if (row == null)
                return 0;
            return row._GetLastNotEmptyColumn(includeMerged);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="includeMerged"></param>
        /// <param name="y"></param>
        /// <returns>1-based, otherwise 0</returns>
        public int GetLastColumnInRow(bool includeMerged, int y)
        {
            IRow row = GetRow(y, false);
            if (row == null)
                return 0;
            return row._GetLastColumn(includeMerged);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="includeMerged"></param>
        /// <param name="y1"></param>
        /// <param name="y2"></param>
        /// <returns>1-based, otherwise 0</returns>
        public int GetLastNotEmptyColumnInRowRange(bool includeMerged, int y1 = 1, int? y2 = null)
        {
            return GetRowsInRange(RowScope.NotNull, y1, y2).Max(a => a._GetLastNotEmptyColumn(includeMerged));
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="includeMerged"></param>
        /// <returns>1-based, otherwise 0</returns>
        public int GetLastNotEmptyRow(bool includeMerged)
        {
            return GetLastNotEmptyRowInColumnRange(includeMerged, 1, null);
        }

        public IRow GetRow(int y, bool createRow)
        {
            return Sheet._GetRow(y, createRow);
        }

        public IRow RemoveRow(int y, bool shiftRemainingRows)
        {
            return Sheet._RemoveRow(y, shiftRemainingRows);
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
            GetRowsInRange(RowScope.NotNull, y1, y2).ForEach(a => a.Height = -1);
        }

        public void AutosizeRows()
        {
            AutosizeRowsInRange();
        }

        public void ClearRow(int y, bool clearMerging)
        {
            if (clearMerging)
                ClearMergingInRow(y);
            var r = GetRow(y, false);
            if (r != null)
                Sheet.RemoveRow(r);
        }

        public void ClearMergingInRow(int y)
        {
            NewRange(y, 1, y, null).ClearMerging();
        }

        public enum RowScope
        {
            /// <summary>
            /// (!)Considerably slow due to checking all the cells' values
            /// </summary>
            NotEmpty,
            /// <summary>
            /// Returns only rows with cells.
            /// </summary>
            WithCells,
            /// <summary>
            /// Returns only rows existing as objects.
            /// </summary>
            NotNull,
            /// <summary>
            /// Returns NULL for non-existing rows within the range.
            /// (!)Might return a huge pile of null and no-cell rows after the last not empty row.  
            /// </summary>
            IncludeNull,
            /// <summary>
            /// (!)When using it, make sure that ISheet::LastRowNum is not huge.
            /// </summary>
            CreateIfNull
        }
        public IEnumerable<IRow> GetRowsInRange(RowScope rowScope, int y1 = 1, int? y2 = null)
        {
            return Sheet._GetRowsInRange(rowScope, y1, y2);
        }

        public void SetStyleInRow(ICellStyle style, bool createCells, int y)
        {
            SetStyleInRowRange(style, createCells, y, y);
        }

        public void SetStyleInRowRange(ICellStyle style, bool createCells, int y1, int? y2 = null)
        {
            NewRange(y1, 1, y2, null).SetStyle(style, createCells);
        }

        public void ReplaceStyleInRowRange(ICellStyle style1, ICellStyle style2, int y1, int? y2 = null)
        {
            NewRange(y1, 1, y2, null).ReplaceStyle(style1, style2);
        }

        public void ClearStyleInRowRange(ICellStyle style, int y1, int? y2 = null)
        {
            ReplaceStyleInRowRange(style, null, y1, y2);
        }

        public IRow AppendRow<T>(IEnumerable<T> values)
        {
            return Sheet._AppendRow(values);
        }

        public IRow AppendRow(params string[] values)
        {
            return Sheet._AppendRow(values);
        }

        public IRow InsertRow<T>(int y, IEnumerable<T> values = null)
        {
            return Sheet._InsertRow(y, values);
        }

        public IRow InsertRow(int y, params string[] values)
        {
            return Sheet._InsertRow(y, values);
        }

        public IRow WriteRow<T>(int y, IEnumerable<T> values)
        {
            return Sheet._WriteRow(y, values);
        }

        public IRow WriteRow<T>(int y, params string[] values)
        {
            return Sheet._WriteRow(y, values);
        }

        public void ShiftRowCellsRight(int y, int x1, int shift, Action<ICell> onFormulaCellMoved = null)
        {
            GetRow(y, false)?._ShiftCellsRight(x1, shift, onFormulaCellMoved);
        }

        public void ShiftRowCellsLeft(int y, int x1, int shift, Action<ICell> onFormulaCellMoved = null)
        {
            GetRow(y, false)?._ShiftCellsLeft(x1, shift, onFormulaCellMoved);
        }
    }
}