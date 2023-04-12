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
            return GetRowsInRange(RowScope.OnlyExisting, y1, y2).Max(a => a._GetLastColumn(includeMerged));
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
            return row._GetLastNotEmptyColumn(includeMerged);
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
            return row._GetLastColumn(includeMerged);
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
            return GetRowsInRange(RowScope.OnlyExisting, y1, y2).Max(a => a._GetLastNotEmptyColumn(includeMerged));
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

        public IRow GetRow(int y, bool createRow)
        {
            return Sheet._GetRow(y, createRow);
        }

        public int GetLastRow(bool includeMerged = true)
        {
            IRow row = Sheet.GetRow(Sheet.LastRowNum);
            if (row == null)
                return 0;
            if (!includeMerged)
                return row._Y();
            int maxY = 0;
            foreach (var c in row.Cells)
            {
                var r = c._GetMergedRange();
                if (r != null && maxY < r.Y2.Value)
                    maxY = r.Y2.Value;
            }
            return maxY;
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
            GetRowsInRange(RowScope.OnlyExisting, y1, y2).ForEach(a => a.Height = -1);
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
            OnlyExisting,
            IncludeNull,
            CreateIfNull
        }
        public IEnumerable<IRow> GetRowsInRange(RowScope rowScope = RowScope.IncludeNull, int y1 = 1, int? y2 = null)
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

        public IEnumerable<IRow> GetRows(RowScope rowScope = RowScope.IncludeNull)
        {
            return Sheet._GetRows(rowScope);
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