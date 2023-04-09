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

namespace Cliver
{
    public partial class Sheet
    {
        public Row GetRow(int y, bool createRow)
        {
            IRow r = _.GetRow(y - 1);
            if (r == null)
            {
                if (!createRow)
                    return null;
                r = _.CreateRow(y - 1);
            }
            return new Row(r, this);
        }

        public IEnumerable<Row> GetRows(RowScope rowScope = RowScope.IncludeNull)
        {
            return GetRowsInRange(rowScope);
        }

        public IEnumerable<Row> GetRowsInRange(RowScope rowScope = RowScope.IncludeNull, int y1 = 1, int? y2 = null)
        {
            if (y2 == null)
                y2 = _.LastRowNum + 1;
            //var rows = Sheet.GetRowEnumerator();//!!!buggy: sometimes misses added rows
            for (int i = y1 - 1; i < y2; i++)
            {
                var r = _.GetRow(i);
                if (r == null)
                {
                    if (rowScope == RowScope.OnlyExisting)
                        continue;
                    if (rowScope == RowScope.CreateIfNull)
                        r = _.CreateRow(i);
                }
                if (r != null)
                    yield return new Row(r, this);
            }
        }

        public Row AppendRow<T>(IEnumerable<T> values)
        {
            int y0 = _.LastRowNum;//(!)it is 0 when no row or 1 row
            int y = y0 + (y0 == 0 && _.GetRow(y0) == null ? 1 : 2);
            return WriteRow(y, values);
        }

        public Row AppendRow<T>(params T[] values)
        {
            return AppendRow(values);
        }

        public Row InsertRow(int y)
        {
            _.ShiftRows(y - 1, _.LastRowNum, 1);
            return new Row(_.GetRow(y - 1), this);
        }

        public Row InsertRow<T>(int y, IEnumerable<T> values = null)
        {
            if (y <= _.LastRowNum)
                _.ShiftRows(y - 1, _.LastRowNum, 1);
            return WriteRow(y, values);
        }

        public Row InsertRow<T>(int y, params T[] values)
        {
            return InsertRow(y, (IEnumerable<T>)values);
        }

        public Row WriteRow<T>(int y, IEnumerable<T> values)
        {
            Row r = GetRow(y, true);
            r.Write(values);
            return r;
        }

        public Row WriteRow<T>(int y, params T[] values)
        {
            return WriteRow(y, (IEnumerable<T>)values);
        }

        public void ShiftRowCellsRight(int y, int x1, int shift, Action<Cell> onFormulaCellMoved = null)
        {
            GetRow(y, false)?.ShiftCellsRight(x1, shift, onFormulaCellMoved);
        }

        public void ShiftRowCellsLeft(int y, int x1, int shift, Action<Cell> onFormulaCellMoved = null)
        {
            GetRow(y, false)?.ShiftCellsLeft(x1, shift, onFormulaCellMoved);
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

        public void AutosizeRowsInRange(int y1 = 1, int? y2 = null)
        {
            GetRowsInRange(RowScope.OnlyExisting, y1, y2).ForEach(a => a._.Height = -1);
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
                _.RemoveRow(r._);
        }

        public void ClearMergingInRow(int y)
        {
            NewRange(y, 1, y, null).ClearMerging();
        }

        public int GetLastRow(bool includeMerged = true)
        {
            Row row = GetRow(_.LastRowNum, false);
            if (row == null)
                return 0;
            if (!includeMerged)
                return row.Y;
            int maxY = 0;
            foreach (var c in row.GetCells())
            {
                var r = c.GetMergedRange();
                if (r != null && maxY < r.Y2.Value)
                    maxY = r.Y2.Value;
            }
            return maxY;
        }

        public int GetLastColumnInRowRange(int y1 = 1, int? y2 = null, bool includeMerged = true)
        {
            return GetRowsInRange(RowScope.OnlyExisting, y1, y2).Max(a => a.GetLastColumn(includeMerged));
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="y"></param>
        /// <param name="includeMerged"></param>
        /// <returns>1-based, otherwise 0</returns>
        public int GetLastNotEmptyColumnInRow(int y, bool includeMerged = true)
        {
            return GetRow(y, false).GetLastNotEmptyColumn(includeMerged);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="y"></param>
        /// <param name="includeMerged"></param>
        /// <returns>1-based, otherwise 0</returns>
        public int GetLastColumnInRow(int y, bool includeMerged = true)
        {
            return GetRow(y, false).GetLastColumn(includeMerged);
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
                y2 = _.LastRowNum + 1;
            return GetRowsInRange(RowScope.OnlyExisting, y1, y2).Max(a => a.GetLastNotEmptyColumn(includeMerged));
        }
    }
}