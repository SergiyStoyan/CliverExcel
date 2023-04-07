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
        /// <summary>
        /// 
        /// </summary>
        /// <param name="includeMerged"></param>
        /// <returns>1-based, otherwise 0</returns>
        public int GetLastNotEmptyRow(bool includeMerged = true)
        {
            return GetLastNotEmptyRowInColumnRange(1, null, includeMerged);
        }

        public IRow GetRow(int y, bool create)
        {
            IRow r = Sheet.GetRow(y - 1);
            if (r == null && create)
                r = Sheet.CreateRow(y - 1);
            return r;
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
            for (int i = Sheet.LastRowNum; i >= 0; i--)
            {
                IRow row = Sheet.GetRow(i);
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
        /// <param name="includeMerged"></param>
        /// <param name="xs"></param>
        /// <returns>1-based, otherwise 0</returns>
        public int GetLastNotEmptyRowInColumns(bool includeMerged, params int[] xs)
        {
            for (int i = Sheet.LastRowNum; i >= 0; i--)
            {
                IRow row = Sheet.GetRow(i);
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
        /// <param name="x"></param>
        /// <param name="includeMerged"></param>
        /// <returns>1-based, otherwise 0</returns>
        public int GetLastNotEmptyRowInColumn(int x, bool includeMerged = true)
        {
            for (int i = Sheet.LastRowNum; i >= 0; i--)
            {
                IRow row = Sheet.GetRow(i);
                if (row == null)
                    continue;
                var c = row.GetCell(x - 1);
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
            for (int i = Sheet.LastRowNum; i >= 0; i--)
            {
                IRow row = Sheet.GetRow(i);
                if (row == null)
                    continue;
                var c = row.GetCell(x - 1);
                if (c == null)
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

        public int GetLastRow(bool includeMerged = true)
        {
            IRow row = Sheet.GetRow(Sheet.LastRowNum);
            if (row == null)
                return 0;
            if (!includeMerged)
                return row.Y();
            int maxY = 0;
            foreach (var c in row.Cells)
            {
                var r = c.GetMergedRange();
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
                ClearMergingForRow(y);
            var r = GetRow(y, false);
            if (r != null)
                Sheet.RemoveRow(r);
        }

        public void ClearMergingForRow(int y)
        {
            Range r = new Range(y, 1, y, null);
            ClearMerging(r);
        }

        public enum RowScope
        {
            OnlyExisting,
            IncludeNull,
            CreateIfNull
        }
        public IEnumerable<IRow> GetRowsInRange(RowScope rowScope = RowScope.IncludeNull, int y1 = 1, int? y2 = null)
        {
            if (y2 == null)
                y2 = Sheet.LastRowNum + 1;
            //var rows = Sheet.GetRowEnumerator();//!!!buggy: sometimes misses added rows
            for (int i = y1 - 1; i < y2; i++)
            {
                var r = Sheet.GetRow(i);
                if (r == null)
                {
                    if (rowScope == RowScope.OnlyExisting)
                        continue;
                    if (rowScope == RowScope.CreateIfNull)
                        r = Sheet.CreateRow(i);
                }
                if (r != null)
                    yield return r;
            }
        }

        public void SetStyleInRow(ICellStyle style, bool createCells, int y)
        {
            SetStyleInRowRange(style, createCells, y, y);
        }

        public void SetStyleInRowRange(ICellStyle style, bool createCells, int y1, int? y2 = null)
        {
            SetStyle(new Range(y1, 1, y2, null), style, createCells);
        }

        public void ReplaceStyleInRowRange(ICellStyle style1, ICellStyle style2, int y1, int? y2 = null)
        {
            ReplaceStyle(new Range(y1, 1, y2, null), style1, style2);
        }

        public void ClearStyleInRowRange(ICellStyle style, int y1, int? y2 = null)
        {
            ReplaceStyleInRowRange(style, null, y1, y2);
        }

        public IEnumerable<IRow> GetRows(RowScope rowScope = RowScope.IncludeNull)
        {
            return GetRowsInRange(rowScope);
        }

        public IRow AppendRow<T>(IEnumerable<T> values)
        {
            int y0 = Sheet.LastRowNum;//(!)it is 0 when no row or 1 row
            int y = y0 + (y0 == 0 && Sheet.GetRow(y0) == null ? 1 : 2);
            return WriteRow(y, values);
        }

        public IRow AppendRow<T>(params T[] values)
        {
            return AppendRow(values);
        }

        public IRow InsertRow<T>(int y, IEnumerable<T> values = null)
        {
            if (y <= Sheet.LastRowNum)
                Sheet.ShiftRows(y - 1, Sheet.LastRowNum, 1);
            return WriteRow(y, values);
        }

        public IRow InsertRow<T>(params T[] values)
        {
            return InsertRow((IEnumerable<T>)values);
        }

        public IRow WriteRow<T>(int y, IEnumerable<T> values)
        {
            IRow r = GetRow(y, true);
            r.Write(values);
            return r;
        }

        public IRow WriteRow<T>(int y, params T[] values)
        {
            return WriteRow(y, (IEnumerable<T>)values);
        }
    }
}