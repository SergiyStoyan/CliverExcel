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
using Newtonsoft.Json.Serialization;
using System.Reflection;
using Newtonsoft.Json;

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
                        return r.Y2;
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
                        return r.Y2;
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
                        return r.Y2;
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
                        return r.Y2;
                }
                return c.RowIndex + 1;
            }
            return 0;
        }

        public void HighlightRow(int y, ICellStyle style, Color color)
        {
            GetRow(y, true).Highlight(style, color);
        }

        public void Highlight(IRow row, ICellStyle style, Color color)
        {
            row.Highlight(style, color);
        }

        public void AutosizeRowsInRange(int y1 = 1, int? y2 = null)
        {
            GetRowsInRange(y1, y2).ForEach(a => a.Height = -1);
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
            Range r = new Range(y, y, 1, int.MaxValue);
            ClearMerging(r);
        }

        public IEnumerable<IRow> GetRowsInRange(int y1 = 1, int? y2 = null)
        {
            if (y2 == null)
                y2 = Sheet.LastRowNum + 1;
            //var rows = Sheet.GetRowEnumerator();//!!!buggy: sometimes misses added rows
            for (int i = y1 - 1; i < y2; i++)
                yield return Sheet.GetRow(i);
        }

        public IEnumerable<IRow> GetRows()
        {
            return GetRowsInRange();
        }

        public IRow AppendRow(IEnumerable<object> values)
        {
            int y = Sheet.LastRowNum + 2;
            return WriteRow(y, values);
        }

        public IRow AppendRow(params object[] values)
        {
            return AppendRow(values);
        }

        public IRow InsertRow(int y, IEnumerable<object> values = null)
        {
            if (y <= Sheet.LastRowNum)
                Sheet.ShiftRows(y - 1, Sheet.LastRowNum, 1);
            return WriteRow(y, values);
        }

        public IRow InsertRow(params object[] values)
        {
            return InsertRow((IEnumerable<object>)values);
        }

        public IRow WriteRow(int y, IEnumerable<object> values)
        {
            IRow r = GetRow(y, true);
            r.WriteRow(values);
            return r;
        }

        public IRow WriteRow(int y, params object[] values)
        {
            return WriteRow(y, (IEnumerable<object>)values);
        }
    }
}