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

//works  
namespace Cliver
{
    public partial class Excel : IDisposable
    {
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

        public int GetLastNotEmptyRowInColumnRange(int x1 = 1, int? x2 = null, bool includeMerged = true)
        {
            if (x2 == null)
                x2 = int.MaxValue;
            //var rows = Sheet.GetRowEnumerator();//!!!buggy: sometimes misses added rows
            //ICell luc = null;
            //while (rows.MoveNext())
            //{
            //    IRow row = (IRow)rows.Current;
            //    var c = row.Cells.Find(a => a.ColumnIndex + 1 >= x1 && a.ColumnIndex < x2 && !string.IsNullOrEmpty(a.GetValueAsString()));
            //    if (c != null)
            //        luc = c;
            //}
            //if (luc == null)
            //    return -1;
            //if (includeMerged)
            //{
            //    var r = luc.GetMergedRange();
            //    if (r != null)
            //        return r.LastY;
            //}
            //return luc.RowIndex + 1;
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
                        return r.LastY;
                }
                return c.RowIndex + 1;
            }
            return -1;
        }

        public int GetLastNotEmptyRowInColumns(bool includeMerged, params int[] xs)
        {
            //var rows = Sheet.GetRowEnumerator();//!!!buggy: sometimes misses added rows
            //ICell luc = null;
            //while (rows.MoveNext())
            //{
            //    IRow row = (IRow)rows.Current;
            //    var c = row.Cells.Find(a => xs.Contains(a.ColumnIndex + 1) && !string.IsNullOrEmpty(a.GetValueAsString()));
            //    if (c != null)
            //        luc = c;
            //}
            //if (luc == null)
            //    return -1;
            //if (includeMerged)
            //{
            //    var r = luc.GetMergedRange();
            //    if (r != null)
            //        return r.LastY;
            //}
            //return luc.RowIndex + 1;
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
                        return r.LastY;
                }
                return c.RowIndex + 1;
            }
            return -1;
        }

        public int GetLastNotEmptyRowInColumn(int x, bool includeMerged = true)
        {
            //ICell luc = null;
            //var rows = Sheet.GetRowEnumerator();//!!!buggy: sometimes misses added rows
            //while (rows.MoveNext())
            //{
            //    IRow row = (IRow)rows.Current;
            //    var c = row.GetCell(x - 1);
            //    if (!string.IsNullOrEmpty(c?.GetValueAsString()))
            //        luc = c;
            //}
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
                        return r.LastY;
                }
                return c.RowIndex + 1;
            }
            return -1;
        }

        public int GetLastRowInColumn(int x, bool includeMerged = true)
        {
            //ICell luc = null;
            //var rows = Sheet.GetRowEnumerator();//!!!buggy: sometimes misses added rows
            //while (rows.MoveNext())
            //{
            //    IRow row = (IRow)rows.Current;
            //    var c = row.GetCell(x - 1);
            //    if (c != null)
            //        luc = c;
            //}
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
                        return r.LastY;
                }
                return c.RowIndex + 1;
            }
            return -1;
        }

        public void HighlightRow(int y, Color color)
        {
            Highlight(GetRow(y, true), color);
        }

        public void Highlight(IRow row, Color color)
        {
            row.RowStyle = highlight(Workbook, row.RowStyle, color);
        }

        public void AutosizeRowsInRange(int y1 = 1, int? y2 = null)
        {
            //var rows = Sheet.GetRowEnumerator();//!!!buggy: sometimes misses added rows
            //while (rows.MoveNext())
            for (int y0 = y1 - 1; y0 < y2; y0++)
            {
                IRow row = Sheet.GetRow(y0);
                if (row == null)
                    continue;
                row.Height = -1;
            }
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
    }
}