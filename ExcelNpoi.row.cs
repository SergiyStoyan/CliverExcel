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
        public int GetLastUsedRow(bool includeMerged = true)
        {
            return GetLastUsedRowInColumnRange(1, null, includeMerged);
        }

        public IRow GetRow(int y, bool create)
        {
            IRow r = Sheet.GetRow(y - 1);
            if (r == null && create)
            {
                r = Sheet.CreateRow(y - 1);
                //ICellStyle cs = Workbook.CreateCellStyle();!!!replace it with GetRegisteredStyle()
                //cs.DataFormat = Workbook.CreateDataFormat().GetFormat("text");
                //r.RowStyle = cs;//!!!Cells must be formatted as text! Otherwise string dates are converted into numbers. (However, if no format set, NPOI presets ' before numeric values to keep them as strings.)
            }
            return r;
        }

        public int GetLastUsedRowInColumnRange(int x1 = 1, int? x2 = null, bool includeMerged = true)
        {
            if (x2 == null)
                x2 = int.MaxValue;
            var rows = Sheet.GetRowEnumerator();
            ICell luc = null;
            while (rows.MoveNext())
            {
                IRow row = (IRow)rows.Current;
                var c = row.Cells.Find(a => a.ColumnIndex + 1 >= x1 && a.ColumnIndex < x2 && !string.IsNullOrEmpty(a.GetValueAsString()));
                if (c != null)
                    luc = c;
            }
            if (luc == null)
                return -1;
            if (includeMerged)
            {
                var r = luc.GetMergedRange();
                if (r != null)
                    return r.LastY;
            }
            return luc.RowIndex + 1;
        }

        public int GetLastUsedRowInColumns(bool includeMerged, params int[] xs)
        {
            var rows = Sheet.GetRowEnumerator();
            ICell luc = null;
            while (rows.MoveNext())
            {
                IRow row = (IRow)rows.Current;
                var c = row.Cells.Find(a => xs.Contains(a.ColumnIndex + 1) && !string.IsNullOrEmpty(a.GetValueAsString()));
                if (c != null)
                    luc = c;
            }
            if (luc == null)
                return -1;
            if (includeMerged)
            {
                var r = luc.GetMergedRange();
                if (r != null)
                    return r.LastY;
            }
            return luc.RowIndex + 1;
        }

        public int GetLastUsedRowInColumn(int x, bool includeMerged = true)
        {
            var rows = Sheet.GetRowEnumerator();
            ICell luc = null;
            while (rows.MoveNext())
            {
                IRow row = (IRow)rows.Current;
                var c = row.GetCell(x);
                if (!string.IsNullOrEmpty(c?.GetValueAsString()))
                    luc = c;
            }
            if (luc == null)
                return -1;
            if (includeMerged)
            {
                var r = luc.GetMergedRange();
                if (r != null)
                    return r.LastY;
            }
            return luc.RowIndex + 1;
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
            var rows = Sheet.GetRowEnumerator();
            while (rows.MoveNext())
            {
                IRow row = (IRow)rows.Current;
                if (row.RowNum + 1 < y1)
                    continue;
                if (row.RowNum >= y2)
                    return;
                row.Height = -1;
            }
        }

        public void AutosizeRows()
        {
            AutosizeRowsInRange();
        }
    }
}