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

//works  
namespace Cliver
{
    public partial class Excel
    {
        public int FindColumnByHeader(Regex header, int headerY = 1)
        {
            for (int x = GetLastUsedColumnInRow(headerY, false); x > 0; x--)
                if (header.IsMatch(this[headerY, x]))
                    return x;
            return -1;
        }

        public void ShiftColumns(int x, int shift, Action<ICell> onFormulaCellMoved = null)
        {
            Dictionary<int, int> columnXs2width = new Dictionary<int, int>();
            int lastColumnX = x;
            columnXs2width[lastColumnX] = Sheet.GetColumnWidth(lastColumnX - 1);
            var rows = Sheet.GetRowEnumerator();
            while (rows.MoveNext())
            {
                IRow row = (IRow)rows.Current;
                int columnX = row.GetLastUsedColumnInRow(true);
                if (lastColumnX < columnX)
                {
                    for (int i = lastColumnX; i < columnX; i++)
                        columnXs2width[i + 1] = Sheet.GetColumnWidth(i);
                    lastColumnX = columnX;
                }
                for (int i = columnX; i >= x; i--)
                    MoveCell(row.RowNum + 1, i, row.RowNum + 1, i + shift, onFormulaCellMoved);
            }
            foreach (int columnX in columnXs2width.Keys.OrderByDescending(a => a))
                Sheet.SetColumnWidth(columnX + shift - 1, columnXs2width[columnX]);
        }

        public void ShiftColumns(IRow row, int x, int shift, Action<ICell> onFormulaCellMoved = null)
        {
            for (int i = row.GetLastUsedColumnInRow(true); i >= x; i--)
                MoveCell(row.RowNum + 1, i, row.RowNum + 1, i + shift, onFormulaCellMoved);
        }

        public int GetLastUsedColumnInRow(int y, bool includeMerged = true)
        {
            IRow row = GetRow(y, false);
            if (row == null)
                return -1;
            return row.GetLastUsedColumnInRow(includeMerged);
        }

        public int GetLastUsedColumnInRowRange(int y1 = 1, int? y2 = null, bool includeMerged = true)
        {
            var rows = Sheet.GetRowEnumerator();
            int luc = -2;
            while (rows.MoveNext())
            {
                IRow row = (IRow)rows.Current;
                if (row.RowNum + 1 < y1)
                    continue;
                if (row.RowNum >= y2)
                    break;
                int i = row.GetLastUsedColumnInRow(includeMerged);
                if (luc < i)
                    luc = i;
            }
            return luc + 1;
        }

        public int GetLastUsedColumn(bool includeMerged)
        {
            return GetLastUsedColumnInRowRange(1, null, includeMerged);
        }

        public void CopyColumn(string columnName, ISheet sourceSheet, ISheet destinationSheet)
        {
            int x = CellReference.ConvertColStringToIndex(columnName);
            CopyColumn(x, sourceSheet, destinationSheet);
        }

        public void CopyColumn(int x, ISheet sourceSheet, ISheet destinationSheet)
        {
            var range = new CellRangeAddress(0, sourceSheet.LastRowNum, x - 1, x - 1);
            CopyRange(range, sourceSheet, destinationSheet);
        }

        public void AutosizeColumns(IEnumerable<int> columnIs, int padding = 0)
        {
            foreach (int i in columnIs)
            {
                Sheet.AutoSizeColumn(i - 1);
                if (padding > 0)
                    Sheet.SetColumnWidth(i - 1, Sheet.GetColumnWidth(i - 1) + padding);
            }
        }

        public void AutosizeColumnsInRange(int x1 = 1, int? x2 = null, int padding = 0)
        {
            if (x2 == null)
                x2 = GetLastUsedColumnInRowRange(x1, null, true);
            for (int x0 = x1 - 1; x0 < x2; x0++)
            {
                Sheet.AutoSizeColumn(x0);
                if (padding > 0)
                    Sheet.SetColumnWidth(x0, Sheet.GetColumnWidth(x0) + padding);
            }
        }

        public void AutosizeColumns(int padding = 0)
        {
            AutosizeColumnsInRange(1, null, padding);
        }

        public void ClearColumn(int x, bool clearMerging)
        {
            if (clearMerging)
                ClearMergingForColumn(x);
            var rows = Sheet.GetRowEnumerator();
            while (rows.MoveNext())
            {
                IRow row = (IRow)rows.Current;
                ICell c = row.GetCell(x);
                if (c != null)
                    row.RemoveCell(c);
            }
        }

        public void ClearMergingForColumn(int x)
        {
            Range r = new Range(1, int.MaxValue, x, x);
            ClearMerging(r);
        }
    }
}