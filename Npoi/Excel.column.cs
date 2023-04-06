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
using NPOI.SS.Formula.Functions;
using static Cliver.Excel;

namespace Cliver
{
    public partial class Excel
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="header"></param>
        /// <param name="headerY"></param>
        /// <returns>1-based, otherwise 0</returns>
        public int FindColumnByHeader(Regex header, int headerY = 1)
        {
            IRow row = GetRow(headerY, false);
            if (row == null)
                return 0;
            for (int x = 1; x <= row.Cells.Count; x++)
                if (header.IsMatch(GetValueAsString(headerY, x, false)))
                    return x;
            return 0;
        }

        public void ShiftColumns(int x, int shift, Action<ICell> onFormulaCellMoved = null)
        {
            Dictionary<int, int> columnXs2width = new Dictionary<int, int>();
            int lastColumnX = x;
            columnXs2width[lastColumnX] = Sheet.GetColumnWidth(lastColumnX - 1);
            //var rows = Sheet.GetRowEnumerator();//!!!buggy: sometimes misses added rows
            //while (rows.MoveNext())
            for (int y0 = Sheet.LastRowNum; y0 >= 0; y0--)
            {
                IRow row = Sheet.GetRow(y0);
                if (row == null)
                    continue;
                int columnX = row.GetLastColumn(true);
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
                SetColumnWidth(columnX + shift, columnXs2width[columnX]);
        }

        public void ShiftColumns(IRow row, int x, int shift, Action<ICell> onFormulaCellMoved = null)
        {
            for (int i = row.GetLastColumn(true); i >= x; i--)
                MoveCell(row.RowNum + 1, i, row.RowNum + 1, i + shift, onFormulaCellMoved);
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
            return row.GetLastNotEmptyColumn(includeMerged);
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
            return row.GetLastColumn(includeMerged);
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
            return GetRowsInRange(false, y1, y2).Max(a => a.GetLastNotEmptyColumn(includeMerged));
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="includeMerged"></param>
        /// <returns>1-based, otherwise 0</returns>
        public int GetLastNotEmptyColumn(bool includeMerged)
        {
            return GetLastNotEmptyColumnInRowRange(1, null, includeMerged);
        }

        public void CopyColumn(string columnName, ISheet destinationSheet)
        {
            int x = CellReference.ConvertColStringToIndex(columnName);
            CopyColumn(x, destinationSheet);
        }

        public void CopyColumn(int x, ISheet destinationSheet)
        {
            var range = new Range(1, x, null, x);
            CopyRange(range, destinationSheet);
        }

        /// <summary>
        /// (!)Very slow on large data.
        /// </summary>
        /// <param name="columnIs"></param>
        /// <param name="padding">a character width</param>
        public void AutosizeColumns(IEnumerable<int> Xs, float padding = 0)
        {
            foreach (int y in Xs)
                AutosizeColumn(y, padding);
        }

        /// <summary>
        /// (!)Very slow on large data.
        /// </summary>
        /// <param name="x"></param>
        /// <param name="padding">a character width</param>
        public void AutosizeColumn(int x, float padding = 0)
        {
            Sheet.AutoSizeColumn(x - 1, false);

            //GetCellsInColumn(x).Max(a => a.GetValueAsString())
            //int width = ((int)(maxNumCharacters * 1.14388)) * 256;
            //sheet.setColumnWidth(i, width);

            if (padding > 0)
                SetColumnWidth(x, Sheet.GetColumnWidth(x - 1) + (int)(padding * 256));
        }

        public IEnumerable<ICell> GetCellsInColumn(int x)
        {
            return GetRows().Select(a => a.GetCell(x));
        }

        /// <summary>
        /// Safe against the API's one
        /// </summary>
        /// <param name="x"></param>
        /// <param name="width">units of 1/256th of a character width</param>
        public void SetColumnWidth(int x, int width)
        {
            const int cellMaxWidth = 256 * 255;
            int w = MathRoutines.Truncate(width, cellMaxWidth);
            Sheet.SetColumnWidth(x - 1, w);
        }

        /// <summary>
        /// Safe against the API's one
        /// </summary>
        /// <param name="x"></param>
        /// <param name="width">a character width</param>
        public void SetColumnWidth(int x, float width)
        {
            SetColumnWidth(x, (int)(width * 255));
        }

        /// <summary>
        /// (!)Very slow on large data.
        /// </summary>
        /// <param name="x1"></param>
        /// <param name="x2"></param>
        /// <param name="padding">a character width</param>
        public void AutosizeColumnsInRange(int x1 = 1, int? x2 = null, float padding = 0)
        {
            if (x2 == null)
                x2 = GetLastColumn();
            for (int x = x1; x <= x2; x++)
                AutosizeColumn(x, padding);
        }

        public int GetLastColumnInRowRange(int y1 = 1, int? y2 = null, bool includeMerged = true)
        {
            return GetRowsInRange(false, y1, y2).Max(a => a.GetLastColumn(includeMerged));
        }

        public int GetLastColumn(bool includeMerged = true)
        {
            return GetLastColumnInRowRange(1, null, includeMerged);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="padding">a character width</param>
        public void AutosizeColumns(float padding = 0)
        {
            AutosizeColumnsInRange(1, null, padding);
        }

        public void ClearColumn(int x, bool clearMerging)
        {
            if (clearMerging)
                ClearMergingInColumn(x);
            //var rows = Sheet.GetRowEnumerator();//!!!buggy: sometimes misses added rows
            //while (rows.MoveNext())
            for (int y0 = Sheet.LastRowNum; y0 >= 0; y0--)
            {
                IRow row = Sheet.GetRow(y0);
                if (row == null)
                    continue;
                ICell c = row.GetCell(x);
                if (c != null)
                    row.RemoveCell(c);
            }
        }

        public void ClearMergingInColumn(int x)
        {
            ClearMerging(new Range(1, x, null, x));
        }

        public void SetStyleInColumn(ICellStyle style, bool createCells, int x)
        {
            SetStyleInColumnRange(style, createCells, x, x);
        }

        public void SetStyleInColumnRange(ICellStyle style, bool createCells, int x1, int? x2 = null)
        {
            SetStyle(new Range(1, x1, null, x2), style, createCells);
        }

        public void ReplaceStyleInColumnRange(ICellStyle style1, ICellStyle style2, int x1, int? x2 = null)
        {
            ReplaceStyle(new Range(1, x1, null, x2), style1, style2);
        }

        public void ClearStyleInColumnRange(ICellStyle style, int x1, int? x2 = null)
        {
            ReplaceStyleInColumnRange(style, null, x1, x2);
        }
    }
}