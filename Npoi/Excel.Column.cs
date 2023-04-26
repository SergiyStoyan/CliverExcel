﻿//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Cliver
{
    public partial class Excel
    {
        public class Column
        {
            internal protected Column(ISheet sheet, int x)
            {
                Sheet = sheet;
                X = x;
            }
            readonly public ISheet Sheet;

            public readonly int X;

            public ICell GetCell(int y, bool createCell)
            {
                return Sheet._GetCell(y, X, createCell);
            }

            public int GetLastRow(bool includeMerged)
            {
                for (int i = Sheet.LastRowNum; i >= 0; i--)
                {
                    IRow row = Sheet.GetRow(i);
                    if (row == null)
                        continue;
                    var c = row.GetCell(X - 1);
                    if (c == null)
                        continue;
                    if (includeMerged)
                    {
                        var r = c._GetMergedRange();
                        if (r != null)
                            return r.Y2.Value;
                    }
                    return c.RowIndex + 1;
                }
                return 0;
            }

            public IEnumerable<ICell> GetCells(bool createCells)
            {
                return GetCellsInRange(createCells);
            }

            public IEnumerable<ICell> GetCellsInRange(bool createCells, int y1 = 1, int? y2 = null)
            {
                if (y2 == null)
                    y2 = GetLastRow(false);
                for (int y = y1; y <= y2; y++)
                    yield return GetCell(y, createCells);
            }

            public void SetStyles(int y1, IEnumerable<ICellStyle> styles)
            {
                SetStyles(y1, styles.ToArray());
            }

            public void SetStyles(int y1, params ICellStyle[] styles)
            {
                var cs = GetCellsInRange(true, y1, styles.Length).ToList();
                for (int i = y1 - 1; i < styles.Length; i++)
                    cs[i].CellStyle = styles[i];
            }

            public void SetStyle(ICellStyle style, bool createCells)
            {
                new Range(Sheet, 1, X, null, X).SetStyle(style, createCells);
            }

            public int GetLastNotEmptyRow(bool includeMerged)
            {
                for (int i = Sheet.LastRowNum; i >= 0; i--)
                {
                    IRow row = Sheet.GetRow(i);
                    if (row == null)
                        continue;
                    var c = row.GetCell(X - 1);
                    if (string.IsNullOrEmpty(c?._GetValueAsString()))
                        continue;
                    if (includeMerged)
                    {
                        var r = c._GetMergedRange();
                        if (r != null)
                            return r.Y2.Value;
                    }
                    return c.RowIndex + 1;
                }
                return 0;
            }

            public void Clear(bool clearMerging)
            {
                if (clearMerging)
                    ClearMerging();
                foreach (var r in Sheet._GetRows(RowScope.ExistingOnly))
                {
                    var c = r.GetCell(X - 1);
                    if (c != null)
                        r.RemoveCell(c);
                }
            }

            public void ClearMerging()
            {
                new Range(Sheet, 1, X, null, X).ClearMerging();
            }

            public void ShiftCellsDown(int y1, int shift, Action<ICell> onFormulaCellMoved = null)
            {
                for (int y = GetLastRow(true); y >= y1; y--)
                    Sheet._MoveCell(y, X, y + shift, X, onFormulaCellMoved);
            }

            public void ShiftCellsUp(int y1, int shift, Action<ICell> onFormulaCellMoved = null)
            {
                int y2 = GetLastRow(true);
                for (int y = y1; y <= y2; y++)
                    Sheet._MoveCell(y, X, y - shift, X, onFormulaCellMoved);
            }

            /// <summary>
            /// Safe against the API's one
            /// </summary>
            /// <param name="x"></param>
            /// <param name="width">units of 1/256th of a character width</param>
            public void SetWidth(int width)
            {
                const int cellMaxWidth = 256 * 255;
                int w = MathRoutines.Truncate(width, cellMaxWidth);
                Sheet.SetColumnWidth(X - 1, w);
            }

            /// <summary>
            /// Safe against the API's one
            /// </summary>
            /// <param name="x"></param>
            /// <param name="width">a character width</param>
            public void SetWidth(float width)
            {
                SetWidth((int)(width * 255));
            }

            public IEnumerable<ICell> GetCells(RowScope rowScope)
            {
                return Sheet._GetRows(rowScope).Select(a => a?.GetCell(X));
            }

            /// <summary>
            /// (!)Very slow on large data.
            /// </summary>
            /// <param name="x"></param>
            /// <param name="padding">a character width</param>
            public void Autosize(float padding = 0)
            {
                Sheet.AutoSizeColumn(X - 1, false);

                //GetCellsInColumn(x).Max(a => a.GetValueAsString())
                //int width = ((int)(maxNumCharacters * 1.14388)) * 256;
                //sheet.setColumnWidth(i, width);

                if (padding > 0)
                    SetWidth(Sheet.GetColumnWidth(X - 1) + (int)(padding * 256));
            }

            public void Copy(ISheet toSheet, string toColumnName = null)
            {
                int toX = toColumnName == null ? X : CellReference.ConvertColStringToIndex(toColumnName);
                Copy(toSheet, toX);
            }

            public void Copy(ISheet toSheet, int toX)
            {
                new Range(Sheet, 1, X, null, X).Copy(1, toX, toSheet);
            }
        }
    }
}