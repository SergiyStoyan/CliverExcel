﻿//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System.Collections.Generic;
using System;

namespace Cliver
{
    public partial class Excel
    {
        /// <summary>
        /// (!) 1-based
        /// </summary>
        public class Range
        {
            public int X1 = 1;
            public int? X2 = null;
            public int Y1 = 1;
            public int? Y2 = null;

            internal Range(ISheet sheet, int y1 = 1, int x1 = 1, int? y2 = null, int? x2 = null)
            {
                Sheet = sheet;
                Y1 = y1;
                Y2 = y2;
                X1 = x1;
                X2 = x2;
            }

            public ISheet Sheet;

            public ICell GetFirstCell(bool createCell)
            {
                return Sheet._GetCell(Y1, X1, createCell);
            }

            public string GetStringAddress()
            {
                return CellReference.ConvertNumToColString(X1 - 1) + Y1 + ":" + CellReference.ConvertNumToColString(X2 != null ? X2.Value - 1 : Sheet.Workbook.SpreadsheetVersion.LastColumnIndex) + Y2;
            }

            /// <summary>
            /// 
            /// </summary>
            /// <returns>(!) 0-based</returns>
            public CellRangeAddress GetCellRangeAddress()
            {
                return new CellRangeAddress(Y1 - 1, Y2 != null ? Y2.Value - 1 : Sheet.Workbook.SpreadsheetVersion.MaxRows - 1, X1 - 1, X2 != null ? X2.Value - 1 : Sheet.Workbook.SpreadsheetVersion.LastColumnIndex);
            }

            public void Clear(bool clearMerging, bool removeComment = true)
            {
                if (clearMerging)
                    ClearMerging();

                int maxY = Y2 != null ? Y2.Value : Sheet.LastRowNum + 1;
                for (int y = Y1; y <= maxY; y++)
                {
                    IRow row = Sheet._GetRow(y, false);
                    if (row == null)
                        continue;
                    int maxX = X2 != null ? X2.Value : row.LastCellNum;
                    for (int x = X1; x <= maxX; x++)
                        row._GetCell(x, false)?._Remove(removeComment);
                }
            }

            public void ClearMerging()
            {
                CellRangeAddress cra = GetCellRangeAddress();
                for (int i = Sheet.MergedRegions.Count - 1; i >= 0; i--)
                    if (Sheet.MergedRegions[i].Intersects(cra))
                        Sheet.RemoveMergedRegion(i);
            }

            public void Merge(bool clearOldMerging = false)
            {
                if (clearOldMerging)
                    ClearMerging();
                Sheet.AddMergedRegion(GetCellRangeAddress());
            }

            public void ReplaceStyle(ICellStyle style1, ICellStyle style2)
            {
                int maxY = Y2 != null ? Y2.Value : Sheet.LastRowNum + 1;
                for (int y = Y1; y <= maxY; y++)
                {
                    IRow row = Sheet._GetRow(y, false);
                    if (row == null)
                        continue;
                    if (Y1 == 1 && Y2 == null
                        && row.RowStyle?.Index == style1.Index
                        )
                        row.RowStyle = style2;
                    int maxX = X2 != null ? X2.Value : row.LastCellNum;
                    for (int x = X1; x <= maxX; x++)
                    {
                        ICell c = row._GetCell(x, false);
                        if (c != null && c.CellStyle?.Index == style1.Index)
                            c.CellStyle = style2;
                    }
                }
            }

            public void SetStyle(ICellStyle style, bool createCells)
            {
                int maxY = Y2 != null ? Y2.Value : Sheet.LastRowNum + 1;
                for (int y = Y1; y <= maxY; y++)
                {
                    IRow row = Sheet._GetRow(y, createCells);
                    if (row == null)
                        continue;
                    if (Y1 == 1 && Y2 == null)
                        row.RowStyle = style;
                    int maxX = X2 != null ? X2.Value : row.LastCellNum;
                    for (int x = X1; x <= maxX; x++)
                    {
                        ICell c = row._GetCell(x, createCells);
                        if (c != null)
                            c.CellStyle = style;
                    }
                }
            }

            public void UnsetStyle(ICellStyle style)
            {
                ReplaceStyle(style, null);
            }

            public void SetAlteredStyles<T>(T alterationKey, Excel.StyleCache.AlterStyle<T> alterStyle, CellScope cellScope, bool reuseUnusedStyle = false) where T : Excel.StyleCache.IKey
            {
                foreach (var c in GetCells(cellScope))
                    c?._SetAlteredStyle(alterationKey, alterStyle, reuseUnusedStyle);
            }

            public IEnumerable<ICell> GetCells(CellScope cellScope)
            {
                int maxY = Y2 != null ? Y2.Value : Sheet.LastRowNum + 1;
                int maxX;
                switch (cellScope)
                {
                    case CellScope.NotEmpty:
                        for (int y = Y1; y <= maxY; y++)
                        {
                            var r = Sheet._GetRow(y, false);
                            if (r == null)
                                continue;
                            maxX = X2 < r.LastCellNum ? X2.Value : r.LastCellNum;
                            for (int x = X1; x <= maxX; x++)
                            {
                                var c = r._GetCell(x, false);
                                if (!string.IsNullOrWhiteSpace(c._GetValueAsString()))
                                    yield return c;
                            }
                        }
                        break;
                    case CellScope.NotNull:
                        for (int y = Y1; y <= maxY; y++)
                        {
                            var r = Sheet._GetRow(y, false);
                            if (r == null)
                                continue;
                            maxX = X2 < r.LastCellNum ? X2.Value : r.LastCellNum;
                            for (int x = X1; x <= maxX; x++)
                            {
                                var c = r._GetCell(x, false);
                                if (c != null)
                                    yield return c;
                            }
                        }
                        break;
                    case CellScope.IncludeNull:
                        if (X2 != null)
                            maxX = X2.Value;
                        else
                        {
                            maxX = 0;
                            for (int y = Y1; y <= maxY; y++)
                            {
                                var r = Sheet._GetRow(y, false);
                                if (r != null && maxX < r.LastCellNum)
                                    maxX = r.LastCellNum;
                            }
                        }
                        for (int y = Y1; y <= maxY; y++)
                        {
                            var r = Sheet._GetRow(y, false);
                            for (int x = X1; x <= maxX; x++)
                            {
                                var c = r?._GetCell(x, false);
                                yield return c;
                            }
                        }
                        break;
                    case CellScope.CreateIfNull:
                        if (X2 != null)
                            maxX = X2.Value;
                        else
                        {
                            maxX = 0;
                            for (int y = Y1; y <= maxY; y++)
                            {
                                var r = Sheet._GetRow(y, false);
                                if (r != null && maxX < r.LastCellNum)
                                    maxX = r.LastCellNum;
                            }
                        }
                        for (int y = Y1; y <= maxY; y++)
                        {
                            var r = Sheet._GetRow(y, true);
                            for (int x = X1; x <= maxX; x++)
                            {
                                var c = r._GetCell(x, true);
                                yield return c;
                            }
                        }
                        break;
                    default:
                        throw new Exception("Unknown option: " + cellScope.ToString());
                }
            }

            ICell[][] copyCutRange(bool cut)
            {
                int maxY = Y2 != null ? Y2.Value : Sheet.LastRowNum + 1;
                ICell[][] rangeCells = new ICell[maxY - Y1 + 1][];
                for (int y = Y1; y <= maxY; y++)
                {
                    IRow row = Sheet.GetRow(y - 1);
                    if (row == null)
                        continue;
                    int maxX = X2 != null ? X2.Value : row.LastCellNum;
                    ICell[] rowCells = new ICell[maxX];
                    for (int x = X1; x <= maxX; x++)
                    {
                        ICell cell = row.GetCell(x - 1);
                        rowCells[x - X1] = cell;
                        if (cut)
                            row.RemoveCell(cell);
                    }
                    if (cut && X1 == 1 && X2 == null)
                        Sheet.RemoveRow(row);
                    rangeCells[y - Y1] = rowCells;
                }
                return rangeCells;
            }

            public ICell[][] Copy()
            {
                return copyCutRange(false);
            }

            public ICell[][] Cut()
            {
                return copyCutRange(true);
            }

            public void Move(int toY, int toX, Excel.CopyCellMode copyCellMode = null)
            {
                PasteRange(Cut(), toY, toX, copyCellMode);
            }

            public void Copy(int toY, int toX, Excel.CopyCellMode copyCellMode = null)
            {
                PasteRange(Copy(), toY, toX, copyCellMode);
            }

            public void SetComment(string comment, string author = null)
            {
                throw new System.Exception("TBD");
                //if (comment == null)
                //{
                //    int maxY = Y2 != null ? Y2.Value : Sheet.LastRowNum + 1;
                //    for (int y = Y1; y <= maxY; y++)
                //    {
                //        IRow row = Sheet._GetRow(y, false);
                //        if (row == null)
                //            continue;
                //        int maxX = X2 != null ? X2.Value : row.LastCellNum;
                //        for (int x = X1; x <= maxX; x++)
                //            row._GetCell(x, false)?.RemoveCellComment();
                //    }
                //}
                //else
                //{
                //    var creationHelper = Sheet.Workbook.GetCreationHelper();
                //    var richTextString = creationHelper.CreateRichTextString(comment);
                //    var clientAnchor = creationHelper.CreateClientAnchor();
                //    //clientAnchor.Col1 = cell.ColumnIndex + 1;
                //    //clientAnchor.Col2 = cell.ColumnIndex + 3;
                //    //clientAnchor.Row1 = cell.RowIndex + 1;
                //    //clientAnchor.Row2 = cell.RowIndex + 5;
                //    var drawingPatriarch = Sheet.CreateDrawingPatriarch();

                //    int maxY = Y2 != null ? Y2.Value : Sheet.LastRowNum + 1;
                //    for (int y = Y1; y <= maxY; y++)
                //    {
                //        IRow row = Sheet._GetRow(y, false);
                //        if (row == null)
                //            continue;
                //        int maxX = X2 != null ? X2.Value : row.LastCellNum;
                //        for (int x = X1; x <= maxX; x++)
                //        {
                //            ICell cell = row._GetCell(x, true);
                //            IComment iComment = drawingPatriarch.CreateCellComment(clientAnchor);
                //            iComment.String = richTextString;
                //            if (!string.IsNullOrWhiteSpace(author))
                //                iComment.Author = author;
                //            cell.CellComment = iComment;
                //        }
                //    }
                //}
            }
        }
    }
}