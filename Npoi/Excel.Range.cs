//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************
using NPOI.SS.UserModel;
using NPOI.SS.Util;

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

            public Range(ISheet sheet, int y1 = 1, int x1 = 1, int? y2 = null, int? x2 = null)
            {
                Sheet = sheet;
                Y1 = y1;
                Y2 = y2;
                X1 = x1;
                X2 = x2;
            }

            public ISheet Sheet;

            public ICell GetMainCell(bool createCell)
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

            ///// <summary>
            ///// (!)When createOnlyUniqueStyles, it is slower. Otherwise, each call registers a new style for non-styled cells.
            ///// </summary>
            //public void Highlight(Range range, bool createOnlyUniqueStyles, Color color, FillPattern fillPattern = FillPattern.SolidForeground)
            //{
            //    ICellStyle newStyle = null;
            //    int maxY = Y2 != null ? Y2.Value : Sheet.LastRowNum + 1;
            //    for (int y = Y1; y <= maxY; y++)
            //    {
            //        IRow row = GetRow(y, color != null);
            //        if (row == null)
            //            continue;
            //        int maxX = X2 != null ? X2.Value : row.LastCellNum;
            //        for (int x = X1; x <= maxX; x++)
            //        {
            //            ICell c = row.GetCell(x, true);
            //            if (c.CellStyle == null)
            //            {
            //                if (color != null)
            //                {
            //                    if (newStyle == null)
            //                        newStyle = highlight(this, null, createOnlyUniqueStyles, color, fillPattern);
            //                    c.CellStyle = newStyle;
            //                }
            //            }
            //            c.CellStyle = highlight(this, c.CellStyle, createOnlyUniqueStyles, color, fillPattern);
            //        }
            //    }
            //}

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

            public void Move(int toY, int toX, OnFormulaCellMoved onFormulaCellMoved = null, ISheet toSheet = null)
            {
                PasteRange(Cut(), toY, toX, onFormulaCellMoved, toSheet);
            }

            public void Copy(int toY, int toX, OnFormulaCellMoved onFormulaCellMoved = null, ISheet toSheet = null)
            {
                PasteRange(Copy(), toY, toX, onFormulaCellMoved, toSheet);
            }
        }
    }
}