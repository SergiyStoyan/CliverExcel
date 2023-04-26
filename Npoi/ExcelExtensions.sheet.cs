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
using NPOI.HSSF.UserModel;

namespace Cliver
{
    static public partial class ExcelExtensions
    {
        /// <summary>
        /// Removes empty rows.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="includeEmptyCellRows">might considerably slow down if TRUE</param>
        /// <param name="shiftRemainingRows">the remaining rows are shifted if TRUE</param>
        public static void _RemoveEmptyRows(this ISheet sheet, bool includeEmptyCellRows, bool shiftRemainingRows)
        {
            int removedRowsCount = 0;
            for (int i = sheet.LastRowNum; i >= 0; i--)
            {
                var r = sheet.GetRow(i);
                if (r != null)
                {
                    if (r.LastCellNum >= 0
                        && (!includeEmptyCellRows || r._GetLastNotEmptyColumn(false) > 0)
                        )
                    {
                        if (removedRowsCount > 0)
                        {
                            if (r.RowNum + 1 + removedRowsCount <= sheet.LastRowNum)
                                sheet.ShiftRows(r.RowNum + 1 + removedRowsCount, sheet.LastRowNum, -removedRowsCount);
                            removedRowsCount = 0;
                        }
                        continue;
                    }
                    sheet.RemoveRow(r);
                }
                if (shiftRemainingRows)
                    removedRowsCount++;
            }
            if (removedRowsCount > 0
                && 1 + removedRowsCount <= sheet.LastRowNum
                )
                sheet.ShiftRows(1 + removedRowsCount, sheet.LastRowNum, -removedRowsCount);
        }

        ///// <summary>
        ///// Safe.
        ///// </summary>
        ///// <param name="sheet"></param>
        ///// <param name="y1"></param>
        ///// <param name="y2"></param>
        ///// <param name="shift"></param>
        ///// <returns></returns>
        //static public IEnumerable<IRow> _ShiftRows(this ISheet sheet, int y1, int y2, int shift)
        //{
        //    if (1 + shift <= sheet.LastRowNum)
        //        sheet.ShiftRows(1 + shift, sheet.LastRowNum, -shift);
        //}

        static public IEnumerable<IRow> _GetRows(this ISheet sheet, RowScope rowScope)
        {
            return sheet._GetRowsInRange(rowScope);
        }

        static public int _GetLastRow(this ISheet sheet, LastRowCondition lastRowCondition, bool includeMerged)
        {
            IRow row = null;
            switch (lastRowCondition)
            {
                case LastRowCondition.NotEmpty:
                    return sheet._GetLastNotEmptyRow(includeMerged);
                case LastRowCondition.HasCells:
                    for (int i = sheet.LastRowNum; i >= 0; i--)
                    {
                        row = sheet.GetRow(i);
                        if (row == null)
                            continue;
                        if (row.LastCellNum >= 0)
                            break;
                    }
                    break;
                case LastRowCondition.NotNull:
                    row = sheet.GetRow(sheet.LastRowNum);
                    break;
                default:
                    throw new Exception("Unknown option: " + lastRowCondition.ToString());
            }
            if (row == null)
                return 0;
            if (!includeMerged)
                return row._Y();
            int maxY = 0;
            foreach (var c in row.Cells)
            {
                var r = c._GetMergedRange();
                if (r != null && maxY < r.Y2.Value)
                    maxY = r.Y2.Value;
            }
            return maxY;
        }

        static public void _ReplaceStyle(this ISheet sheet, ICellStyle style1, ICellStyle style2)
        {
            new Range(sheet).ReplaceStyle(style1, style2);
        }

        static public void _SetStyle(this ISheet sheet, ICellStyle style, bool createCells)
        {
            new Range(sheet).SetStyle(style, createCells);
        }

        static public void _UnsetStyle(this ISheet sheet, ICellStyle style)
        {
            new Range(sheet).UnsetStyle(style);
        }

        static public IEnumerable<Image> _GetImages(this ISheet sheet, int y, int x)
        {
            if (sheet.Workbook is XSSFWorkbook xSSFWorkbook)
            {
                XSSFDrawing dp = (XSSFDrawing)sheet.CreateDrawingPatriarch();
                foreach (XSSFShape s in dp.GetShapes())
                {
                    XSSFPicture p = s as XSSFPicture;
                    if (p == null)
                        continue;
                    var a = p.ClientAnchor;
                    if (y - 1 >= a.Row1 && y - 1 <= a.Row2 && x - 1 >= a.Col1 && x - 1 <= a.Col2)
                    {
                        IPictureData pictureData = p.PictureData;
                        yield return new Image { Data = pictureData.Data, Name = null, Type = pictureData.PictureType, X = a.Col1, Y = a.Row1/*, Anchor = a*/ };
                    }
                }
            }
            else if (sheet.Workbook is HSSFWorkbook hSSFWorkbook)
            {
                HSSFPatriarch g;

            }
            else
                throw new Exception("Unsupported workbook type: " + sheet.Workbook.GetType().FullName);
        }

        static public Range _NewRange(this ISheet sheet, int y1 = 1, int x1 = 1, int? y2 = null, int? x2 = null)
        {
            return new Range(sheet, y1, x1, y2, x2);
        }

        static public Range _GetMergedRange(this ISheet sheet, int y, int x)
        {
            return sheet._getMergedRange(y, x);
        }
    }
}