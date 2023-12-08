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
using NPOI.Util;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.PTG;
using NPOI.SS.Formula;

namespace Cliver
{
    static public partial class ExcelExtensions
    {
        static public void _SetComment(this ISheet sheet, int y, int x, string comment, string author = null, IClientAnchor anchor = null)
        {
            sheet._GetCell(y, x, true)._SetComment(comment, author, anchor);
        }

        static public void _AppendOrSetComment(this ISheet sheet, int y, int x, string comment, string author = null, string separator = "\r\n\r\n", IClientAnchor anchor = null)
        {
            sheet._GetCell(y, x, true)._AppendOrSetComment(comment, author, separator, anchor);
        }

        static public void _SetLink(this ISheet sheet, int y, int x, string link)
        {
            sheet._GetCell(y, x, true)._SetLink(link);
        }

        static public string _GetLink(this ISheet sheet, int y, int x)
        {
            return sheet?._GetCell(y, x, false)?._GetLink();
        }

        static public void _ShiftCellsRight(this ISheet sheet, int x1, int y1, int y2, int shift, OnFormulaCellMoved onFormulaCellMoved = null)
        {
            for (int y = y1; y <= y2; y++)
                sheet._GetRow(y, false)?._ShiftCellsRight(x1, shift, onFormulaCellMoved);
        }

        static public void _ShiftCellsLeft(this ISheet sheet, int x1, int y1, int y2, int shift, OnFormulaCellMoved onFormulaCellMoved = null)
        {
            for (int y = y1; y <= y2; y++)
                sheet._GetRow(y, false)?._ShiftCellsLeft(x1, shift, onFormulaCellMoved);
        }

        static public void _ShiftCellsDown(this ISheet sheet, int y1, int x1, int x2, int shift, OnFormulaCellMoved onFormulaCellMoved = null)
        {
            for (int x = x1; x <= x2; x++)
                sheet._GetColumn(x)?.ShiftCellsDown(y1, shift, onFormulaCellMoved);
        }

        static public void _ShiftCellsUp(this ISheet sheet, int y1, int x1, int x2, int shift, OnFormulaCellMoved onFormulaCellMoved = null)
        {
            for (int x = x1; x <= x2; x++)
                sheet._GetColumn(x)?.ShiftCellsUp(y1, shift, onFormulaCellMoved);
        }

        static public ICell _CopyCell(this ISheet sheet, int fromCellY, int fromCellX, int toCellY, int toCellX, OnFormulaCellMoved onFormulaCellMoved = null, ISheet toSheet = null, StyleMap toStyleMap = null)
        {
            ICell sourceCell = sheet._GetCell(fromCellY, fromCellX, false);
            return sourceCell._Copy(toCellY, toCellX, onFormulaCellMoved, toSheet, toStyleMap);
        }

        static public string _GetValueAsString(this ISheet sheet, int y, int x, bool allowNull = false)
        {
            ICell c = sheet._GetCell(y, x, false);
            if (c == null)
                return allowNull ? null : string.Empty;
            return c?._GetValueAsString(allowNull);
        }

        static public string _GetValueAsString(this ISheet sheet, string cellAddress, bool allowNull = false)
        {
            ICell c = sheet._GetCell(cellAddress, false);
            if (c == null)
                return allowNull ? null : string.Empty;
            return c?._GetValueAsString(allowNull);
        }

        static public object _GetValue(this ISheet sheet, int y, int x)
        {
            ICell c = sheet._GetCell(y, x, false);
            return c?._GetValue();
        }

        static public void _SetValue(this ISheet sheet, int y, int x, object value)
        {
            ICell c = sheet._GetCell(y, x, true);
            c._SetValue(value);
        }

        static public void _SetValue(this ISheet sheet, string cellAddress, object value)
        {
            ICell c = sheet._GetCell(cellAddress, true);
            c._SetValue(value);
        }

        static public ICell _MoveCell(this ISheet sheet, int fromCellY, int fromCellX, int toCellY, int toCellX, OnFormulaCellMoved onFormulaCellMoved = null, ISheet toSheet = null, StyleMap toStyleMap = null)
        {
            ICell fromCell = sheet._GetCell(fromCellY, fromCellX, false);
            return fromCell._Move(toCellY, toCellX, onFormulaCellMoved, toSheet, toStyleMap);
        }

        static public ICell _GetCell(this ISheet sheet, int y, int x, bool createCell)
        {
            IRow r = sheet._GetRow(y, createCell);
            if (r == null)
                return null;
            return r._GetCell(x, createCell);
        }

        static public ICell _GetCell(this ISheet sheet, string cellAddress, bool createCell)
        {
            var cs = GetCoordinates(cellAddress);
            IRow r = sheet._GetRow(cs.Y, createCell);
            if (r == null)
                return null;
            return r._GetCell(cs.X, createCell);
        }

        static public void _RemoveCell(this ISheet sheet, int y, int x)
        {
            sheet._GetCell(y, x, false)?._Remove();
        }

        static public void _UpdateFormulaRange(this ISheet sheet, int y, int x, int rangeY1Shift, int rangeX1Shift, int? rangeY2Shift = null, int? rangeX2Shift = null)
        {
            sheet._GetCell(y, x, false)?._UpdateFormulaRange(rangeY1Shift, rangeX1Shift, rangeY2Shift, rangeX2Shift);
        }

        static public void _ClearMerging(this ISheet sheet, int y, int x)
        {
            sheet._GetCell(y, x, false)?._ClearMerging();
        }

        static public void _CreateDropdown<T>(this ISheet sheet, int y, int x, IEnumerable<T> values, T value, bool allowBlank = true)
        {
            sheet._GetCell(y, x, true)._CreateDropdown(values, value, allowBlank);
        }

        /// <summary>
        /// !!!sizing seems to work not correctly when Image is obtained from Tesseract (check sizing of the input bitmap?)
        /// </summary>
        /// <exception cref="Exception"></exception>
        static public void _AddImage(this ISheet sheet, Image image)
        {
            int imageId = sheet.Workbook.AddPicture(image.Data, image.Type);
            IDrawing d = sheet.CreateDrawingPatriarch();
            IClientAnchor a = d.CreateAnchor(0, 0, 0, 0, image.X - 1, image.Y - 1, image.X - 1, image.Y - 1);
            a.AnchorType = AnchorType.MoveDontResize;
            IPicture p = d.CreatePicture(a, imageId);
            p.Resize(1);
            //p.Resize(1, 1);
        }

        static public Range _GetMergedRange(this ISheet sheet, int y, int x)
        {
            foreach (var mr in sheet.MergedRegions)
                if (mr.IsInRange(y - 1, x - 1))
                    return new Range(sheet, mr.FirstRow + 1, mr.FirstColumn + 1, mr.LastRow + 1, mr.LastColumn + 1);
            return null;
        }

        /// <summary>        
        /// Images anchored in the specified cell coordinates. The cell may possibly not exist.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="y"></param>
        /// <param name="x"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
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
                //HSSFPatriarch g;
                throw new Exception("TBD for: " + sheet.Workbook.GetType().FullName);
            }
            else
                throw new Exception("Unsupported workbook type: " + sheet.Workbook.GetType().FullName);
        }

    }
}