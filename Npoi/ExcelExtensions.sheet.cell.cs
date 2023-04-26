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

namespace Cliver
{
    static public partial class ExcelExtensions
    {
        static public void _SetLink(this ISheet sheet, int y, int x, Uri uri)
        {
            sheet._GetCell(y, x, true)._SetLink(uri);
        }

        static public Uri _GetLink(this ISheet sheet, int y, int x)
        {
            return sheet._GetCell(y, x, false)?._GetLink();
        }

        static public void _ShiftCellsRight(this ISheet sheet, int x1, int y1, int y2, int shift, Action<ICell> onFormulaCellMoved = null)
        {
            for (int y = y1; y <= y2; y++)
            {
                for (int x = sheet._GetLastNotEmptyColumnInRow(false, y); x >= x1; x--)
                    sheet._MoveCell(y, x, y, x + shift, onFormulaCellMoved);
                sheet._GetCell(y, x1, false)?.SetBlank();
            }
        }

        static public void _ShiftCellsLeft(this ISheet sheet, int x1, int y1, int y2, int shift, Action<ICell> onFormulaCellMoved = null)
        {
            for (int y = y1; y <= y2; y++)
            {
                for (int x = 1; x <= x1; x++)
                    sheet._MoveCell(y, x, y, x - shift, onFormulaCellMoved);
                sheet._GetCell(y, x1, false)?.SetBlank();
            }
        }

        static public void _ShiftCellsDown(this ISheet sheet, int y1, int x1, int x2, int shift, Action<ICell> onFormulaCellMoved = null)
        {
            for (int x = x1; x <= x2; x++)
            {
                for (int y = sheet._GetLastNotEmptyRowInColumn(false, x); y >= y1; y--)
                    sheet._MoveCell(y, x, y + shift, x, onFormulaCellMoved);
                sheet._GetCell(y1, x, false)?.SetBlank();
            }
        }

        static public void _ShiftCellsUp(this ISheet sheet, int y1, int x1, int x2, int shift, Action<ICell> onFormulaCellMoved = null)
        {
            for (int x = x1; x <= x2; x++)
            {
                for (int y = 1; y <= y1; y++)
                    sheet._MoveCell(y, x, y - shift, x, onFormulaCellMoved);
                sheet._GetCell(y1, x, false)?.SetBlank();
            }
        }

        static public void _CopyCell(this ISheet sheet, int fromCellY, int fromCellX, int toCellY, int toCellX)
        {
            ICell sourceCell = sheet._GetCell(fromCellY, fromCellX, false);
            sourceCell._Copy(toCellY, toCellX);
        }

        static public string _GetValueAsString(this ISheet sheet, int y, int x, bool allowNull = false)
        {
            ICell c = sheet._GetCell(y, x, false);
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

        static public void _MoveCell(this ISheet sheet, int fromCellY, int fromCellX, int toCellY, int toCellX, Action<ICell> onFormulaCellMoved = null)
        {
            ICell fromCell = sheet._GetCell(fromCellY, fromCellX, false);
            fromCell._Move(toCellY, toCellX, onFormulaCellMoved);
        }

        static public ICell _GetCell(this ISheet sheet, int y, int x, bool createCell)
        {
            IRow r = sheet._GetRow(y, createCell);
            if (r == null)
                return null;
            return r._GetCell(x, createCell);
        }

        static public ICell _GetCell(this ISheet sheet, string address, bool createCell)
        {
            var cs = GetCoordinates(address);
            IRow r = sheet._GetRow(cs.Y, createCell);
            if (r == null)
                return null;
            return r._GetCell(cs.X, createCell);
        }

        static public void RemoveCell(this ISheet sheet, int y, int x)
        {
            IRow r = sheet.GetRow(y);
            if (r == null)
                return;
            ICell c = r.GetCell(x);
            if (c == null)
                return;
            r.RemoveCell(c);
        }

        static internal Range _getMergedRange(this ISheet sheet, int y, int x)
        {
            foreach (var mr in sheet.MergedRegions)
                if (mr.IsInRange(y - 1, x - 1))
                    return new Range(sheet, mr.FirstRow + 1, mr.FirstColumn + 1, mr.LastRow + 1, mr.LastColumn + 1);
            return null;
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
            //if (p is XSSFPicture xSSFPicture)
            //    xSSFPicture.IsNoFill = true;
            p.Resize(1);
            //p.Resize(1, 1);
        }
    }
}