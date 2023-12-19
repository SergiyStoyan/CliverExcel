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
        static public void _SetComment(this ISheet sheet, int y, int x, string comment, Excel.CommentStyle commentStyle = null)
        {
            sheet._GetCell(y, x, true)._SetComment(comment, commentStyle);
        }

        static public void _AppendOrSetComment(this ISheet sheet, int y, int x, string comment, Excel.CommentStyle commentStyle = null)
        {
            sheet._GetCell(y, x, true)._AppendOrSetComment(comment, commentStyle);
        }

        static public void _SetLink(this ISheet sheet, int y, int x, string link)
        {
            sheet._GetCell(y, x, true)._SetLink(link);
        }

        static public string _GetLink(this ISheet sheet, int y, int x)
        {
            return sheet?._GetCell(y, x, false)?._GetLink();
        }

        static public void _ShiftCellsRight(this ISheet sheet, int x1, int y1, int y2, int shift, CopyCellMode copyCellMode = null)
        {
            for (int y = y1; y <= y2; y++)
                sheet._GetRow(y, false)?._ShiftCellsRight(x1, shift, copyCellMode);
        }

        static public void _ShiftCellsLeft(this ISheet sheet, int x1, int y1, int y2, int shift, CopyCellMode copyCellMode = null)
        {
            for (int y = y1; y <= y2; y++)
                sheet._GetRow(y, false)?._ShiftCellsLeft(x1, shift, copyCellMode);
        }

        static public void _ShiftCellsDown(this ISheet sheet, int y1, int x1, int x2, int shift, CopyCellMode copyCellMode = null)
        {
            for (int x = x1; x <= x2; x++)
                sheet._GetColumn(x)?.ShiftCellsDown(y1, shift, copyCellMode);
        }

        static public void _ShiftCellsUp(this ISheet sheet, int y1, int x1, int x2, int shift, CopyCellMode copyCellMode = null)
        {
            for (int x = x1; x <= x2; x++)
                sheet._GetColumn(x)?.ShiftCellsUp(y1, shift, copyCellMode);
        }

        static public ICell _MoveCell(this ISheet sheet, int y1, int x1, int y2, int x2, CopyCellMode copyCellMode = null, ISheet sheet2 = null, StyleMap styleMap = null)
        {
            sheet2 = sheet2 ?? sheet;
            if (sheet2 == sheet && y1 == y2 && x1 == x2)//(!)otherwise it will remove the cell
                return sheet._GetCell(y1, x1, false);

            CopyCellMode ccm;
            if (copyCellMode != null && sheet is HSSFSheet)//!!!REMOVE WHEN FIXED IN NPOI: NPOI_BUG_ShapeId_duplication_when_creating_a_comment
            {
                ccm = copyCellMode.Clone();
                ccm.CopyComment = false;//(!)done due to the bug in in NPOI: ShapeId duplication when creating a comment.
            }
            else
                ccm = null;
            ICell cell2 = sheet._CopyCell(y1, x1, y2, x2, ccm, sheet2, styleMap);
            if (ccm != null && copyCellMode?.CopyComment == true && cell2 != null)
            {
                cell2.RemoveCellComment();
                cell2.CellComment = sheet._GetCell(y1, x1, false)?.CellComment;
            }

            bool removeComment = false;
            if (copyCellMode?.CopyComment == true)
                removeComment = sheet.GetCellComment(new CellAddress(x1 - 1, y1 - 1)) == null;
            sheet2._RemoveCell(y1, x1, removeComment);
            return cell2;
        }

        static public ICell _CopyCell(this ISheet sheet, int y1, int x1, int y2, int x2, CopyCellMode copyCellMode = null, ISheet sheet2 = null, StyleMap styleMap = null)
        {
            sheet2 = sheet2 ?? sheet;
            if (sheet == sheet2 && y1 == y2 && x1 == x2)
                return null;

            ICell cell1 = sheet._GetCell(y1, x1, false);
            if (cell1 == null)
            {
                var comment = sheet.GetCellComment(new CellAddress(x1 - 1, y1 - 1));
                sheet2._RemoveCell(y2, x2, comment == null && copyCellMode?.CopyComment == true);
                return null;
            }

            ICell cell2 = sheet2._GetCell(y2, x2, true);
            if (cell2.CellType != cell1.CellType)
            {
                cell2.SetBlank();//necessary if changing type
                cell2.SetCellType(cell1.CellType);
            }

            if (cell1.Sheet.Workbook != cell2.Sheet.Workbook)
            {
                if (styleMap == null)
                    throw new Exception("StyleMap must be specified when copying cell to another workbook.");
                if (cell2.Sheet.Workbook != styleMap.Workbook2)
                    throw new Exception("cell2 does not belong to StyleMap's workbook.");
                cell2.CellStyle = styleMap.GetMappedStyle(cell1.CellStyle);
            }
            else
                cell2.CellStyle = cell1.CellStyle;

            switch (cell1.CellType)
            {
                case CellType.Formula:
                    cell2.CellFormula = cell1.CellFormula;
                    break;
                case CellType.Numeric:
                    cell2.SetCellValue(cell1.NumericCellValue);
                    break;
                case CellType.String:
                    cell2.SetCellValue(cell1.StringCellValue);
                    break;
                case CellType.Boolean:
                    cell2.SetCellValue(cell1.BooleanCellValue);
                    break;
                case CellType.Error:
                    cell2.SetCellErrorValue(cell1.ErrorCellValue);
                    break;
                case CellType.Blank:
                    cell2.SetBlank();
                    break;
                default:
                    throw new Exception("Unknown cell type: " + cell1.CellType);
            }

            if (copyCellMode?.CopyComment == true)
            {
                cell2.RemoveCellComment();
                if (cell1.CellComment != null)
                {
                    //cell1.Sheet.CopyComment(cell1, cell2);!!!on HSSF it moves the comment, not copies; on XSSF it copies but does not preserve box size
                    var drawingPatriarch2 = /*cell1.Sheet.DrawingPatriarch != null ? cell1.Sheet.DrawingPatriarch :*/ cell2.Sheet.CreateDrawingPatriarch();
                    (int Y, int X) shift = (cell2._Y() - cell1._Y(), cell2._X() - cell1._X());
                    IClientAnchor anchor2 = drawingPatriarch2.CreateAnchor(
                        cell1.CellComment.ClientAnchor.Dx1,
                        cell1.CellComment.ClientAnchor.Dy1,
                        cell1.CellComment.ClientAnchor.Dx2,
                        cell1.CellComment.ClientAnchor.Dy2,
                        cell1.CellComment.ClientAnchor.Col1 + shift.X,
                        cell1.CellComment.ClientAnchor.Row1 + shift.Y,
                        cell1.CellComment.ClientAnchor.Col2 + shift.X,
                        cell1.CellComment.ClientAnchor.Row2 + shift.Y
                    );
                    IComment comment2;
                    try
                    {
                        comment2 = drawingPatriarch2.CreateCellComment(anchor2);
                    }
                    catch (Exception e)
                    {//!!!when fixed, search for label NPOI_BUG_ShapeId_duplication_when_creating_a_comment and update the code
                        throw new Exception("A bug in HSSFPatriarch implementation: ShapeId duplication when creating a comment.", e);
                    }
                    if (cell1.CellComment.Author != null)
                        comment2.Author = cell1.CellComment.Author;
                    comment2.String = cell1.CellComment.String.Copy();
                    cell2.CellComment = comment2;
                }
            }

            if (!(copyCellMode?.CopyLink == false))
                cell2.Hyperlink = cell1.Hyperlink;

            if (cell2?.CellType == CellType.Formula)
                copyCellMode?.OnFormulaCellMoved?.Invoke(cell1, cell2);

            return cell2;
        }

        static public void _RemoveCell(this ISheet sheet, int y, int x, bool removeComment)
        {
            ICell cell = sheet._GetCell(y, x, false);
            if (removeComment)
            {
                cell = cell ?? sheet._GetCell(y, x, true);
                cell.RemoveCellComment();
            }
            cell?.Row.RemoveCell(cell);
        }

        static public string _GetValueAsString(this ISheet sheet, int y, int x, StringMode stringMode = DefaultStringMode)
        {
            ICell c = sheet._GetCell(y, x, false);
            return c._GetValueAsString(stringMode);
        }

        static public string _GetValueAsString(this ISheet sheet, string cellAddress, StringMode stringMode = DefaultStringMode)
        {
            ICell c = sheet._GetCell(cellAddress, false);
            return c._GetValueAsString(stringMode);
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
            return sheet._GetCell(cs.Y, cs.X, createCell);
        }

        static public ICell _GetCell(this ISheet sheet, CellAddress cellAddress, bool createCell)
        {
            return sheet._GetCell(cellAddress.Row + 1, cellAddress.Column + 1, createCell);
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
            var drawingPatriarch = /*sheet.DrawingPatriarch != null ? sheet.DrawingPatriarch :*/ sheet.CreateDrawingPatriarch();
            IClientAnchor a = drawingPatriarch.CreateAnchor(0, 0, 0, 0, image.X - 1, image.Y - 1, image.X - 1, image.Y - 1);
            a.AnchorType = AnchorType.MoveDontResize;
            IPicture p = drawingPatriarch.CreatePicture(a, imageId);
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