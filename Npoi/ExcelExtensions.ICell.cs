﻿//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************
using NPOI.HSSF.UserModel;
using NPOI.OpenXml4Net.OPC;
using NPOI.SS.Extractor;
using NPOI.SS.Formula;
using NPOI.SS.Formula.Functions;
using NPOI.SS.Formula.PTG;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.Extractor;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.Extractor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using static Cliver.Excel;

namespace Cliver
{
    static public partial class ExcelExtensions
    {
        static public IClientAnchor _SetComment(this ICell cell, string comment, string author = null, IClientAnchor anchor = null)
        {
            cell.RemoveCellComment();//!!!adding multiple comments brings to error
            if (string.IsNullOrWhiteSpace(comment))
                return null;

            //if (anchor == null)
            //{
            //    anchor = creationHelper.CreateClientAnchor();
            //    anchor.Col1 = cell.ColumnIndex;
            //    anchor.Col2 = cell.ColumnIndex + 3;
            //    anchor.Row1 = cell.RowIndex;
            //    anchor.Row2 = cell.RowIndex + Regex.Matches(comment, @"^", RegexOptions.Multiline).Count + 3;
            //}
            var drawingPatriarch = cell.Sheet.DrawingPatriarch != null ? cell.Sheet.DrawingPatriarch : cell.Sheet.CreateDrawingPatriarch();
            if (anchor == null)
                anchor = drawingPatriarch.CreateAnchor(0, 0, 0, 0, cell.ColumnIndex, cell.RowIndex, cell.ColumnIndex + 3, cell.RowIndex + Regex.Matches(comment, @"^", RegexOptions.Multiline).Count + 3);
            IComment iComment = null;
            //try
            //{
            iComment = drawingPatriarch.CreateCellComment(anchor);
            //}
            //catch (Exception e)//!!!sometimes occured for unknown reason
            //{
            //    //var cs = cell.Sheet.GetCellComments(); //!!!this throws an exception too
            //}
            if (author != null)
                iComment.Author = author;
            comment = addCommentAuthor(comment, author);
            iComment.String = cell.Sheet.Workbook.GetCreationHelper().CreateRichTextString(comment);
            cell.CellComment = iComment;

            return cell.CellComment.ClientAnchor;
        }
        static string addCommentAuthor(string comment, string author)
        {
            if (!string.IsNullOrWhiteSpace(author))
                comment = "[" + author + "]:\r\n" + comment;
            return comment;
        }

        static public IClientAnchor _AppendOrSetComment(this ICell cell, string comment, string author = null, string delimiter = "\r\n\r\n", IClientAnchor anchor = null)
        {
            if (string.IsNullOrWhiteSpace(comment))
                return cell?.CellComment?.ClientAnchor;

            if (string.IsNullOrEmpty(cell?.CellComment?.String?.String))
                return cell._SetComment(comment, author, anchor);

            IComment iComment = cell.CellComment;
            comment = delimiter + addCommentAuthor(comment, author);
            iComment.String = cell.Sheet.Workbook.GetCreationHelper().CreateRichTextString(iComment.String.String + comment);
            int r2 = iComment.ClientAnchor.Row2;
            iComment.ClientAnchor.Row2 += Regex.Matches(comment, @"^", RegexOptions.Multiline).Count - 1;
            if (iComment.ClientAnchor.Row2 <= r2)
            {//!!!sometimes happens for unknown reason
                //throw new Exception("Could not increase ClientAnchor");
                return cell._SetComment(iComment.String.String);
            }
            cell.CellComment = iComment;
            return cell.CellComment.ClientAnchor;
        }

        static public string _GetAddress(this ICell cell)
        {
            return cell?.Address.ToString();
        }

        /// Remove the cell from its row.
        static public void _Remove(this ICell cell)
        {
            cell.Row.RemoveCell(cell);
        }

        static public ICell _Move(this ICell cell1, int cell2Y, int cell2X, OnFormulaCellMoved onFormulaCellMoved = null, ISheet sheet2 = null, StyleMap styleMap = null)
        {
            ICell cell2 = _Copy(cell1, cell2Y, cell2X, onFormulaCellMoved, sheet2, styleMap);
            cell1?._Remove();
            return cell2;
        }

        static public void _Move(this ICell cell1, ICell cell2, OnFormulaCellMoved onFormulaCellMoved = null, StyleMap styleMap = null)
        {
            _Copy(cell1, cell2, onFormulaCellMoved, styleMap);
            cell1?._Remove();
        }

        static public ICell _Copy(this ICell cell1, int cell2Y, int cell2X, OnFormulaCellMoved onFormulaCellMoved = null, ISheet sheet2 = null, StyleMap styleMap = null)
        {
            if (sheet2 == null)
                sheet2 = cell1.Sheet;
            if (cell1 == null)
            {
                ICell cell2 = sheet2._GetCell(cell2Y, cell2X, false);
                cell2?._Remove();
                return null;
            }
            else
            {
                ICell cell2 = sheet2._GetCell(cell2Y, cell2X, true);
                _Copy(cell1, cell2, onFormulaCellMoved, styleMap);
                return cell2;
            }
        }

        static public void _Copy(this ICell cell1, ICell cell2, OnFormulaCellMoved onFormulaCellMoved = null, StyleMap styleMap = null)
        {
            if (cell1 == null)
            {
                cell2?._Remove();
                return;
            }

            cell2.SetBlank();
            cell2.SetCellType(cell1.CellType);

            if (cell1.Sheet.Workbook != cell2.Sheet.Workbook)
            {
                if (styleMap == null)
                    throw new Exception("styleMap must be specified when copying cell to another workbook.");
                if (cell2.Sheet.Workbook != styleMap.ToWorkbook)
                    throw new Exception("cell2 does not belong to styleMap's workbook.");
                cell2.CellStyle = styleMap.GetMappedStyle(cell1.CellStyle);
            }
            else
                cell2.CellStyle = cell1.CellStyle;

            cell2.RemoveCellComment();
            cell2.CellComment = cell1.CellComment;
            //cell2._SetLink(cell1.Hyperlink?.Address);
            cell2.Hyperlink = cell1.Hyperlink;
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
            if (cell2?.CellType == CellType.Formula)
                onFormulaCellMoved?.Invoke(cell1, cell2);
        }

        /// <summary>
        /// NULL- and type-safe.
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="allowNull"></param>
        /// <returns></returns>
        static public string _GetValueAsString(this ICell cell, bool allowNull = false)
        {
            object o = cell?._GetValue();
            if (o == null)
                return allowNull ? null : string.Empty;
            if (o is DateTime dt)
                return dt.ToString("yyyy-MM-dd hh:mm:ss");
            return o?.ToString();
        }

        static public object _GetValue(this ICell cell)
        {
            if (cell == null)
                return null;
            switch (cell.CellType)
            {
                case CellType.Unknown:
                    //return cell.ToString();
                    throw new Exception("Needs debugging for this cell type: " + cell.CellType);
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        try
                        {
                            return cell.DateCellValue;
                        }
                        catch /*(Exception e)*///!!!bug in NPOI2.5.1: after called Save(), it throws NullReferenceException: GetLocaleCalendar()  https://github.com/nissl-lab/npoi/issues/358
                        {
                            //Log.Warning("NPOI bug", e);
                            return DateTime.FromOADate(cell.NumericCellValue);
                        }
                        //return formatter.FormatCellValue(c);
                    }
                    return cell.NumericCellValue;
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Boolean:
                    return cell.BooleanCellValue;
                case CellType.Formula:
                    //return c.CellFormula;
                    IFormulaEvaluator formulaEvaluator;
                    if (cell.Sheet.Workbook is XSSFWorkbook)
                        formulaEvaluator = new XSSFFormulaEvaluator(cell.Sheet.Workbook);
                    else if (cell.Sheet.Workbook is HSSFWorkbook)
                        formulaEvaluator = new HSSFFormulaEvaluator(cell.Sheet.Workbook);
                    else
                        throw new Exception("Unsupported workbook type: " + cell.Sheet.Workbook.GetType().FullName);
                    var cv = formulaEvaluator.Evaluate(cell);
                    switch (cv.CellType)
                    {
                        case CellType.Unknown:
                            //return cv.ToString();
                            throw new Exception("Needs debugging for this cell type: " + cell.CellType);
                        case CellType.Numeric:
                            return cv.NumberValue;
                        case CellType.String:
                            return cv.StringValue;
                        case CellType.Boolean:
                            return cv.BooleanValue;
                        case CellType.Error:
                            return FormulaError.ForInt(cv.ErrorValue).String;
                        case CellType.Blank:
                            return null;
                        default:
                            throw new Exception("Unknown type: " + cv.CellType);
                    }
                case CellType.Error:
                    //return c.ErrorCellValue.ToString();
                    return FormulaError.ForInt(cell.ErrorCellValue).String;
                case CellType.Blank:
                    return null;
                default:
                    throw new Exception("Unknown type: " + cell.CellType);
            }
        }

        /// <summary>
        /// NULL- and type-safe.
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        static public void _SetValue(this ICell cell, object value)
        {
            if (value == null)
                cell.SetBlank();
            else if (value is sbyte
                        || value is byte
                        || value is short
                        || value is ushort
                        || value is int
                        || value is uint
                        || value is long
                        || value is ulong
                        || value is float
                        || value is double
                        || value is decimal
                )
                cell.SetCellValue(Convert.ToDouble(value));
            else if (value is bool b)
                cell.SetCellValue(b);
            else if (value is DateTime dt)
                cell.SetCellValue(dt);
            else
                cell.SetCellValue(value?.ToString());
        }

        static public string _GetLink(this ICell cell)
        {
            return cell?.Hyperlink?.Address;
        }

        static public void _SetLink(this ICell cell, string link, HyperlinkType hyperlinkType = HyperlinkType.Unknown)
        {
            while (cell.Hyperlink != null)//it might be more than 1 link in the table
                cell.RemoveHyperlink();//(!)seems to be necessary in any case to get rid of the old link. Otherwise sometimes the old link is not overriden by the new one.
            if (link == null)
            {
                //if (cell.GetValueAsString() == LinkEmptyValueFiller)
                //    cell.SetCellValue("");
                return;
            }
            if (string.IsNullOrEmpty(cell._GetValueAsString()))
                cell.SetCellValue(Excel.LinkEmptyValueFiller);

            if (hyperlinkType == HyperlinkType.Unknown)
            {
                if (Regex.IsMatch(link, @"^\s*(https?|ftps?)\:", RegexOptions.IgnoreCase))
                    hyperlinkType = HyperlinkType.Url;
                else if (Regex.IsMatch(link, @"^\s*[a-z]\:", RegexOptions.IgnoreCase))
                    hyperlinkType = HyperlinkType.File;
                else if (Regex.IsMatch(link, @"\@", RegexOptions.IgnoreCase))
                    hyperlinkType = HyperlinkType.Email;
                else
                    hyperlinkType = HyperlinkType.Document;
            }

            if (cell.Sheet.Workbook is XSSFWorkbook)
                cell.Hyperlink = new XSSFHyperlink(hyperlinkType) { Address = link };
            else if (cell.Sheet.Workbook is HSSFWorkbook)
                cell.Hyperlink = new HSSFHyperlink(hyperlinkType) { Address = link };
            else
                throw new Exception("Unsupported workbook type: " + cell.Sheet.Workbook.GetType().FullName);
        }

        /// <summary>
        /// It automatically updates the ranges in the cell formula.
        /// It is expected to work properly for trivial formulas. 
        /// (!)You have to check if it works as you need. 
        /// </summary>
        /// <param name="formulaCell"></param>
        /// <param name="rangeY1Shift"></param>
        /// <param name="rangeX1Shift"></param>
        /// <param name="rangeY2Shift"></param>
        /// <param name="rangeX2Shift"></param>
        /// <exception cref="Exception"></exception>
        static public void _UpdateFormulaRange(this ICell formulaCell, int rangeY1Shift, int rangeX1Shift, int? rangeY2Shift = null, int? rangeX2Shift = null)
        {
            if (formulaCell?.CellType != CellType.Formula)
                return;

            if (rangeY2Shift == null)
                rangeY2Shift = rangeY1Shift;
            if (rangeX2Shift == null)
                rangeX2Shift = rangeX1Shift;

            IFormulaParsingWorkbook evaluationWorkbook;
            if (formulaCell.Sheet.Workbook is XSSFWorkbook)
                evaluationWorkbook = XSSFEvaluationWorkbook.Create(formulaCell.Sheet.Workbook);
            else if (formulaCell.Sheet.Workbook is HSSFWorkbook)
                evaluationWorkbook = HSSFEvaluationWorkbook.Create(formulaCell.Sheet.Workbook);
            //else if (sheet is SXSSFWorkbook)
            //{
            //    evaluationWorkbook = SXSSFEvaluationWorkbook.Create((SXSSFWorkbook)Workbook);
            else
                throw new Exception("Unsupported workbook type: " + formulaCell.Sheet.Workbook.GetType().FullName);

            var ptgs = FormulaParser.Parse(formulaCell.CellFormula, evaluationWorkbook, FormulaType.Cell, formulaCell.Sheet.Workbook.GetSheetIndex(formulaCell.Sheet));
            foreach (Ptg ptg in ptgs)
            {
                if (ptg is RefPtgBase rpb)
                {
                    if (rpb.IsRowRelative)
                        rpb.Row = rpb.Row + rangeY1Shift;
                    if (rpb.Row < 0)
                        rpb.Row = 0;
                    if (rpb.IsColRelative)
                        rpb.Column = rpb.Column + rangeX1Shift;
                    if (rpb.Column < 0)
                        rpb.Column = 0;
                }
                else if (ptg is AreaPtgBase apb)
                {
                    if (apb.IsFirstRowRelative)
                        apb.FirstRow += rangeY1Shift;
                    if (apb.FirstRow < 0)
                        apb.FirstRow = 0;
                    if (apb.IsLastRowRelative)
                        apb.LastRow += rangeY2Shift.Value;
                    if (apb.LastRow < 0)
                        apb.LastRow = 0;
                    if (apb.IsFirstColRelative)
                        apb.FirstColumn += rangeX1Shift;
                    if (apb.FirstColumn < 0)
                        apb.FirstColumn = 0;
                    if (apb.IsLastColRelative)
                        apb.LastColumn += rangeX2Shift.Value;
                    if (apb.LastColumn < 0)
                        apb.LastColumn = 0;
                }
                //else
                //    throw new Exception("Unexpected ptg type: " + ptg.GetType());
            }
            formulaCell.CellFormula = FormulaRenderer.ToFormulaString((IFormulaRenderingWorkbook)evaluationWorkbook, ptgs);
        }

        static public Excel.Range _GetMergedRange(this ICell cell)
        {
            return cell.Sheet._GetMergedRange(cell.RowIndex + 1, cell.ColumnIndex + 1);
        }

        static public void _ClearMerging(this ICell cell)
        {
            for (int i = cell.Sheet.MergedRegions.Count - 1; i >= 0; i--)
                if (cell.Sheet.MergedRegions[i].IsInRange(cell.RowIndex, cell.ColumnIndex))
                {
                    cell.Sheet.RemoveMergedRegion(i);
                    return;//there can be only one MergedRegion
                }
        }

        /// <summary>
        /// Cell's 1-based row index on the sheet.
        /// </summary>
        /// <param name="cell"></param>
        /// <returns>1-based</returns>
        static public int _Y(this ICell cell)
        {
            return cell.RowIndex + 1;
        }

        /// <summary>
        /// Cell's 1-based column index on the sheet.
        /// </summary>
        /// <param name="cell"></param>
        /// <returns>1-based</returns>
        static public int _X(this ICell cell)
        {
            return cell.ColumnIndex + 1;
        }

        static public void _CreateDropdown<T>(this ICell cell, IEnumerable<T> values, T value, bool allowBlank = true)
        {
            List<string> vs = new List<string>();
            foreach (object v in values)
                vs.Add(v?.ToString());

            IDataValidationHelper dvh;
            if (cell.Sheet is XSSFSheet)
                dvh = new XSSFDataValidationHelper((XSSFSheet)cell.Sheet);
            else if (cell.Sheet is HSSFSheet)
                dvh = new HSSFDataValidationHelper((HSSFSheet)cell.Sheet);
            else
                throw new Exception("Unsupported workbook type: " + cell.Sheet.Workbook.GetType().FullName);
            //string dvs = string.Join(",", vs);
            //IDataValidationConstraint dvc = Sheet.GetDataValidations().Find(a => string.Join(",", a.ValidationConstraint.ExplicitListValues) == dvs)?.ValidationConstraint;
            //if (dvc == null)
            //dvc = dvh.CreateCustomConstraint(dvs);
            IDataValidationConstraint dvc = dvh.CreateExplicitListConstraint(vs.ToArray());
            CellRangeAddressList cral = new CellRangeAddressList(cell.RowIndex, cell.RowIndex, cell.ColumnIndex, cell.ColumnIndex);
            IDataValidation dv = dvh.CreateValidation(dvc, cral);
            dv.SuppressDropDownArrow = true;
            dv.EmptyCellAllowed = allowBlank;
            ((XSSFSheet)cell.Sheet).AddValidationData(dv);

            cell.SetCellValue(value?.ToString());
        }

        static public IEnumerable<Excel.Image> _GetImages(this ICell cell)
        {
            return cell.Sheet._GetImages(cell._Y(), cell._X());
        }
    }
}