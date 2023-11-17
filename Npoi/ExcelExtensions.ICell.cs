//********************************************************************************************
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
using System.Text.RegularExpressions;
using static Cliver.Excel;

namespace Cliver
{
    static public partial class ExcelExtensions
    {
        static public string _GetAddress(this ICell cell)
        {
            return cell?.Address.ToString();
        }

        static public void _Remove(this ICell cell)
        {
            cell.Row.RemoveCell(cell);
        }

        static public ICell _Move(this ICell fromCell, int toCellY, int toCellX, OnFormulaCellMoved onFormulaCellMoved = null, ISheet toSheet = null)
        {
            ICell toCell = fromCell._Copy(toCellY, toCellX, onFormulaCellMoved, toSheet);
            if (fromCell != null)
                fromCell.Row.RemoveCell(fromCell);
            //if (toCell?.CellType == CellType.Formula)
            //    onFormulaCellMoved?.Invoke(fromCell, toCell);
            return toCell;
        }

        static public ICell _Copy(this ICell fromCell, int toCellY, int toCellX, OnFormulaCellMoved onFormulaCellMoved = null, ISheet toSheet = null)
        {
            if (toSheet == null)
            {
                if (fromCell == null)
                    return null;
                toSheet = fromCell.Sheet;
            }
            if (fromCell == null)
            {
                IRow toRow = toSheet._GetRow(toCellY, false);
                if (toRow == null)
                    return null;
                ICell toCell = toRow._GetCell(toCellX, false);
                if (toCell == null)
                    return toCell;
                toRow.RemoveCell(toCell);
                return toCell;
            }
            else
            {
                ICell toCell = toSheet._GetCell(toCellY, toCellX, true);
                _Copy(fromCell, toCell, onFormulaCellMoved);
                return toCell;
            }
        }

        static public void _Copy(this ICell fromCell, ICell toCell, OnFormulaCellMoved onFormulaCellMoved = null)
        {
            toCell.SetBlank();
            toCell.SetCellType(fromCell.CellType);
            toCell.CellStyle = fromCell.CellStyle;
            toCell.CellComment = fromCell.CellComment;
            //toCell._SetLink(fromCell.Hyperlink?.Address);
            toCell.Hyperlink = fromCell.Hyperlink;
            switch (fromCell.CellType)
            {
                case CellType.Formula:
                    toCell.CellFormula = fromCell.CellFormula;
                    break;
                case CellType.Numeric:
                    toCell.SetCellValue(fromCell.NumericCellValue);
                    break;
                case CellType.String:
                    toCell.SetCellValue(fromCell.StringCellValue);
                    break;
                case CellType.Boolean:
                    toCell.SetCellValue(fromCell.BooleanCellValue);
                    break;
                case CellType.Error:
                    toCell.SetCellErrorValue(fromCell.ErrorCellValue);
                    break;
                case CellType.Blank:
                    toCell.SetBlank();
                    break;
                default:
                    throw new Exception("Unknown cell type: " + fromCell.CellType);
            }
            if (toCell?.CellType == CellType.Formula)
                onFormulaCellMoved?.Invoke(fromCell, toCell);
        }

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
                cell.SetCellValue((double)value);
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