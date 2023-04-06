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

namespace Cliver
{
    static public partial class ExcelExtensions
    {
        static public string GetValueAsString(this ICell cell, bool allowNull = false)
        {
            object o = cell?.GetValue();
            if (!allowNull && o == null)
                return string.Empty;
            if (o is DateTime dt)
                return dt.ToString("yyyy-MM-dd hh:mm:ss");
            return o?.ToString();
        }

        static public object GetValue(this ICell cell)
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
                        catch (Exception e)//!!!bug in NPOI2.5.1: after called Save(), it throws NullReferenceException: GetLocaleCalendar()  https://github.com/nissl-lab/npoi/issues/358
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

        static public void SetValue(this ICell cell, object value)
        {
            if (value == null)
            {
                cell.SetBlank();
                return;
            }
            if (value is double d)
            {
                cell.SetCellValue(d);
                return;
            }
            if (value is bool b)
            {
                cell.SetCellValue(b);
                return;
            }
            if (value is DateTime dt)
            {
                cell.SetCellValue(dt);
                return;
            }
            cell.SetCellValue(value?.ToString());
        }

        static public Uri GetLink(this ICell cell)
        {
            if (cell == null)
                return null;
            if (cell.Hyperlink == null)
                return null;
            return new Uri(cell.Hyperlink.Address, UriKind.RelativeOrAbsolute);
        }

        static public void SetLink(this ICell cell, Uri uri)
        {
            if (uri == null)
            {
                //if (cell.GetValueAsString() == LinkEmptyValueFiller)
                //    cell.SetCellValue("");
                cell.Hyperlink = null;
                return;
            }
            if (string.IsNullOrEmpty(cell.GetValueAsString()))
                cell.SetCellValue(LinkEmptyValueFiller);
            if (cell.Sheet.Workbook is XSSFWorkbook)
                cell.Hyperlink = new XSSFHyperlink(HyperlinkType.Url) { Address = uri.ToString() };
            else if (cell.Sheet.Workbook is HSSFWorkbook)
                cell.Hyperlink = new HSSFHyperlink(HyperlinkType.Url) { Address = uri.ToString() };
            else
                throw new Exception("Unsupported workbook type: " + cell.Sheet.Workbook.GetType().FullName);
        }
        public static string LinkEmptyValueFiller = "           ";

        static public void UpdateFormulaRange(this ICell formulaCell, int rangeY1Shift, int rangeX1Shift, int? rangeY2Shift = null, int? rangeX2Shift = null)
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
                    if (rpb.IsColRelative)
                        rpb.Column = rpb.Column + rangeX1Shift;
                }
                else if (ptg is AreaPtgBase apb)
                {
                    if (apb.IsFirstRowRelative)
                        apb.FirstRow += rangeY1Shift;
                    if (apb.IsLastRowRelative)
                        apb.LastRow += rangeY2Shift.Value;
                    if (apb.IsFirstColRelative)
                        apb.FirstColumn += rangeX1Shift;
                    if (apb.IsLastColRelative)
                        apb.LastColumn += rangeX2Shift.Value;
                }
                //else
                //    throw new Exception("Unexpected ptg type: " + ptg.GetType());
            }
            formulaCell.CellFormula = FormulaRenderer.ToFormulaString((IFormulaRenderingWorkbook)evaluationWorkbook, ptgs);
        }

        //static public void Highlight(this ICell cell, ICellStyle style, Excel.Color color)
        //{
        //    cell.CellStyle = Excel.highlight(cell.Sheet.Workbook, style, color);
        //}

        static public Excel.Range GetMergedRange(this ICell cell)
        {
            return Excel.getMergedRange(cell.Row.Sheet, cell.RowIndex + 1, cell.ColumnIndex + 1);
        }

        static public void ClearMerging(this ICell cell)
        {
            for (int i = cell.Sheet.MergedRegions.Count - 1; i >= 0; i--)
                if (cell.Sheet.MergedRegions[i].IsInRange(cell.RowIndex, cell.ColumnIndex))
                {
                    cell.Sheet.RemoveMergedRegion(i);
                    return;//there can be only one MergedRegion
                }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cell"></param>
        /// <returns>1-based</returns>
        static public int Y(this ICell cell)
        {
            return cell.RowIndex + 1;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cell"></param>
        /// <returns>1-based</returns>
        static public int X(this ICell cell)
        {
            return cell.ColumnIndex + 1;
        }

        static public void CreateDropdown(this ICell cell, IEnumerable<object> values, object value, bool allowBlank = true)
        {
            List<string> vs = new List<string>();
            foreach (object v in values)
            {
                string s;
                if (v is string)
                    s = (string)v;
                else if (v != null)
                    s = v.ToString();
                else
                    s = null;
                vs.Add(s);
            }

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

            {
                string s;
                if (value is string)
                    s = (string)value;
                else if (value != null)
                    s = value.ToString();
                else
                    s = null;
                cell.SetCellValue(s);
            }
        }
    }
}