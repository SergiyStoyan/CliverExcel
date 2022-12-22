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

//works  
namespace Cliver
{
    static public partial class ExcelExtensions
    {
        static public string GetValueAsString(this ICell cell)
        {
            if (cell == null)
                return null;
            switch (cell.CellType)
            {
                case CellType.Unknown:
                    return cell.ToString();
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        try
                        {
                            return cell.DateCellValue.ToString("yyyy-MM-dd hh:mm:ss");
                        }
                        catch (Exception e)//!!!bug in NPOI2.5.1: after called Save(), it throws NullReferenceException: GetLocaleCalendar()  https://github.com/nissl-lab/npoi/issues/358
                        {
                            //Log.Warning("NPOI bug", e);
                            return DateTime.FromOADate(cell.NumericCellValue).ToString("yyyy-MM-dd hh:mm:ss");
                        }
                        //return formatter.FormatCellValue(c);
                    }
                    return cell.NumericCellValue.ToString();
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Boolean:
                    return cell.BooleanCellValue.ToString().ToUpper();
                case CellType.Formula:
                    //return c.CellFormula;
                    IFormulaEvaluator formulaEvaluator;
                    if (cell.Sheet.Workbook is XSSFWorkbook)
                        formulaEvaluator = new XSSFFormulaEvaluator(cell.Sheet.Workbook);
                    else if (cell.Sheet.Workbook is HSSFWorkbook)
                        formulaEvaluator = new HSSFFormulaEvaluator(cell.Sheet.Workbook);
                    else
                        throw new Exception("Unexpected Workbook type: " + cell.Sheet.Workbook.GetType());
                    var cv = formulaEvaluator.Evaluate(cell);
                    switch (cv.CellType)
                    {
                        case CellType.Unknown:
                            return cv.ToString();
                        case CellType.Numeric:
                            return cv.NumberValue.ToString();
                        case CellType.String:
                            return cv.StringValue;
                        case CellType.Boolean:
                            return cv.BooleanValue.ToString().ToUpper();
                        case CellType.Error:
                            return FormulaError.ForInt(cv.ErrorValue).String;
                        case CellType.Blank:
                            return string.Empty;
                        default:
                            throw new Exception("Unknown type: " + cv.CellType);
                    }
                case CellType.Error:
                    //return c.ErrorCellValue.ToString();
                    return FormulaError.ForInt(cell.ErrorCellValue).String;
                case CellType.Blank:
                    return string.Empty;
                default:
                    throw new Exception("Unknown type: " + cell.CellType);
            }
        }

        static public void SetLink(this ICell cell, Uri uri)
        {
            if (string.IsNullOrEmpty(cell.GetValueAsString()))
                cell.SetCellValue(LinkEmptyValueFiller);
            if (cell.Sheet.Workbook is XSSFWorkbook)
                cell.Hyperlink = new XSSFHyperlink(HyperlinkType.Url) { Address = uri.ToString() };
            else if (cell.Sheet.Workbook is HSSFWorkbook)
                cell.Hyperlink = new HSSFHyperlink(HyperlinkType.Url) { Address = uri.ToString() };
        }
        public static string LinkEmptyValueFiller = "           ";

        static public Uri GetLink(this ICell cell)
        {
            if (cell == null)
                return null;
            if (cell.Hyperlink == null)
                return null;
            return new Uri(cell.Hyperlink.Address, UriKind.RelativeOrAbsolute);
        }

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
                throw new Exception("Unexpected Workbook type: " + formulaCell.Sheet.Workbook.GetType());

            var ptgs = FormulaParser.Parse(formulaCell.CellFormula, evaluationWorkbook, FormulaType.Cell, formulaCell.Sheet.Workbook.GetSheetIndex(formulaCell.Sheet));
            foreach (Ptg ptg in ptgs)
            {
                if (ptg is RefPtgBase)
                {
                    RefPtgBase ref2 = (RefPtgBase)ptg;
                    if (ref2.IsRowRelative)
                        ref2.Row = ref2.Row + rangeY1Shift;
                    if (ref2.IsColRelative)
                        ref2.Column = ref2.Column + rangeX1Shift;
                }
                else if (ptg is AreaPtgBase)
                {
                    AreaPtgBase ref2 = (AreaPtgBase)ptg;
                    if (ref2.IsFirstRowRelative)
                        ref2.FirstRow += rangeY1Shift;
                    if (ref2.IsLastRowRelative)
                        ref2.LastRow += rangeY2Shift.Value;
                    if (ref2.IsFirstColRelative)
                        ref2.FirstColumn += rangeX1Shift;
                    if (ref2.IsLastColRelative)
                        ref2.LastColumn += rangeX2Shift.Value;
                }
                //else
                //    throw new Exception("Unexpected ptg type: " + ptg.GetType());
            }
            formulaCell.CellFormula = FormulaRenderer.ToFormulaString((IFormulaRenderingWorkbook)evaluationWorkbook, ptgs);
        }

        static public void Highlight(this ICell cell, Excel.Color color)
        {
            cell.CellStyle = Excel.highlight(cell.Sheet.Workbook, cell.CellStyle, color);
        }

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
    }
}