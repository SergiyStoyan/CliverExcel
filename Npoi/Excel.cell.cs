//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula;
using NPOI.SS.Formula.PTG;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
<<<<<<< Updated upstream
=======
using System.Collections.Generic;
using static Cliver.Excel;
>>>>>>> Stashed changes

namespace Cliver
{
    public partial class Cell
    {
<<<<<<< Updated upstream
        public void ShiftCellsDown(int cellsY, int firstCellX, int lastCellX, int rowCount, Action<ICell> updateFormula = null)
        {
            for (int x = firstCellX; x <= lastCellX; x++)
            {
                for (int y = GetLastNotEmptyRowInColumn(x); y >= cellsY; y--)
                {
                    CopyCell(y, x, y + rowCount, x);
                    if (updateFormula == null)
                        continue;
                    ICell formulaCell = GetCell(y + rowCount, x, false);
                    if (formulaCell?.CellType != CellType.Formula)
                        continue;
                    updateFormula(formulaCell);
                }
                GetCell(cellsY, x, false)?.SetBlank();
            }
        }

        public void CopyCell(ICell source, ICell destination)
        {
            destination.SetBlank();
            destination.SetCellType(source.CellType);
            destination.CellStyle = source.CellStyle;
            destination.CellComment = source.CellComment;
            destination.Hyperlink = source.Hyperlink;
            switch (source.CellType)
            {
                case CellType.Formula:
                    destination.CellFormula = source.CellFormula;
                    break;
                case CellType.Numeric:
                    destination.SetCellValue(source.NumericCellValue);
                    break;
                case CellType.String:
                    destination.SetCellValue(source.StringCellValue);
                    break;
                case CellType.Boolean:
                    destination.SetCellValue(source.BooleanCellValue);
                    break;
                case CellType.Error:
                    destination.SetCellErrorValue(source.ErrorCellValue);
                    break;
                case CellType.Blank:
                    destination.SetBlank();
                    break;
                default:
                    throw new Exception("Unknown cell type: " + source.CellType);
            }
        }

        public ICell CopyCell(ICell sourceCell, int destinationY, int destinationX)
        {
            if (sourceCell == null)
            {
                IRow destinationRow = GetRow(destinationY, false);
                if (destinationRow == null)
                    return null;
                ICell destinationCell = destinationRow.GetCell(destinationX, false);
                if (destinationCell == null)
                    return destinationCell;
                destinationRow.RemoveCell(destinationCell);
                return destinationCell;
            }
            else
            {
                ICell destinationCell = GetCell(destinationY, destinationX, true);
                CopyCell(sourceCell, destinationCell);
                return destinationCell;
            }
        }

        public void MoveCell(ICell sourceCell, int destinationY, int destinationX, Action<ICell> onFormulaCellMoved = null)
        {
            ICell destinationCell = CopyCell(sourceCell, destinationY, destinationX);
            if (sourceCell != null)
                sourceCell.Row.RemoveCell(sourceCell);
            if (destinationCell?.CellType == CellType.Formula)
                onFormulaCellMoved?.Invoke(destinationCell);
        }

        public void CopyCell(int sourceY, int sourceX, int destinationY, int destinationX)
        {
            ICell sourceCell = GetCell(sourceY, sourceX, false);
            CopyCell(sourceCell, destinationY, destinationX);
        }

        public void MoveCell(int sourceY, int sourceX, int destinationY, int destinationX, Action<ICell> onFormulaCellMoved = null)
        {
            ICell sourceCell = GetCell(sourceY, sourceX, false);
            MoveCell(sourceCell, destinationY, destinationX, onFormulaCellMoved);
        }

        public ICell GetCell(int y, int x, bool create)
        {
            IRow r = GetRow(y, create);
            if (r == null)
                return null;
            return r.GetCell(x, create);
        }

        public ICell GetCell(string address, bool create)
        {
            var cs = GetCoordinates(address);
            IRow r = GetRow(cs.Y, create);
            if (r == null)
                return null;
            return r.GetCell(cs.X, create);
=======
        public Cell(ICell cell)
        {
            _ = cell;
        }
        public ICell _ { get; private set; }

        public Sheet GetSheet()
        {
            return new Sheet(_.Sheet);
        }

        public void Move(int toCellY, int toCellX, Action<Cell> onFormulaCellMoved = null, Sheet toSheet = null)
        {
            Cell toCell = Copy(toCellY, toCellX, toSheet);
            _.Row.RemoveCell(_);
            if (toCell?._.CellType == CellType.Formula)
                onFormulaCellMoved?.Invoke(toCell);
        }

        public Cell Copy(int toCellY, int toCellX, Sheet toSheet = null)
        {
            if (toSheet == null)
                toSheet = GetSheet();
            //if (fromCell == null)
            //{
            //    IRow toRow = toSheet.GetRow(toCellY, false);
            //    if (toRow == null)
            //        return null;
            //    ICell toCell = toRow.GetCell(toCellX, false);
            //    if (toCell == null)
            //        return toCell;
            //    toRow.RemoveCell(toCell);
            //    return toCell;
            //}
            Cell toCell = toSheet.GetCell(toCellY, toCellX, true);
            Copy(toCell);
            return toCell;
        }

        public void Copy(Cell toCell)
        {
            toCell._.SetBlank();
            toCell._.SetCellType(_.CellType);
            toCell._.CellStyle = _.CellStyle;
            toCell._.CellComment = _.CellComment;
            toCell._.Hyperlink = _.Hyperlink;
            switch (_.CellType)
            {
                case CellType.Formula:
                    toCell._.CellFormula = _.CellFormula;
                    break;
                case CellType.Numeric:
                    toCell._.SetCellValue(_.NumericCellValue);
                    break;
                case CellType.String:
                    toCell._.SetCellValue(_.StringCellValue);
                    break;
                case CellType.Boolean:
                    toCell._.SetCellValue(_.BooleanCellValue);
                    break;
                case CellType.Error:
                    toCell._.SetCellErrorValue(_.ErrorCellValue);
                    break;
                case CellType.Blank:
                    toCell._.SetBlank();
                    break;
                default:
                    throw new Exception("Unknown _ type: " + _.CellType);
            }
        }

        public string GetValueAsString( bool allowNull = false)
        {
            object o = GetValue();
            if (!allowNull && o == null)
                return string.Empty;
            if (o is DateTime dt)
                return dt.ToString("yyyy-MM-dd hh:mm:ss");
            return o?.ToString();
        }

        public object GetValue()
        {
            switch (_.CellType)
            {
                case CellType.Unknown:
                    //return _.ToString();
                    throw new Exception("Needs debugging for this _ type: " + _.CellType);
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(_))
                    {
                        try
                        {
                            return _.DateCellValue;
                        }
                        catch (Exception e)//!!!bug in NPOI2.5.1: after called Save(), it throws NullReferenceException: GetLocaleCalendar()  https://github.com/nissl-lab/npoi/issues/358
                        {
                            //Log.Warning("NPOI bug", e);
                            return DateTime.FromOADate(_.NumericCellValue);
                        }
                        //return formatter.FormatCellValue(c);
                    }
                    return _.NumericCellValue;
                case CellType.String:
                    return _.StringCellValue;
                case CellType.Boolean:
                    return _.BooleanCellValue;
                case CellType.Formula:
                    //return c.CellFormula;
                    IFormulaEvaluator formulaEvaluator;
                    if (_.Sheet.Workbook is XSSFWorkbook)
                        formulaEvaluator = new XSSFFormulaEvaluator(_.Sheet.Workbook);
                    else if (_.Sheet.Workbook is HSSFWorkbook)
                        formulaEvaluator = new HSSFFormulaEvaluator(_.Sheet.Workbook);
                    else
                        throw new Exception("Unsupported workbook type: " + _.Sheet.Workbook.GetType().FullName);
                    var cv = formulaEvaluator.Evaluate(_);
                    switch (cv.CellType)
                    {
                        case CellType.Unknown:
                            //return cv.ToString();
                            throw new Exception("Needs debugging for this _ type: " + _.CellType);
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
                    return FormulaError.ForInt(_.ErrorCellValue).String;
                case CellType.Blank:
                    return null;
                default:
                    throw new Exception("Unknown type: " + _.CellType);
            }
        }

        public void SetValue( object value)
        {
            if (value == null)
            {
                _.SetBlank();
                return;
            }
            if (value is double d)
            {
                _.SetCellValue(d);
                return;
            }
            if (value is bool b)
            {
                _.SetCellValue(b);
                return;
            }
            if (value is DateTime dt)
            {
                _.SetCellValue(dt);
                return;
            }
            _.SetCellValue(value?.ToString());
        }

        public Uri GetLink()
        {
            if (_ == null)
                return null;
            if (_.Hyperlink == null)
                return null;
            return new Uri(_.Hyperlink.Address, UriKind.RelativeOrAbsolute);
>>>>>>> Stashed changes
        }

        public void SetLink( Uri uri)
        {
            if (uri == null)
            {
                //if (_.GetValueAsString() == LinkEmptyValueFiller)
                //    _.SetCellValue("");
                _.Hyperlink = null;
                return;
            }
            if (string.IsNullOrEmpty(GetValueAsString()))
                _.SetCellValue(LinkEmptyValueFiller);
            if (_.Sheet.Workbook is XSSFWorkbook)
                _.Hyperlink = new XSSFHyperlink(HyperlinkType.Url) { Address = uri.ToString() };
            else if (_.Sheet.Workbook is HSSFWorkbook)
                _.Hyperlink = new HSSFHyperlink(HyperlinkType.Url) { Address = uri.ToString() };
            else
                throw new Exception("Unsupported workbook type: " + _.Sheet.Workbook.GetType().FullName);
        }
        public string LinkEmptyValueFiller = "           ";

        public void UpdateFormulaRange(int rangeY1Shift, int rangeX1Shift, int? rangeY2Shift = null, int? rangeX2Shift = null)
        {
            if (_.CellType != CellType.Formula)
                return;

            if (rangeY2Shift == null)
                rangeY2Shift = rangeY1Shift;
            if (rangeX2Shift == null)
                rangeX2Shift = rangeX1Shift;

            IFormulaParsingWorkbook evaluationWorkbook;
            if (_.Sheet.Workbook is XSSFWorkbook)
                evaluationWorkbook = XSSFEvaluationWorkbook.Create(_.Sheet.Workbook);
            else if (_.Sheet.Workbook is HSSFWorkbook)
                evaluationWorkbook = HSSFEvaluationWorkbook.Create(_.Sheet.Workbook);
            //else if (sheet is SXSSFWorkbook)
            //{
            //    evaluationWorkbook = SXSSFEvaluationWorkbook.Create((SXSSFWorkbook)Workbook);
            else
                throw new Exception("Unsupported workbook type: " + _.Sheet.Workbook.GetType().FullName);

            var ptgs = FormulaParser.Parse(_.CellFormula, evaluationWorkbook, FormulaType.Cell, _.Sheet.Workbook.GetSheetIndex(_.Sheet));
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
            _.CellFormula = FormulaRenderer.ToFormulaString((IFormulaRenderingWorkbook)evaluationWorkbook, ptgs);
        }

        // public void Highlight( ICellStyle style, Excel.Color color)
        //{
        //    _.CellStyle = Excel.highlight(_.Sheet.Workbook, style, color);
        //}

<<<<<<< Updated upstream
        //public void Highlight(ICell cell, ICellStyle style, Color color)
        //{
        //    cell.Highlight(style, color);
        //}
=======
        public Excel.Range GetMergedRange()
        {
            return GetSheet().getMergedRange(Y(), X());
        }

        public void ClearMerging()
        {
            for (int i = _.Sheet.MergedRegions.Count - 1; i >= 0; i--)
                if (_.Sheet.MergedRegions[i].IsInRange(_.RowIndex, _.ColumnIndex))
                {
                    _.Sheet.RemoveMergedRegion(i);
                    return;//there can be only one MergedRegion
                }
        }

        /// <summary>
        /// Cell's 1-based row index on the sheet.
        /// </summary>
        /// <param name="_"></param>
        /// <returns>1-based</returns>
        public int Y()
        {
            return _.RowIndex + 1;
        }

        /// <summary>
        /// Cell's 1-based column index on the sheet.
        /// </summary>
        /// <param name="_"></param>
        /// <returns>1-based</returns>
        public int X()
        {
            return _.ColumnIndex + 1;
        }

        public void CreateDropdown<T>( IEnumerable<T> values, T value, bool allowBlank = true)
        {
            List<string> vs = new List<string>();
            foreach (object v in values)
                vs.Add(v?.ToString());

            IDataValidationHelper dvh;
            if (_.Sheet is XSSFSheet)
                dvh = new XSSFDataValidationHelper((XSSFSheet)_.Sheet);
            else if (_.Sheet is HSSFSheet)
                dvh = new HSSFDataValidationHelper((HSSFSheet)_.Sheet);
            else
                throw new Exception("Unsupported workbook type: " + _.Sheet.Workbook.GetType().FullName);
            //string dvs = string.Join(",", vs);
            //IDataValidationConstraint dvc = Sheet.GetDataValidations().Find(a => string.Join(",", a.ValidationConstraint.ExplicitListValues) == dvs)?.ValidationConstraint;
            //if (dvc == null)
            //dvc = dvh.CreateCustomConstraint(dvs);
            IDataValidationConstraint dvc = dvh.CreateExplicitListConstraint(vs.ToArray());
            CellRangeAddressList cral = new CellRangeAddressList(_.RowIndex, _.RowIndex, _.ColumnIndex, _.ColumnIndex);
            IDataValidation dv = dvh.CreateValidation(dvc, cral);
            dv.SuppressDropDownArrow = true;
            dv.EmptyCellAllowed = allowBlank;
            ((XSSFSheet)_.Sheet).AddValidationData(dv);

            _.SetCellValue(value?.ToString());
        }

        public IEnumerable<Excel.Image> GetImages()
        {
            return GetSheet().GetImages(Y(), X());
        }
>>>>>>> Stashed changes
    }
}