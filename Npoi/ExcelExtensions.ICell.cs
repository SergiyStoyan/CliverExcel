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
using NPOI.Util;
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
        static public void _SetAlteredStyle<T>(this ICell cell, T alterationKey, Excel.StyleCache.AlterStyle<T> alterStyle, bool reuseUnusedStyle = false) where T : Excel.StyleCache.IKey
        {
            _ = cell ?? throw new ArgumentNullException(nameof(cell));
            cell.CellStyle = cell.Sheet.Workbook._Excel().OneWorkbookStyleCache.GetAlteredStyle(cell.CellStyle, alterationKey, alterStyle, reuseUnusedStyle);
        }

        static public string _GetAddress(this ICell cell)
        {
            return cell?.Address.ToString();
        }

        static public void _Remove(this ICell cell, bool removeComment)
        {
            _ = cell ?? throw new ArgumentNullException(nameof(cell));
            cell.Sheet._RemoveCell(cell._Y(), cell._X(), removeComment);
        }

        static public ICell _Move(this ICell cell1, int y2, int x2, CopyCellMode copyCellMode = null, ISheet sheet2 = null, StyleMap styleMap = null)
        {
            _ = cell1 ?? throw new ArgumentNullException(nameof(cell1));
            return cell1.Sheet._MoveCell(cell1._Y(), cell1._X(), y2, x2, copyCellMode, sheet2, styleMap);
        }

        static public void _Move(this ICell cell1, ICell cell2, CopyCellMode copyCellMode = null, StyleMap styleMap = null)
        {
            _ = cell1 ?? throw new ArgumentNullException(nameof(cell1));
            _ = cell2 ?? throw new ArgumentNullException(nameof(cell2));
            cell1.Sheet._MoveCell(cell1._Y(), cell1._X(), cell2._Y(), cell2._X(), copyCellMode, cell2.Sheet, styleMap);
        }

        static public ICell _Copy(this ICell cell1, int y2, int x2, CopyCellMode copyCellMode = null, ISheet sheet2 = null, StyleMap styleMap = null)
        {
            _ = cell1 ?? throw new ArgumentNullException(nameof(cell1));
            return cell1.Sheet._CopyCell(cell1._Y(), cell1._X(), y2, x2, copyCellMode, sheet2, styleMap);
        }

        static public void _Copy(this ICell cell1, ICell cell2, CopyCellMode copyCellMode = null, StyleMap styleMap = null)
        {
            _ = cell1 ?? throw new ArgumentNullException(nameof(cell1));
            _ = cell2 ?? throw new ArgumentNullException(nameof(cell2));
            cell1.Sheet._CopyCell(cell1._Y(), cell1._X(), cell2._Y(), cell2._X(), copyCellMode, cell2.Sheet, styleMap);
        }

        /// <summary>
        /// NULL- and type-safe.
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="allowNull"></param>
        /// <returns></returns>
        static public string _GetValueAsString(this ICell cell, StringMode stringMode = DefaultStringMode)
        {
            object o = cell?._GetValue();
            if (o == null)
                return stringMode.HasFlag(StringMode.NotNull) ? string.Empty : null;
            if (o is DateTime dt)
                return dt.ToString("yyyy-MM-dd HH:mm:ss");
            string s = o?.ToString();
            if (s == null && stringMode.HasFlag(StringMode.NotNull))
                s = string.Empty;
            if (stringMode.HasFlag(StringMode.Trim))
                return s?.Trim();
            return s;
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
            _ = cell ?? throw new ArgumentNullException(nameof(cell));

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

        ///// <summary>
        ///// (!)Some cells (made by a thrid-part app?) can have multiple links. NPOI gets the first one, while Excel gets the last one which is considered correct.
        ///// This methods gets the last one.
        ///// </summary>
        ///// <param name="cell"></param>
        ///// <returns></returns>
        //static public string _GetLink(this ICell cell)
        //{
        //    return cell?.Sheet.GetHyperlinkList()
        //            .LastOrDefault(a => a.FirstColumn == cell.ColumnIndex && a.FirstRow == cell.RowIndex && a.LastColumn == cell.ColumnIndex && a.LastRow == cell.RowIndex)
        //            ?.Address;//(!)HACK: third-party files can have multiple links where the last one seems to be correct
        //}

        /// <summary>
        /// (!)Some cells (made by a thrid-part app?) can have multiple links. NPOI gets the first one, while Excel seems to get the last one which is considered correct.
        /// (!)This method follows NPOI routine because it is faster. To get links corrected once, call ISheet._FixLinks().
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="urlUnescapeFileType">
        /// Usually you expect that file links are url-escaped and need to be unescaped.
        /// It is TRUE by default because _SetLink() does url-escape and Excel too does so.
        /// However, some app can set links as they are, dave the file and then retrieve them as they are.
        /// Also, Excel still might treat unescaped links properly but then it will url-escape them when saving the file.
        /// So, links can be either encoded or not encoded!
        /// Unfortunately it is impossible to say if a link is url-escaped with 100% confidence. So, decision is left to the caller. 
        ///</param>
        /// <returns></returns>
        static public string _GetLink(this ICell cell, bool urlUnescapeFileType = true)
        {
            var h = cell?.Hyperlink;
            string link = h?.Address;
            //return cell?.Sheet.GetHyperlinkList()
            //        .LastOrDefault(a => a.FirstColumn == cell.ColumnIndex && a.FirstRow == cell.RowIndex && a.LastColumn == cell.ColumnIndex && a.LastRow == cell.RowIndex)
            //        ?.Address;//(!)HACK: third-party files can have multiple links where the last one seems to be correct
 
            if (urlUnescapeFileType && h.Type == HyperlinkType.File && link.Contains('%'))//(!)# cannot be used in excel links (too, unescaped % may lead to confusion) so such paths are url-escaped.
                link = Uri.UnescapeDataString(link);
            return link;
        }

        /// <summary>
        /// (!)Removes all the old links for the cell, if any.
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="link"></param>
        /// <param name="hyperlinkType">by default the type is detected by the method</param>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="Exception"></exception>
        static public void _SetLink(this ICell cell, string link, HyperlinkType hyperlinkType = HyperlinkType.Unknown)
        {
            _ = cell ?? throw new ArgumentNullException(nameof(cell));

            while (cell.Hyperlink != null)//(!)BUG: there can be multiple links per cell
                cell.RemoveHyperlink();//(!)necessary to get rid of all the old links if any. Otherwise sometimes the old link is not overriden by the new one.

            if (link == null)
            {
                //if (cell.GetValueAsString() == LinkEmptyValueFiller)
                //    cell.SetCellValue("");
                return;
            }

            if (string.IsNullOrWhiteSpace(cell._GetValueAsString()))
                cell.SetCellValue(cell.Sheet.Workbook._Excel().LinkEmptyValueFiller);

            if (hyperlinkType == HyperlinkType.Unknown)
            {
                if (Regex.IsMatch(link, @"^\s*(https?|ftps?)\:", RegexOptions.IgnoreCase))
                    hyperlinkType = HyperlinkType.Url;
                else if (Regex.IsMatch(link, @"^\s*(file\:\/\/\/)?[a-z]\:", RegexOptions.IgnoreCase))
                {
                    hyperlinkType = HyperlinkType.File;

                    if (link.Contains('%')//(!)Testing showed that Excel can properly treat non-url-encoded links with % but when it saves the file, it url-encodes them. So, GetLink() gets them url-encoded!
                         || link.Contains('#')//(!)# cannot be used in excel links. As escaped links look wrong in the popup, escape only when necessary.
                    //https://support.microsoft.com/en-gb/topic/you-cannot-use-a-pound-character-in-a-file-name-for-a-hyperlink-in-an-office-program-3dc41767-a82e-fc9b-c405-de8b1166be92
                        )
                    {
                        //string[] ps = Regex.Split(link, Regex.Escape(System.IO.Path.DirectorySeparatorChar.ToString()));
                        //for (int i = 0; i < ps.Length; i++)
                        //{
                        //    string p = ps[i];
                        //    if (p.Contains('%') || p.Contains('#'))
                        //        ps[i] = Uri.EscapeDataString(p);
                        //}
                        //link = string.Join(System.IO.Path.DirectorySeparatorChar.ToString(), ps);
                        link = Uri.EscapeDataString(link);
                        //link = Regex.Replace(link, @"\#", "%23");!!!does not work if there are spaces
                    }
                }
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

            //if (link != cell.Hyperlink.Address)
            //{
            //    var ls2 = cell.Sheet.GetHyperlinkList().Where(a => a.FirstColumn == cell.ColumnIndex && a.FirstRow == cell.RowIndex).ToList();
            //    throw new Exception("Could not set link: " + link);
            //}
        }

        /// <summary>
        /// It automatically updates the cell formula when moving a range of cells.
        /// It is expected to work properly for trivial formulas.
        /// (!)Check carefully if it does what you need. If does not, copy this method and customize.
        /// </summary>
        /// <param name="formulaCell"></param>
        /// <param name="rangeY1"></param>
        /// <param name="rangeX1"></param>
        /// <param name="rangeY2"></param>
        /// <param name="rangeX2"></param>
        /// <param name="yShift"></param>
        /// <param name="xShift"></param>
        /// <exception cref="Exception"></exception>
        static public void _UpdateFormulaOnMovingCellRange(this ICell formulaCell, int rangeY1, int rangeX1, int rangeY2, int rangeX2, int yShift, int xShift)
        {
            if (formulaCell?.CellType != CellType.Formula)
                return;

            IFormulaParsingWorkbook evaluationWorkbook;
            if (formulaCell.Sheet.Workbook is XSSFWorkbook)
                evaluationWorkbook = XSSFEvaluationWorkbook.Create(formulaCell.Sheet.Workbook);
            else if (formulaCell.Sheet.Workbook is HSSFWorkbook)
                evaluationWorkbook = HSSFEvaluationWorkbook.Create(formulaCell.Sheet.Workbook);
            else
                throw new Exception("Unsupported workbook type: " + formulaCell.Sheet.Workbook.GetType().FullName);

            int r1 = rangeY1 - 1;
            int c1 = rangeX1 - 1;
            int r2 = rangeY2 - 1;
            int c2 = rangeX2 - 1;

            var ptgs = FormulaParser.Parse(formulaCell.CellFormula, evaluationWorkbook, FormulaType.Cell, formulaCell.Sheet.Workbook.GetSheetIndex(formulaCell.Sheet));
            foreach (Ptg ptg in ptgs)
            {
                if (ptg is RefPtgBase rpb)
                {
                    if (rpb.Row >= r1 && rpb.Row <= r2 && rpb.Column >= c1 && rpb.Column <= c2)
                    {
                        if (rpb.IsRowRelative)
                        {
                            rpb.Row = rpb.Row + yShift;
                            if (rpb.Row < 0)
                                rpb.Row = 0;
                        }
                        if (rpb.IsColRelative)
                        {
                            rpb.Column = rpb.Column + xShift;
                            if (rpb.Column < 0)
                                rpb.Column = 0;
                        }
                    }
                }
                else if (ptg is AreaPtgBase apb)
                {
                    if (apb.FirstRow >= r1 && apb.FirstRow <= r2 && apb.FirstColumn >= c1 && apb.FirstColumn <= c2)
                    {
                        if (apb.IsFirstRowRelative)
                        {
                            apb.FirstRow += yShift;
                            if (apb.FirstRow < 0)
                                apb.FirstRow = 0;
                        }
                        if (apb.IsFirstColRelative)
                        {
                            apb.FirstColumn += xShift;
                            if (apb.FirstColumn < 0)
                                apb.FirstColumn = 0;
                        }
                    }
                    if (apb.LastRow >= r1 && apb.LastRow <= r2 && apb.LastColumn >= c1 && apb.LastColumn <= c2)
                    {
                        if (apb.IsLastRowRelative)
                        {
                            apb.LastRow += yShift;
                            if (apb.LastRow < 0)
                                apb.LastRow = 0;
                        }
                        if (apb.IsLastColRelative)
                        {
                            apb.LastColumn += xShift;
                            if (apb.LastColumn < 0)
                                apb.LastColumn = 0;
                        }
                    }
                }
                else
                {
                    //throw new Exception("Unexpected ptg type: " + ptg.GetType());
                }
            }
            formulaCell.CellFormula = FormulaRenderer.ToFormulaString((IFormulaRenderingWorkbook)evaluationWorkbook, ptgs);
        }

        static public Excel.Range _GetMergedRange(this ICell cell)
        {
            return cell?.Sheet._GetMergedRange(cell.RowIndex + 1, cell.ColumnIndex + 1);
        }

        static public void _ClearMerging(this ICell cell)
        {
            _ = cell ?? throw new ArgumentNullException(nameof(cell));
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
            _ = cell ?? throw new ArgumentNullException(nameof(cell));
            return cell.RowIndex + 1;
        }

        /// <summary>
        /// Cell's 1-based column index on the sheet.
        /// </summary>
        /// <param name="cell"></param>
        /// <returns>1-based</returns>
        static public int _X(this ICell cell)
        {
            _ = cell ?? throw new ArgumentNullException(nameof(cell));
            return cell.ColumnIndex + 1;
        }

        static public void _CreateDropdown<T>(this ICell cell, IEnumerable<T> values, T value, bool allowBlank = true)
        {
            _ = cell ?? throw new ArgumentNullException(nameof(cell));

            List<string> vs = new List<string>();
            foreach (object v in values)
                vs.Add(v?.ToString());

            IDataValidationHelper dvh = cell.Sheet.GetDataValidationHelper();
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

        static public void _RemoveDropdown<T>(this ICell cell)
        {
            _ = cell ?? throw new ArgumentNullException(nameof(cell));

            throw new NotImplementedException();
        }

        static public IEnumerable<Excel.Image> _GetImages(this ICell cell)
        {
            _ = cell ?? throw new ArgumentNullException(nameof(cell));

            return cell.Sheet._GetImages(cell._Y(), cell._X());
        }
    }
}