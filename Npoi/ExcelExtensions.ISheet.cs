﻿//********************************************************************************************
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
        /// (!)Some cells can have (made mistakenly by a thrid-party app?) multiple links.
        /// NPOI gets the first one while Excel seems to get the last one which is considered correct.
        /// This method removes all the links for each cell except the last one.
        /// </summary>
        /// <param name="sheet"></param>
        static public void _FixLinks(this ISheet sheet)
        {
            var ls = sheet.GetHyperlinkList().Where(a => a.FirstColumn == a.LastColumn && a.FirstRow == a.LastRow)
                  .OrderBy(a => a.FirstColumn * 1000 + a.FirstRow).ToList();
            if (ls.Count < 1)
                return;
            IHyperlink lastLink = ls[0];
            int cellLinkCount = 1;
            for (int i = 1; i < ls.Count; i++)
            {
                var l = ls[i];
                if (lastLink.FirstColumn != l.FirstColumn
                    || lastLink.FirstRow != l.FirstRow
                   )
                {
                    setLastLink();
                    cellLinkCount = 1;
                }
                lastLink = l;
                cellLinkCount++;
            }
            setLastLink();
            void setLastLink()
            {
                var r = sheet.GetRow(lastLink.FirstRow);
                if (r == null)
                {
                    r = sheet.CreateRow(lastLink.FirstRow);
                    var c = r.CreateCell(lastLink.FirstColumn);
                    while (c.Hyperlink != null)
                        c.RemoveHyperlink();//the only way to remove stray link
                    sheet.RemoveRow(r);
                    return;
                }
                {
                    var c = r.GetCell(lastLink.FirstColumn);
                    if (c == null)
                    {
                        c = r.CreateCell(lastLink.FirstColumn);
                        while (c.Hyperlink != null)
                            c.RemoveHyperlink();//the only way to remove stray link
                        r.RemoveCell(c);
                        return;
                    }
                    if (c.Hyperlink?.Address != lastLink.Address || cellLinkCount > 1)
                        c._SetLink(lastLink.Address, lastLink.Type);
                }
            }
        }

        static public void _Remove(this ISheet sheet)
        {
            sheet.Workbook.RemoveSheetAt(sheet._GetIndex() - 1);
        }

        /// <summary>
        /// (!)The name will be corrected by altering unacceptable symbols.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="name"></param>
        static public void _SetName(this ISheet sheet, string name)
        {
            sheet.Workbook.SetSheetName(sheet._GetIndex() - 1, Excel.GetSafeSheetName(name));
        }

        static public string _GetName(this ISheet sheet)
        {
            return sheet.SheetName;
        }

        static public int _GetIndex(this ISheet sheet)
        {
            return sheet.Workbook.GetSheetIndex(sheet.SheetName) + 1;
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

        static public Range _GetRange(this ISheet sheet, int y1 = 1, int x1 = 1, int? y2 = null, int? x2 = null)
        {
            return new Range(sheet, y1, x1, y2, x2);
        }

        /// <summary>
        /// It automatically updates all the formula cells in the sheet when moving a range of cells.
        /// It is expected to work properly for trivial formulas.
        /// (!)Check carefully if it does what you need. If does not, copy this method and customize.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rangeY1"></param>
        /// <param name="rangeX1"></param>
        /// <param name="rangeY2"></param>
        /// <param name="rangeX2"></param>
        /// <param name="excludeRange">set it TRUE if formula cells within the moving range has been updated already by the shifting method</param>
        /// <param name="yShift"></param>
        /// <param name="xShift"></param>
        static public void _UpdateFormulasOnMovingCellRange(this ISheet sheet, int rangeY1, int rangeX1, int rangeY2, int rangeX2, bool excludeRange, int yShift, int xShift)
        {
            foreach (IRow r in sheet._GetRows(RowScope.NotNull))
                foreach (ICell c in r.Cells.Where(a => a.CellType == CellType.Formula))
                {
                    if (excludeRange && c._Y() >= rangeY1 && c._Y() <= rangeY2 && c._X() >= rangeX1 && c._X() <= rangeX2)
                        continue;
                    c._UpdateFormulaOnMovingCellRange(rangeY1, rangeX1, rangeY2, rangeX2, yShift, xShift);
                }
        }
    }
}