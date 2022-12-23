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
        static public ICell GetCell(this IRow r, int x, bool create)
        {
            ICell c = r.GetCell(x - 1);
            if (c != null)
                return c;
            if (create)
                return r.CreateCell(x - 1);
            return null;
        }

        static public void Highlight(this IRow row, Excel.Color color)
        {
            row.RowStyle = Excel.highlight(row.Sheet.Workbook, row.RowStyle, color);
        }

        static public int GetLastUsedColumnInRow(this IRow row, bool includeMerged = true)
        {
            if (row == null || row.Cells.Count < 1)
                return -1;
            for (int x0 = row.Cells.Count - 1; x0 >= 0; x0--)
            {
                var c = row.Cells[x0];
                if (!string.IsNullOrWhiteSpace(c.GetValueAsString()))
                {
                    if (includeMerged)
                    {
                        var r = c.GetMergedRange();
                        if (r != null)
                            return r.LastX;
                    }
                    return c.ColumnIndex + 1;
                }
            }
            return -1;
        }
    }
}