//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text.RegularExpressions;

namespace Cliver
{
    public partial class Excel : IDisposable
    {
        public partial class Table
        {
            public int LastRowY()
            {
                return Sheet.LastRowNum + 1;
            }

            public int GetLastNotEmptyRow(bool includeMerged = false)
            {
                return Sheet._GetLastNotEmptyRowInColumnRange(includeMerged, 1, Columns[Columns.Count - 1].X);
            }

            //Can be used as a framework
            //public void Highlight(IRow row, Excel.Color color, int checkX = 0)
            //{
            //    ICell c0 = row.GetCell(checkX);
            //    if (Excel.AreColorsEqual(color, c0?.CellStyle?.FillForegroundColorColor))
            //        return;
            //    foreach (Column column in Columns)
            //    {
            //        ICell c = GetCell(row, column, true);
            //        ICellStyle style = c.CellStyle;
            //        style = style == null ? Excel.CreateUnregisteredStyle() : Excel.CloneUnregisteredStyle(style);
            //        Excel.Highlight(style, color);
            //        c.CellStyle = Excel.GetRegisteredStyle(style);
            //    }
            //}
        }
    }
}