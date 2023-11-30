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
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text.RegularExpressions;

namespace Cliver
{
    public static class ExcelStyleCacheExtensions
    {
        static public void _SetStyles(this IRow row, Excel.StyleCache styleCache, int alterationKey, Action<ICellStyle> alterStyle)
        {
            foreach (ICell cell in row.Cells)
                cell._SetStyle(styleCache, alterationKey, alterStyle);
        }

        static public void _SetStyle(this ICell cell, Excel.StyleCache styleCache, int alterationKey, Action<ICellStyle> alterStyle)
        {
            cell.CellStyle = styleCache.GetAlteredStyle(cell.CellStyle, alterationKey, alterStyle);
        }
    }
}