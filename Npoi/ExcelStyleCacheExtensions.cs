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
        static public void _SetAlteredStyles<T>(this IRow row, Excel.StyleCache styleCache, T alterationKey, Excel.StyleCache.AlterStyle<T> alterStyle) where T : Excel.StyleCache.IKey
        {
            foreach (ICell cell in row.Cells)
                cell._SetAlteredStyle(styleCache, alterationKey, alterStyle);
        }

        static public void _SetAlteredStyle<T>(this ICell cell, Excel.StyleCache styleCache, T alterationKey, Excel.StyleCache.AlterStyle<T> alterStyle) where T : Excel.StyleCache.IKey
        {
            if (cell.Sheet.Workbook != styleCache.ToWorkbook)
                throw new Exception("Cell does not belong to styleCache's workbook.");
            cell.CellStyle = styleCache.GetAlteredStyle(cell.CellStyle, alterationKey, alterStyle);
        }
    }
}