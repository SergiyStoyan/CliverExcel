//********************************************************************************************
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
        static public void _Remove(this ISheet sheet)
        {
            sheet.Workbook.RemoveSheetAt(sheet._GetIndex() - 1);
        }

        /// <summary>
        /// (!)The name will be corrected by altering unacceptable symbols.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="name"></param>
        static public void _Rename(this ISheet sheet, string name)
        {
            sheet.Workbook.SetSheetName(sheet._GetIndex() - 1, Excel.GetSafeSheetName(name));
        }

        static public string _Name(this ISheet sheet)
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

        static public Range _NewRange(this ISheet sheet, int y1 = 1, int x1 = 1, int? y2 = null, int? x2 = null)
        {
            return new Range(sheet, y1, x1, y2, x2);
        }
    }
}