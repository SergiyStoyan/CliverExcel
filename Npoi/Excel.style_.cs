//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace Cliver
{
    public partial class Excel 
    {
        //public void Unhighlight(Color color)
        //{
        //    if (Workbook is XSSFWorkbook)
        //    {
        //        if (color == null)
        //        {
        //            foreach (XSSFCellStyle s in GetStyles())
        //                s.SetFillForegroundColor(null);
        //            return;
        //        }
        //        XSSFColor c = new XSSFColor(color.RGB);
        //        foreach (XSSFCellStyle s in GetStyles())
        //        {
        //            if (AreColorsEqual(s.FillForegroundColorColor, c))
        //                s.SetFillForegroundColor(null);
        //        }
        //    }
        //    else if (Workbook is HSSFWorkbook hw)
        //    {
        //        if (color == null)
        //        {
        //            foreach (HSSFCellStyle s in GetStyles())
        //                s.FillForegroundColor = 0;
        //            return;
        //        }
        //        HSSFPalette palette = hw.GetCustomPalette();
        //        HSSFColor c = palette.FindColor(color.RGB[0], color.RGB[1], color.RGB[2]);
        //        foreach (HSSFCellStyle s in GetStyles())
        //        {
        //            if (AreColorsEqual(color, c))
        //                s.FillForegroundColor = 0;
        //        }
        //    }
        //    else
        //        throw new Exception("Unsupported workbook type: " + Workbook.GetType().FullName);
        //}

        public ICellStyle Highlight(ICellStyle style, Color color, FillPattern fillPattern = FillPattern.SolidForeground, bool createOnlyUniqueStyle = true)
        {
            return highlight(this, style, createOnlyUniqueStyle, color, fillPattern);
        }

        /// <summary>
        /// Intended for either adding or removing backgound color.
        /// (!)When createUniqueStyleOnly, it is slow.
        /// </summary>
        /// <param name="excel"></param>
        /// <param name="style"></param>
        /// <param name="createUniqueStyleOnly"></param>
        /// <param name="color"></param>
        /// <param name="fillPattern"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        static internal ICellStyle highlight(Excel excel, ICellStyle style, bool createOnlyUniqueStyle, Color color, FillPattern fillPattern = FillPattern.SolidForeground)
        {
            return excel.Workbook._highlight(style, createOnlyUniqueStyle, color, fillPattern);
        }

        /// <summary>
        /// Looks for an equal style in the workbook and, if it does not exists, creates a new one.
        /// (!)Incidentally, there is a somewhat analogous method NPOI.SS.Util.CellUtil.SetCellStyleProperties() which is not as handy in use though.
        /// </summary>
        /// <param name="style">it is a style created by CreateUnregisteredStyle() and then modified as needed. But it can be a registered style, too.</param>
        /// <param name="unregisteredStyleWorkbook"></param>
        /// <returns></returns>
        public ICellStyle GetRegisteredStyle(ICellStyle unregisteredStyle, IWorkbook unregisteredStyleWorkbook = null)
        {
            return Workbook._GetRegisteredStyle(unregisteredStyle, unregisteredStyleWorkbook);
        }

        public IEnumerable<ICellStyle> FindEqualStyles(ICellStyle style, IWorkbook styleWorkbook = null)
        {
            return Workbook._FindEqualStyles(style, styleWorkbook);
        }

        /// <summary>
        /// Both styles can be unregistered. (!)However, font and format used by them must be registered in the respective workbooks.
        /// </summary>
        /// <param name="fromStyle"></param>
        /// <param name="toStyle"></param>
        /// <param name="toStyleWorkbook"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public ICellStyle CopyStyle(ICellStyle fromStyle, ICellStyle toStyle, IWorkbook toStyleWorkbook = null)
        {
            return Workbook._CopyStyle(fromStyle, toStyle, toStyleWorkbook);
        }

        public ICellStyle CreateUnregisteredStyle()
        {
            return Workbook._CreateUnregisteredStyle();
        }

        /// <summary>
        /// Creates an unregistered copy of a style.
        /// </summary>
        /// <param name="fromStyle"></param>
        /// <param name="cloneStyleWorkbook"></param>
        /// <returns></returns>
        public ICellStyle CloneUnregisteredStyle(ICellStyle fromStyle, IWorkbook cloneStyleWorkbook = null)
        {
            return Workbook._CloneUnregisteredStyle(fromStyle, cloneStyleWorkbook);

        }

        /// <summary>
        /// If the font does not exists, it is created.
        /// </summary>
        /// <param name="bold"></param>
        /// <param name="color"></param>
        /// <param name="fontHeight"></param>
        /// <param name="name"></param>
        /// <param name="italic"></param>
        /// <param name="strikeout"></param>
        /// <param name="fontSuperScript"></param>
        /// <param name="fontUnderlineType"></param>
        /// <returns></returns>
        public IFont GetRegisteredFont(bool bold, IndexedColors color, short fontHeightInPoints, string name, bool italic = false, bool strikeout = false, FontSuperScript typeOffset = FontSuperScript.None, FontUnderlineType underline = FontUnderlineType.None)
        {
            return Workbook._GetRegisteredFont(bold, color, fontHeightInPoints, name, italic, strikeout, typeOffset, underline);
        }

        public IEnumerable<ICellStyle> GetStyles()
        {
            return Workbook._GetStyles();
        }

        public IEnumerable<ICellStyle> GetUnusedStyles(params short[] ignoredStyleIds)
        {
            return Workbook._GetUnusedStyles(ignoredStyleIds);
        }

        /// <summary>
        /// Makes all the duplicated styles unused. Call GetUnusedStyles() after this method to re-use styles.
        /// </summary>
        /// <exception cref="Exception"></exception>
        public void OptimiseStyles()
        {
            Workbook._OptimiseStyles();
        }

        public void ReplaceStyle(ICellStyle style1, ICellStyle style2)
        {
            Sheet._ReplaceStyle(style1, style2);
        }

        public void SetStyle(ICellStyle style, bool createCells)
        {
            Sheet._SetStyle(style, createCells);
        }

        public void UnsetStyle(ICellStyle style)
        {
            Sheet._UnsetStyle(style);
        }
    }
}