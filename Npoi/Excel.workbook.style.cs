//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System.Collections.Generic;

namespace Cliver
{
    public partial class Excel
    {
        /// <summary>
        /// Intended for either adding or removing backgound color.
        /// The style can be unregistered but on HSSFWorkbook the color will be added to the palette.
        /// </summary>
        /// <param name="style"></param>
        /// <param name="color"></param>
        /// <param name="fillPattern"></param>
        /// <returns></returns>
        public void Highlight(ICellStyle style, Excel.Color color, FillPattern fillPattern = FillPattern.SolidForeground)
        {
            Workbook._Highlight(style, color, fillPattern);
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
        /// Font and format, if do not exist in the destination workbook, will be created there.
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

        public IFont _CreateUnregisteredFont()
        {
            return Workbook._CreateUnregisteredFont();
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
        /// Creates an unregistered copy of a font.
        /// </summary>
        /// <param name="font"></param>
        /// <returns></returns>
        public IFont CloneUnregisteredFont(IFont font)
        {
            return Workbook._CloneUnregisteredFont(font);
        }

        /// <summary>
        /// If the font does not exists, it is created.
        /// </summary>
        /// <param name="font"></param>
        /// <returns></returns>
        public IFont GetRegisteredFont(IFont font)
        {
            return Workbook._GetRegisteredFont(font);
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
        public IFont GetRegisteredFont(bool bold, short color, short fontHeightInPoints, string name, bool italic = false, bool strikeout = false, FontSuperScript typeOffset = FontSuperScript.None, FontUnderlineType underline = FontUnderlineType.None)
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
        /// (!)Tends to be slow on big sheets.
        /// </summary>
        /// <exception cref="Exception"></exception>
        public void OptimiseStyles()
        {
            Workbook._OptimiseStyles();
        }
    }
}