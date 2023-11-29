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
    partial class Excel
    {
        /// <summary>
        /// Collection of styles that are automaticlly created by altering existing styles.
        /// It helps when you need to alter certain style parameters in some cells, e.g. set a new color but you do not know in advance which styles you will alter.
        /// It takes care about registering and caching all the styles needed in the workbook during editing.
        /// Used for:
        /// - not to trouble about duplicating styles in the workbook;
        /// - to increase performance by avoiding matching styles in the workbook;
        /// </summary>
        public class StyleCache
        {
            public StyleCache(Excel excel)
            {
                Excel = excel;
            }

            public Excel Excel { get; private set; }

            Dictionary<(long, short), ICellStyle> alternation_style1Keys2style2 = new Dictionary<(long, short), ICellStyle>();

            ICellStyle getAlteredStyle(ICellStyle style, (long, short) alterationKey, Action<ICellStyle> alterStyle)
            {
                if (!alternation_style1Keys2style2.TryGetValue(alterationKey, out ICellStyle s2))
                {
                    s2 = Excel.CloneUnregisteredStyle(style);
                    alterStyle(s2);
                    s2 = Excel.GetRegisteredStyle(s2);
                    alternation_style1Keys2style2[alterationKey] = s2;
                }
                return s2;
            }

            /// <summary>
            /// Get a resgistered and cached style that is obtained by altering the input style.
            /// </summary>
            /// <param name="style">the style to be altered</param>
            /// <param name="getAlterationKey">provides a key that is unique for the given style alteration, e.g. changing to a font. (!)It must be unique for all the planned alterations.</param>
            /// <param name="alterStyle">performs style alteration. (!)The passed in style is unregistered and must remain so.</param>
            /// <returns></returns>
            public ICellStyle GetAlteredStyle(ICellStyle style, Func<long> getAlterationKey, Action<ICellStyle> alterStyle)
            {
                return getAlteredStyle(style, (getAlterationKey(), style.Index), alterStyle);
            }

            /// <summary>
            /// Get a resgistered and cached style that is obtained by altering the input style.
            /// </summary>
            /// <param name="style">the style to be altered</param>
            /// <param name="getAlterationKey">provides a key that is unique for the given style alteration, e.g. changing to a font. (!)It must be unique for all the planned alterations.</param>
            /// <param name="alterStyle">performs style alteration. (!)The passed in style is unregistered and must remain so.</param>
            /// <returns></returns>
            public ICellStyle GetAlteredStyle(ICellStyle style, Func<string> getAlterationKey, Action<ICellStyle> alterStyle)
            {
                return getAlteredStyle(style, (getAlterationKey().GetHashCode(), style.Index), alterStyle);
            }
        }
    }
}