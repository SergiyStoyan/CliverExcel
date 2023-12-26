//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;

namespace Cliver
{
    partial class Excel
    {
        /// <summary>
        /// Used to map styles between 2 workbooks with optional style altering.
        /// </summary>
        public class StyleMap : StyleCache
        {
            public StyleMap(IWorkbook fromWorkbook, IWorkbook toWorkbook) : base(fromWorkbook, toWorkbook)
            {
                if (fromWorkbook == toWorkbook)
                    throw new Exception("This class is intended for working with 2 different workbooks.");
            }

            /// <summary>
            /// Used for mappping styles between 2 workbooks
            /// </summary>
            /// <param name="style"></param>
            /// <param name="reuseUnusedStyle">(!)slows down performance. It makes sense ony when styles need optimization</param>
            /// <returns></returns>
            public ICellStyle GetMappedStyle(ICellStyle style, bool reuseUnusedStyle = false)
            {
                const long d = 1 << 48 - 1;
                long styleKey = (((long)style.Index) << 48) + d;//it uses octets not used by GetAlteredStyle()

                if (!style1Keys2style2.TryGetValue(styleKey, out ICellStyle s2))
                {
                    s2 = Workbook1._GetRegisteredStyle(s2, reuseUnusedStyle, Workbook2);
                    style1Keys2style2[styleKey] = s2;
                }
                return s2;
            }
        }
    }
}