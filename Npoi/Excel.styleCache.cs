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
        /// Collection of styles that are automaticlly created by altering existing styles or copying styles to another workbook.
        /// It helps when you need to alter certain style parameters in some cells, e.g. set a new color but you do not know in advance which styles you will alter/copy.
        /// It takes care about registering and caching all the styles needed in the workbook during editing.
        /// Used for:
        /// - not to trouble about duplicating styles in the workbook;
        /// - to eliminate need for matching styles in the workbook;
        /// </summary>
        public class StyleCache
        {
            public StyleCache(IWorkbook workbook1, IWorkbook workbook2 = null)
            {
                if (workbook2 == null)
                    workbook2 = workbook1;
                Workbook1 = workbook1;
                Workbook2 = workbook2;
            }

            readonly public IWorkbook Workbook1;
            readonly public IWorkbook Workbook2;

            Dictionary<long, ICellStyle> alternation_style1Keys2style2 = new Dictionary<long, ICellStyle>();

            /// <summary>
            /// Get a resgistered and cached style that is obtained by altering the input style.
            /// </summary>
            /// <param name="style">the style to be altered</param>
            /// <param name="alterationKey"></param>
            /// <param name="alterStyle">performs style alteration. (!)The passed in style is unregistered and must remain so.</param>
            /// <returns></returns>
            public ICellStyle GetAlteredStyle(ICellStyle style, IKey alterationKey, Action<ICellStyle> alterStyle)
            {
                long alteration_styleKey = (((long)alterationKey.Get()) << 16) + style.Index;

                if (!alternation_style1Keys2style2.TryGetValue(alteration_styleKey, out ICellStyle s2))
                {
                    s2 = Workbook1._CloneUnregisteredStyle(style, Workbook2);
                    alterStyle(s2);
                    s2 = Workbook1._GetRegisteredStyle(s2, Workbook2);
                    alternation_style1Keys2style2[alteration_styleKey] = s2;
                }
                return s2;
            }

            /// <summary>
            /// Used for mappping styles between workbooks
            /// </summary>
            /// <param name="style"></param>
            /// <returns></returns>
            public ICellStyle GetMappedStyle(ICellStyle style)
            {
                const long d = 1 << 48 - 1;
                long alteration_styleKey = (((long)style.Index) << 48) + d;//it uses octets not used by GetAlteredStyle()

                if (!alternation_style1Keys2style2.TryGetValue(alteration_styleKey, out ICellStyle s2))
                {
                    s2 = Workbook1._CloneUnregisteredStyle(style, Workbook2);
                    s2 = Workbook1._GetRegisteredStyle(s2, Workbook2);
                    alternation_style1Keys2style2[alteration_styleKey] = s2;
                }
                return s2;
            }

            public interface IKey
            {
                int Get();
            }

            /// <summary>
            /// Deafault implementation of IKey intended for altering styles in cells within one workbook.
            /// </summary>
            public class Key : IKey
            {
                public void Add(params byte[] subkeys)
                {
                    this.subkeys.AddRange(subkeys);
                }

                public void Add(params string[] subkeys)
                {
                    subkeys.ForEach(a => this.subkeys.AddRange(System.Text.Encoding.ASCII.GetBytes(a)));
                }

                protected List<byte> subkeys = new List<byte>();

                virtual public int Get()
                {
                    //return get64BitHash(subkeys);
                    return get32BitHash(subkeys);
                }

                protected int get32BitHash(List<byte> bytes)
                {
                    unchecked
                    {
                        const int p = 16777619;
                        uint hash = 2166136261;
                        foreach (var d in bytes)
                            hash = (hash ^ d) * p;
                        return (int)hash;
                    }
                }

                protected long get64BitHash(List<byte> bytes)
                {
                    unchecked
                    {
                        const ulong p = 1099511628211UL;
                        var hash = 14695981039346656037UL;
                        foreach (var d in bytes)
                            hash = (hash ^ d) * p;
                        return (long)hash;
                    }
                }
            }
        }

        protected StyleCache styleCache = null;

        public void SetStyles(IRow row, Excel.StyleCache.IKey alterationKey, Action<ICellStyle> alterStyle)
        {
            row._SetStyles(styleCache, alterationKey, alterStyle);
        }

        public void SetStyle(ICell cell, Excel.StyleCache.IKey alterationKey, Action<ICellStyle> alterStyle)
        {
            cell._SetStyle(styleCache, alterationKey, alterStyle);
        }
    }
}