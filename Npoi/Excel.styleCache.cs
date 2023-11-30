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

            //Dictionary<(long, short), ICellStyle> alternation_style1Keys2style2 = new Dictionary<(long, short), ICellStyle>();
            Dictionary<long, ICellStyle> alternation_style1Keys2style2 = new Dictionary<long, ICellStyle>();

            /// <summary>
            /// Get a resgistered and cached style that is obtained by altering the input style.
            /// </summary>
            /// <param name="style">the style to be altered</param>
            /// <param name="alterationKey"></param>
            /// <param name="alterStyle">performs style alteration. (!)The passed in style is unregistered and must remain so.</param>
            /// <returns></returns>
            public ICellStyle GetAlteredStyle(ICellStyle style, int alterationKey, Action<ICellStyle> alterStyle)
            {
                //var alteration_styleKey = (alterationKey, style.Index);
                var alteration_styleKey = (alterationKey << 16) + style.Index;

                if (!alternation_style1Keys2style2.TryGetValue(alteration_styleKey, out ICellStyle s2))
                {
                    s2 = Excel.CloneUnregisteredStyle(style);
                    alterStyle(s2);
                    s2 = Excel.GetRegisteredStyle(s2);
                    alternation_style1Keys2style2[alteration_styleKey] = s2;
                }
                return s2;
            }

            public class KeyBuilder
            {
                public void Add(params byte[] subkeys)
                {
                    this.subkeys.AddRange(subkeys);
                }

                public void Add(params string[] subkeys)
                {
                    subkeys.ForEach(a => this.subkeys.AddRange(System.Text.Encoding.ASCII.GetBytes(a)));
                }

                internal List<byte> subkeys = new List<byte>();

                //public long Get()
                //{
                //    return get64BitHash(subkeys);
                //}
                public int Get()
                {
                    return get32BitHash(subkeys);
                }

                int get32BitHash(List<byte> bytes)
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

                long get64BitHash(List<byte> bytes)
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

        public void SetStyles(IRow row, int alterationKey, Action<ICellStyle> alterStyle)
        {
            row._SetStyles(styleCache, alterationKey, alterStyle);
        }

        public void SetStyle(ICell cell, int alterationKey, Action<ICellStyle> alterStyle)
        {
            cell._SetStyle(styleCache, alterationKey, alterStyle);
        }
    }
}