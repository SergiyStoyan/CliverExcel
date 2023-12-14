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
            public StyleCache(IWorkbook fromWorkbook, IWorkbook toWorkbook = null)
            {
                if (toWorkbook == null)
                    toWorkbook = fromWorkbook;
                FromWorkbook = fromWorkbook;
                ToWorkbook = toWorkbook;
            }

            readonly public IWorkbook FromWorkbook;
            readonly public IWorkbook ToWorkbook;

            protected Dictionary<long, ICellStyle> style1Keys2style2 = new Dictionary<long, ICellStyle>();

            /// <summary>
            /// Performs style alteration.
            /// (!)Style is unregistered and must remain so.
            /// </summary>
            /// <typeparam name="T"></typeparam>
            /// <param name="style"></param>
            /// <param name="alterationKey"></param>
            public delegate void AlterStyle<T>(ICellStyle style, T alterationKey) where T : Excel.StyleCache.IKey;

            /// <summary>
            /// Get a resgistered and cached style that is obtained by altering the input style.
            /// </summary>
            /// <param name="style">the style to be altered</param>
            /// <param name="alterationKey"></param>
            /// <param name="alterStyle">performs style alteration. (!)Style is unregistered and must remain so.</param>
            /// <param name="reuseUnusedStyle">(!)slows down performance. It makes sense ony when styles need optimization</param>
            /// <returns></returns>T
            public ICellStyle GetAlteredStyle<T>(ICellStyle style, T alterationKey, AlterStyle<T> alterStyle, bool reuseUnusedStyle = false) where T : Excel.StyleCache.IKey
            {
                long alteration_styleKey = (((long)alterationKey.Get()) << 16) + style.Index;

                if (!style1Keys2style2.TryGetValue(alteration_styleKey, out ICellStyle s2))
                {
                    s2 = FromWorkbook._CloneUnregisteredStyle(style, ToWorkbook);
                    alterStyle(s2, alterationKey);
                    s2 = FromWorkbook._GetRegisteredStyle(s2, reuseUnusedStyle, ToWorkbook);
                    style1Keys2style2[alteration_styleKey] = s2;
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
                    s2 = FromWorkbook._GetRegisteredStyle(s2, reuseUnusedStyle, ToWorkbook);
                    style1Keys2style2[styleKey] = s2;
                }
                return s2;
            }
        }

        ///// <summary>
        ///// Used only to cache styles within 1 workbook with optional style altering.
        ///// </summary>
        //public class StyleCache : StyleCacheBase
        //{
        //    public StyleCache(IWorkbook workbook) : base(workbook)
        //    { }

        //    ///// <summary>
        //    ///// Should be used only to cache styles within the same workbook (FromWorkbook=ToWorkbook2). 
        //    ///// Its goal is to increase performance by avoiding of calling _GetRegisteredStyle() each time.
        //    ///// (!)Input style with index>=0 is considered as registered in the actual workbook.
        //    ///// </summary>
        //    ///// <param name="style"></param>
        //    ///// <returns></returns>
        //    //public ICellStyle GetCachedStyle(ICellStyle style)
        //    //{
        //    //    if (style == null)
        //    //        return style;

        //    //    if (style.Index < 0)
        //    //        style = FromWorkbook._GetRegisteredStyle(style);

        //    //    const long d = 1 << 48 - 1;
        //    //    long styleKey = (((long)style.Index) << 48) + d;//it uses octets not used by GetAlteredStyle()

        //    //    if (!style1Keys2style2.TryGetValue(styleKey, out ICellStyle s2))
        //    //    {
        //    //        s2 = FromWorkbook._GetRegisteredStyle(s2);
        //    //        style1Keys2style2[styleKey] = s2;
        //    //    }
        //    //    return s2;
        //    //}
        //}
    }
}