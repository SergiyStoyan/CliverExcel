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
                    s2 = Workbook1._CloneUnregisteredStyle(style, Workbook2);
                    alterStyle(s2, alterationKey);
                    s2 = Workbook1._GetRegisteredStyle(s2, reuseUnusedStyle, Workbook2);
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
                    s2 = Workbook1._GetRegisteredStyle(s2, reuseUnusedStyle, Workbook2);
                    style1Keys2style2[styleKey] = s2;
                }
                return s2;
            }
        }

        public class EasyStyleCache : StyleCache
        {
            public EasyStyleCache(IWorkbook workbook) : base(workbook)
            { }

            /// <summary>
            /// Copy the listed properties from unregistered maskStyle to style.
            /// </summary>
            /// <param name="style"></param>
            /// <param name="stylePropertyNames"></param>
            /// <param name="maskStyle"></param>
            /// <returns></returns>
            public ICellStyle GetAlteredStyle(ICellStyle style, IEnumerable<StyleProperty> stylePropertyNames, ICellStyle maskStyle)
            {
                Key alterationKey = new Key();
                Dictionary<StyleProperty, object> stylePropertieNames2Value = getStyleProperties(Workbook1, maskStyle, stylePropertyNames);
                stylePropertieNames2Value.ForEach(a => alterationKey.Add(a.Key, a.Value));
                void alterStyle(ICellStyle s, Key ak)
                {
                    setStyleProperties(Workbook1, stylePropertieNames2Value, s);
                }
                return GetAlteredStyle(style, alterationKey, alterStyle);
            }

            /// <summary>
            /// Set the listed properties to style.
            /// </summary>
            /// <param name="style"></param>
            /// <param name="stylePropertieNames2Value"></param>
            /// <returns></returns>
            public ICellStyle GetAlteredStyle(ICellStyle style, Dictionary<StyleProperty, object> stylePropertieNames2Value)
            {
                Key alterationKey = new Key();
                stylePropertieNames2Value.ForEach(a => alterationKey.Add(a.Key, a.Value));
                void alterStyle(ICellStyle s, Key ak)
                {
                    setStyleProperties(Workbook1, stylePropertieNames2Value, s);
                }
                return GetAlteredStyle(style, alterationKey, alterStyle);
            }

            public enum StyleProperty
            {
                Alignment,
                BorderBottom,
                BorderDiagonal,
                BorderDiagonalLineStyle,
                BorderLeft,
                BorderRight,
                BorderTop,
                DataFormat,
                FillForegroundColor,
                FillBackgroundColor,
                BorderDiagonalColor,
                BottomBorderColor,
                LeftBorderColor,
                RightBorderColor,
                TopBorderColor,
                FillPattern,
                Indention,
                IsHidden,
                IsLocked,
                Rotation,
                ShrinkToFit,
                VerticalAlignment,
                WrapText,
                Font,
            }

            static Dictionary<StyleProperty, object> getStyleProperties(IWorkbook workbook, ICellStyle style, IEnumerable<StyleProperty> propertyNames)
            {
                HashSet<StyleProperty> spns = new HashSet<StyleProperty>(propertyNames);

                Dictionary<StyleProperty, object> spns2V = new Dictionary<StyleProperty, object>();

                if (spns.Contains(StyleProperty.Alignment))
                    spns2V[StyleProperty.Alignment] = style.Alignment;
                if (spns.Contains(StyleProperty.BorderBottom))
                    spns2V[StyleProperty.BorderBottom] = style.BorderBottom;
                if (spns.Contains(StyleProperty.BorderDiagonal))
                    spns2V[StyleProperty.BorderDiagonal] = style.BorderDiagonal;
                if (spns.Contains(StyleProperty.BorderDiagonalLineStyle))
                    spns2V[StyleProperty.BorderDiagonalLineStyle] = style.BorderDiagonalLineStyle;
                if (spns.Contains(StyleProperty.BorderLeft))
                    spns2V[StyleProperty.BorderLeft] = style.BorderLeft;
                if (spns.Contains(StyleProperty.BorderRight))
                    spns2V[StyleProperty.BorderRight] = style.BorderRight;
                if (spns.Contains(StyleProperty.BorderTop))
                    spns2V[StyleProperty.BorderTop] = style.BorderTop;
                if (spns.Contains(StyleProperty.DataFormat))
                    spns2V[StyleProperty.DataFormat] = style.DataFormat;

                if (style is XSSFCellStyle xcs)
                {
                    if (spns.Contains(StyleProperty.FillForegroundColor))
                        spns2V[StyleProperty.FillForegroundColor] = xcs.FillForegroundColorColor;
                    if (spns.Contains(StyleProperty.FillBackgroundColor))
                        spns2V[StyleProperty.FillBackgroundColor] = xcs.FillBackgroundColorColor;
                    if (spns.Contains(StyleProperty.BorderDiagonalColor))
                        spns2V[StyleProperty.BorderDiagonalColor] = xcs.DiagonalBorderXSSFColor;
                    if (spns.Contains(StyleProperty.BottomBorderColor))
                        spns2V[StyleProperty.BottomBorderColor] = xcs.BottomBorderXSSFColor;
                    if (spns.Contains(StyleProperty.LeftBorderColor))
                        spns2V[StyleProperty.LeftBorderColor] = xcs.LeftBorderXSSFColor;
                    if (spns.Contains(StyleProperty.RightBorderColor))
                        spns2V[StyleProperty.RightBorderColor] = xcs.RightBorderXSSFColor;
                    if (spns.Contains(StyleProperty.TopBorderColor))
                        spns2V[StyleProperty.TopBorderColor] = xcs.TopBorderXSSFColor;
                }
                else if (style is HSSFCellStyle)
                {
                    if (spns.Contains(StyleProperty.FillForegroundColor))
                        spns2V[StyleProperty.FillForegroundColor] = style.FillForegroundColor;
                    if (spns.Contains(StyleProperty.FillBackgroundColor))
                        spns2V[StyleProperty.FillBackgroundColor] = style.FillBackgroundColor;
                    if (spns.Contains(StyleProperty.BorderDiagonalColor))
                        spns2V[StyleProperty.BorderDiagonalColor] = style.BorderDiagonalColor;
                    if (spns.Contains(StyleProperty.BottomBorderColor))
                        spns2V[StyleProperty.BottomBorderColor] = style.BottomBorderColor;
                    if (spns.Contains(StyleProperty.LeftBorderColor))
                        spns2V[StyleProperty.LeftBorderColor] = style.LeftBorderColor;
                    if (spns.Contains(StyleProperty.RightBorderColor))
                        spns2V[StyleProperty.RightBorderColor] = style.RightBorderColor;
                    if (spns.Contains(StyleProperty.TopBorderColor))
                        spns2V[StyleProperty.TopBorderColor] = style.TopBorderColor;
                }
                else
                    throw new Exception("Unsupported workbook type: " + workbook.GetType().FullName);

                if (spns.Contains(StyleProperty.FillPattern))
                    spns2V[StyleProperty.FillPattern] = style.FillPattern;
                if (spns.Contains(StyleProperty.Indention))
                    spns2V[StyleProperty.Indention] = style.Indention;
                if (spns.Contains(StyleProperty.IsHidden))
                    spns2V[StyleProperty.IsHidden] = style.IsHidden;
                if (spns.Contains(StyleProperty.IsLocked))
                    spns2V[StyleProperty.IsLocked] = style.IsLocked;
                if (spns.Contains(StyleProperty.Rotation))
                    spns2V[StyleProperty.Rotation] = style.Rotation;
                if (spns.Contains(StyleProperty.ShrinkToFit))
                    spns2V[StyleProperty.ShrinkToFit] = style.ShrinkToFit;
                if (spns.Contains(StyleProperty.VerticalAlignment))
                    spns2V[StyleProperty.VerticalAlignment] = style.VerticalAlignment;
                if (spns.Contains(StyleProperty.WrapText))
                    spns2V[StyleProperty.WrapText] = style.WrapText;
                if (spns.Contains(StyleProperty.Font))
                    spns2V[StyleProperty.Font] = workbook._GetFont(style);

                return spns2V;
            }

            static void setStyleProperties(IWorkbook workbook, Dictionary<StyleProperty, object> styleProperties2Value, ICellStyle style)
            {
                if (styleProperties2Value.TryGetValue(StyleProperty.Alignment, out object alignment))
                    style.Alignment = (HorizontalAlignment)alignment;
                if (styleProperties2Value.TryGetValue(StyleProperty.BorderBottom, out object borderBottom))
                    style.BorderBottom = (BorderStyle)borderBottom;
                if (styleProperties2Value.TryGetValue(StyleProperty.BorderDiagonal, out object borderDiagonal))
                    style.BorderDiagonal = (BorderDiagonal)borderDiagonal;
                if (styleProperties2Value.TryGetValue(StyleProperty.BorderDiagonalLineStyle, out object borderDiagonalLineStyle))
                    style.BorderDiagonalLineStyle = (BorderStyle)borderDiagonalLineStyle;
                if (styleProperties2Value.TryGetValue(StyleProperty.BorderLeft, out object borderLeft))
                    style.BorderLeft = (BorderStyle)borderLeft;
                if (styleProperties2Value.TryGetValue(StyleProperty.BorderRight, out object borderRight))
                    style.BorderRight = (BorderStyle)borderRight;
                if (styleProperties2Value.TryGetValue(StyleProperty.BorderTop, out object borderTop))
                    style.BorderTop = (BorderStyle)borderTop;
                if (styleProperties2Value.TryGetValue(StyleProperty.DataFormat, out object dataFormat))
                    style.DataFormat = (short)dataFormat;

                if (style is XSSFCellStyle xcs)
                {
                    if (styleProperties2Value.TryGetValue(StyleProperty.FillForegroundColor, out object fillForegroundColor))
                        xcs.FillForegroundColorColor = (IColor)fillForegroundColor;
                    if (styleProperties2Value.TryGetValue(StyleProperty.FillBackgroundColor, out object fillBackgroundColor))
                        xcs.FillBackgroundColorColor = (IColor)fillBackgroundColor;
                    if (styleProperties2Value.TryGetValue(StyleProperty.BorderDiagonalColor, out object borderDiagonalColor))
                        xcs.SetDiagonalBorderColor((XSSFColor)borderDiagonalColor);
                    if (styleProperties2Value.TryGetValue(StyleProperty.BottomBorderColor, out object bottomBorderColor))
                        xcs.SetBottomBorderColor((XSSFColor)bottomBorderColor);
                    if (styleProperties2Value.TryGetValue(StyleProperty.LeftBorderColor, out object leftBorderColor))
                        xcs.SetLeftBorderColor((XSSFColor)leftBorderColor);
                    if (styleProperties2Value.TryGetValue(StyleProperty.RightBorderColor, out object rightBorderColor))
                        xcs.SetRightBorderColor((XSSFColor)rightBorderColor);
                    if (styleProperties2Value.TryGetValue(StyleProperty.TopBorderColor, out object topBorderColor))
                        xcs.SetTopBorderColor((XSSFColor)topBorderColor);
                }
                else if (style is HSSFCellStyle)
                {
                    if (styleProperties2Value.TryGetValue(StyleProperty.FillForegroundColor, out object fillForegroundColor))
                        style.FillForegroundColor = (short)fillForegroundColor;
                    if (styleProperties2Value.TryGetValue(StyleProperty.FillBackgroundColor, out object fillBackgroundColor))
                        style.FillBackgroundColor = (short)fillBackgroundColor;
                    if (styleProperties2Value.TryGetValue(StyleProperty.BorderDiagonalColor, out object borderDiagonalColor))
                        style.BorderDiagonalColor = (short)borderDiagonalColor;
                    if (styleProperties2Value.TryGetValue(StyleProperty.BottomBorderColor, out object bottomBorderColor))
                        style.BottomBorderColor = (short)bottomBorderColor;
                    if (styleProperties2Value.TryGetValue(StyleProperty.LeftBorderColor, out object leftBorderColor))
                        style.LeftBorderColor = (short)leftBorderColor;
                    if (styleProperties2Value.TryGetValue(StyleProperty.RightBorderColor, out object rightBorderColor))
                        style.RightBorderColor = (short)rightBorderColor;
                    if (styleProperties2Value.TryGetValue(StyleProperty.TopBorderColor, out object topBorderColor))
                        style.TopBorderColor = (short)topBorderColor;
                }
                else
                    throw new Exception("Unsupported workbook type: " + workbook.GetType().FullName);

                if (styleProperties2Value.TryGetValue(StyleProperty.FillPattern, out object fillPattern))
                    style.FillPattern = (FillPattern)fillPattern;
                if (styleProperties2Value.TryGetValue(StyleProperty.Indention, out object indention))
                    style.Indention = (short)indention;
                if (styleProperties2Value.TryGetValue(StyleProperty.IsHidden, out object isHidden))
                    style.IsHidden = (bool)isHidden;
                if (styleProperties2Value.TryGetValue(StyleProperty.IsLocked, out object isLocked))
                    style.IsLocked = (bool)isLocked;
                if (styleProperties2Value.TryGetValue(StyleProperty.Rotation, out object rotation))
                    style.Rotation = (short)rotation;
                if (styleProperties2Value.TryGetValue(StyleProperty.ShrinkToFit, out object shrinkToFit))
                    style.ShrinkToFit = (bool)shrinkToFit;
                if (styleProperties2Value.TryGetValue(StyleProperty.VerticalAlignment, out object verticalAlignment))
                    style.VerticalAlignment = (VerticalAlignment)verticalAlignment;
                if (styleProperties2Value.TryGetValue(StyleProperty.WrapText, out object wrapText))
                    style.WrapText = (bool)wrapText;
                if (styleProperties2Value.TryGetValue(StyleProperty.Font, out object font))
                    style.SetFont((IFont)font);
            }

            new class Key : StyleCache.IKey
            {
                public void Add(StyleProperty stylePropertieName, object value)
                {
                    //stylePropertieNames2Value[stylePropertieName] = value;

                    subkeys.AddRange(System.Text.Encoding.ASCII.GetBytes(stylePropertieName + value.ToString()));
                }

                protected List<byte> subkeys = new List<byte>();

                public int Get()
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
            }
        }
    }
}