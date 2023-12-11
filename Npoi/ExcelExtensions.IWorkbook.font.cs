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
    static public partial class ExcelExtensions
    {
        //    static public void _SetCommentDefaultFont(this IWorkbook workbook, IFont font)
        //    {
        //        workbooks2CommentDefaultFont[workbook] = new CommentFonts { Workbook = workbook, Font = workbook._GetRegisteredFont(font) };
        //    }
        //    public class CommentFonts
        //    {
        //        public IFont Font { get; internal set; }
        //        public IFont BoldFont
        //        {
        //            get
        //            {
        //                if (boldFont == null)
        //                {
        //                    boldFont = Workbook._CloneUnregisteredFont(Font);
        //                    boldFont.IsBold = true;
        //                    boldFont = Workbook._GetRegisteredFont(boldFont);
        //                }
        //                return boldFont;
        //            }
        //        }
        //        IFont boldFont = null;
        //        public IWorkbook Workbook { get; internal set; }
        //    }
        //    public class CommentFonts1
        //    {
        //                    public IFont DefaultFont { get; internal set; }

        //        Dictionary<IFont, IFont> fonts2BoldFont = new Dictionary<IFont, IFont>();
        //        public IFont GetBoldFont(IFont font)
        //        {
        //            if(!fonts2BoldFont.TryGetValue(font, out IFont boldFont))
        //            {
        //                boldFont = Workbook._CloneUnregisteredFont(font);
        //                boldFont.IsBold = true;
        //                boldFont = Workbook._GetRegisteredFont(boldFont);
        //                fonts2BoldFont[font] = boldFont;
        //            }
        //            return boldFont;
        //        }
        //        public IWorkbook Workbook { get; internal set; }
        //    }
        //    static Dictionary<IWorkbook, CommentFonts> workbooks2CommentDefaultFont = new Dictionary<IWorkbook, CommentFonts>();

        //    static public CommentFonts _GetCommentDefaultFonts(this IWorkbook workbook)//getting the default Font for comments
        //    {
        //        if (!workbooks2CommentDefaultFont.TryGetValue(workbook, out CommentFonts cfs))
        //        {
        //            IFont f = workbook._CreateUnregisteredFont();
        //            f.FontName = Excel.CommentDefaultFontName;
        //            f.FontHeight = Excel.CommentDefaultFontSize * 20;
        //            cfs = new CommentFonts { Workbook = workbook, Font = workbook._GetRegisteredFont(f) };
        //            workbooks2CommentDefaultFont[workbook] = cfs;
        //        }
        //        return cfs;
        //    }

        static public void _SetCommentDefaultFont(this IWorkbook workbook, IFont font)
        {
            workbooks2CommentDefaultFont[workbook] = workbook._GetRegisteredFont(font);
        }
        static Dictionary<IWorkbook, IFont> workbooks2CommentDefaultFont = new Dictionary<IWorkbook, IFont>();

        static public IFont _GetCommentDefaultFont(this IWorkbook workbook)//getting the default Font for comments
        {
            if (!workbooks2CommentDefaultFont.TryGetValue(workbook, out IFont f))
            {
                f = workbook._CreateUnregisteredFont();
                f.FontName = Excel.CommentDefaultFontName;
                f.FontHeight = Excel.CommentDefaultFontSize * 20;
                f = workbook._GetRegisteredFont(f);
                workbooks2CommentDefaultFont[workbook] = f;
            }
            return f;
        }

        /// <summary>
        /// Creates an unregistered copy of a font.
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="font"></param>
        /// <returns></returns>
        static public IFont _CloneUnregisteredFont(this IWorkbook workbook, IFont font)
        {
            IFont f = workbook._CreateUnregisteredFont();
            f.IsBold = font.IsBold;
            f.Color = font.Color;
            f.FontHeight = font.FontHeight;
            f.FontName = font.FontName;
            f.IsItalic = font.IsItalic;
            f.IsStrikeout = font.IsStrikeout;
            f.TypeOffset = font.TypeOffset;
            f.Underline = font.Underline;
            return f;
        }

        /// <summary>
        /// If the font does not exists, it is created.
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="font"></param>
        /// <returns></returns>
        static public IFont _GetRegisteredFont(this IWorkbook workbook, IFont font, bool reuseUnusedFont = false)
        {
            IFont f = workbook.FindFont(font.IsBold, font.Color, (short)font.FontHeight, font.FontName, font.IsItalic, font.IsStrikeout, font.TypeOffset, font.Underline);
            if (f == null)
            {
                if (reuseUnusedFont)
                {
                    f = workbook._GetUnusedFonts().FirstOrDefault();
                    if (f == null)
                        f = workbook.CreateFont();
                }
                else
                    f = workbook.CreateFont();
                f.IsBold = font.IsBold;
                f.Color = font.Color;
                f.FontHeight = font.FontHeight;
                f.FontName = font.FontName;
                f.IsItalic = font.IsItalic;
                f.IsStrikeout = font.IsStrikeout;
                f.TypeOffset = font.TypeOffset;
                f.Underline = font.Underline;
            }
            return f;
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
        static public IFont _GetRegisteredFont(this IWorkbook workbook, bool bold, short color, short fontHeightInPoints, string name, bool italic = false, bool strikeout = false, FontSuperScript typeOffset = FontSuperScript.None, FontUnderlineType underline = FontUnderlineType.None)
        {
            short fontHeight = (short)(20 * fontHeightInPoints);
            IFont f = workbook.FindFont(bold, color, fontHeight, name, italic, strikeout, typeOffset, underline);
            if (f == null)
            {
                f = workbook.CreateFont();
                f.IsBold = bold;
                f.Color = color;
                f.FontHeight = fontHeight;
                f.FontName = name;
                f.IsItalic = italic;
                f.IsStrikeout = strikeout;
                f.TypeOffset = typeOffset;
                f.Underline = underline;
            }
            return f;
        }

        static public IEnumerable<IFont> _GetFonts(this IWorkbook workbook)
        {
            for (short i = (short)(workbook.NumberOfFonts - 1); i >= 0; i--)
                yield return workbook.GetFontAt(i);
        }

        /// <summary>
        /// Works for unregistered styles too.
        /// (!)Font must be registered in the workbook though.
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="style"></param>
        /// <exception cref="Exception"></exception>
        public static IFont _GetFont(this IWorkbook workbook, ICellStyle style)
        {
            try
            {
                return workbook.GetFontAt(style.FontIndex);
            }
            catch (Exception e)
            {
                throw new Exception("Could not get font[ID=" + style.FontIndex + " for style[ID=" + style.Index + "]. The font is not registered in the workbook.", e);
            }
        }

        static public IEnumerable<IFont> _FindEqualFonts(this IWorkbook workbook, IFont font, IWorkbook searchWorkbook = null)
        {
            if (searchWorkbook == null)
                searchWorkbook = workbook;
            foreach (IFont font2 in searchWorkbook._GetFonts())
                if (Excel.AreFontsEqual(font, font2))
                    yield return font2;
        }

        /// <summary>
        /// Finds fonts in the workbook that are not used and hence can be used as new.
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="ignoredFontIds"></param>
        /// <returns></returns>
        static public IEnumerable<IFont> _GetUnusedFonts(this IWorkbook workbook, params short[] ignoredFontIds)
        {
            var usedFontIds = workbook._GetStyles().Select(a => a.FontIndex).ToList();
            foreach (var fi in workbook._GetFonts().Where(a => !ignoredFontIds.Contains(a.Index) && !usedFontIds.Contains(a.Index)))
                yield return fi;
        }

        /// <summary>
        /// Makes all the duplicated fonts unused so they can be used as new.
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="unusedFonts"></param>
        static public void _OptimizeFonts(this IWorkbook workbook, out List<IFont> unusedFonts)
        {
            unusedFonts = new List<IFont>();
            var fonts = workbook._GetFonts().ToList();
            var styles = workbook._GetStyles().ToList();
            while (fonts.Count > 0)
            {
                var font = fonts[0];
                fonts.RemoveAt(0);
                List<IFont> font2s = Excel.FindEqualFonts(font, fonts).Where(a => a.Index != font.Index).ToList();
                styles.Where(a => font2s.Find(b => b.Index == a.FontIndex) != null).ForEach(a => a.SetFont(font));
                fonts = fonts.Except(font2s).ToList();
                unusedFonts.AddRange(font2s);
            }
        }
    }
}
