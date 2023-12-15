//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;

namespace Cliver
{
    /// <summary>
    /// (!)Row and column numbers, indexes of objects like sheets, are 1-based, when native NPOI objects tend to use 0-based indexes.
    /// </summary>
    public partial class Excel : IDisposable
    {
        public Excel(string file = null)
        {
            File = file;

            if (System.IO.File.Exists(File))
                using (FileStream fs = new FileStream(File, FileMode.Open, FileAccess.Read))
                {
                    try
                    {
                        fs.Position = 0;//!!!prevents occasional error: EOF in header
                        Workbook = new XSSFWorkbook(fs);
                    }
                    catch (ICSharpCode.SharpZipLib.Zip.ZipException)
                    {
                        fs.Position = 0;//!!!prevents error: EOF in header
                        Workbook = new HSSFWorkbook(fs);//old Excel 97-2003
                    }
                }
            else
            {
                if (PathRoutines.GetFileExtension(File).ToLower() != "xls")
                    Workbook = new XSSFWorkbook();
                else
                    Workbook = new HSSFWorkbook();
            }

            OneWorkbookStyleCache = new StyleCache(Workbook);
            DefaultCommentStyle = new CommentStyle(Workbook);

            //workbooks2Excel[Workbook] = new WeakReference<Excel>(this);
            workbooks2Excel.Add(Workbook, this);
        }
        //readonly static Dictionary<IWorkbook, WeakReference<Excel>> workbooks2Excel = new Dictionary<IWorkbook, WeakReference<Excel>>();
        static System.Runtime.CompilerServices.ConditionalWeakTable<IWorkbook, Excel> workbooks2Excel = new System.Runtime.CompilerServices.ConditionalWeakTable<IWorkbook, Excel>();


        public IWorkbook Workbook { get; private set; }

        /// <summary>
        /// Workbook (alias)
        /// </summary>
        public IWorkbook _ { get { return Workbook; } }

        public string File { get; private set; }
        readonly internal StyleCache OneWorkbookStyleCache;
        public string LinkEmptyValueFiller = "           ";
        public CommentStyle DefaultCommentStyle;

        ~Excel()
        {
            Dispose();
        }

        public void Dispose()
        {
            lock (this)
            {
                if (Workbook != null)
                {
                    Workbook.Close();
                    Workbook.Dispose();
                    workbooks2Excel.Remove(Workbook);
                    Workbook = null;
                }
            }
        }

        public bool Disposed { get { return Workbook == null; } }

        public void Save(string file = null)
        {
            if (file != null)
                File = file;
            Workbook._Save(File);
        }

        /// <summary>
        /// Makes sure that the file is not mangled in the case of a error.
        /// </summary>
        /// <param name="file"></param>
        public void SafeSave(string file = null)
        {
            if (file != null)
                File = file;

            string tempFile = PathRoutines.InsertSuffixBeforeFileExtension(File, "_TEMP");
            try
            {
                Workbook._Save(tempFile);
            }
            catch
            {
                if (System.IO.File.Exists(File))
                    System.IO.File.Delete(tempFile);
                throw;
            }
            FileSystemRoutines.MoveFile(tempFile, File, true);
        }

        ///// <summary> 
        ///// (!)Never returns NULL. Row is created if does not exist.
        ///// </summary>
        ///// <param name="y">1-based</param>
        ///// <returns></returns>
        //public IRow this[int y]!!!do not do that: it is ambiguous: row or column?
        //{
        //    get
        //    {
        //        return Sheet._GetRow(y, true);
        //    }
        //}
        //public Column this[string columnName]//!!!do not do that: it is used for Sheet
        //{
        //    get
        //    {
        //        return Sheet._GetColumn(columnName);
        //    }
        //}

        public class Image
        {
            //public IClientAnchor Anchor;
            /// <summary>
            /// 1-based
            /// </summary>
            public int Y;
            /// <summary>
            /// 1-based
            /// </summary>
            public int X;
            public string Name;
            public PictureType Type;
            public byte[] Data;
        }

        public class Color
        {
            public readonly byte R;
            public readonly byte G;
            public readonly byte B;
            public byte[] RGB { get { return new byte[] { R, G, B }; } }

            public Color(byte r, byte g, byte b)
            {
                R = r;
                G = g;
                B = b;
                RGB[0] = R;
                RGB[1] = G;
                RGB[2] = B;
            }

            public Color(int aRGB) : this((byte)((aRGB >> 16) & 0xFF), (byte)((aRGB >> 8) & 0xFF), (byte)(aRGB & 0xFF))
            {
            }

            public Color(byte[] RGB) : this(RGB[0], RGB[1], RGB[2])
            {
            }

            public Color(IColor c) : this(c.RGB[0], c.RGB[1], c.RGB[2])
            {
            }

            public Color(System.Drawing.Color color) : this(color.ToArgb())
            {
            }
        }

        public class RichTextStringFormattingRun
        {
            public int Start;
            public int ExcludedEnd;
            public IFont Font;
            public RichTextStringFormattingRun(int start, int excludedEnd, IFont font)
            {
                Start = start;
                ExcludedEnd = excludedEnd;
                Font = font;
            }
        }

        public delegate void OnFormulaCellMoved(ICell fromCell, ICell toCell);

        //public enum ColumnScope
        //{
        //    /// <summary>
        //    /// Returns only columns with at least one not empty cell.
        //    /// (!)Slow due to checking all the cells' values.
        //    /// </summary>
        //    NotEmpty,
        //    /// <summary>
        //    /// Returns only columns with cells.
        //    /// </summary>
        //    WithCells,
        //    /// <summary>
        //    /// Returns only rows existing as objects.
        //    /// </summary>
        //    NotNull,
        //    /// <summary>
        //    /// Returns all the rows withing the range with non-existing rows represented as NULL. 
        //    /// (!)Might return a huge pile of null and no-cell rows after the last not empty row.  
        //    /// </summary>
        //    IncludeNull,
        //    /// <summary>
        //    /// Returns all the rows withing the range with non-existing rows having been created.
        //    /// </summary>
        //    CreateIfNull
        //}

        public enum RowScope
        {
            /// <summary>
            /// Returns only rows with at least one not empty cell.
            /// (!)Slow due to checking all the cells' values.
            /// </summary>
            NotEmpty,
            /// <summary>
            /// Returns only rows with cells.
            /// </summary>
            WithCells,
            /// <summary>
            /// Returns only rows existing as objects.
            /// </summary>
            NotNull,
            /// <summary>
            /// Returns all the rows withing the range with non-existing rows represented as NULL. 
            /// (!)Might return a huge pile of null and no-cell rows after the last not empty row.  
            /// </summary>
            IncludeNull,
            /// <summary>
            /// Returns all the rows withing the range with non-existing rows having been created.
            /// </summary>
            CreateIfNull
        }

        public enum LastRowCondition
        {
            /// <summary>
            /// (!)Considerably slow due to checking all the cells' values
            /// </summary>
            NotEmpty,
            /// <summary>
            /// Row with cells.
            /// </summary>
            HasCells,
            /// <summary>
            /// Row existing as an object.
            /// </summary>
            NotNull,
        }

        public enum RowStyleMode
        {
            /// <summary>
            /// Set the row default style.
            /// </summary>
            Row = 1,
            /// <summary>
            /// Set style to the existing cells.
            /// </summary>
            ExistingCells = 2,
            /// <summary>
            /// Set style to all the cells with no gaps. When need, blank cells are created.
            /// </summary>
            NoGapCells = 4,
        }

        public class CopyCellMode
        {
            public bool CopyComment { get; set; } = false;
            public bool CopyLink { get; set; } = true;
            public OnFormulaCellMoved OnFormulaCellMoved { get; set; } = null;

            public CopyCellMode Clone()
            {
                CopyCellMode ccm = new CopyCellMode
                {
                    OnFormulaCellMoved = OnFormulaCellMoved,
                    CopyComment = CopyComment,
                    CopyLink = CopyLink
                };
                return ccm;
            }
        }

        public class MoveRegionMode : CopyCellMode
        {
            public bool UpdateMergedRegions { get; set; } = false;
        }

        public class CommentStyle
        {
            public IFont Font
            {
                get
                {
                    if (font == null)
                    {
                        font = Workbook._CreateUnregisteredFont();
                        font.FontName = DefaultFontName;
                        font.FontHeight = DefaultFontSize * 20;
                        font = Workbook._GetRegisteredFont(font);
                    }
                    return font;
                }
                set
                {
                    font = value;
                }
            }
            IFont font = null;

            public IFont AuthorFont
            {
                get
                {
                    if (authorFont == null)
                    {
                        authorFont = Workbook._CloneUnregisteredFont(Font);
                        authorFont.IsBold = !Font.IsBold;
                        authorFont = Workbook._GetRegisteredFont(authorFont);
                    }
                    return authorFont;
                }
                set
                {
                    authorFont = value;
                }
            }
            IFont authorFont = null;

            public string Author;
            public string AuthorDelimiter = "\r\n";

            public Excel.Color Background;

            public int PaddingRows = 2;
            public int AppendPaddingRows = 0;
            public int Columns = 3;
            public string AppendDelimiter = "\r\n";

            public string DefaultFontName = "Tahoma";
            public int DefaultFontSize = 9;
            public Excel.Color DefaultFontColor;
            public Excel.Color AuthorDefaultFontColor;

            public IWorkbook Workbook { get; private set; }

            internal CommentStyle(IWorkbook workbook)
            {
                Workbook = workbook;
            }
        }
    }
}