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
    /// It can have only one sheet active at time. Changing the active sheet is done by OpenSheet().
    /// (!)Row and column numbers, indexes of objects like sheets, are 1-based, when native NPOI objects tend to use 0-based indexes.
    /// </summary>
    public partial class Excel : IDisposable
    {
        static Excel()
        {
        }

        /// <summary>
        /// (!)No sheet auto-created. If the given sheetIndex does not exist, Sheet will be NULL.
        /// </summary>
        /// <param name="file"></param>
        /// <param name="sheetIndex"></param>
        public Excel(string file, int sheetIndex = 1)
        {
            init(file);
            OpenSheet(sheetIndex);
        }

        /// <summary>
        /// Use it when you need the sheet auto-created. If the given sheetName does not exist and createSheet=TRUE, the sheet will be created.
        /// </summary>
        /// <param name="file"></param>
        /// <param name="sheetName"></param>
        /// <param name="createSheet"></param>
        public Excel(string file, string sheetName, bool createSheet = true)
        {
            init(file);
            OpenSheet(sheetName, createSheet);
        }

        void init(string file)
        {
            File = file;
            styleCache = new StyleCache(this);

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
        }

        public IWorkbook Workbook { get; private set; }

        public string File { get; private set; }

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
                    Workbook = null;
                }
            }
        }

        public bool Disposed { get { return Workbook == null; } }

        /// <summary>         
        /// NULL- and type-safe.
        /// (!)Never returns NULL.
        /// </summary>
        /// <param name="y"></param>
        /// <param name="x"></param>
        /// <returns></returns>
        public string this[int y, int x]
        {
            get
            {
                return GetValueAsString(y, x, false);
            }
            set
            {
                GetCell(y, x, true).SetCellValue(value);
            }
        }

        /// <summary>
        /// NULL- and type-safe.
        /// (!)Never returns NULL.
        /// </summary>
        public string this[string cellAddress]
        {
            get
            {
                return Sheet._GetValueAsString(cellAddress, false);
            }
            set
            {
                Sheet._SetValue(cellAddress, value);
            }
        }

        ///// <summary> 
        ///// (!)Never returns NULL. Row is created if does not exist.
        ///// </summary>
        ///// <param name="y">1-based</param>
        ///// <returns></returns>
        //public IRow this[int y]!!!do not do it: it is ambiguous: row or column?
        //{
        //    get
        //    {
        //        return Sheet._GetRow(y, true);
        //    }
        //}
        //public Column this[string columnName]//!!!do not do it: it is used for Sheet
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
    }
}