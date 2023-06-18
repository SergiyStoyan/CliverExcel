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

        public Excel(string file, int sheetIndex = 1)
        {
            File = file;
            init();
            OpenSheet(sheetIndex);
        }

        public Excel(string file, string sheetName)
        {
            File = file;
            init();
            OpenSheet(sheetName);
        }

        void init()
        {
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
                //System.IO.File.Create(File).Dispose();
                Workbook = new XSSFWorkbook();
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
        /// (!)Never returns NULL.
        /// </summary>
        /// <param name="y"></param>
        /// <param name="x"></param>
        /// <returns></returns>
        public string this[int y, int x]
        {
            get
            {
                return Sheet._GetValueAsString(y, x, false);
            }
            set
            {
                Sheet._SetValue(y, x, value);
            }
        }

        /// <summary>
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
            readonly public byte[] RGB = new byte[3];

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