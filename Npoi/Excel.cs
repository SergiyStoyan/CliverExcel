//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text.RegularExpressions;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.SS.Formula.PTG;
using NPOI.SS.Formula;

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
                        //FormulaEvaluator = new XSSFFormulaEvaluator(Workbook);
                    }
                    catch (ICSharpCode.SharpZipLib.Zip.ZipException)
                    {
                        fs.Position = 0;//!!!prevents error: EOF in header
                        Workbook = new HSSFWorkbook(fs);//old Excel 97-2003
                        //FormulaEvaluator = new HSSFFormulaEvaluator(Workbook);
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
                    Workbook = null;
                }
            }
        }

        public bool Disposed { get { return Workbook == null; } }

        /// <summary>
        /// Set the active sheet. If no sheet with such name exists, a new sheet is created.
        /// </summary>
        /// <param name="name">(!)name can be auto-corrected</param>
        public void OpenSheet(string name)
        {
            Sheet = Workbook.GetSheet(name);
            if (Sheet == null)
                Sheet = Workbook.CreateSheet(GetSafeSheetName(name));
        }

        /// <summary>
        /// Set the active sheet.
        /// </summary>
        /// <param name="index">1-based</param>
        /// <returns>true if the index exists, otherwise false</returns>
        public bool OpenSheet(int index)
        {
            if (Workbook.NumberOfSheets > 0 && Workbook.NumberOfSheets >= index)
            {
                Sheet = Workbook.GetSheetAt(index - 1);
                return true;
            }
            return false;
        }

        public ISheet Sheet { get; private set; }

        /// <summary>
        /// Get name/rename the active sheet.
        /// (!)When setting, name can be auto-corrected.
        /// </summary>
        public string SheetName
        {
            get
            {
                return Sheet?.SheetName;
            }
            set
            {
                if (Sheet != null)
                    Workbook.SetSheetName(Workbook.GetSheetIndex(Sheet), GetSafeSheetName(value));
            }
        }

        public void Save(string file = null)
        {
            if (file != null)
                File = file;
            using (var fileData = new FileStream(File, FileMode.Create))
            {
                Workbook.Write(fileData, true);
            }
        }

        public string HyperlinkBase
        {
            get
            {
                XSSFWorkbook xSSFWorkbook = Workbook as XSSFWorkbook;
                if (xSSFWorkbook == null)
                    throw new Exception("TBD");
                NPOI.OpenXmlFormats.CT_Property p = xSSFWorkbook.GetProperties().CustomProperties.GetProperty("HyperlinkBase");//so is in Epplus
                return p?.Item?.ToString();
            }
            set
            {
                XSSFWorkbook xSSFWorkbook = Workbook as XSSFWorkbook;
                if (xSSFWorkbook == null)
                    throw new Exception("TBD");
                List<NPOI.OpenXmlFormats.CT_Property> ps = xSSFWorkbook.GetProperties().CustomProperties.GetUnderlyingProperties().property;
                NPOI.OpenXmlFormats.CT_Property p = ps.Find(a => a.name == "HyperlinkBase");//so is in Epplus
                if (value == null)
                {
                    if (p != null)
                        ps.Remove(p);
                    return;
                }
                if (p == null)
                    xSSFWorkbook.GetProperties().CustomProperties.AddProperty("HyperlinkBase", value);
                else
                    p.Item = value;
            }
        }

        public void SetLink(int y, int x, Uri uri)
        {
            GetCell(y, x, true).SetLink(uri);
        }

        public Uri GetLink(int y, int x)
        {
            return GetCell(y, x, false)?.GetLink();
        }

        public string GetValueAsString(int y, int x, bool allowNull = false)
        {
            ICell c = GetCell(y, x, false);
            return c?.GetValueAsString(allowNull);
        }

        public object GetValue(int y, int x)
        {
            ICell c = GetCell(y, x, false);
            return c?.GetValue();
        }

        public void SetValue(int y, int x, object value)
        {
            ICell c = GetCell(y, x, true);
            c.SetValue(value);
        }

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
                return GetValueAsString(y, x, false);
            }
            set
            {
                ICell c = GetCell(y, x, true);
                c.SetCellValue(value);
            }
        }

        public void CreateDropdown(int y, int x, IEnumerable<object> values, object value, bool allowBlank = true)
        {
            GetCell(y, x, true).CreateDropdown(values, value, allowBlank);
        }

        /// <summary>
        /// !!!BUGGY!!!
        /// </summary>
        /// <param name="y"></param>
        /// <param name="x"></param>
        /// <param name="name"></param>
        /// <param name="pngImage"></param>
        /// <exception cref="Exception"></exception>
        public void AddImage(int y, int x, string name, byte[] pngImage)
        {
            int i = Workbook.AddPicture(pngImage, PictureType.PNG);
            ICreationHelper h = Workbook.GetCreationHelper();
            IClientAnchor a = h.CreateClientAnchor();
            a.AnchorType = AnchorType.MoveDontResize;
            a.Col1 = x - 1;//0 index based column
            a.Row1 = y - 1;//0 index based row
            if (Workbook is XSSFWorkbook xSSFWorkbook)
            {
                XSSFDrawing d = (XSSFDrawing)Sheet.CreateDrawingPatriarch();
                XSSFPicture p = (XSSFPicture)d.CreatePicture(a, i);
                p.IsNoFill = true;
                p.Resize();
            }
            else if (Workbook is HSSFWorkbook hSSFWorkbook)
            {
            }
            else
                throw new Exception("Unsupported workbook type: " + Workbook.GetType().FullName);
        }

        //public static byte[] ImageToPngByteArray(Image img)
        //{
        //    using (var stream = new MemoryStream())
        //    {
        //        img.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
        //        return stream.ToArray();
        //    }
        //}

        /// <summary>
        /// !!!BUGGY!!!
        /// </summary>
        /// <param name="y"></param>
        /// <param name="x"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public byte[] GetImage(int y, int x, out string name)
        {
            name = null;
            return null;

            XSSFDrawing d = Sheet.CreateDrawingPatriarch() as XSSFDrawing;
            foreach (XSSFShape s in d.GetShapes())
            {
                XSSFPicture p = s as XSSFPicture;
                if (p == null)
                    continue;
                IClientAnchor a = p.GetPreferredSize();
                if (y - 1 >= a.Row1 && y - 1 <= a.Row2 && x - 1 >= a.Col1 && x - 1 <= a.Col2)
                {
                    XSSFPictureData pd = p.PictureData as XSSFPictureData;
                    return pd.Data;
                }
            }
            //name = null;
            return null;

            //var lst = Workbook.GetAllPictures();
            //for (int i = 0; i < lst.Count; i++)
            //{
            //    var pd = (XSSFPictureData)lst[i];
            //    pd.RelationParts.Add[]
            //    using (Stream s = new MemoryStream(pd.Data))
            //        return new Bitmap(s);
            //}

            //foreach (NPOI.POIXMLDocumentPart dp in Workbook.GetRelations())
            //{
            //    if (dp is XSSFDrawing)
            //    {
            //        NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_Drawing d = ((XSSFDrawing)dp).GetCTDrawing();
            //        foreach (NPOI.OpenXmlFormats.Dml.Spreadsheet.IEG_Anchor a in d.CellAnchors)
            //        {
            //            NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_TwoCellAnchor aa = a as NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_TwoCellAnchor;
            //            if (aa == null)
            //                continue;
            //            NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_Marker m = aa.from;
            //            if (m.row == y && m.col == x)
            //                //using (Stream s = new MemoryStream(((XSSFPicture)aa.picture).PictureData))
            //                    return new Bitmap(0,0);
            //            //CTMarker to = anchor.getTo();
            //            //int row2 = to.GetRow();
            //            //int col2 = to.getCol();

            //            // do something here
            //        }
            //    }
            //}



            //foreach (XSSFPictureData pd in Workbook.GetAllPictures())
            //{
            //    NPOI.OpenXml4Net.OPC.PackagePart pp = pd.GetPackagePart();
            //    pp.GetInputStream
            //  }
            //foreach (NPOI.POIXMLDocumentPart dp in Workbook.GetRelations())
            //{
            //    NPOI.OpenXml4Net.OPC.PackagePart pp = dp.GetPackagePart();
            //    pp.GetInputStream
            //  }
        }

        public void SetStyle(ICellStyle style, bool createCells)
        {
            SetStyleInRawRange(style, createCells);
        }

        public void ReplaceStyle(ICellStyle style1, ICellStyle style2)
        {
            ReplaceStyleInRawRange(style1, style2);
        }
    }
}