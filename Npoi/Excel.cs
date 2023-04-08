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
                if (Workbook is XSSFWorkbook xSSFWorkbook)
                {
                    NPOI.OpenXmlFormats.CT_Property p = xSSFWorkbook.GetProperties().CustomProperties.GetProperty("HyperlinkBase");//so is in Epplus
                    return p?.Item?.ToString();
                }
                else if (Workbook is HSSFWorkbook hSSFWorkbook)
                {
                    hSSFWorkbook.CreateInformationProperties();
                    return hSSFWorkbook.DocumentSummaryInformation.CustomProperties["HyperlinkBase"]?.ToString();
                }
                else
                    throw new Exception("Unsupported workbook type: " + Workbook.GetType().FullName);
            }
            set
            {
                if (Workbook is XSSFWorkbook xSSFWorkbook)
                {
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
                else if (Workbook is HSSFWorkbook hSSFWorkbook)
                {
                    hSSFWorkbook.CreateInformationProperties();
                    if (value == null)
                    {
                        hSSFWorkbook.DocumentSummaryInformation.CustomProperties.Remove("HyperlinkBase");
                        return;
                    }
                    hSSFWorkbook.DocumentSummaryInformation.CustomProperties.Put("HyperlinkBase", value);//so is in Epplus
                }
                else
                    throw new Exception("Unsupported workbook type: " + Workbook.GetType().FullName);
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

        public void CreateDropdown<T>(int y, int x, IEnumerable<T> values, T value, bool allowBlank = true)
        {
            GetCell(y, x, true).CreateDropdown(values, value, allowBlank);
        }

        /// <summary>
        /// !!!BUGGY!!!
        /// </summary>
        /// <exception cref="Exception"></exception>
        public void AddImage(Image image)
        {
            int i = Workbook.AddPicture(image.Data, image.Type);
            ICreationHelper h = Workbook.GetCreationHelper();
            IClientAnchor a = h.CreateClientAnchor();
            a.AnchorType = AnchorType.MoveDontResize;
            a.Col1 = image.X - 1;//0-based column index
            a.Row1 = image.Y - 1;//0-based row index
            IDrawing d = Sheet.CreateDrawingPatriarch();
            IPicture p = d.CreatePicture(a, i);
            if (p is XSSFPicture xSSFPicture)
                xSSFPicture.IsNoFill = true;
            //p.Resize();
        }

        public class Image
        {
            //public IClientAnchor Anchor;
            public int Y;
            public int X;
            public string Name;
            public PictureType Type;
            public byte[] Data;
        }

        /// <summary>
        /// !!!BUGGY!!!
        /// </summary>
        /// <param name="y"></param>
        /// <param name="x"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public IEnumerable<Image> GetImages(int y, int x)
        {
            foreach (IPicture p in Workbook.GetAllPictures())
            {
                if (p == null)
                    continue;
                IClientAnchor a = p.GetPreferredSize();
                if (y - 1 >= a.Row1 && y - 1 <= a.Row2 && x - 1 >= a.Col1 && x - 1 <= a.Col2)
                {
                    IPictureData pictureData = p.PictureData;
                    yield return new Image { Data = pictureData.Data, Name = null, Type = pictureData.PictureType, X = a.Col1, Y = a.Row1/*, Anchor = a*/ };
                }
            }

            //XSSFDrawing d = Sheet.CreateDrawingPatriarch() as XSSFDrawing;
            //foreach (XSSFShape s in d.GetShapes())
            //{
            //    XSSFPicture p = s as XSSFPicture;
            //    if (p == null)
            //        continue;
            //    IClientAnchor a = p.GetPreferredSize();
            //    if (y - 1 >= a.Row1 && y - 1 <= a.Row2 && x - 1 >= a.Col1 && x - 1 <= a.Col2)
            //    {
            //        XSSFPictureData pd = p.PictureData as XSSFPictureData;
            //        pictureType = pd.PictureType;
            //        return pd.Data;
            //    }
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
        }
    }
}