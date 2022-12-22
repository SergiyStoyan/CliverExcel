﻿//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text.RegularExpressions;
using System.Drawing;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.SS.Formula.PTG;
using NPOI.SS.Formula;
using Newtonsoft.Json.Serialization;
using System.Reflection;
using Newtonsoft.Json;

//works  
namespace Cliver
{
    public partial class Excel : IDisposable
    {
        static Excel()
        {
        }

        public Excel(string file, int worksheetId = 1)
        {
            File = file;
            init();
            OpenWorksheet(worksheetId);
        }

        public Excel(string file, string worksheetName)
        {
            File = file;
            init();
            OpenWorksheet(worksheetName);
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
                Workbook = new XSSFWorkbook();
        }

        public IWorkbook Workbook;
        //public IFormulaEvaluator FormulaEvaluator = null;

        public readonly string File;

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

        public void OpenWorksheet(string name)
        {
            Sheet = Workbook.GetSheet(name);
            if (Sheet == null)
            {
                name = Regex.Replace(name, @"\:", "-");//npoi does not accept :
                Sheet = Workbook.CreateSheet(WorkbookUtil.CreateSafeSheetName(name));

                //!!!All Sheet must be formatted as text! Otherwise string dates are converted into numbers.
                //!!!No way found to set default style for a whole sheet. However, NPOI presets ' before numeric values to keep them as strings.
                //ICellStyle defaultStyle = (XSSFCellStyle)Workbook.CreateCellStyle();
                //defaultStyle.DataFormat = Workbook.CreateDataFormat().GetFormat("text");
                //ICell c = GetCell(0, 0, true);
                //Sheet.SetDefaultColumnStyle(0, defaultStyle);
            }
        }
        //ICellStyle defaultStyle;

        public bool OpenWorksheet(int index)
        {
            index--;
            if (Workbook.NumberOfSheets > 0 && Workbook.NumberOfSheets > index)
            {
                Sheet = Workbook.GetSheetAt(index);
                return true;
            }
            return false;
        }
        public ISheet Sheet;

        public string SheetName
        {
            get
            {
                return Sheet.SheetName;
            }
            set
            {
                if (Sheet != null)
                    Workbook.SetSheetName(Workbook.GetSheetIndex(Sheet), value);
            }
        }

        public void Save()
        {
            using (var fileData = new FileStream(File, FileMode.Create))
            {
                Workbook.Write(fileData, true);
            }
        }

        public int AppendLine(IEnumerable<object> values)
        {
            int y = GetLastUsedRow(true) + 1;
            int i = 1;
            foreach (object v in values)
            {
                string s;
                if (v is string)
                    s = (string)v;
                else if (v != null)
                    s = v.ToString();
                else
                    s = null;

                this[y, i++] = s;
            }
            return y;
        }

        public void SetLink(int y, int x, Uri uri)
        {
            ICell c = GetCell(y, x, true);
            if (string.IsNullOrEmpty(this[y, x]))
                c.SetCellValue(LinkEmptyValueFiller);
            if (Workbook is XSSFWorkbook)
                c.Hyperlink = new XSSFHyperlink(HyperlinkType.Url) { Address = uri.ToString() };
            else if (Workbook is HSSFWorkbook)
                c.Hyperlink = new HSSFHyperlink(HyperlinkType.Url) { Address = uri.ToString() };
        }
        public string LinkEmptyValueFiller = "           ";

        public Uri GetLink(int y, int x)
        {
            ICell c = GetCell(y, x, false);
            if (c == null)
                return null;
            if (c.Hyperlink == null)
                return null;
            return new Uri(c.Hyperlink.Address, UriKind.RelativeOrAbsolute);
        }

        public string this[int y, int x]
        {
            get
            {
                ICell c = GetCell(y, x, false);
                return ExcelExtensions.GetValueAsString(c/*, FormulaEvaluator*/);
            }
            set
            {
                ICell c = GetCell(y, x, true);
                //c.SetBlank();
                //c.SetCellType(CellType.String);
                c.SetCellValue(value);
            }
        }

        public void InsertLine(int y, IEnumerable<object> values = null)
        {
            if (y <= Sheet.LastRowNum)
                Sheet.ShiftRows(y - 1, Sheet.LastRowNum, 1);
            GetRow(y, true);
            if (values != null)
                WriteLine(y, values);
        }

        public void WriteLine(int y, IEnumerable<object> values)
        {
            IRow r = GetRow(y, true);

            int x = 1;
            foreach (object v in values)
            {
                string s;
                if (v is string)
                    s = (string)v;
                else if (v != null)
                    s = v.ToString();
                else
                    s = null;

                r.GetCell(x++, true).SetCellValue(s);
            }
        }

        public void CreateDropdown(int y, int x, IEnumerable<object> values, object value, bool allowBlank = true)
        {
            List<string> vs = new List<string>();
            foreach (object v in values)
            {
                string s;
                if (v is string)
                    s = (string)v;
                else if (v != null)
                    s = v.ToString();
                else
                    s = null;
                vs.Add(s);
            }
            IDataValidationHelper dvh = new XSSFDataValidationHelper((XSSFSheet)Sheet);
            //string dvs = string.Join(",", vs);
            //IDataValidationConstraint dvc = Sheet.GetDataValidations().Find(a => string.Join(",", a.ValidationConstraint.ExplicitListValues) == dvs)?.ValidationConstraint;
            //if (dvc == null)
            //dvc = dvh.CreateCustomConstraint(dvs);
            IDataValidationConstraint dvc = dvh.CreateExplicitListConstraint(vs.ToArray());
            CellRangeAddressList cral = new CellRangeAddressList(y - 1, y - 1, x - 1, x - 1);
            IDataValidation dv = dvh.CreateValidation(dvc, cral);
            dv.SuppressDropDownArrow = true;
            dv.EmptyCellAllowed = allowBlank;
            ((XSSFSheet)Sheet).AddValidationData(dv);

            {
                string s;
                if (value is string)
                    s = (string)value;
                else if (value != null)
                    s = value.ToString();
                else
                    s = null;
                GetCell(y, x, true).SetCellValue(s);
            }
        }

        public void AddImage(int y, int x, /*string name,*/ byte[] pngImage)//!!!!buggy
        {
            throw new Exception("TBD");
            int i = Workbook.AddPicture(pngImage, PictureType.PNG);
            ICreationHelper h = Workbook.GetCreationHelper();
            IClientAnchor a = h.CreateClientAnchor();
            a.AnchorType = AnchorType.MoveDontResize;
            a.Col1 = x - 1;//0 index based column
            a.Row1 = y - 1;//0 index based row
            XSSFDrawing d = (XSSFDrawing)Sheet.CreateDrawingPatriarch();
            XSSFPicture p = (XSSFPicture)d.CreatePicture(a, i);
            p.IsNoFill = true;
            p.Resize();
        }

        //public static byte[] ImageToPngByteArray(Image img)
        //{
        //    using (var stream = new MemoryStream())
        //    {
        //        img.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
        //        return stream.ToArray();
        //    }
        //}

        public byte[] GetImage(int y, int x/*, out string name*/)//!!!!buggy
        {
            throw new Exception("TBD");
            //name = null;

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
    }
}