////********************************************************************************************
////Author: Sergiy Stoyan
////        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
////        http://www.cliversoft.com
////********************************************************************************************
//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.IO;
//using System.Text.RegularExpressions;
//using System.Drawing;
//using ClosedXML.Excel;
//using ClosedXML.Excel.Drawings;

//namespace Cliver.HawkeyeInvoiceParser
//{
//    public class Excel : IDisposable
//    {
//        //static public void TestBug(string file = @"c:\test\test.xlsx")
//        //{
//        //    using (XLWorkbook workbook = new XLWorkbook())
//        //    {
//        //        IXLWorksheet worksheet = workbook.Worksheets.Add("test");
//        //        worksheet.Cell(1, 1).Value = "test";
//        //        worksheet.Cell(1, 1).Hyperlink = new XLHyperlink("http://test");
//        //        workbook.SaveAs(file);
//        //    }
//        //    using (XLWorkbook workbook = new XLWorkbook(file))
//        //    {
//        //        IXLWorksheet worksheet = workbook.Worksheet("test");
//        //        worksheet.Row(1).InsertRowsAbove(1);
//        //        worksheet.Cell(1, 1).Value = "test2";
//        //        worksheet.Cell(1, 1).Hyperlink = new XLHyperlink("http://test");
//        //        //workbook.Save();
//        //    }
//        //}

//        static Excel()
//        {
//        }

//        public Excel(string file, int worksheetId = 1)
//        {
//            File = file;
//            init();
//            OpenWorksheet(worksheetId);
//        }

//        public Excel(string file, string worksheetName)
//        {
//            File = file;
//            init();
//            OpenWorksheet(worksheetName);
//        }

//        void init()
//        {
//            if (System.IO.File.Exists(File))
//                workbook = new XLWorkbook(File);
//            else
//                workbook = new XLWorkbook();
//        }

//        XLWorkbook workbook;

//        public readonly string File;

//        ~Excel()
//        {
//            Dispose();
//        }

//        public void Dispose()
//        {
//            lock (this)
//            {
//                if (workbook != null)
//                {
//                    workbook.Dispose();
//                    workbook = null;
//                }
//            }
//        }

//        public string HyperlinkBase
//        {
//            get
//            {
//                IXLCustomProperty p = workbook.CustomProperties.CustomProperty("HyperlinkBase");//so is in Epplus
//                //if (p == null)
//                //    p = workbook.CustomProperties.CustomProperty("Hyperlink Base");
//                if (p == null)
//                    return null;
//                return p.GetValue<string>();
//            }
//            set
//            {
//                if (value == null)
//                    workbook.CustomProperties.Delete("HyperlinkBase");
//                else
//                    workbook.CustomProperties.Add("HyperlinkBase", value);//so is in Epplus
//                //workbook.CustomProperties.Add("Hyperlink Base", value?.ToString());
//            }
//        }

//        public void OpenWorksheet(string name)
//        {
//            worksheet = workbook.Worksheets.Where(a => a.Name == name).FirstOrDefault();
//            if (worksheet == null)
//                worksheet = workbook.Worksheets.Add(name);
//        }

//        public bool OpenWorksheet(int index)
//        {
//            if (workbook.Worksheets.Count > 0)
//            {
//                worksheet = workbook.Worksheet(index);
//                return true;
//            }
//            return false;
//        }
//        IXLWorksheet worksheet;

//        public string WorksheetName
//        {
//            get
//            {
//                return worksheet.Name;
//            }
//            set
//            {
//                if (worksheet != null)
//                    worksheet.Name = value;
//            }
//        }

//        public void Save()
//        {
//            //if (System.IO.File.Exists(File))
//            //    workbook.Save();
//            //else
//            workbook.SaveAs(File); //, new SaveOptions { ValidatePackage = false, GenerateCalculationChain = false, ConsolidateConditionalFormatRanges = false, ConsolidateDataValidationRanges = false, EvaluateFormulasBeforeSaving = false });
//        }

//        public int GetLastUsedRow()
//        {
//            if (worksheet == null)
//                throw new Exception("No active sheet.");
//            IXLRow r = worksheet.LastRowUsed();
//            return r == null ? 0 : r.RowNumber();
//        }

//        public int AppendLine(IEnumerable<object> values)
//        {
//            int y = GetLastUsedRow() + 1;
//            int i = 1;
//            foreach (object v in values)
//            {
//                string s;
//                if (v is string)
//                    s = (string)v;
//                else if (v != null)
//                    s = v.ToString();
//                else
//                    s = null;

//                this[y, i++] = s;
//                //sheet.Cells[currentRow, i++].Style.Numberformat.Format = "@";
//            }
//            return y;
//        }

//        public void SetLink(int y, int x, Uri uri)
//        {
//            IXLCell c = worksheet.Cell(y, x);
//            string v = c.GetValue<string>();
//            if (string.IsNullOrEmpty(v))
//                c.SetValue(LinkEmptyValueFiller);

//            XLHyperlink h;
//            try
//            {
//                h = c.Hyperlink;//!!!workaround for the bug in ClosedXML
//            }
//            catch (Exception e)
//            {
//                h = c.Hyperlink;
//            }
//            h?.Delete();
//            c.Hyperlink = new XLHyperlink(uri);
//        }
//        public static string LinkEmptyValueFiller = "           ";

//        public Uri GetLink(int y, int x)
//        {
//            return worksheet.Cell(y, x).GetHyperlink()?.ExternalAddress;
//        }

//        public string this[int y, int x]
//        {
//            get
//            {
//                return worksheet.Cell(y, x).GetValue<string>();
//            }
//            set
//            {
//                worksheet.Cell(y, x).SetValue(value);
//            }
//        }

//        public void InsertLine(int y, IEnumerable<object> values = null)
//        {
//            worksheet.Row(y).InsertRowsAbove(1);
//            if (values != null)
//                WriteLine(y, values);
//        }

//        public void WriteLine(int y, IEnumerable<object> values)
//        {
//            int i = 1;
//            foreach (object v in values)
//            {
//                string s;
//                if (v is string)
//                    s = (string)v;
//                else if (v != null)
//                    s = v.ToString();
//                else
//                    s = null;

//                this[y, i++] = s;
//                //sheet.Cells[y, i++].Style.Numberformat.Format = "@";
//            }
//        }

//        public void CreateDropdown(int y, int x, IEnumerable<object> values, object value, bool allowBlank = true)
//        {
//            List<string> vs = new List<string>();
//            foreach (object v in values)
//            {
//                string s;
//                if (v is string)
//                    s = (string)v;
//                else if (v != null)
//                    s = v.ToString();
//                else
//                    s = null;
//                vs.Add(s);
//            }
//            IXLDataValidation dv = worksheet.Cell(y, x).SetDataValidation();
//            dv.List("\"" + string.Join(",", vs) + "\"");
//            dv.IgnoreBlanks = allowBlank;
//            dv.InCellDropdown = true;

//            {
//                string s;
//                if (value is string)
//                    s = (string)value;
//                else if (value != null)
//                    s = value.ToString();
//                else
//                    s = null;
//                this[y, x] = s;
//            }
//        }

//        public void AddImage(int y, int x, string name, Bitmap image)
//        {
//            using (MemoryStream ms = new MemoryStream())
//            {
//                image.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
//                IXLPicture p = worksheet.AddPicture(ms, name);
//                p.MoveTo(worksheet.Cell(y, x));//.Scale(0.5); // optional: resize picture
//            }
//        }

//        public Bitmap GetImage(int y, int x)
//        {
//            IXLCell c = worksheet.Cell(y, x);
//            IXLPicture p = worksheet.Pictures.Where(a => a.TopLeftCell == c).FirstOrDefault();
//            if (p == null)
//                return null;
//            return new Bitmap(p.ImageStream);
//        }

//        public void FitColumnsWidth(params int[] columnIs)
//        {
//            foreach (int i in columnIs)
//                worksheet.Column(i).AdjustToContents();
//        }

//        public void FitColumnsWidth(int column1I, int column2I)
//        {
//            worksheet.Columns(column1I.ToString() + ":" + column2I).AdjustToContents();
//        }

//        public void HighlightRow(int y, Color color)
//        {
//            worksheet.Row(y).Style.Fill.BackgroundColor = XLColor.FromColor(color);
//        }

//        public void ClearHighlighting()
//        {
//            worksheet.Style.Fill.BackgroundColor = XLColor.NoColor;
//        }
//    }
//}