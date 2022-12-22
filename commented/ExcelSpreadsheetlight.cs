////********************************************************************************************
////Author: Sergiy Stoyan
////        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
////        http://www.cliversoft.com
////********************************************************************************************
//using System;
//using System.Collections.Generic;
//using System.Linq;
////using OfficeOpenXml;
//using System.IO;
//using System.Text.RegularExpressions;
////using OfficeOpenXml.Drawing;
//using System.Drawing;
//using SpreadsheetLight;
//using DocumentFormat.OpenXml;
//using SpreadsheetLight.Drawing;

//!!!old, bad supported
//namespace Cliver.HawkeyeInvoiceParser
//{
//    public class Excel : IDisposable
//    {
//        static Excel()
//        { }

//        public Excel(string file)
//        {
//            File = file;
//            document = new SLDocument(file);
//            OpenWorksheet(0);
//        }

//        public Excel(string file, string worksheetName)
//        {
//            File = file;
//            document = new SLDocument(file);
//            OpenWorksheet(worksheetName);
//        }
//        SLDocument document;

//        public readonly string File;

//        ~Excel()
//        {
//            Dispose();
//        }

//        public void Dispose()
//        {
//            lock (this)
//            {
//                if (document != null)
//                {
//                    try
//                    {
//                        document.Dispose();
//                    }
//                    catch (Exception e)//unclear error here
//                    {
//                    }
//                    document = null;
//                }
//            }
//        }

//        public Uri HyperlinkBase
//        {
//            get
//            {
//                return package.Workbook.Properties.HyperlinkBase;
//            }
//            set
//            {
//                package.Workbook.Properties.HyperlinkBase = value;
//            }
//        }

//        public void OpenWorksheet(string name)
//        {
//            if (!document.SelectWorksheet(name))
//                if (!document.AddWorksheet(name))
//                    throw new Exception("Could not add worksheet '" + name + "'");
//        }

//        public bool OpenWorksheet(int index)
//        {
//            List<string> ws = document.GetWorksheetNames();
//            if (ws.Count <= index)
//                return false;
//            return document.SelectWorksheet(ws[index]);
//        }

//        public void Save()
//        {
//            document.Save();
//        }

//        public void SetAppendLineToBeginning()
//        {
//            currentRow = 1;
//        }
//        public void SetAppendLineToEnd()
//        {
//            currentRow = getLastUsedRow() + 1;
//        }
//        int getLastUsedRow()
//        {
//          return  document.GetWorksheetStatistics().EndRowIndex;
//        }

//        public int AppendLine(IEnumerable<object> values)
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

//                this[currentRow, i++] = s;
//            }
//            return currentRow++;
//        }
//        int currentRow = 1;

//        public void SetLink(int y, int x, Uri uri)
//        {
//            document.InsertHyperlink(y,x, uri.IsFile? SLHyperlinkTypeValues.FilePath: SLHyperlinkTypeValues.Url, uri.ToString(), false);
//            if (string.IsNullOrEmpty(this[y, x]))
//                this[y, x] = LinkEmptyValueFiller;

//            string g= document.GetCellFormula(y, x);
//        }
//        public static string LinkEmptyValueFiller = "           ";

//        public Uri GetLink(int y, int x)
//        {
//            return document.GetCellFormula(y,x);
//        }

//        public string this[int y, int x]
//        {
//            get
//            {
//                return document.GetCellValueAsString(y,x);
//            }
//            set
//            {
//                document.SetCellValue(y, x, value);
//            }
//        }

//        public void InsertLine(int y, IEnumerable<object> values = null)
//        {
//            document.InsertRow(y, 1);
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
//            SLDataValidation dv = document.CreateDataValidation(y,x);
//            dv.AllowList(string.Join("\r\n",vs), allowBlank, true);
//            document.AddDataValidation(dv);
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
//            ImageConverter converter = new ImageConverter();
//            SLPicture p = new SLPicture((byte[])converter.ConvertTo(image, typeof(byte[])), DocumentFormat.OpenXml.Packaging.ImagePartType.Bmp);
//                p.SetPosition()
//            ExcelPicture ep = worksheet.Drawings.AddPicture(name, image);
//            ep.SetPosition(y - 1, 0, x - 1, 0);//it seems to set to the next cell
//        }

//        public Bitmap GetImage(int y, int x)
//        {
//            y--;
//            x--;//it seems to get from the next cell
//            ExcelDrawing ed = worksheet.Drawings.FirstOrDefault(a => a.From.Row == y && a.From.Column == x);
//            if (ed == null)
//                return null;
//            return (Bitmap)((ExcelPicture)ed).Image;
//        }

//        public void FitColumnsWidth(params int[] columnIs)
//        {
//            //worksheet.Cells[[Worksheet.Dimension.Address].AutoFitColumns();
//            foreach (int i in columnIs)
//                worksheet.Column(i).AutoFit();
//        }

//        public void Highlight(int y, Color color)
//        {
//            if (worksheet.Row(y).Style.Fill.PatternType == OfficeOpenXml.Style.ExcelFillStyle.None)
//                worksheet.Row(y).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
//            worksheet.Row(y).Style.Fill.BackgroundColor.SetColor(color);
//        }
//    }
//}