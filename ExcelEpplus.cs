////********************************************************************************************
////Author: Sergiy Stoyan
////        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
////        http://www.cliversoft.com
////********************************************************************************************
//using System;
//using System.Collections.Generic;
//using System.Linq;
//using OfficeOpenXml;
//using System.IO;
//using System.Text.RegularExpressions;
//using OfficeOpenXml.Drawing;
//using System.Drawing;

//!!!buggy!!! file size inflating!!!
//namespace Cliver.HawkeyeInvoiceParser
//{
//    public class Excel : IDisposable
//    {
//        static Excel()//!!!5.3.2 version has changed index-base from 1 to 0 (at least for worksheets), against 4.*.*
//        {
//            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
//        }

//        public Excel(string file)
//        {
//            File = file;
//            package = new ExcelPackage(new FileInfo(file));
//            OpenWorksheet(0);
//        }

//        public Excel(string file, string worksheetName)
//        {
//            File = file;
//            package = new ExcelPackage(new FileInfo(file));
//            OpenWorksheet(worksheetName);
//        }
//        ExcelPackage package;

//        public readonly string File;

//        ~Excel()
//        {
//            Dispose();
//        }

//        public void Dispose()
//        {
//            lock (this)
//            {
//                if (package != null)
//                {
//                    try
//                    {
//                        package.Dispose();
//                    }
//                    catch (Exception e)//unclear error here
//                    {
//                    }
//                    package = null;
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
//            worksheet = package.Workbook.Worksheets.Where(x => x.Name == name).FirstOrDefault();
//            if (worksheet == null)
//                worksheet = package.Workbook.Worksheets.Add(name);
//            worksheet.Cells.Style.Numberformat.Format = "@";
//            //package.Workbook..ErrorCheckingOptions.BackgroundChecking = false;
//        }

//        public bool OpenWorksheet(int index)
//        {
//            package.Compatibility.IsWorksheets1Based = false;
//            worksheet = package.Workbook.Worksheets.Where(x => x.Index == index).FirstOrDefault();
//            if (worksheet != null)
//            {
//                worksheet.Cells.Style.Numberformat.Format = "@";
//                //package.Workbook..ErrorCheckingOptions.BackgroundChecking = false;
//            }
//            return worksheet != null;
//        }
//        ExcelWorksheet worksheet;

//        public void Save()
//        {
//            //try
//            //{
//            package.Save();
//            //package.Dispose();//!!!re-open due to epplus bug inflating file
//            //package = new ExcelPackage(new FileInfo(File));
//            //OpenWorksheet(worksheet.Name);
//            //}
//            //catch (System.InvalidOperationException e)
//            //{
//            //if (e.InnerException != null && Regex.IsMatch(e.InnerException.Message, Regex.Escape("Part does not exist."), RegexOptions.IgnoreCase))
//            //    //error when a reading performed
//            //else
//            //    throw e;
//            //}
//        }

//        public int GetLastUsedRow()
//        {
//            if (worksheet == null)
//                throw new Exception("No active worksheet.");
//            for (int row = worksheet.Dimension.End.Row; row > 0; row--)
//                if (worksheet.Cells[row, 1, row, worksheet.Dimension.End.Column].Any(c => !string.IsNullOrEmpty(c.Text)))
//                    return row;
//            return 0;
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

//                worksheet.Cells[y, i++].Value = s;
//                //worksheet.Cells[currentRow, i++].Style.Numberformat.Format = "@";
//            }
//            return y;
//        }

//        public void SetLink(int y, int x, Uri uri)
//        {
//            if (string.IsNullOrEmpty(this[y, x]))
//                this[y, x] = LinkEmptyValueFiller;
//            //string v = this[y, x];
//            worksheet.Cells[y, x].Hyperlink = uri;
//            //if (string.IsNullOrEmpty(this[y, x]))//!!!sometimes it looses value after setting Hyperlink (may be depends on cell format?)
//            //    this[y, x] = string.IsNullOrEmpty(v) ? LinkEmptyValueFiller : v;
//        }
//        public static string LinkEmptyValueFiller = "           ";

//        public Uri GetLink(int y, int x)
//        {
//            return worksheet.Cells[y, x].Hyperlink;
//        }

//        public string this[int y, int x]
//        {
//            get
//            {
//                return worksheet.Cells[y, x].Value?.ToString();
//            }
//            set
//            {
//                worksheet.Cells[y, x].Value = value;
//            }
//        }

//        public void InsertLine(int y, IEnumerable<object> values = null)
//        {
//            worksheet.InsertRow(y, 1);
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

//                worksheet.Cells[y, i++].Value = s;
//                //worksheet.Cells[y, i++].Style.Numberformat.Format = "@";
//            }
//        }

//        public void CreateDropdown(int y, int x, IEnumerable<object> values, object value, bool allowBlank = true)
//        {
//            var dv = worksheet.DataValidations[worksheet.Cells[y, x].Address];
//            if (dv != null)
//                worksheet.DataValidations.Remove(dv);
//            OfficeOpenXml.DataValidation.ExcelDataValidationList dvl = (OfficeOpenXml.DataValidation.ExcelDataValidationList)worksheet.Cells[y, x].DataValidation.AddListDataValidation();
//            dvl.AllowBlank = allowBlank;
//            foreach (object v in values)
//            {
//                string s;
//                if (v is string)
//                    s = (string)v;
//                else if (v != null)
//                    s = v.ToString();
//                else
//                    s = null;
//                dvl.Formula.Values.Add(s);
//            }
//            {
//                string s;
//                if (value is string)
//                    s = (string)value;
//                else if (value != null)
//                    s = value.ToString();
//                else
//                    s = null;
//                worksheet.Cells[y, x].Value = s;
//            }
//        }

//        public void AddImage(int y, int x, string name, Bitmap image)
//        {
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

//        public void HighlightRow(int y, Color color)
//        {
//            if (worksheet.Row(y).Style.Fill.PatternType == OfficeOpenXml.Style.ExcelFillStyle.None)
//                worksheet.Row(y).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
//            worksheet.Row(y).Style.Fill.BackgroundColor.SetColor(color);
//        }
//    }
//}