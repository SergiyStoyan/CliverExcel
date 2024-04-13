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
//using Microsoft.Office.Interop.Excel;
//using System.Runtime.InteropServices;

////TBD
////- test NPOI;
////- improve AutoFit;
////- implement AddImage/GetImage;
////- fix Dispose;
///

//!!!works but very slow!!!
//namespace Cliver
//{
//    public class Excel : IDisposable
//    {
//        static Excel()
//        {
//        }

//        public Excel(string file)
//        {
//            init(file);
//            OpenWorksheet(1);
//        }

//        void init(string file)
//        {
//            File = file;
//            baseUri = new Uri(File);
//            lock (staticLockObject)
//            {
//                if (excel == null)
//                {
//                    excel = new Application();

//                    GetWindowThreadProcessId(excel.Hwnd, out int id);
//                    Win.ProcessRoutines.AntiZombieGuard.This.Track(Win.ProcessRoutines.GetProcess(id));

//                    excel.Visible = false;
//                    excel.DisplayAlerts = false;
//                    excel.UserControl = false;
//                    excel.ScreenUpdating = false;
//                    //excel.AskToUpdateLinks
//                }

//                if (System.IO.File.Exists(File))
//                    workbook = excel.Workbooks.Open(File);
//                else
//                    workbook = excel.Workbooks.Add(Type.Missing);
//            }
//        }

//        static Application excel;
//        static object staticLockObject = new object();

//        Workbook workbook;
//        Worksheet worksheet;

//        public Excel(string file, string worksheetName)
//        {
//            init(file);
//            OpenWorksheet(worksheetName);
//        }

//        public string File;

//        ~Excel()
//        {
//            Dispose();
//        }

//        public void Dispose()
//        {
//            lock (staticLockObject)
//            {
//                if (excel == null)
//                    return;

//                //if (cell != null)
//                //    while (System.Runtime.InteropServices.Marshal.ReleaseComObject(cell) > 0) ;
//                //cell = null;

//                disposeCurrentWorksheet();

//                if (workbook != null)
//                {
//                    //while (Marshal.ReleaseComObject(workbook.Worksheets) > 0) ;
//                    try
//                    {
//                        workbook.Close(false, File, null);//Workbook.close SaveChanges, filename, routeworkbook 
//                    }
//                    catch (Exception e)
//                    { }
//                    while (Marshal.ReleaseComObject(workbook) > 0) ;
//                    workbook = null;
//                }

//                if (excel.Workbooks.Count < 1)
//                {
//                    try
//                    {
//                        excel.Workbooks.Close();
//                    }
//                    catch (Exception e)
//                    { }
//                    while (Marshal.ReleaseComObject(excel.Workbooks) > 0) ;

//                    GetWindowThreadProcessId(excel.Hwnd, out int id);
//                    excel.Quit();
//                    try
//                    {
//                        while (Marshal.ReleaseComObject(excel) > 0) ;
//                    }
//                    catch (Exception e)
//                    { }
//                    if (id > 0 && Win.ProcessRoutines.GetProcess(id) != null)
//                    {
//                        Log.Warning2("Killing Excel...");
//                        Win.ProcessRoutines.TryKillProcessTree(id);
//                    }

//                    excel = null;
//                }
//            }
//        }
//        [DllImport("user32.dll")]
//        static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);

//        void disposeCurrentWorksheet()
//        {
//            if (worksheet != null)
//                while (Marshal.ReleaseComObject(worksheet) > 0) ;
//            worksheet = null;
//        }

//        public string HyperlinkBase
//        {
//            get
//            {
//                dynamic dps = workbook.BuiltinDocumentProperties;
//                dynamic dp = dps["Hyperlink Base"];
//                //var c = dp.Creator;
//                //var t = dp.Type;
//                //var a = dp.Application;
//                //var n = dp.Name;
//                ////var d = dp.LinkSource;
//                var v = dp.Value;
//                if (string.IsNullOrWhiteSpace(v))
//                    return null;
//                return (string)v;
//            }
//            set
//            {
//                //DocumentProperties dps = (DocumentProperties)workbook.BuiltinDocumentProperties;
//                //dynamic dps = workbook.CustomDocumentProperties;
//                //dps["HyperlinkBase"].Value = value.ToString();
//                dynamic dps = workbook.BuiltinDocumentProperties;
//                dps["Hyperlink Base"].Value = value;
//            }
//        }

//        public void OpenWorksheet(string name)
//        {
//            disposeCurrentWorksheet();
//            worksheet = (Worksheet)workbook.Worksheets[name];
//            if (worksheet == null)
//                worksheet = (Worksheet)workbook.Worksheets.Add(name);
//        }

//        public bool OpenWorksheet(int index)
//        {
//            if (index == worksheet?.Index)
//                return true;
//            disposeCurrentWorksheet();
//            worksheet = (Worksheet)workbook.Worksheets[index];
//            return worksheet != null;
//        }

//        public string WorksheetName
//        {
//            get
//            {
//                return worksheet?.Name;
//            }
//            set
//            {
//                if (worksheet != null)
//                    worksheet.Name = value;
//            }
//        }

//        public void Save()
//        {
//            workbook.SaveAs(File);
//        }

//        public int GetLastUsedRow()
//        {
//            if (worksheet == null)
//                throw new Exception("No active worksheet.");
//            //Range r = worksheet.Cells.Find("*", System.Reflection.Missing.Value, XlFindLookIn.xlValues, XlLookAt.xlWhole, XlSearchOrder.xlByRows, XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
//            //int i1 = r == null ? 0 : r.Row;
//            //int i2 = getLastRow();
//            //if (i1 != i2)
//            //    return i1 < i2 ? i1 : i2;
//            return getLastUsedRow();
//        }

//        int getLastUsedRow()
//        {
//            for (int y = worksheet.UsedRange.Rows.Count; y > 0; y--)
//                if (!string.IsNullOrEmpty(((Range)worksheet.Cells[y, 1]).End[XlDirection.xlToRight].Text?.ToString()))
//                    return y;
//            return 0;
//        }

//        public int AppendLine(IEnumerable<object> values)
//        {
//            int y = GetLastUsedRow() + 1;
//            WriteLine(y, values);
//            return y;
//        }

//        public void SetLink(int y, int x, Uri uri)
//        {
//            Range r = (Range)worksheet.Cells[y, x];
//            if (string.IsNullOrEmpty(r.Value2?.ToString()))
//                r.Value2 = LinkEmptyValueFiller;
//            r.Hyperlinks.Delete();
//            r.Hyperlinks.Add(worksheet.Cells[y, x], uri.ToString());
//        }
//        public static string LinkEmptyValueFiller = "           ";

//        public Uri GetLink(int y, int x)
//        {
//            Hyperlinks hs = ((Range)worksheet.Cells[y, x]).Hyperlinks;
//            if (hs.Count > 0)
//            {
//                Uri u = new Uri(hs[1].Address, UriKind.RelativeOrAbsolute);
//                if (u.IsAbsoluteUri)
//                    return u;
//                return new Uri(baseUri, u);
//            }
//            return null;
//        }
//        Uri baseUri;

//        public string this[int y, int x]
//        {
//            get
//            {
//                return ((Range)worksheet.Cells[y, x]).Text?.ToString();
//            }
//            set
//            {
//                worksheet.Cells[y, x] = value;
//            }
//        }

//        public void InsertLine(int y, IEnumerable<object> values = null)
//        {
//            ((Range)worksheet.Rows[y]).Insert(XlInsertShiftDirection.xlShiftDown);
//            if (values != null)
//                WriteLine(y, values);
//        }

//        public void WriteLine(int y, IEnumerable<object> values)
//        {
//            //worksheet.Cells[y, 1] = values.ToArray();
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

//                worksheet.Cells[y, i++] = s;
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

//            Range cell = (Range)worksheet.Cells[y, x];
//            cell.Validation.Delete();
//            cell.Validation.Add(
//               XlDVType.xlValidateList,
//               XlDVAlertStyle.xlValidAlertInformation,
//               XlFormatConditionOperator.xlBetween,
//               string.Join(",", vs),
//               Type.Missing
//               );
//            cell.Validation.IgnoreBlank = allowBlank;
//            cell.Validation.InCellDropdown = true;

//            {
//                string s;
//                if (value is string)
//                    s = (string)value;
//                else if (value != null)
//                    s = value.ToString();
//                else
//                    s = null;
//                cell.Value2 = s;
//            }
//        }

//        public void AddImage(int y, int x, string name, Bitmap image)
//        {
//            Range c = (Range)worksheet.Cells[y, x];
//            string f = Program.TempDir + "\\" + name;
//            image.Save(f);
//            Shape s = worksheet.Shapes.AddPicture(f, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, (float)c.Left, (float)c.Top, image.Width, image.Height);
//            s.Name = name;
//            s.Placement = XlPlacement.xlMove;
//        }

//        public Bitmap GetImage(int y, int x)
//        {
//            //foreach (Shape s in worksheet.Shapes)s. ;.c .shap; w.pic ws;ws.pi
//            return null;
//        }

//        public void FitColumnsWidth(params int[] columnIs)
//        {
//            foreach (int i in columnIs)
//            {
//                Range cell1 = (Range)worksheet.Cells[2, i];
//                worksheet.Range[cell1, cell1.End[XlDirection.xlDown]].Columns.AutoFit();
//            }
//        }

//        public void FitColumnsWidth(int column1I, int column2I)
//        {
//            Range cell1 = (Range)worksheet.Cells[2, column1I];
//            Range cell2 = ((Range)worksheet.Cells[2, column2I]).End[XlDirection.xlDown];
//            worksheet.Range[cell1, cell2].Columns.AutoFit();
//        }

//        public void HighlightRow(int y, Color color)
//        {
//            ((Range)worksheet.Rows[y]).Interior.Color = ColorTranslator.ToOle(color);
//        }

//        public void ClearHighlighting()
//        {
//            Range cell1 = (Range)worksheet.Cells[1, 1];
//            worksheet.Range[cell1, cell1.End[XlDirection.xlToRight].End[XlDirection.xlDown]].Interior.ColorIndex = 0;
//        }
//    }
//}