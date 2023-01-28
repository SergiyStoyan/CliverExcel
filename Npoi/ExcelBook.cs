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
//using NPOI.XSSF.UserModel;
//using NPOI.HSSF.UserModel;
//using NPOI.SS.UserModel;
//using NPOI.SS.Util;
//using NPOI.SS.Formula.PTG;
//using NPOI.SS.Formula;

//namespace Cliver
//{
//    public partial class ExcelBook : IDisposable
//    {
//        readonly static List<ExcelBook> excelBooks = new List<ExcelBook>();

//        static public ExcelBook Get(string file)
//        {
//            lock (excelBooks)
//            {
//                var eb = excelBooks.FirstOrDefault(a => a.File == file);
//                if (eb?.Disposed == true)
//                {
//                    excelBooks.Remove(eb);
//                    eb = null;
//                }
//                if (eb == null)
//                {
//                    eb = new ExcelBook(file);
//                    excelBooks.Add(eb);
//                }
//                return eb;
//            }
//        }

//        ExcelBook(string file)
//        {
//            File = file;

//            if (System.IO.File.Exists(File))
//                using (FileStream fs = new FileStream(File, FileMode.Open, FileAccess.Read))
//                {
//                    try
//                    {
//                        fs.Position = 0;//!!!prevents occasional error: EOF in header
//                        Workbook = new XSSFWorkbook(fs);
//                        //FormulaEvaluator = new XSSFFormulaEvaluator(Workbook);
//                    }
//                    catch (ICSharpCode.SharpZipLib.Zip.ZipException)
//                    {
//                        fs.Position = 0;//!!!prevents error: EOF in header
//                        Workbook = new HSSFWorkbook(fs);//old Excel 97-2003
//                        //FormulaEvaluator = new HSSFFormulaEvaluator(Workbook);
//                    }
//                }
//            else
//                Workbook = new XSSFWorkbook();
//        }

//        public IWorkbook Workbook { get; private set; }

//        public string File { get; private set; }

//        ~ExcelBook()
//        {
//            Dispose();
//        }

//        public void Dispose()
//        {
//            lock (this)
//            {
//                if (Workbook != null)
//                {
//                    Workbook.Close();
//                    Workbook = null;
//                }
//            }
//        }

//        public bool Disposed { get { return Workbook == null; } }

//        public void Save(string file = null)
//        {
//            if (file != null)
//                File = file;
//            using (var fileData = new FileStream(File, FileMode.Create))
//            {
//                Workbook.Write(fileData, true);
//            }
//        }
//    }
//}