//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************

using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.IO;
using System;
using System.Linq;

namespace Cliver
{
    static public partial class ExcelExtensions
    {
        static public Excel _Excel(this IWorkbook workbook)
        {
            return Cliver.Excel.Get(workbook);
        }

        public static void _RemoveSheet(this IWorkbook workbook, ISheet sheet)
        {
            int si = workbook.GetSheetIndex(sheet);
            if (si >= 0)
                workbook.RemoveSheetAt(si);
        }

        public static void _RemoveSheet(this IWorkbook workbook, string sheetName)
        {
            int si = workbook.GetSheetIndex(sheetName);
            if (si >= 0)
                workbook.RemoveSheetAt(si);
        }

        public static void _RemoveSheet(this IWorkbook workbook, int sheetIndex)
        {
            workbook.RemoveSheetAt(sheetIndex - 1);
        }

        public static void _SetAuthor(this IWorkbook workbook, string author)
        {
            if (workbook is XSSFWorkbook xSSFWorkbook)
                xSSFWorkbook.GetProperties().CoreProperties.Creator = author;
            else if (workbook is HSSFWorkbook hSSFWorkbook)
                hSSFWorkbook.SummaryInformation.Author = author;
            else
                throw new Exception("Unsupported workbook type: " + workbook.GetType().FullName);
        }

        public static string _GetAuthor(this IWorkbook workbook)
        {
            if (workbook is XSSFWorkbook xSSFWorkbook)
                return xSSFWorkbook.GetProperties().CoreProperties.Creator;
            else if (workbook is HSSFWorkbook hSSFWorkbook)
                return hSSFWorkbook.SummaryInformation.Author;
            else
                throw new Exception("Unsupported workbook type: " + workbook.GetType().FullName);
        }

        public static IEnumerable<ISheet> _GetSheets(this IWorkbook workbook)
        {
            for (int i = workbook.NumberOfSheets - 1; i >= 0; i--)
                yield return workbook.GetSheetAt(i);
        }

        /// <summary>
        /// (!)The name will be corrected by altering unacceptable symbols.
        /// </summary>
        /// <param name="name"></param>
        /// <param name="createSheet"></param>
        /// <returns></returns>
        static public ISheet _GetSheet(this IWorkbook workbook, string name, bool createSheet = true)
        {
            ISheet sheet = workbook.GetSheet(name);
            if (sheet == null && createSheet)
                sheet = workbook.CreateSheet(Excel.GetSafeSheetName(name));
            workbook.SetActiveSheet(sheet._GetIndex() - 1);
            return sheet;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="index">1-based</param>
        /// <returns></returns>
        static public ISheet _GetSheet(this IWorkbook workbook, int index)
        {
            if (workbook.NumberOfSheets < 1 || workbook.NumberOfSheets < index)
                return null;
            //workbook.SetActiveSheet(index - 1);
            return workbook.GetSheetAt(index - 1);
        }

        static public void _Save(this IWorkbook workbook, string file)
        {
            using (var fileData = new FileStream(file, FileMode.Create))
            {
                workbook.Write(fileData, true);
            }
        }

        static public void _SetHyperlinkBase(this IWorkbook workbook, string value)
        {
            if (workbook is XSSFWorkbook xSSFWorkbook)
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
            else if (workbook is HSSFWorkbook hSSFWorkbook)
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
                throw new Exception("Unsupported workbook type: " + workbook.GetType().FullName);
        }

        static public string _GetHyperlinkBase(this IWorkbook workbook)
        {
            if (workbook is XSSFWorkbook xSSFWorkbook)
            {
                NPOI.OpenXmlFormats.CT_Property p = xSSFWorkbook.GetProperties().CustomProperties.GetProperty("HyperlinkBase");//so is in Epplus
                return p?.Item?.ToString();
            }
            else if (workbook is HSSFWorkbook hSSFWorkbook)
            {
                hSSFWorkbook.CreateInformationProperties();
                return hSSFWorkbook.DocumentSummaryInformation.CustomProperties["HyperlinkBase"]?.ToString();
            }
            else
                throw new Exception("Unsupported workbook type: " + workbook.GetType().FullName);
        }

        public static IEnumerable<Excel.RichTextStringFormattingRun> _GetRichTextStringFormattingRuns(this IWorkbook workbook, IRichTextString text)
        {
            int nfrsCount = text.NumFormattingRuns;
            List<int> nfrs = new List<int>();
            for (int i = 0; i < nfrsCount; i++)
                nfrs.Add(text.GetIndexOfFormattingRun(i));
            nfrs.Add(text.Length);
            if (text is HSSFRichTextString hText)
            {
                for (int i = 0; i < nfrsCount; i++)
                {
                    var f = workbook.GetFontAt(hText.GetFontOfFormattingRun(i));
                    if (f != null)
                        yield return new Excel.RichTextStringFormattingRun(nfrs[i], nfrs[i + 1], f);
                }
            }
            else if (text is XSSFRichTextString xText)
            {
                //if (nfrsCount < 1)//(!)it can have no FormattingRun
                //{
                //    xText.Append("");// makes NumFormattingRuns > 0
                //    for (int i = 0; i < nfrsCount; i++)
                //        nfrs.Add(text.GetIndexOfFormattingRun(i));
                //    nfrs.Add(text.Length);
                //}
                for (int i = 0; i < nfrsCount; i++)
                {
                    var f = xText.GetFontOfFormattingRun(i);
                    if (f != null)
                        yield return new Excel.RichTextStringFormattingRun(nfrs[i], nfrs[i + 1], f);
                }
            }
            else
                throw new Exception("Unsupported type of IRichTextString: " + text.GetType());
        }

        public static IRichTextString _GetRichTextString(this IWorkbook workbook, string @string, IEnumerable<Excel.RichTextStringFormattingRun> richTextStringFormattingRuns)
        {
            if (workbook is XSSFWorkbook) //(!)NPOI bug work around
            {
                XSSFRichTextString text = new XSSFRichTextString();
                int p = 0;
                foreach (var rtsfr in richTextStringFormattingRuns?.Where(a => a?.Font != null && a.Start < @string.Length).OrderBy(a => a.Start))
                {
                    if (p < rtsfr.Start)
                        text.Append(@string.Substring(p, rtsfr.Start - p));
                    text.Append(@string.Substring(rtsfr.Start, (rtsfr.ExcludedEnd < @string.Length ? rtsfr.ExcludedEnd : @string.Length) - rtsfr.Start), (XSSFFont)rtsfr.Font);
                    p = rtsfr.ExcludedEnd;
                }
                if (p < @string.Length)
                    text.Append(@string.Substring(p, @string.Length - p));
                return text;
            }
            else
            {
                IRichTextString text = workbook.GetCreationHelper().CreateRichTextString(@string);
                richTextStringFormattingRuns?.Where(a => a?.Font != null).ForEach(a => text.ApplyFont(a.Start, a.ExcludedEnd, a.Font));
                return text;
            }
        }
    }
}