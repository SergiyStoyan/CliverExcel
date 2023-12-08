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

namespace Cliver
{
    static public partial class ExcelExtensions
    {
        public static void _RemoveSheet(this IWorkbook workbook, ISheet sheet)
        {
            workbook.RemoveSheetAt(workbook.GetSheetIndex(sheet));
        }

        public static void _RemoveSheet(this IWorkbook workbook, string sheetName)
        {
            workbook.RemoveSheetAt(workbook.GetSheetIndex(sheetName));
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
        /// (!)The name will be corrected to remove unacceptable symbols.
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="name"></param>
        /// <param name="createSheet"></param>
        /// <returns></returns>
        static public ISheet _OpenSheet(this IWorkbook workbook, string name, bool createSheet = true)
        {
            ISheet sheet = workbook.GetSheet(name);
            if (sheet == null && createSheet)
                sheet = workbook.CreateSheet(Excel.GetSafeSheetName(name));
            workbook.SetActiveSheet(sheet._GetIndex() - 1);
            return sheet;
        }

        /// <summary>
        /// 1-based
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        static public ISheet _OpenSheet(this IWorkbook workbook, int index)
        {
            if (workbook.NumberOfSheets < 1 || workbook.NumberOfSheets < index)
                return null;
            workbook.SetActiveSheet(index - 1);
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
    }
}