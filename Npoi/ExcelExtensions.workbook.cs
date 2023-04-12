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
        public static IEnumerable<ISheet> _GetSheets(this IWorkbook workbook)
        {
            for (int i = 0; i < workbook.ActiveSheetIndex; i++)
            {
                yield return workbook.GetSheetAt(i);
            }
        }

        /// <summary>
        /// If no sheet with such name exists, a new sheet is created. 
        /// (!)The name will be corrected to remove unacceptable symbols.
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        static public ISheet _OpenSheet(this IWorkbook workbook, string name)
        {
            ISheet sheet = workbook.GetSheet(name);
            if (sheet == null)
                sheet = workbook.CreateSheet(Excel.GetSafeSheetName(name));
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
            if (workbook.NumberOfSheets > 0 && workbook.NumberOfSheets >= index)
            {
                return workbook.GetSheetAt(index - 1);
            }
            return null;
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