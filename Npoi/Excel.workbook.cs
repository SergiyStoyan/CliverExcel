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
    public partial class Excel : IDisposable
    {
        public IEnumerable<ISheet> GetSheets()
        {
            return Workbook._GetSheets();
        }

        /// <summary>
        /// Set the active sheet. If no sheet with such name exists, a new sheet is created.
        /// (!)The name will be corrected to remove unacceptable symbols.
        /// </summary>
        /// <param name="name"></param>
        /// <param name="createSheet"></param>
        public void OpenSheet(string name, bool createSheet = true)
        {
            Sheet = Workbook._OpenSheet(name, createSheet);
        }

        /// <summary>
        /// Set the active sheet.
        /// </summary>
        /// <param name="index">1-based</param>
        /// <returns>true if the index exists, otherwise false</returns>
        public bool OpenSheet(int index)
        {
            Sheet = Workbook._OpenSheet(index);
            return Sheet != null;
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
            Workbook._Save(File);
        }

        public string HyperlinkBase
        {
            get
            {
                return Workbook._GetHyperlinkBase();
            }
            set
            {
                Workbook._SetHyperlinkBase(value);
            }
        }

        public string Author
        {
            get
            {
                return Workbook._GetAuthor();
            }
            set
            {
                Workbook._SetAuthor(value);
            }
        }

        /// <summary>
        /// Set it to make Excel keep links absolute
        /// </summary>
        public const string AbsoluteLinksHyperlinkBase = "x";
    }
}