//********************************************************************************************
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

//works  
namespace Cliver
{
    public partial class Excel
    {
        static public string GetSafeSheetName(string name)
        {
            name = Regex.Replace(name, @"\:", "-");//npoi does not accept :
            return WorkbookUtil.CreateSafeSheetName(name);
        }

        static public string GetColumnName(int x)
        {
            return CellReference.ConvertNumToColString(x - 1);
        }

        static public int GetX(string columnName)
        {
            return CellReference.ConvertColStringToIndex(columnName) + 1;
        }

        static public (int Y, int X) GetCoordinates(string address)
        {
            var a = ParseAddress(address);
            return (a.Y, GetX(a.ColumnName));
        }

        static public (int Y, string ColumnName) ParseAddress(string address)
        {
            Match m = Regex.Match(address, @"^\s*([a-z]+)(\d+)\s*$", RegexOptions.IgnoreCase);
            if (!m.Success)
                throw new Exception("Address is not parsable: " + address);
            return (int.Parse(m.Groups[2].Value), m.Groups[1].Value);
        }
    }
}