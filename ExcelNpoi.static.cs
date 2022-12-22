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
using Newtonsoft.Json.Serialization;
using System.Reflection;
using Newtonsoft.Json;

//works  
namespace Cliver
{
    public partial class Excel : IDisposable
    {
        static public string GetColumnName(int x)
        {
            return CellReference.ConvertNumToColString(x - 1);
        }
    }
}