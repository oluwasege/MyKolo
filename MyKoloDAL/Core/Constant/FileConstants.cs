using System;
using System.Collections.Generic;
using System.IO;
using System.Text;


namespace MyKoloDAL.Core.Constant
{
    public class FileConstants
    {
        public readonly static string DBFOLDER= Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "MyKoloDb");
    }
}
