using System;
using System.Collections.Generic;
using System.Text;


namespace MyKoloDAL.Core.FileProcessing.Interfaces
{
    internal interface IFileProcessor
    {
        bool ReadFromFile();
        bool WriteToFile();           
    }
}
