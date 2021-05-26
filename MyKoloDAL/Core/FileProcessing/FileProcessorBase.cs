using MyKoloDAL.Core.Constant;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;


namespace MyKoloDAL.Core.FileProcessing
{
    public class FileProcessorBase
    {
        protected string folderName = FileConstants.DBFOLDER;
        
        public FileProcessorBase()
        {
            
           

            if (!Directory.Exists(folderName))
            {
                Directory.CreateDirectory(folderName);
            }
        }
    }
}
