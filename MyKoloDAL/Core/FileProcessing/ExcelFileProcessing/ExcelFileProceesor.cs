using System;
using System.Collections.Generic;
using System.Text;
using MyKoloDAL.Core.FileProcessing.Interfaces;
using IronXL;
using System.IO;
using MyKoloDAL.Core.Constant;

namespace MyKoloDAL.Core.FileProcessing.ExcelFileProcessing
{
    public class ExcelFileProceesor : FileProcessorBase,IFileProcessor
    {
        private WorkBook dbFile;
        private string dBFileName = "MyKoloDb.xls";

        public ExcelFileProceesor():base()
        {
                        

            //Create new Excel WorkBook document. 
            //The default file format is XLSX, but we can override that for legacy support
            WorkBook xlsWorkbook = WorkBook.Create(ExcelFileFormat.XLS);
            xlsWorkbook.Metadata.Author = "IronXL";
            //Add a blank WorkSheet
            WorkSheet xlsSheet = xlsWorkbook.CreateWorkSheet("new_sheet");
            //Add data and styles to the new worksheet
            xlsSheet["A1"].Value = "Hello World";
            xlsSheet["A2"].Style.BottomBorder.SetColor("#ff6600");
            xlsSheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Double;
            //Save the excel file
            xlsWorkbook.SaveAs("NewExcelFile.xls");
        }


        public bool ReadToFile()
        {
            return false;
        }

        public bool WriteToFile()
        {
            return false;
        }
        
    }
}
