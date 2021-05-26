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
            if(!File.Exists(Path.Combine(folderName,dBFileName)))
            {
                               
                this.dbFile= WorkBook.Create(ExcelFileFormat.XLS);
                dbFile.Metadata.Author = "SBSC Segun";
                
                WorkSheet SavingsTable = dbFile.CreateWorkSheet("Saving");
                WorkSheet ExpensesTable = dbFile.CreateWorkSheet("Expense");
                
                SavingsTable["A1"].Value = "Id";
                SavingsTable["B1"].Value = "CreatedDateTime";
                SavingsTable["C1"].Value = "Amount";
                SavingsTable["D1"].Value = "Description";

                ExpensesTable["A1"].Value = "Id";
                ExpensesTable["B1"].Value = "CreatedDateTime";
                ExpensesTable["C1"].Value = "Amount";
                ExpensesTable["D1"].Value = "Description";

                dbFile.Save();
                
            }

           
        }


        public bool ReadFromFile()
        {
            return false;
        }

        public bool WriteToFile()
        {
            return false;
        }
        
    }
}
