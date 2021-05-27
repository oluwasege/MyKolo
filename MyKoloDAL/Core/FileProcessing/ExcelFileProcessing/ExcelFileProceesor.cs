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
        private WorkSheet SavingsTable;
        private WorkSheet ExpensesTable;
        private readonly string dBFileName = "MyKoloDb.xls";
        private int startingRow = 2;



        public ExcelFileProceesor() : base()
        {
            if (!File.Exists(Path.Combine(folderName, dBFileName)))
            {

                this.dbFile = WorkBook.Create(ExcelFileFormat.XLS);
                dbFile.Metadata.Author = "SBSC Segun";

                SavingsTable = dbFile.CreateWorkSheet("Saving");
                ExpensesTable = dbFile.CreateWorkSheet("Expense");

                SavingsTable["A1"].Value = "Id";
                SavingsTable["B1"].Value = "CreatedDateTime";
                SavingsTable["C1"].Value = "Amount";
                SavingsTable["D1"].Value = "Description";

                ExpensesTable["A1"].Value = "Id";
                ExpensesTable["B1"].Value = "CreatedDateTime";
                ExpensesTable["C1"].Value = "Amount";
                ExpensesTable["D1"].Value = "Description";

                dbFile.SaveAs(Path.Combine(folderName, dBFileName));

            }
        }
            // Database class 
            // work for all models. Expense and Savings
            // --> Add Single

            // Expenses[A2] = Expense.Id
            // Expenses[B2] = Expense.CreatedDateTime
            // Expenses[C2] = Expense.Amount
            // Expenses[D2] = Expense.Description.

            /**
             *  ExpenseService
             *  AddExpense(Expense e){
             *  
             *    // we can do some validation on the expense before saving to the database.
             *    database.AddRecord(e);
             *    
             *  }
             *  
             *  Savings 
             *  
             *  class SingleFieldMeta {
             *        public string ColumnName;
             *        public Object ColumnValue; 
             *        
             *  }
             *  
             *  Expense --> [
             *  {columnName:"ModelType", ColumnValue: "User"}
             *  {columnName:"Id", ColumnValue:20},
             *  {columnName:"CreatedDateTime", ColumnValue:12/12/2021}
             *  {columnName:"Amount", ColumnValue:100}
             *  {columnName:"Description", ColumnValue:"Valid Description}]
             *  



             *  
             *  Add Savings(Saving s){
             *    database.AddRecord(s);
             *    we can have expense and Savings inherit from a base model,
             *    and then 
             *    database.AddRecord(List<SingleFieldMeta> record){
             *      string TargetTable = "";
             *      
             *      if(record[0].columnValue = "Expense"){
             *       
             *      }        
             *      
             *      
             *          
             *    }
             *  
             *  }
             * 
             */



            // --> Delete Single
            // --> Update Single
            // --> Read Single


            // -- Add Multiple
            // -- Delete Multiple
            // -- Update Multiple 
            // -- Read Multiple

            // Logic for adding an expense
            // Expense Controller --> ExpenseService --> Database add a New Expense






        


        public bool ReadFromFile()
        {
            return false;
        }

        public bool WriteToFile()
        {
            return false;
        }

        /// <summary>
        /// This writes a new Record of Any Model Type that has been converted. To 
        /// the appropriate table (sheet) in the excel file.
        /// 
        /// LOGIC: 
        /// our write to file method because of its location in our application (inside the CORE),
        /// Needs to be callable from Any business Logic module. due to this need, we created 
        /// a single object that represente any Field and its corresponding value in Any object.
        /// 
        /// if we can collect these fields and values (which represent a record) 
        /// in a single List, 
        /// then we can write a general process on this List to save the values to the Database
        /// 
        /// </summary>
        /// <param name="convertedSingleRecordData">This is a List of objects representing
        /// each field in the record with its own Value
        /// </param>
        /// <returns></returns>

        public bool WriteToFile(List<SingleFieldMeta> convertedSingleRecordData)
        {
            WorkSheet TargetTable = null;
            switch(convertedSingleRecordData[0].ColunValue.ToString())
            {
                case "Expenses":
                    TargetTable = ExpensesTable;
                    break;
                case "Savings":
                    TargetTable = SavingsTable;
                    break;
            }

            int recordCount = (int)TargetTable["F2"].Value;
            int rowToInsertRecord = recordCount + startingRow;

          for(int index=1;index<=convertedSingleRecordData.Count;index++)
            {
                SingleFieldMeta targetField = convertedSingleRecordData[index];
                //unicode character reference
                char targetColumn = Convert.ToChar(64 + index);
                string targetCell = string.Concat(targetColumn.ToString(), rowToInsertRecord.ToString());

                //set the value after we have found the target cell.
                IronXL.Range targetRange = TargetTable[targetCell];
                targetRange.Value = targetField.ColunValue;
            }

            TargetTable["F2"].Value = (int)recordCount++;

            return true;
        }
    }
}
