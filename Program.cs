using System;
using CSV.Models;
namespace Office_File_Formats
{
    class Program
    {
        static void Main(string[] args)
        {
            //Done
            //Models.wordFile.CreateWordprocessingDocument($"{ Constants.Locations.DataFolder}\\info.docx");
            
            Models.excelFile.CreateSpreadsheetWorkbook($"{ Constants.Locations.DataFolder}\\info.xlsx");
            Models.excelFile.InsertText($"{ Constants.Locations.DataFolder}\\info.xlsx");
        }
    }
}
