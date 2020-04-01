using System;
using Assignment4.Models;

namespace Assignment4
{
    class Program
    {
        static void Main(string[] args)
        {
            Word.CreateWordprocessingDocument($"{ Constants.Locations.DataFolder}\\info.docx");

            Excel.CreateSpreadsheetWorkbook($"{ Constants.Locations.DataFolder}\\info.xlsx");
            Excel.InsertText($"{ Constants.Locations.DataFolder}\\info.xlsx");

            ppt.CreatePresentation(@"C:\Users\Nipin's HP\Desktop\info.pptx");
            insertPpt.InsertNewSlide(@"C:\Users\Nipin's HP\Desktop\info.pptx", 1, "My new slide");


        }
    }
}
