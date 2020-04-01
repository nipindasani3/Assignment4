using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using System;
using System.IO;
using System.Collections.Generic;
using static Assignment4.Models.Constants;

namespace Assignment4.Models
{
    class Excel
    {
        public static void CreateSpreadsheetWorkbook(string filepath)
        {
            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook))
            {
                // Add a WorkbookPart to the document.
                WorkbookPart workbookpart = spreadSheet.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();

                // Add a WorksheetPart to the WorkbookPart.
                WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                // Add Sheets to the Workbook.
                Sheets sheets = spreadSheet.WorkbookPart.Workbook.
                    AppendChild<Sheets>(new Sheets());

                // Append a new worksheet and associate it with the workbook.
                Sheet sheet = new Sheet()
                {
                    Id = spreadSheet.WorkbookPart.
                    GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "Sheet1"
                };
                sheets.Append(sheet);

                workbookpart.Workbook.Save();
                Cell cell = InsertCellInWorksheet("A", 2, worksheetPart);

                // Set the value of cell A1.
                cell.CellValue = new CellValue("My Name is Nipin Dasani");
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                // Save the new worksheet.
                worksheetPart.Worksheet.Save();
            }
        }
        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }
        public static void InsertText(string docName)
        {
            // Open the document for editing.
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true))
            {
                //UniqueId(generate a Guid from your application using Guid.NewGuid() in C#)
                // Insert a new worksheet.
                WorksheetPart worksheetPart = InsertWorksheet(spreadSheet.WorkbookPart);
                Cell cell = InsertCellInWorksheet("A", 1, worksheetPart);
                // Set the value of cell A.
                cell.CellValue = new CellValue("UniqueId");
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell = InsertCellInWorksheet("B", 1, worksheetPart);
                // Set the value of cell B.
                cell.CellValue = new CellValue("StudentId");
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                cell = InsertCellInWorksheet("C", 1, worksheetPart);
                // Set the value of cell C.
                cell.CellValue = new CellValue("FirstName");
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                cell = InsertCellInWorksheet("D", 1, worksheetPart);
                // Set the value of cell D.
                cell.CellValue = new CellValue("LastName");
                cell.DataType = new EnumValue<CellValues>(CellValues.Date);
                cell = InsertCellInWorksheet("E", 1, worksheetPart);
                cell.CellValue = new CellValue("DateOfBirth");
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell = InsertCellInWorksheet("F", 1, worksheetPart);
                // Set the value of cell F.
                cell.CellValue = new CellValue("IsMe");
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell = InsertCellInWorksheet("G", 1, worksheetPart);
                // Set the value of cell F.
                cell.CellValue = new CellValue("Age");
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                List<Student> studentlist = new List<Student>();
               
                uint i = 2;
                foreach (var student in studentlist)
                {
                    Guid g = Guid.NewGuid();
                    cell = InsertCellInWorksheet("A", i, worksheetPart);
                    // Set the value of cell A.
                    cell.CellValue = new CellValue(g.ToString());
                    // Insert cell A1 into the new worksheet.
                    cell = InsertCellInWorksheet("B", i, worksheetPart);
                    // Set the value of cell A.
                    cell.CellValue = new CellValue(student.StudentId);
                    cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                    cell = InsertCellInWorksheet("C", i, worksheetPart);
                    // Set the value of cell B.
                    cell.CellValue = new CellValue(student.FirstName);
                    cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                    cell = InsertCellInWorksheet("D", i, worksheetPart);
                    // Set the value of cell C.
                    cell.CellValue = new CellValue(student.LastName);
                    cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                    cell = InsertCellInWorksheet("E", i, worksheetPart);
                    // Set the value of cell D.
                    cell.CellValue = new CellValue(student.DateOfBirth);
                    cell.DataType = new EnumValue<CellValues>(CellValues.Date);
                    cell = InsertCellInWorksheet("F", i, worksheetPart);
                    setMyRecord(student);
                    // Set the value of cell E.
                    if (student.IsMe == true)
                    {
                        cell.CellValue = new CellValue("1");
                    }
                    else
                    {
                        cell.CellValue = new CellValue("0");
                    }
                    cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                    cell = InsertCellInWorksheet("G", i, worksheetPart);
                    // Set the value of cell F.
                    cell.CellValue = new CellValue(Convert.ToString(student.Age));
                    cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                    i = i + 1;
                }

                // Save the new worksheet.
                worksheetPart.Worksheet.Save();
            }
        }

        private static WorksheetPart InsertWorksheet(WorkbookPart workbookPart)
        {
            // Add a new worksheet part to the workbook.
            WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());
            newWorksheetPart.Worksheet.Save();

            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

            // Get a unique ID for the new sheet.
            uint sheetId = 1;

            if (sheets.Elements<Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }



            string sheetName = "Sheet" + sheetId;

            // Append the new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);
            workbookPart.Workbook.Save();

            return newWorksheetPart;
        }

        public static void setMyRecord(Student studennt)
        {

            if (studennt.StudentId == "200447887")
            {
                studennt.IsMe = true;
            }

        }


    }
}

