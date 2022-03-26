using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;

namespace DocxToXLSX
{
    public partial class Form2 : Form
    {
        static string mycontent = "";

        public Form2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(Application.StartupPath + @"\BILL#1-2-ch.docx", true);

            //Body body = wordprocessingDocument.MainDocumentPart.Document.Body;

            //MessageBox.Show(wordprocessingDocument.GetAllParts().ToString());


            //MessageBox.Show(body.InnerText);

            getParagraphs(Application.StartupPath + @"\BILL#1-2-ch.docx");



            //wordprocessingDocument.SaveAs(Application.StartupPath + @"\BILL#1-2-ch.xlsx");

            //wordprocessingDocument.Close();
            //body = null;
            //wordprocessingDocument = null;


        }



        static void getParagraphs(string filename)
        {

            CreateSpreadsheetWorkbook(Application.StartupPath + @"\BILL#1-2-ch.xlsx");

            using (WordprocessingDocument myDoc = WordprocessingDocument.Open(filename, true))
            {
                Body body = myDoc.MainDocumentPart.Document.Body;

                string chapter = "";
                string subHeading = "";
                string content = "";


                foreach (var paragraph in body)
                {
                    if (!string.IsNullOrEmpty(paragraph.InnerText.Trim()))
                    {
                        if (paragraph.InnerText.Trim().Contains("Chapter"))
                        {
                            chapter = paragraph.InnerText.Trim();
                        }
                        else
                        {
                            if (paragraph.InnerText.Trim().StartsWith("-"))
                            {
                                subHeading = paragraph.InnerText.Trim();
                            }
                            else
                            {
                                content += paragraph.InnerText.Trim() + "*";

                                //AddUpdateCellValue(SpreadsheetDocument spreadSheet, "Sheet1", 1, "A", content);

                                //WriteText(Application.StartupPath + @"\BILL#1-2-ch.xlsx", content, 1, "A");
                            }
                        }
                    }
                }

                mycontent = content;

                InsertText(Application.StartupPath + @"\BILL#1-2-ch.xlsx", content);

            }
        }



        public static void CreateSpreadsheetWorkbook(string filepath)
        {
            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook);

            // Add a WorkbookPart to the document.
            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            //Sheet sheet = new Sheet()
            //{
            //    Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
            //    SheetId = 1,
            //    Name = "mySheet"
            //};
            //sheets.Append(sheet);

            workbookpart.Workbook.Save();

            // Close the document.
            spreadsheetDocument.Close();
        }



        // Given a document name and text, 
        // inserts a new worksheet and writes the text to cell "A1" of the new worksheet.
        public static void WriteText(string docName, string text, uint rowIndex, string columnName)
        {
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true))
            {
                WorksheetPart worksheetPart = GetWorksheetPartByName(spreadSheet, "");
                if (worksheetPart != null)
                {
                    // Create new Worksheet
                    Worksheet worksheet = new Worksheet();
                    worksheetPart.Worksheet = worksheet;

                    // Create new SheetData
                    SheetData sheetData = new SheetData();

                    // Create new row
                    Row row = new Row() { RowIndex = rowIndex };

                    // Create new cell
                    Cell cell = new Cell() { CellReference = columnName + rowIndex, DataType = CellValues.Number, CellValue = new CellValue(text) };

                    // Append cell to row
                    row.Append(cell);

                    // Append row to sheetData
                    sheetData.Append(row);

                    // Append sheetData to worksheet
                    worksheet.Append(sheetData);

                    worksheetPart.Worksheet.Save();
                }
                spreadSheet.WorkbookPart.Workbook.Save();
            }
        }





        // Given a document name and text, 
        // inserts a new worksheet and writes the text to cell "A1" of the new worksheet.
        public static void InsertText(string docName, string text)
        {
            // Open the document for editing.
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true))
            {

                //AddUpdateCellValue(spreadSheet, "Sheet1", 1, "A", text);
                

                //InsertRow("Sheet1", spreadSheet.WorkbookPart, 1);


                // Get the SharedStringTablePart. If it does not exist, create a new one.
                SharedStringTablePart shareStringPart;
                if (spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                {
                    shareStringPart = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                }
                else
                {
                    shareStringPart = spreadSheet.WorkbookPart.AddNewPart<SharedStringTablePart>();
                }



                //var arr = text.Split('*');


                //Insert the text into the SharedStringTablePart.
                int index = InsertSharedStringItem(text, shareStringPart);

                //Insert a new worksheet.
               WorksheetPart worksheetPart = AddWorksheetOnceOnly(spreadSheet.WorkbookPart);





                if (index == 0)
                {
                    // Insert cell A1 into the new worksheet.
                    Cell cell = InsertCellInWorksheet1("A", 1, worksheetPart);
                    // Set the value of cell A1.
                    cell.CellValue = new CellValue(index.ToString());
                    cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                }
                else
                {
                    // Insert cell A1 into the new worksheet.
                    Cell cell = InsertCellInWorksheet1("B", (uint)index, worksheetPart);
                    // Set the value of cell A1.
                    cell.CellValue = new CellValue(index.ToString());
                    cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                }


                // Save the new worksheet.
                worksheetPart.Worksheet.Save();

            }
        }



        // Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
        // and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
        private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.

            //var arr = text.Split('*');

            //for (int a = 0; a <= arr.Length - 1; a++)
            //{
            //    shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(arr[a])));
            //    shareStringPart.SharedStringTable.Save();
            //}


            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }


        private static WorksheetPart GetWorksheetPartByName(SpreadsheetDocument document, string sheetName)
        {
            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().
                            Elements<Sheet>().Where(s => s.Name == sheetName);
            if (sheets.Count() == 0)
            {
                return null;
            }
            string relationshipId = sheets.First().Id.Value;
            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);
            return worksheetPart;
        }


        public static void UpdateCell(string docName, string text, uint rowIndex, string columnName)
        {
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true))
            {
                WorksheetPart worksheetPart = GetWorksheetPartByName(spreadSheet, "Sheet1");
                if (worksheetPart != null)
                {
                    // Create new Worksheet
                    Worksheet worksheet = new Worksheet();
                    worksheetPart.Worksheet = worksheet;

                    // Create new SheetData
                    SheetData sheetData = new SheetData();

                    // Create new row
                    Row row = new Row() { RowIndex = rowIndex };

                    // Create new cell
                    Cell cell = new Cell() { CellReference = columnName + rowIndex, DataType = CellValues.Number, CellValue = new CellValue(text) };

                    // Append cell to row
                    row.Append(cell);

                    // Append row to sheetData
                    sheetData.Append(row);

                    // Append sheetData to worksheet
                    worksheet.Append(sheetData);

                    worksheetPart.Worksheet.Save();
                }
                spreadSheet.WorkbookPart.Workbook.Save();
            }

        }



        // Given a WorkbookPart, inserts a new worksheet.
        private static WorksheetPart AddWorksheetOnceOnly(WorkbookPart workbookPart)
        {
            // Add a new worksheet part to the workbook.
            WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());
            newWorksheetPart.Worksheet.Save();

            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();

            if (sheets.Elements<Sheet>().Count() == 0)
            {
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
            }

            return newWorksheetPart;
        }



        // Given a WorkbookPart, inserts a new worksheet.
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



        // Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
        // If the cell already exists, returns it. 
        private static Cell InsertCellInWorksheet1(string columnName, uint rowIndex, WorksheetPart worksheetPart)
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
                    if (cell.CellReference.Value.Length == cellReference.Length)
                    {
                        if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                        {
                            refCell = cell;
                            break;
                        }
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }




        // Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
        // If the cell already exists, returns it. 
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
                    if (cell.CellReference.Value.Length == cellReference.Length)
                    {
                        if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                        {
                            refCell = cell;
                            break;
                        }
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }



        static void RemoveSectionBreaks(string filename)
        {

            using (WordprocessingDocument myDoc = WordprocessingDocument.Open(filename, true))
            {

                MainDocumentPart mainPart = myDoc.MainDocumentPart;

                List<ParagraphProperties> paraProps = mainPart.Document.Descendants<ParagraphProperties>()

                .Where(pPr => IsSectionProps(pPr)).ToList();

                foreach (ParagraphProperties pPr in paraProps)
                {
                    pPr.RemoveChild<SectionProperties>(pPr.GetFirstChild<SectionProperties>());
                }

                mainPart.Document.Save();

            }

        }



        static bool IsSectionProps(ParagraphProperties pPr)
        {

            SectionProperties sectPr = pPr.GetFirstChild<SectionProperties>();

            if (sectPr == null)

                return false;

            else

                return true;

        }



        private static bool SetCellValue(DocumentFormat.OpenXml.Packaging.SpreadsheetDocument spreadsheet, DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet, uint columnIndex, uint rowIndex, DocumentFormat.OpenXml.Spreadsheet.CellValues valueType, string value, uint? styleIndex, bool save = true)
        {
            DocumentFormat.OpenXml.Spreadsheet.SheetData sheetData = worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>();
            DocumentFormat.OpenXml.Spreadsheet.Row row;
            DocumentFormat.OpenXml.Spreadsheet.Row previousRow = null;
            DocumentFormat.OpenXml.Spreadsheet.Cell cell;
            DocumentFormat.OpenXml.Spreadsheet.Cell previousCell = null;
            DocumentFormat.OpenXml.Spreadsheet.Columns columns;
            DocumentFormat.OpenXml.Spreadsheet.Column previousColumn = null;
            string cellAddress = Excel.ColumnNameFromIndex(columnIndex) + rowIndex;

            // Check if the row exists, create if necessary
            if (sheetData.Elements<DocumentFormat.OpenXml.Spreadsheet.Row>().Where(item => item.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<DocumentFormat.OpenXml.Spreadsheet.Row>().Where(item => item.RowIndex == rowIndex).First();
            }
            else
            {
                row = new DocumentFormat.OpenXml.Spreadsheet.Row() { RowIndex = rowIndex };
                //sheetData.Append(row);
                for (uint counter = rowIndex - 1; counter > 0; counter--)
                {
                    previousRow = sheetData.Elements<DocumentFormat.OpenXml.Spreadsheet.Row>().Where(item => item.RowIndex == counter).FirstOrDefault();
                    if (previousRow != null)
                    {
                        break;
                    }
                }
                sheetData.InsertAfter(row, previousRow);
            }

            // Check if the cell exists, create if necessary
            if (row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>().Where(item => item.CellReference.Value == cellAddress).Count() > 0)
            {
                cell = row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>().Where(item => item.CellReference.Value == cellAddress).First();
            }
            else
            {
                // Find the previous existing cell in the row
                for (uint counter = columnIndex - 1; counter > 0; counter--)
                {
                    previousCell = row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>().Where(item => item.CellReference.Value == Excel.ColumnNameFromIndex(counter) + rowIndex).FirstOrDefault();
                    if (previousCell != null)
                    {
                        break;
                    }
                }
                cell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = cellAddress };
                row.InsertAfter(cell, previousCell);
            }

            // Check if the column collection exists
            columns = worksheet.Elements<DocumentFormat.OpenXml.Spreadsheet.Columns>().FirstOrDefault();
            if (columns == null)
            {
                columns = worksheet.InsertAt(new DocumentFormat.OpenXml.Spreadsheet.Columns(), 0);
            }
            // Check if the column exists
            if (columns.Elements<DocumentFormat.OpenXml.Spreadsheet.Column>().Where(item => item.Min == columnIndex).Count() == 0)
            {
                // Find the previous existing column in the columns
                for (uint counter = columnIndex - 1; counter > 0; counter--)
                {
                    previousColumn = columns.Elements<DocumentFormat.OpenXml.Spreadsheet.Column>().Where(item => item.Min == counter).FirstOrDefault();
                    if (previousColumn != null)
                    {
                        break;
                    }
                }
                columns.InsertAfter(
                   new DocumentFormat.OpenXml.Spreadsheet.Column()
                   {
                       Min = columnIndex,
                       Max = columnIndex,
                       CustomWidth = true,
                       Width = 9
                   }, previousColumn);
            }

            // Add the value
            cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(value);
            if (styleIndex != null)
            {
                cell.StyleIndex = styleIndex;
            }
            if (valueType != DocumentFormat.OpenXml.Spreadsheet.CellValues.Date)
            {
                cell.DataType = new DocumentFormat.OpenXml.EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(valueType);
            }

            if (save)
            {
                worksheet.Save();
            }

            return true;
        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }



        public static void AddUpdateCellValue(SpreadsheetDocument spreadSheet, string sheetname, uint rowIndex, string columnName, string text)
        {
            // Opening document for editing            
            WorksheetPart worksheetPart = RetrieveSheetPartByName(spreadSheet, sheetname);

            if (worksheetPart != null)
            {
                Cell cell = InsertCellInSheet(columnName, (rowIndex + 1), worksheetPart);
                cell.CellValue = new CellValue(text);
                //cell datatype
                cell.DataType = new EnumValue<CellValues>(CellValues.String);
                // Save the worksheet.
                worksheetPart.Worksheet.Save();
            }
        }

        //retrieve sheetpart
        public static WorksheetPart RetrieveSheetPartByName(SpreadsheetDocument document, string sheetName)
        {
            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().
            Elements<Sheet>().Where(s => s.Name == sheetName);

            if (sheets.Count() == 0)
                return null;

            string relationshipId = sheets.First().Id.Value;
            WorksheetPart worksheetPart = (WorksheetPart)
            document.WorkbookPart.GetPartById(relationshipId);
            return worksheetPart;
        }

        //insert cell in sheet based on column and row index            
        public static Cell InsertCellInSheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;
            Row row;
            //check whether row exist or not            
            //if row exist            
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            //if row does not exist then it will create new row            
            else
            {
                row = new Row()
                {
                    RowIndex = rowIndex
                };
                sheetData.Append(row);
            }
            //check whether cell exist or not            
            //if cell exist            
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            //if cell does not exist            
            else
            {
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }
                Cell newCell = new Cell()
                {
                    CellReference = cellReference
                };
                row.InsertBefore(newCell, refCell);
                worksheet.Save();
                return newCell;
            }
        }

        // retrieve cell based on column and row index            
        public static Cell RetreiveCell(Worksheet worksheet,
         string columnName, uint rowIndex)
        {
            Row row = RetrieveRow(worksheet, rowIndex);
            var newRow = new Row()
            {
                RowIndex = (uint)rowIndex + 1
            };
            //adding new row            
            worksheet.InsertAt(newRow, Convert.ToInt32(rowIndex + 1));
            //create cell with value            
            Cell cell = new Cell();
            cell.CellValue = new CellValue("");
            cell.DataType =
             new EnumValue<CellValues>(CellValues.String);
            newRow.AddAnnotation(cell);
            worksheet.Save();

            row = newRow;
            if (row == null)
                return null;
            return row.Elements<Cell>().Where(c => string.Compare(c.CellReference.Value, columnName +
         (rowIndex + 1), true) == 0).First();
        }

        // it will return a row based on worksheet and rowindex            
        public static Row RetrieveRow(Worksheet worksheet, uint rowIndex)
        {
            return worksheet.GetFirstChild<SheetData>().
            Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
        }



        static void InsertRow(string sheetName, WorkbookPart wbPart, uint rowIndex)
        {
            Sheet sheet = wbPart.Workbook.Descendants<Sheet>().Where((s) => s.Name == sheetName).FirstOrDefault();

            if (sheet != null)
            {
                Worksheet ws = ((WorksheetPart)(wbPart.GetPartById(sheet.Id))).Worksheet;
                SheetData sheetData = ws.WorksheetPart.Worksheet.GetFirstChild<SheetData>();
                Row refRow = GetRow(sheetData, rowIndex);
                ++rowIndex;

                Cell cell1 = new Cell() { CellReference = "A" + rowIndex };
                CellValue cellValue1 = new CellValue();
                cellValue1.Text = "";
                cell1.Append(cellValue1);
                Row newRow = new Row()
                {
                    RowIndex = rowIndex
                };
                newRow.Append(cell1);
                for (int i = (int)rowIndex; i <= sheetData.Elements<Row>().Count(); i++)
                {
                    var row = sheetData.Elements<Row>().Where(r => r.RowIndex.Value == i).FirstOrDefault();
                    row.RowIndex++;
                    foreach (Cell c in row.Elements<Cell>())
                    {
                        string refer = c.CellReference.Value;
                        int num = Convert.ToInt32(Regex.Replace(refer, @"[^\d]*", ""));
                        num++;
                        string letters = Regex.Replace(refer, @"[^A-Z]*", "");
                        c.CellReference.Value = letters + num;
                    }
                }
                sheetData.InsertAfter(newRow, refRow);
                //ws.Save();
            }
        }

        static Row GetRow(SheetData wsData, UInt32 rowIndex)
        {
            var row = wsData.Elements<Row>().
            Where(r => r.RowIndex.Value == rowIndex).FirstOrDefault();
            if (row == null)
            {
                row = new Row();
                row.RowIndex = rowIndex;
                wsData.Append(row);
            }
            return row;
        }




    }
}
