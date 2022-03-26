using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;
using System.Linq;

namespace DocxToXLSX
{
    public partial class Form3 : Form
    {

        public Form3()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            getParagraphs(Application.StartupPath + @"\BILL#1-2-ch.docx");
            MessageBox.Show("Excel file created successfully", "File Created", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }


        static void getParagraphs(string filename)
        {

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
                                if (paragraph.InnerText.IndexOf("Normal Range") != -1 || paragraph.InnerText.IndexOf("Result") != -1)
                                {

                                }
                                else
                                {
                                    content += paragraph.InnerText.Trim() + "*";
                                }

                                //AddUpdateCellValue(SpreadsheetDocument spreadSheet, "Sheet1", 1, "A", content);

                                //WriteText(Application.StartupPath + @"\BILL#1-2-ch.xlsx", content, 1, "A");
                            }
                        }
                    }
                }




                var docPart = myDoc.MainDocumentPart;
                var doc = docPart.Document;
                DocumentFormat.OpenXml.Wordprocessing.Table myTable = doc.Body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Table>().FirstOrDefault();

                int heading_no = 0;

                foreach (TableRow row in myTable.Elements<TableRow>())
                {
                    Paragraph p = row.Descendants<Paragraph>().ElementAt(1);

                    heading_no++;

                    content += "tblhead_" + heading_no + "_" + p.InnerText + "*";

                    Paragraph p1 = row.Descendants<Paragraph>().ElementAt(2);

                    heading_no++;

                    content += "tblhead_" + heading_no + "_" + p1.InnerText + "*";

                }


                CreateExcelFile(Application.StartupPath + @"\BILL#1-2-ch.xlsx", content);


            }
        }


        private static int ProcessList(List<string> tempRows, List<List<string>> totalRows, int MaxCol)
        {
            if (tempRows.Count > MaxCol)
            {
                MaxCol = tempRows.Count;
            }

            totalRows.Add(tempRows);
            return MaxCol;
        }


        static void CreateExcelFile(string filePath, string content)
        {
            using (SpreadsheetDocument spreedDoc = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart wbPart = spreedDoc.WorkbookPart;
                if (wbPart == null)
                {
                    wbPart = spreedDoc.AddWorkbookPart();
                    wbPart.Workbook = new Workbook();
                }

                string sheetName = "Sheet1";
                WorksheetPart worksheetPart = null;
                worksheetPart = wbPart.AddNewPart<WorksheetPart>();
                var sheetData = new SheetData();

                worksheetPart.Worksheet = new Worksheet(sheetData);

                if (wbPart.Workbook.Sheets == null)
                {
                    wbPart.Workbook.AppendChild<Sheets>(new Sheets());
                }

                var sheet = new Sheet()
                {
                    Id = wbPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = sheetName
                };

                var workingSheet = ((WorksheetPart)wbPart.GetPartById(sheet.Id)).Worksheet;


                var arr = content.Split('*');

                int rowindex = 1;





                for (int a = 0; a <= arr.Length - 1; a++)
                {

                    Row row = new Row();
                    row.RowIndex = (UInt32)rowindex;

                    if (rowindex == 1) //Header
                    {
                        //row.AppendChild(AddCellWithText1("Name"));
                        //row.AppendChild(AddCellWithText1("Email"));
                    }
                    else //Data
                    {

                        List<Package> packages = new List<Package>
                        {
                            new Package { Company = arr[0] },
                            new Package { Company = arr[1] },
                            new Package { Company = arr[2] },
                            new Package { Company = arr[3] },
                            new Package { Company = arr[4] },
                            new Package { Company = arr[5] },
                            new Package { Company = arr[6] },
                            new Package { Company = arr[7] },
                            new Package { Company = arr[8] },
                            new Package { Company = arr[9] },
                            new Package { Company = arr[10] },
                            new Package { Company = arr[11] }
                        };



                        List<string> headerNames = new List<string> { "Company" };

                        ExcelFacade excelFacade = new ExcelFacade();
                        excelFacade.Create<Package>(Application.StartupPath + @"\output1.xlsx", packages, "Packages", headerNames);


                        //javed      row.AppendChild(AddCellWithText1(arr[a]));
                    }

                    sheetData.AppendChild(row);
                    rowindex++;

                }




                //int rowindex = 1;
                //foreach (var emp in lstEmps)
                //{
                //    Row row = new Row();
                //    row.RowIndex = (UInt32)rowindex;

                //    if (rowindex == 1) //Header
                //    {
                //        row.AppendChild(AddCellWithText1("Name"));
                //        row.AppendChild(AddCellWithText1("Email"));
                //    }
                //    else //Data
                //    {
                //        row.AppendChild(AddCellWithText1("safsdfsd"));
                //        row.AppendChild(AddCellWithText1("sdfsdfsd"));
                //    }

                //    sheetData.AppendChild(row);
                //    rowindex++;
                //}

                wbPart.Workbook.Sheets.AppendChild(sheet);

                //Set Border
                //wbPark

                wbPart.Workbook.Save();
            }
        }



        //public static Table DefineTable(WorksheetPart worksheetPart, int rowMin, int rowMax, int colMin, int colMax)
        //{
        //    TableDefinitionPart tableDefinitionPart = worksheetPart.AddNewPart<TableDefinitionPart>("rId" + (worksheetPart.TableDefinitionParts.Count() + 1));
        //    //int tableNo = worksheetPart.TableDefinitionParts.Count();
        //    int tableNo = 1;


        //    string reference = ((char)(64 + colMin)).ToString() + rowMin + ":" + ((char)(64 + colMax)).ToString() + rowMax;

        //    Table table = new Table() { Id = (UInt32)tableNo, Name = "Table" + tableNo, DisplayName = "Table" + tableNo, Reference = reference, TotalsRowShown = false };
        //    AutoFilter autoFilter = new AutoFilter() { Reference = reference };

        //    TableColumns tableColumns = new TableColumns() { Count = (UInt32)(colMax - colMin + 1) };
        //    for (int i = 0; i < (colMax - colMin + 1); i++)
        //    {
        //        tableColumns.Append(new TableColumn() { Id = (UInt32)(colMin + i), Name = "Column" + i }); //changed i+1 -> colMin + i
        //                                                                                                   //Add cell values (shared string)
        //    }

        //    TableStyleInfo tableStyleInfo = new TableStyleInfo() { Name = "TableStyleLight1", ShowFirstColumn = false, ShowLastColumn = false, ShowRowStripes = true, ShowColumnStripes = false };

        //    table.Append(autoFilter);
        //    table.Append(tableColumns);
        //    table.Append(tableStyleInfo);

        //    tableDefinitionPart.Table = table;

        //    TableParts tableParts = (TableParts)worksheetPart.Worksheet.ChildElements.Where(ce => ce is TableParts).FirstOrDefault(); // Add table parts only once
        //    if (tableParts is null)
        //    {
        //        tableParts = new TableParts();
        //        tableParts.Count = (UInt32)0;
        //        worksheetPart.Worksheet.Append(tableParts);
        //    }

        //    tableParts.Count += (UInt32)1;
        //    TablePart tablePart = new TablePart() { Id = "rId" + tableNo };

        //    tableParts.Append(tablePart);

        //    return table;
        //}



        static Cell AddCellWithText1(string text)
        {
            CustomStylesheet obj = new CustomStylesheet();


            Cell c1 = new Cell();
            c1.DataType = CellValues.InlineString;

            InlineString inlineString = new InlineString();
            DocumentFormat.OpenXml.Spreadsheet.Text t = new DocumentFormat.OpenXml.Spreadsheet.Text(text);
            t.Text = text;
            inlineString.AppendChild(t);
            c1.AppendChild(inlineString);




            return c1;
        }


        static Cell AddCellWithText(string text)
        {
            Cell c1 = new Cell();
            c1.DataType = CellValues.InlineString;

            InlineString inlineString = new InlineString();
            DocumentFormat.OpenXml.Spreadsheet.Text t = new DocumentFormat.OpenXml.Spreadsheet.Text("wqrweqrqwe");
            t.Text = text;
            inlineString.AppendChild(t);

            c1.AppendChild(inlineString);

            return c1;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            List<Package1> packages =
                    new List<Package1>
                        { new Package1 { Company = "Coho Vineyard", Weight = 25.2, TrackingNumber = 89453312L, DateOrder = DateTime.Today, HasCompleted = false },
                          new Package1 { Company = "Lucerne Publishing", Weight = 18.7, TrackingNumber = 89112755L, DateOrder = DateTime.Today, HasCompleted = false },
                          new Package1 { Company = "Wingtip Toys", Weight = 6.0, TrackingNumber = 299456122L, DateOrder = DateTime.Today, HasCompleted = false },
                          new Package1 { Company = "Adventure Works", Weight = 33.8, TrackingNumber = 4665518773L, DateOrder =  DateTime.Today.AddDays(-4), HasCompleted = true },
                          new Package1 { Company = "Test Works", Weight = 35.8, TrackingNumber = 4665518774L, DateOrder =  DateTime.Today.AddDays(-2), HasCompleted = true },
                          new Package1 { Company = "Good Works", Weight = 48.8, TrackingNumber = 4665518775L, DateOrder =  DateTime.Today.AddDays(-1), HasCompleted = true },

                        };

            List<string> headerNames = new List<string> { "Company", "Weight", "Tracking Number", "Date Order", "Completed" };

            ExcelFacade excelFacade = new ExcelFacade();
            excelFacade.Create<Package1>(Application.StartupPath + @"\output1.xlsx", packages, "Packages", headerNames);

            MessageBox.Show("Excel file created successfully", "File Created", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }




    public class Package
    {
        public string Company { get; set; }
    }


    public class ResultTable
    {
        public string NormalRange { get; set; }
        public string NormalRangeValue { get; set; }
        public string Result { get; set; }
        public string ResultValue { get; set; }
    }


    public class Package1
    {
        public string Company { get; set; }
        public double Weight { get; set; }
        public long TrackingNumber { get; set; }
        public DateTime DateOrder { get; set; }
        public bool HasCompleted { get; set; }
    }


}
