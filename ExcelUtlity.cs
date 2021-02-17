using DataMigrationSQLtoXML;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace DataMigrationSQLtoXLSM
{
    public class ExcelUtlity
    {
        public async Task WriteDataToOpenExcel(string worksheetName, string saveAsLocation)
        {
            try
            {
                DataMigrationService dataMigrationService = new DataMigrationService();
                Utility utility = new Utility();

                List<Task> tasks = new List<Task>();

                using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(saveAsLocation, true))
                {
                    string Sheet2_SummaryTable = ConfigurationManager.AppSettings["Sheet2_SummaryTable"];
                    string Sheet3_ProductTable = ConfigurationManager.AppSettings["Sheet3_ProductTable"];
                    string Sheet10_ParticularQtyTable = ConfigurationManager.AppSettings["Sheet10_ParticularQtyTable"];
                    string Sheet7_ParticularValTable = ConfigurationManager.AppSettings["Sheet7_ParticularValTable"];

                    List<string> excelSheets = new List<string>();
                    excelSheets.Add(Sheet2_SummaryTable);
                    excelSheets.Add(Sheet3_ProductTable);
                    excelSheets.Add(Sheet10_ParticularQtyTable);
                    excelSheets.Add(Sheet7_ParticularValTable);

                    foreach (string sheet in excelSheets)
                    {
                        try
                        {
                            WorksheetPart worksheetPart = utility.RetrieveSheetPartByName(spreadSheet, sheet);
                            if (worksheetPart != null)
                            {

                                DataTable dataTable = new DataTable();
                                if (sheet == Sheet2_SummaryTable)
                                {
                                    dataTable = await dataMigrationService.GetPrincipalSummaryTableAsync();
                                }
                                else if (sheet == Sheet3_ProductTable)
                                {
                                    dataTable = await dataMigrationService.getPrincipalProductSummaryTableAsync();
                                }
                                else if (sheet == Sheet10_ParticularQtyTable)
                                {
                                    dataTable = await dataMigrationService.GetCustomerQtyTableAsync();
                                }
                                else if (sheet == Sheet7_ParticularValTable)
                                {
                                    dataTable = await dataMigrationService.getPrincipalParticularValAsync();
                                }

                                //int sheetIndex = 0;
                                // utility.AddUpdateCellValue(spreadSheet, "test sheet1", 8, "A", "test data1");

                                List<String> columns = new List<string>();
                                foreach (System.Data.DataColumn column in dataTable.Columns)
                                {
                                    columns.Add(column.ColumnName);
                                }
                                Worksheet worksheet = worksheetPart.Worksheet;
                                SheetData sheetData = worksheet.GetFirstChild<SheetData>();

                                foreach (DataRow dtrow in dataTable.Rows)
                                {
                                    try
                                    {
                                        Row newRow = new Row();
                                        foreach (String col in columns)
                                        {
                                            string objdtDataType = dtrow[col].GetType().ToString();
                                            Cell cell = new Cell();
                                            //cell.DataType = CellValues.String;
                                            //cell.CellValue = new CellValue(dtrow[col].ToString()); //

                                            //Add text to text cell
                                            if (objdtDataType.Contains(TypeCode.Int32.ToString()) || objdtDataType.Contains(TypeCode.Int64.ToString()) || objdtDataType.Contains(TypeCode.Decimal.ToString()))
                                            {
                                                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                                                cell.CellValue = new CellValue(dtrow[col].ToString());
                                            }
                                            else
                                            {
                                                cell.DataType = new EnumValue<CellValues>(CellValues.String);
                                                cell.CellValue = new CellValue(dtrow[col].ToString());
                                            }
                                            newRow.AppendChild(cell);
                                        }
                                        sheetData.AppendChild(newRow);
                                    }
                                    catch (Exception ex)
                                    {

                                    }
                                }

                                worksheetPart.Worksheet.Save();

                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(string.Format("Error in worksheetPart. Error Message: {0}, Inner Exception :{1}", ex.Message, ex.InnerException));
                            Log.error("Error in worksheetPart", ex);
                          
                        }
                    }

                    //WorkbookPart wbPart = spreadSheet.WorkbookPart;
                    //Sheets theSheets = wbPart.Workbook.Sheets;
                    //foreach (OpenXmlElement sheet in theSheets)
                    //{
                    //    foreach (OpenXmlAttribute attr in sheet.GetAttributes())
                    //    {
                    //        Console.WriteLine("{0}: {1}", attr.LocalName, attr.Value);
                    //    }
                    //}
                    //int sheetIndex = 0;
                    //foreach (WorksheetPart excelSheet in spreadSheet.WorkbookPart.WorksheetParts)
                    //{
                    //    string Sheet2 = ConfigurationManager.AppSettings["Sheet2"];
                    //    string Sheet3 = ConfigurationManager.AppSettings["Sheet3"];
                    //    string Sheet10 = ConfigurationManager.AppSettings["Sheet10"];
                    //    string Sheet7 = ConfigurationManager.AppSettings["Sheet7"];

                    //    Log.info("ExcelSheet Name :" + excelSheet.Worksheet.XName);
                    //    Console.WriteLine("ExcelSheet Name :" + excelSheet.Worksheet.XName);
                    //    //if (excelSheet.Name == Sheet2 || excelSheet.Name == Sheet3 || excelSheet.Name == Sheet10 || excelSheet.Name == Sheet7)
                    //    //{
                    //    //    tasks.Add(ProcessExcel(dataMigrationService, excelSheet, Sheet2, Sheet3, Sheet10, Sheet7, xlWorkBook));
                    //    //}
                    //    sheetIndex++;
                    //}
                    //cell.CellValue = new CellValue(text);
                    //cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                    // worksheetPart.Worksheet.Save();
                }

                await Task.WhenAll(tasks);

            }
            catch (Exception ex)
            {
            }
            finally
            {

            }
        }



        public async Task WriteDataToExcel(string worksheetName, string saveAsLocation)
        {
            Microsoft.Office.Interop.Excel.Application xlexcel;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            Microsoft.Office.Interop.Excel.Range xlRange;

            try
            {
                DataMigrationService dataMigrationService = new DataMigrationService();
                object misValue = System.Reflection.Missing.Value;

                xlexcel = new Microsoft.Office.Interop.Excel.Application();
                xlexcel.Visible = true;

                // Open a File
                xlWorkBook = xlexcel.Workbooks.Open(saveAsLocation);

                //  foreach (Microsoft.Office.Interop.Excel.Worksheet excelSheet in xlWorkBook.Worksheets)
                //Parallel.ForEach(xlWorkBook.Worksheets.Cast<Microsoft.Office.Interop.Excel.Worksheet>(),
                //    new ParallelOptions { MaxDegreeOfParallelism = 3 }, excelSheet=>{
                //        string Sheet2 = ConfigurationManager.AppSettings["Sheet2"];
                //        string Sheet3 = ConfigurationManager.AppSettings["Sheet3"];
                //        string Sheet10 = ConfigurationManager.AppSettings["Sheet10"];
                //        string Sheet7 = ConfigurationManager.AppSettings["Sheet7"];

                //        Log.info("ExcelSheet Name :" + excelSheet.Name);
                //        if (excelSheet.Name == Sheet2)
                //        {
                //            DataTable sheet2Table =  dataMigrationService.GetPrincipalSummaryTable();
                //            int rowcount = 1;
                //            foreach (DataRow datarow in sheet2Table.Rows)
                //            {
                //                rowcount += 1;
                //                for (int i = 1; i <= sheet2Table.Columns.Count; i++)
                //                {
                //                    excelSheet.Cells[rowcount, i] = datarow[i - 1].ToString();
                //                }
                //            }
                //            Log.info(" Load end ExcelSheet :" + excelSheet.Name);
                //        }
                //        else if (excelSheet.Name == Sheet3)
                //        {
                //            DataTable sheet3Table = dataMigrationService.getPrincipalProductSummaryTable();

                //            int rowcount = 1;
                //            foreach (DataRow datarow in sheet3Table.Rows)
                //            {
                //                rowcount += 1;
                //                for (int i = 1; i <= sheet3Table.Columns.Count; i++)
                //                {
                //                    excelSheet.Cells[rowcount, i] = datarow[i - 1].ToString();
                //                }
                //            }
                //            Log.info(" Load end ExcelSheet :" + excelSheet.Name);
                //        }
                //        else if (excelSheet.Name == Sheet10)
                //        {
                //            DataTable sheet10Table = dataMigrationService.GetCustomerQtyTable();
                //            int rowcount = 1;
                //            foreach (DataRow datarow in sheet10Table.Rows)
                //            {
                //                rowcount += 1;
                //                for (int i = 1; i <= sheet10Table.Columns.Count; i++)
                //                {
                //                    excelSheet.Cells[rowcount, i] = datarow[i - 1].ToString();
                //                }
                //            }
                //            Log.info(" Load end ExcelSheet :" + excelSheet.Name);
                //        }
                //        else if (excelSheet.Name == Sheet7)
                //        {
                //            DataTable sheet7Table = dataMigrationService.GetCustomerQtyTable();
                //            int rowcount = 1;
                //            foreach (DataRow datarow in sheet7Table.Rows)
                //            {
                //                rowcount += 1;
                //                for (int i = 1; i <= sheet7Table.Columns.Count; i++)
                //                {
                //                    excelSheet.Cells[rowcount, i] = datarow[i - 1].ToString();
                //                }
                //            }
                //            Log.info(" Load end ExcelSheet :" + excelSheet.Name);
                //        }
                //    });
                List<Task> tasks = new List<Task>();

                foreach (Microsoft.Office.Interop.Excel.Worksheet excelSheet in xlWorkBook.Worksheets)
                {
                    string Sheet2 = ConfigurationManager.AppSettings["Sheet2"];
                    string Sheet3 = ConfigurationManager.AppSettings["Sheet3"];
                    string Sheet10 = ConfigurationManager.AppSettings["Sheet10"];
                    string Sheet7 = ConfigurationManager.AppSettings["Sheet7"];
                    Log.info("ExcelSheet Name :" + excelSheet.Name);
                    Console.WriteLine("ExcelSheet Name :" + excelSheet.Name);
                    if (excelSheet.Name == Sheet2 || excelSheet.Name == Sheet3 || excelSheet.Name == Sheet10 || excelSheet.Name == Sheet7)
                    {
                        tasks.Add(ProcessExcel(dataMigrationService, excelSheet, Sheet2, Sheet3, Sheet10, Sheet7, xlWorkBook));
                    }

                }

                await Task.WhenAll(tasks);
                xlWorkBook.Save();
                xlWorkBook.Close(false, misValue, misValue);
                xlexcel.Quit();

                //releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlexcel);
            }
            catch (Exception ex)
            {
            }
            finally
            {
                xlexcel = null;
                xlRange = null;
                xlWorkBook = null;
            }

        }

        private static async Task ProcessExcel(DataMigrationService dataMigrationService, Microsoft.Office.Interop.Excel.Worksheet excelSheet, string Sheet2, string Sheet3, string Sheet10, string Sheet7, Microsoft.Office.Interop.Excel.Workbook xlWorkBook)
        {
            try
            {
                if (excelSheet.Name == Sheet2)
                {
                    DataTable sheet2Table = await dataMigrationService.GetPrincipalSummaryTableAsync();
                    SaveToExcel(excelSheet, xlWorkBook, sheet2Table);
                    Log.info("Load end ExcelSheet :" + excelSheet.Name);
                    Console.WriteLine("Load end ExcelSheet :" + excelSheet.Name);
                }
                else if (excelSheet.Name == Sheet3)
                {
                    DataTable sheet3Table = await dataMigrationService.getPrincipalProductSummaryTableAsync();
                    SaveToExcel(excelSheet, xlWorkBook, sheet3Table);
                    Log.info("Load end ExcelSheet :" + excelSheet.Name);
                    Console.WriteLine("Load end ExcelSheet :" + excelSheet.Name);
                }
                else if (excelSheet.Name == Sheet10)
                {
                    DataTable sheet10Table = await dataMigrationService.GetCustomerQtyTableAsync();
                    SaveToExcel(excelSheet, xlWorkBook, sheet10Table);
                    Log.info("Load end ExcelSheet :" + excelSheet.Name);
                    Console.WriteLine("Load end ExcelSheet :" + excelSheet.Name);
                }
                else if (excelSheet.Name == Sheet7)
                {
                    DataTable sheet7Table = await dataMigrationService.GetCustomerQtyTableAsync();
                    SaveToExcel(excelSheet, xlWorkBook, sheet7Table);
                    Log.info("Load end ExcelSheet :" + excelSheet.Name);
                    Console.WriteLine("Load end ExcelSheet :" + excelSheet.Name);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format("Error in ProcessExcel. Error Message: {0}, Inner Exception :{1}", ex.Message, ex.InnerException));
                Log.error("Error in ProcessExcel", ex);
                // throw;
            }
        }

        private static void SaveToExcel(Microsoft.Office.Interop.Excel.Worksheet excelSheet, Microsoft.Office.Interop.Excel.Workbook xlWorkBook, DataTable sheet10Table)
        {
            int rowcount = 1;
            foreach (DataRow datarow in sheet10Table.Rows)
            {
                rowcount += 1;
                for (int i = 1; i <= sheet10Table.Columns.Count; i++)
                {
                    try
                    {

                        excelSheet.Cells[rowcount, i] = datarow[i - 1];
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(string.Format("Error in SaveToExcel. Error Message: {0}, Inner Exception :{1}, rowcount:{2}, Cell: {3} ", ex.Message, ex.InnerException, rowcount, i));
                        Log.error(String.Format("Error in SaveToExcel. rowcount:{0}, Cell: {1} ", rowcount, i), ex);

                    }
                }
            }
            xlWorkBook.Save();
        }

        public bool WriteDataTableToExcel(DataTable dataTable, string worksheetName, string saveAsLocation)
        {
            Microsoft.Office.Interop.Excel.Application xlexcel;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            Microsoft.Office.Interop.Excel.Range xlRange;

            try
            {
                object misValue = System.Reflection.Missing.Value;

                xlexcel = new Microsoft.Office.Interop.Excel.Application();
                xlexcel.Visible = true;

                // Open a File
                xlWorkBook = xlexcel.Workbooks.Open(saveAsLocation);

                foreach (Microsoft.Office.Interop.Excel.Worksheet excelSheet in xlWorkBook.Worksheets)
                {
                    if (excelSheet.Name == worksheetName)
                    {
                        // loop through each row and add values to our sheet
                        int rowcount = 1;
                        foreach (DataRow datarow in dataTable.Rows)
                        {
                            rowcount += 1;
                            //Microsoft.Office.Interop.Excel.Range row = excelSheet.Rows[rowcount];
                            //row.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);

                            for (int i = 1; i <= dataTable.Columns.Count; i++)
                            {

                                excelSheet.Cells[rowcount, i] = datarow[i - 1].ToString();
                            }
                        }
                        //foreach (Microsoft.Office.Interop.Excel.Range row in sheet.UsedRange.Rows)
                        //{

                        //    row.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);
                        //}
                    }
                }

                //Once done close and quit Excel
                xlWorkBook.Save();
                xlWorkBook.Close(false, misValue, misValue);
                xlexcel.Quit();


                //releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlexcel);
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
            finally
            {
                xlexcel = null;
                xlRange = null;
                xlWorkBook = null;
            }

        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }

        public void FormattingExcelCells(Microsoft.Office.Interop.Excel.Range range, string HTMLcolorCode, System.Drawing.Color fontColor, bool IsFontbool)
        {
            range.Interior.Color = System.Drawing.ColorTranslator.FromHtml(HTMLcolorCode);
            range.Font.Color = System.Drawing.ColorTranslator.ToOle(fontColor);
            if (IsFontbool == true)
            {
                range.Font.Bold = IsFontbool;
            }
        }

        //public bool WriteDataTableToExcel(DataTable dataTable, string worksheetName, string saveAsLocation)
        //{
        //    Microsoft.Office.Interop.Excel.Application excel;
        //    Microsoft.Office.Interop.Excel.Workbook excelworkBook;
        //    Microsoft.Office.Interop.Excel.Worksheet excelSheet;
        //    Microsoft.Office.Interop.Excel.Range excelCellrange;

        //    try
        //    {
        //        // Start Excel and get Application object.
        //        excel = new Microsoft.Office.Interop.Excel.Application();

        //        // for making Excel visible
        //        excel.Visible = false;
        //        excel.DisplayAlerts = false;

        //        // Creation a new Workbook
        //        excelworkBook = excel.Workbooks.Add(Type.Missing);


        //        // Workk sheet
        //        excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.ActiveSheet;
        //        excelSheet.Name = worksheetName;

        //        // loop through each row and add values to our sheet
        //        int rowcount = 1;

        //        foreach (DataRow datarow in dataTable.Rows)
        //        {
        //            rowcount += 1;
        //            for (int i = 1; i <= dataTable.Columns.Count; i++)
        //            {
        //                //// on the first iteration we add the column headers
        //                //if (rowcount == 3)
        //                //{
        //                //    excelSheet.Cells[2, i] = dataTable.Columns[i - 1].ColumnName;
        //                //    excelSheet.Cells.Font.Color = System.Drawing.Color.Black;

        //                //}

        //                excelSheet.Cells[rowcount, i] = datarow[i - 1].ToString();

        //                ////for alternate rows
        //                //if (rowcount > 2)
        //                //{
        //                //    if (i == dataTable.Columns.Count)
        //                //    {
        //                //        if (rowcount % 2 == 0)
        //                //        {
        //                //            excelCellrange = excelSheet.Range[excelSheet.Cells[rowcount, 1], excelSheet.Cells[rowcount, dataTable.Columns.Count]];
        //                //            FormattingExcelCells(excelCellrange, "#CCCCFF", System.Drawing.Color.Black, false);
        //                //        }

        //                //    }
        //                //}

        //            }

        //        }

        //        // now we resize the columns
        //        //excelCellrange = excelSheet.Range[excelSheet.Cells[1, 1], excelSheet.Cells[rowcount, dataTable.Columns.Count]];
        //        //excelCellrange.EntireColumn.AutoFit();
        //        //Microsoft.Office.Interop.Excel.Borders border = excelCellrange.Borders;
        //        //border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
        //        //border.Weight = 2d;


        //        //excelCellrange = excelSheet.Range[excelSheet.Cells[1, 1], excelSheet.Cells[2, dataTable.Columns.Count]];
        //        //FormattingExcelCells(excelCellrange, "#000099", System.Drawing.Color.White, true);


        //        //now save the workbook and exit Excel


        //        excelworkBook.SaveAs(saveAsLocation); ;
        //        excelworkBook.Close();
        //        excel.Quit();
        //        return true;
        //    }
        //    catch (Exception ex)
        //    {
        //        return false;
        //    }
        //    finally
        //    {
        //        excelSheet = null;
        //        excelCellrange = null;
        //        excelworkBook = null;
        //    }

        //}
    }
}
