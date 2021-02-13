using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace DataMigrationSQLtoXML
{
    public class ExcelUtlity
    {
        /// <summary>
        /// FUNCTION FOR EXPORT TO EXCEL
        /// </summary>
        /// <param name="dataTable"></param>
        /// <param name="worksheetName"></param>
        /// <param name="saveAsLocation"></param>
        /// <returns></returns>
        public bool WriteDataTableToExcel(DataTable dataTable, string worksheetName, string saveAsLocation)
        {
            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook excelworkBook;
            Microsoft.Office.Interop.Excel.Worksheet excelSheet;
            Microsoft.Office.Interop.Excel.Range excelCellrange;

            try
            {
                // Start Excel and get Application object.
                excel = new Microsoft.Office.Interop.Excel.Application();

                // for making Excel visible
                excel.Visible = false;
                excel.DisplayAlerts = false;

                // Creation a new Workbook
                excelworkBook = excel.Workbooks.Add(Type.Missing);


                // Workk sheet
                excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.ActiveSheet;
                excelSheet.Name = worksheetName;

                // loop through each row and add values to our sheet
                int rowcount = 1;

                foreach (DataRow datarow in dataTable.Rows)
                {
                    rowcount += 1;
                    for (int i = 1; i <= dataTable.Columns.Count; i++)
                    {
                        //// on the first iteration we add the column headers
                        //if (rowcount == 3)
                        //{
                        //    excelSheet.Cells[2, i] = dataTable.Columns[i - 1].ColumnName;
                        //    excelSheet.Cells.Font.Color = System.Drawing.Color.Black;

                        //}

                        excelSheet.Cells[rowcount, i] = datarow[i - 1].ToString();

                        ////for alternate rows
                        //if (rowcount > 2)
                        //{
                        //    if (i == dataTable.Columns.Count)
                        //    {
                        //        if (rowcount % 2 == 0)
                        //        {
                        //            excelCellrange = excelSheet.Range[excelSheet.Cells[rowcount, 1], excelSheet.Cells[rowcount, dataTable.Columns.Count]];
                        //            FormattingExcelCells(excelCellrange, "#CCCCFF", System.Drawing.Color.Black, false);
                        //        }

                        //    }
                        //}

                    }

                }

                // now we resize the columns
                //excelCellrange = excelSheet.Range[excelSheet.Cells[1, 1], excelSheet.Cells[rowcount, dataTable.Columns.Count]];
                //excelCellrange.EntireColumn.AutoFit();
                //Microsoft.Office.Interop.Excel.Borders border = excelCellrange.Borders;
                //border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                //border.Weight = 2d;


                //excelCellrange = excelSheet.Range[excelSheet.Cells[1, 1], excelSheet.Cells[2, dataTable.Columns.Count]];
                //FormattingExcelCells(excelCellrange, "#000099", System.Drawing.Color.White, true);


                //now save the workbook and exit Excel


                excelworkBook.SaveAs(saveAsLocation); ;
                excelworkBook.Close();
                excel.Quit();
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
            finally
            {
                excelSheet = null;
                excelCellrange = null;
                excelworkBook = null;
            }

        }


        public bool WriteDataTableToExcel1(DataTable dataTable, string worksheetName, string saveAsLocation)
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
        /// <summary>
        /// FUNCTION FOR FORMATTING EXCEL CELLS
        /// </summary>
        /// <param name="range"></param>
        /// <param name="HTMLcolorCode"></param>
        /// <param name="fontColor"></param>
        /// <param name="IsFontbool"></param>
        public void FormattingExcelCells(Microsoft.Office.Interop.Excel.Range range, string HTMLcolorCode, System.Drawing.Color fontColor, bool IsFontbool)
        {
            range.Interior.Color = System.Drawing.ColorTranslator.FromHtml(HTMLcolorCode);
            range.Font.Color = System.Drawing.ColorTranslator.ToOle(fontColor);
            if (IsFontbool == true)
            {
                range.Font.Bold = IsFontbool;
            }
        }

    }
}
