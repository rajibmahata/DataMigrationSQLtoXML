using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataMigrationSQLtoXML
{
    public class DataMigrationService
    {

        public async Task ProcessDataMigrationSQLtoXML()
        {
            try
            {

                await processPrincipalSummaryTable();

            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format("Error in DataMigrationService. Error Message: {0}, Inner Exception :{1}", ex.Message, ex.InnerException.Message));
                Log.error("Error in DataMigrationService", ex);
                throw ex;
            }
        }

        private async Task processPrincipalSummaryTable()
        {
            try
            {
                var PrincipalSummaryTable = new DataTable();
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["PPTLConnectionString"].ConnectionString))
                {
                    using (SqlDataAdapter SqlDataAdapter = new SqlDataAdapter("SELECT * FROM [dbo].[Z_Sales_Written_Manager_SummaryTable]", sqlConnection))
                    {
                        SqlDataAdapter.Fill(PrincipalSummaryTable);
                        ExcelUtlity excelUtlity = new ExcelUtlity();
                        string Sheet2 = ConfigurationManager.AppSettings["Sheet2"];
                        string ExcelPath = ConfigurationManager.AppSettings["ExcelPath"];
                        excelUtlity.WriteDataTableToExcel1(PrincipalSummaryTable, Sheet2, ExcelPath);

                        ////Instantiate a Workbook object
                        //Microsoft.Office.Interop.Excel.Workbook workbook = new Microsoft.Office.Interop.Excel.Workbook();
                        ////Load the Excel file
                        //workbook.LoadFromFile("Input.xlsx");

                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format("Error in DataMigrationService. Error Message: {0}, Inner Exception :{1}", ex.Message, ex.InnerException.Message));
                Log.error("Error in DataMigrationService", ex);
                throw ex;
            }
        }
    }
}
