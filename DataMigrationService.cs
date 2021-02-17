using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace DataMigrationSQLtoXLSM
{
    public class DataMigrationService
    {

        public async Task ProcessDataMigrationSQLtoXML()
        {
            try
            {
                //Task taskSummary = processPrincipalSummaryTable();
                //Task taskProductSummary = processPrincipalProductSummaryTable();
                //Task taskQtyTable = processCustomerQtyTable();
                //Task taskParticularVal = processPrincipalParticularVal();
                //await Task.WhenAll(taskSummary, taskProductSummary, taskQtyTable, taskParticularVal);

                //await processPrincipalSummaryTable();
                //await processPrincipalProductSummaryTable();
                //await processCustomerQtyTable();
                //await processPrincipalParticularVal();

                ExcelUtlity excelUtlity = new ExcelUtlity();
                string Sheet2 = ConfigurationManager.AppSettings["Sheet2"];
                string ExcelPath = ConfigurationManager.AppSettings["ExcelPath"];
                string UpdatedExcelPath = ConfigurationManager.AppSettings["UpdatedExcelPath"];
                if(File.Exists(UpdatedExcelPath))
                {
                    File.Delete(UpdatedExcelPath);
                }
                File.Copy(ExcelPath, UpdatedExcelPath, true);
                await excelUtlity.WriteDataToOpenExcel(Sheet2, UpdatedExcelPath);

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
                    string SummaryTable = ConfigurationManager.AppSettings["SummaryTable"];

                    using (SqlDataAdapter SqlDataAdapter = new SqlDataAdapter("SELECT * FROM " + SummaryTable, sqlConnection))
                    {
                        SqlDataAdapter.Fill(PrincipalSummaryTable);
                        ExcelUtlity excelUtlity = new ExcelUtlity();
                        string Sheet2 = ConfigurationManager.AppSettings["Sheet2"];
                        string ExcelPath = ConfigurationManager.AppSettings["ExcelPath"];
                        excelUtlity.WriteDataTableToExcel(PrincipalSummaryTable, Sheet2, ExcelPath);
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

        private async Task processPrincipalProductSummaryTable()
        {
            try
            {
                var PrincipalSummaryTable = new DataTable();
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["PPTLConnectionString"].ConnectionString))
                {
                    string ProductTable = ConfigurationManager.AppSettings["ProductTable"];
                    using (SqlDataAdapter SqlDataAdapter = new SqlDataAdapter("SELECT * FROM " + ProductTable, sqlConnection))
                    {
                        SqlDataAdapter.Fill(PrincipalSummaryTable);
                        ExcelUtlity excelUtlity = new ExcelUtlity();
                        string Sheet2 = ConfigurationManager.AppSettings["Sheet3"];
                        string ExcelPath = ConfigurationManager.AppSettings["ExcelPath"];
                        excelUtlity.WriteDataTableToExcel(PrincipalSummaryTable, Sheet2, ExcelPath);
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

        private async Task processCustomerQtyTable()
        {
            try
            {
                var PrincipalSummaryTable = new DataTable();
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["PPTLConnectionString"].ConnectionString))
                {
                    string ParticularQtyTable = ConfigurationManager.AppSettings["ParticularQtyTable"];
                    using (SqlDataAdapter SqlDataAdapter = new SqlDataAdapter("SELECT * FROM " + ParticularQtyTable, sqlConnection))
                    {
                        SqlDataAdapter.Fill(PrincipalSummaryTable);
                        ExcelUtlity excelUtlity = new ExcelUtlity();
                        string Sheet2 = ConfigurationManager.AppSettings["Sheet10"];
                        string ExcelPath = ConfigurationManager.AppSettings["ExcelPath"];
                        excelUtlity.WriteDataTableToExcel(PrincipalSummaryTable, Sheet2, ExcelPath);
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

        private async Task processPrincipalParticularVal()
        {
            try
            {
                var PrincipalSummaryTable = new DataTable();
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["PPTLConnectionString"].ConnectionString))
                {
                    string ParticularValTable = ConfigurationManager.AppSettings["ParticularValTable"];

                    using (SqlDataAdapter SqlDataAdapter = new SqlDataAdapter("SELECT * FROM " + ParticularValTable, sqlConnection))
                    {
                        SqlDataAdapter.Fill(PrincipalSummaryTable);
                        ExcelUtlity excelUtlity = new ExcelUtlity();
                        string Sheet2 = ConfigurationManager.AppSettings["Sheet7"];
                        string ExcelPath = ConfigurationManager.AppSettings["ExcelPath"];
                        excelUtlity.WriteDataTableToExcel(PrincipalSummaryTable, Sheet2, ExcelPath);
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

        public async Task<DataTable> GetPrincipalSummaryTableAsync()
        {
            try
            {
                var PrincipalSummaryTable = new DataTable();
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["PPTLConnectionString"].ConnectionString))
                {
                    string SummaryTable = ConfigurationManager.AppSettings["SummaryTable"];

                    using (SqlDataAdapter SqlDataAdapter = new SqlDataAdapter("SELECT * FROM " + SummaryTable, sqlConnection))
                    {
                        SqlDataAdapter.Fill(PrincipalSummaryTable);

                    }
                }

                return PrincipalSummaryTable;
            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format("Error in DataMigrationService. Error Message: {0}, Inner Exception :{1}", ex.Message, ex.InnerException.Message));
                Log.error("Error in DataMigrationService", ex);
                throw ex;
            }
        }

        public async Task<DataTable> getPrincipalProductSummaryTableAsync()
        {
            try
            {
                var PrincipalProductSummaryTable = new DataTable();
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["PPTLConnectionString"].ConnectionString))
                {
                    string ProductTable = ConfigurationManager.AppSettings["ProductTable"];

                    using (SqlDataAdapter SqlDataAdapter = new SqlDataAdapter("SELECT * FROM " + ProductTable, sqlConnection))
                    {
                        SqlDataAdapter.Fill(PrincipalProductSummaryTable);
                    }
                }
                return PrincipalProductSummaryTable;
            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format("Error in DataMigrationService. Error Message: {0}, Inner Exception :{1}", ex.Message, ex.InnerException.Message));
                Log.error("Error in DataMigrationService", ex);
                throw ex;
            }
        }

        public async Task<DataTable> GetCustomerQtyTableAsync()
        {
            try
            {
                var CustomerQtyTable = new DataTable();
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["PPTLConnectionString"].ConnectionString))
                {
                    string ParticularQtyTable = ConfigurationManager.AppSettings["ParticularQtyTable"];
                    using (SqlDataAdapter SqlDataAdapter = new SqlDataAdapter("SELECT * FROM " + ParticularQtyTable, sqlConnection))
                    {
                        SqlDataAdapter.Fill(CustomerQtyTable);
                    }
                }
                return CustomerQtyTable;
            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format("Error in DataMigrationService. Error Message: {0}, Inner Exception :{1}", ex.Message, ex.InnerException.Message));
                Log.error("Error in DataMigrationService", ex);
                throw ex;
            }
        }

        public async Task<DataTable> getPrincipalParticularValAsync()
        {
            try
            {
                var PrincipalParticularVal = new DataTable();
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["PPTLConnectionString"].ConnectionString))
                {
                    string ParticularValTable = ConfigurationManager.AppSettings["ParticularValTable"];

                    using (SqlDataAdapter SqlDataAdapter = new SqlDataAdapter("SELECT * FROM " + ParticularValTable, sqlConnection))
                    {
                        SqlDataAdapter.Fill(PrincipalParticularVal);
                    }
                }
                return PrincipalParticularVal;
            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format("Error in DataMigrationService. Error Message: {0}, Inner Exception :{1}", ex.Message, ex.InnerException.Message));
                Log.error("Error in DataMigrationService", ex);
                throw ex;
            }
        }


        public DataTable GetPrincipalSummaryTable()
        {
            try
            {
                var PrincipalSummaryTable = new DataTable();
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["PPTLConnectionString"].ConnectionString))
                {
                    string SummaryTable = ConfigurationManager.AppSettings["SummaryTable"];

                    using (SqlDataAdapter SqlDataAdapter = new SqlDataAdapter("SELECT * FROM " + SummaryTable, sqlConnection))
                    {
                        SqlDataAdapter.Fill(PrincipalSummaryTable);

                    }
                }

                return PrincipalSummaryTable;
            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format("Error in DataMigrationService. Error Message: {0}, Inner Exception :{1}", ex.Message, ex.InnerException.Message));
                Log.error("Error in DataMigrationService", ex);
                throw ex;
            }
        }

        public DataTable getPrincipalProductSummaryTable()
        {
            try
            {
                var PrincipalProductSummaryTable = new DataTable();
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["PPTLConnectionString"].ConnectionString))
                {
                    string ProductTable = ConfigurationManager.AppSettings["ProductTable"];

                    using (SqlDataAdapter SqlDataAdapter = new SqlDataAdapter("SELECT * FROM " + ProductTable, sqlConnection))
                    {
                        SqlDataAdapter.Fill(PrincipalProductSummaryTable);
                    }
                }
                return PrincipalProductSummaryTable;
            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format("Error in DataMigrationService. Error Message: {0}, Inner Exception :{1}", ex.Message, ex.InnerException.Message));
                Log.error("Error in DataMigrationService", ex);
                throw ex;
            }
        }

        public DataTable GetCustomerQtyTable()
        {
            try
            {
                var CustomerQtyTable = new DataTable();
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["PPTLConnectionString"].ConnectionString))
                {
                    string ParticularQtyTable = ConfigurationManager.AppSettings["ParticularQtyTable"];
                    using (SqlDataAdapter SqlDataAdapter = new SqlDataAdapter("SELECT * FROM " + ParticularQtyTable, sqlConnection))
                    {
                        SqlDataAdapter.Fill(CustomerQtyTable);
                    }
                }
                return CustomerQtyTable;
            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format("Error in DataMigrationService. Error Message: {0}, Inner Exception :{1}", ex.Message, ex.InnerException.Message));
                Log.error("Error in DataMigrationService", ex);
                throw ex;
            }
        }

        public DataTable getPrincipalParticularVal()
        {
            try
            {
                var PrincipalParticularVal = new DataTable();
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["PPTLConnectionString"].ConnectionString))
                {
                    string ParticularValTable = ConfigurationManager.AppSettings["ParticularValTable"];

                    using (SqlDataAdapter SqlDataAdapter = new SqlDataAdapter("SELECT * FROM " + ParticularValTable, sqlConnection))
                    {
                        SqlDataAdapter.Fill(PrincipalParticularVal);
                    }
                }
                return PrincipalParticularVal;
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
