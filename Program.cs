using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataMigrationSQLtoXML
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("*** DataMigrationSQLtoXML Service Started***");
                Log.info("***IBM_SPP_ExtractVM Service Started***");
                DataMigrationService dataMigrationService = new DataMigrationService();
                dataMigrationService.ProcessDataMigrationSQLtoXML();
                Console.WriteLine("***DataMigrationSQLtoXML Service End***");
                Log.info("***DataMigrationSQLtoXML Service End***");

            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format("Error in DataMigrationSQLtoXML Service. Error Message: {0}, Inner Exception :{1}", ex.Message, ex.InnerException.Message));
                Log.error("Error in DataMigrationSQLtoXML Service", ex);
                //  throw;
            }
        }
    }
}
