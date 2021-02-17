using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataMigrationSQLtoXLSM
{
    class Program
    {

        static async Task Main(string[] args)
        {
            try
            {

                Console.WriteLine("*** DataMigrationSQLtoXML Service Started***");
                Log.info("***DataMigrationSQLtoXML Service Started***");
                var watch = new System.Diagnostics.Stopwatch();
                watch.Start();
                DataMigrationService dataMigrationService = new DataMigrationService();
                await dataMigrationService.ProcessDataMigrationSQLtoXML();
                watch.Stop();

                Console.WriteLine($"Execution Time: {watch.ElapsedMilliseconds} ms");
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
