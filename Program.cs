using Microsoft.Win32;
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
                //int period = 2; // trial period
                //string keyName = "Software\\DataMigrationSQLtoXLSM";
                //long ticks = DateTime.Today.Ticks;

                //RegistryKey rootKey = Registry.CurrentUser;
                //RegistryKey regKey = rootKey.OpenSubKey(keyName);
                //if (regKey == null) // first time app has been used
                //{
                //    regKey = rootKey.CreateSubKey(keyName);
                //    long expiry = DateTime.Today.AddDays(period).Ticks;
                //    regKey.SetValue("expiry", expiry, RegistryValueKind.QWord);
                //    regKey.Close();
                //}
                //else
                //{
                //    long expiry = (long)regKey.GetValue("expiry");
                //    regKey.Close();
                //    long today = DateTime.Today.Ticks;
                //    if (today > expiry)
                //    {
                //        Console.WriteLine("Your free trial has expired. Please register to continue using the application");
                //        return;
                //    }
                //}


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
