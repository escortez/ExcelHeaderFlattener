using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChangeHeaders
{
    class Program
    {
        static void Main(string[] args)
        {

            try
            {
                //Sample 3, 4 and 12 uses the Adventureworks database. Enter then name of your SQL server into the variable below...
                //Leave this blank if you don't have access to the Adventureworks database 
                string SqlServerName = "";

                // change this line to contain the path to the output folder
                DirectoryInfo outputDir = new DirectoryInfo(@"c:\temp");
                if (!outputDir.Exists) throw new Exception("outputDir does not exist!");

                // Sample 1 - simply creates a new workbook from scratch
                // containing a worksheet that adds a few numbers together 
                Console.WriteLine("Running UpdatHeader");
               
                string output = UpdateHeader.Run(args.FirstOrDefault(), outputDir);
                Console.WriteLine("New File created: {0}", output);
                Console.WriteLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: {0}", ex.Message);
            }
            Console.WriteLine();
            Console.WriteLine("Press the return key to exit...");
            Console.Read();
        }
    }
}
