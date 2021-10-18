using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Diagnostics;

namespace DecryptExcelFile
{
    class Program
    {
        static void Main(string[] args)
        {
            string srcDir = "C:\\testsource";
            string desDir = "C:\\testdest";
            List<string> srcFileList = Directory.GetFiles(srcDir).ToList();

            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            //Parallel.ForEach(srcFileList, sourceFile =>
            //{
            //    string destFile = Path.Combine(desDir, Path.GetFileName(sourceFile));

            //    DecryptExcelFile(sourceFile, destFile);
            //    Console.WriteLine(string.Format("{0} is done on {1}", sourceFile, DateTime.Now.ToString()));

            //});
            for (int i = 0; i < srcFileList.Count; i++)
            {
                string sourceFile = srcFileList[i];
                string destFile = Path.Combine(desDir, Path.GetFileName(sourceFile));

                DecryptExcelFile(sourceFile, destFile);
                Console.WriteLine(string.Format("{0} is done", sourceFile));

            }
            stopWatch.Stop();
            Console.WriteLine(string.Format("Elasped time in seconds {0}", stopWatch.Elapsed.TotalSeconds));
            Console.ReadLine();
            

        }

        private static void DecryptExcelFile(string srcExcelFilePath, string desExcelFilePath)
        {
            if(File.Exists(desExcelFilePath))
            {
                File.Delete(desExcelFilePath);
            }

            var excelApp = new Excel.Application();
            // Make the object visible.
            excelApp.Visible = false;
            Excel.Workbook workBook = excelApp.Workbooks.Open(srcExcelFilePath, Type.Missing, Type.Missing, Type.Missing, "123456");
            workBook.Password = "";
            workBook.WritePassword = "";
            workBook.SaveAs(desExcelFilePath);
            excelApp.Quit();
        }

        private static void ParseInput(string[] args, out string sourceDir, out string destDir, out string password)
        {
            sourceDir = "";
            destDir = "";
            password = "";

            for(int i = 0; i < args.Length; i++)
            {
                string key = args[i];
                switch (key.ToLower())
                {
                    case "-s":
                        sourceDir = args[i + 1];
                        break;
                    case "-d":
                        destDir = args[i + 1];
                        break;
                    case "-p":
                        password = args[i + 1];
                        break;
                    default:
                        break;
                }
            }
        }
    }
}
