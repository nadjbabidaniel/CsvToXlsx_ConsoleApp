using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CsvToXlsxNetFramework
{
    public class Program
    {
        static void Main(string[] args)
        {
            string repeat;
            Console.WriteLine(@"Please insert path where csv files are (in format C:\Users\Operations Intern 5\Desktop\Michael\): ");
            var path = Console.ReadLine();
            
            if(!path.EndsWith(@"\")) path += @"\";

            path = path.Trim(' ');

            do
            {
                List<string> filePaths = Directory.GetFiles(path, "*.csv").ToList();
                foreach (string csvPath in filePaths)
                {
                    try
                    {
                        Application app = new Application();
                        Workbook wb = app.Workbooks.Open(csvPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                        string name = path + Path.GetFileNameWithoutExtension(csvPath);
                        wb.SaveAs(name, XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        wb.Close();
                        app.Quit();
                    }
                    catch (Exception exception)
                    {
                        Console.WriteLine(exception);
                    }
                }

                Console.Write("Do you want to continue on the same path (Y)");
                repeat = Console.ReadLine().ToLower();

            } while (repeat.Equals("y"));
        }
    }
}
