using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;
using System.Text;

namespace Excel2txt
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Please specify a file!");
                System.Environment.Exit(1);
            }

            var filePath = args[0];
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"The file '{filePath}' doesn't exist!");
                System.Environment.Exit(2);
            }

            Console.Write($"Extracting data from file: {filePath}");

            var outputFile = args.Length == 1 ? "output.txt" : args[1];

            Console.Write($" to file: {outputFile}");

            //TODO: Not locking reading of excel file!!!
            IWorkbook book1 = new XSSFWorkbook(new FileStream(filePath, FileMode.Open));
            IWorkbook product = new XSSFWorkbook();

            var output = new StringBuilder();

            if (book1.NumberOfSheets == 1)
            {
                SerializeSheet(output, book1.GetSheetAt(0));
            }
            else
            {
                for (int i = 0; i < book1.NumberOfSheets; i++)
                {
                    ISheet sheet = book1.GetSheetAt(i);
                    output.AppendLine($"================== Worksheet:{sheet.SheetName}");
                    SerializeSheet(output, sheet);
                }
            }

            File.WriteAllText(outputFile, output.ToString());

            static void SerializeSheet(StringBuilder output, ISheet sheet)
            {
                for (int iRow = 0; iRow <= sheet.LastRowNum; iRow++)
                {
                    if (sheet.GetRow(iRow) != null) //null is when the row only contains empty cells 
                    {
                        var row = sheet.GetRow(iRow);
                        for (int iCell = 0; iCell <= row.LastCellNum; iCell++)
                        {
                            var cell = row.GetCell(iCell);
                            if (cell != null)
                            {
                                output.Append(cell.RichStringCellValue);
                            }

                            output.Append(";");
                        }
                    }
                    output.AppendLine(string.Empty);
                }
            }
        }
    }
}
