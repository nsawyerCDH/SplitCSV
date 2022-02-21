using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using ADSFramework.Excel;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.Write("Enter the Fully qualified File path to Split: ");
            string InputFilePath = Console.ReadLine().Replace("\"", "");

            Console.Write("Enter the number of records to split on: ");
            int.TryParse(Console.ReadLine(), out int SplitNum);

            if (SplitNum <= 0)
            {
                Console.WriteLine("Split number invalid! Aborting...");
                return;
            }

            var dt = CsvConverter.ToDatatable(InputFilePath);
            var outs = new List<DataTable>();

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (i % SplitNum == 0 || i == 0)
                    outs.Add(dt.Clone());

                outs[outs.Count - 1].Rows.Add(dt.Rows[i].ItemArray);
            }

            foreach (DataTable outDt in outs)
            {

                string OutputFilePath = Path.ChangeExtension(Path.Combine(Path.GetDirectoryName(InputFilePath), Path.GetFileNameWithoutExtension(InputFilePath) + "_" + (outs.IndexOf(outDt) + 1)), "csv");

                if (File.Exists(OutputFilePath))
                {
                    Console.Write($"Cannot overwrite existing file {OutputFilePath}!  Aborting...");
                    return;
                }

                CsvConverter.ToCsv(outDt, ',', OutputFilePath);
            }

            Console.WriteLine("Split Complete!");
            Console.Write("Press any key to exit. ");
            Console.ReadLine();
        }
    }
}
