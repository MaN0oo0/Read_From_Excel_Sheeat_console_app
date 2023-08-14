using Ganss.Excel;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Excel;
using Read_From_Excel_Sheeat_console_app.Models;
using System.Data.OleDb;
using System.IO;

namespace Read_From_Excel_Sheeat_console_app
{
    public class Program
    {
         static   async Task Main(string[] args)
        {

            string path = @"C:\file.xlsx";

            Console.Write("Enter seating_no: ");
            string seatinNo = Console.ReadLine();
            var numbers = "";
            foreach (var num in seatinNo)
            {
                if (!char.IsDigit(num))
                {
                    Console.WriteLine($"{num} is not number");
                }
                else
                {
                    numbers += num;
                    continue;
                }
            }
            var watch = new System.Diagnostics.Stopwatch();

            watch.Start();

            if (seatinNo.Length!=numbers.Length)
            {
                Console.WriteLine("not correct Numbers");
            }
            else
            {
                Console.WriteLine("Plz Wait it will take a 1 min ..");
                string excelFile = path;
                var excel = new ExcelMapper();
                var res = (await excel.FetchAsync<ModelRes>(excelFile)).Where(x => x.seating_no == Convert.ToInt32(numbers)).ToList();
                Console.OutputEncoding = System.Text.Encoding.Unicode;
                foreach (var item in res)
                {
                    Console.WriteLine(item);
                }

            }
           

            watch.Stop();

            Console.WriteLine($"Execution Time: {watch.ElapsedMilliseconds} ms");

        }



    }
}
