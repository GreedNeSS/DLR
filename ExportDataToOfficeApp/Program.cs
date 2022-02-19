using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExportDataToOfficeApp
{
    class Program
    {
        static void Main(string[] args)
        {
            List<Car> carsInStock = new List<Car>
            {
                new Car {Color="Green", Make="VW", PetName="Mary"},
                new Car {Color="Red", Make="Saab", PetName="Mel"},
                new Car {Color="Black", Make="Ford", PetName="Hank"},
                new Car {Color="Yellow", Make="BMW", PetName="Davie"}
            };

            //ExportToExcelManual(carsInStock);
            ExportToExcel(carsInStock);
        }

        static void ExportToExcel(List<Car> carsInStock)
        {
            Excel.Application excelApp = new Excel.Application();
            excelApp.Workbooks.Add();

            Excel._Worksheet worksheet = excelApp.ActiveSheet;

            worksheet.Cells[1, "A"] = "Make";
            worksheet.Cells[1, "B"] = "Color";
            worksheet.Cells[1, "C"] = "Pet Name";

            int row = 1;
            foreach (var c in carsInStock)
            {
                row++;
                worksheet.Cells[row, "A"] = c.Make;
                worksheet.Cells[row, "B"] = c.Color;
                worksheet.Cells[row, "C"] = c.PetName;
            }

            worksheet.Range["A1"].AutoFormat(Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic2);

            worksheet.SaveAs($@"{Environment.CurrentDirectory}\Inventory.xlsx");
            excelApp.Quit();
            Console.WriteLine("The Inventory.xlsx file has been saved to your app folder");
        }

        static void ExportToExcelManual(List<Car> carsInStock)
        {
            Excel.Application excelApp = new Excel.Application();
            excelApp.Workbooks.Add(Type.Missing);

            Excel._Worksheet worksheet = (Excel._Worksheet)excelApp.ActiveSheet;

            ((Excel.Range)worksheet.Cells[1, "A"]).Value2 = "Make";
            ((Excel.Range)worksheet.Cells[1, "B"]).Value2 = "Color";
            ((Excel.Range)worksheet.Cells[1, "C"]).Value2 = "Pet Name";

            int row = 1;
            foreach (Car c in carsInStock)
            {
                row++;
                ((Excel.Range)worksheet.Cells[row, "A"]).Value2 = c.Make;
                ((Excel.Range)worksheet.Cells[row, "B"]).Value2 = c.Color;
                ((Excel.Range)worksheet.Cells[row, "C"]).Value2 = c.PetName;
            }

            excelApp.Range["A1", Type.Missing].AutoFormat(Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic2,
                Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing);

            worksheet.SaveAs($@"{Environment.CurrentDirectory}\Inventory.xlsx",
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing);
            excelApp.Quit();
            Console.WriteLine("The Inventory.xlsx file has been saved to your app folder");
        }
    }
}
