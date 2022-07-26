using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace SimpleEPR.Entities
{
    internal class InventoryControl
    {
   
        public static List<Product> Products { get; private set; } = new List<Product>();
        public  static void ReadInventory()
        {

            ExcelPackage Inventory = null;

            try
            {
                FileInfo existingFile  = new FileInfo(@"C:\VsCodeProjects\SimpleERP\SimpleErpFiles\Inventory.xlsx");
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                Inventory = new ExcelPackage(existingFile);

                ExcelWorksheet worksheet = Inventory.Workbook.Worksheets[0];

                int colCount = worksheet.Dimension.End.Column;
                int rowCount = worksheet.Dimension.End.Row;

                for(int row = 2; row <= rowCount; row++)
                {
                    int Id = int.Parse(worksheet.Cells[row, 1].Value.ToString());
                    string Name = worksheet.Cells[row, 2].Value.ToString();
                    double Price = double.Parse(worksheet.Cells[row, 3].Value.ToString());
                    int Quantity = int.Parse(worksheet.Cells[row, 4].Value.ToString());
                   

                    Product p = new Product(Id, Name, Price, Quantity);

                    Products.Add(p);
                 
                }

            }
            catch(IOException e)
            {
                Console.WriteLine("An error occurred");
                Console.WriteLine(e.Message);
            }
            finally
            {
                if (Inventory != null)
                {
                    
                    Inventory.Dispose();
                }
            }
        }

        public static void PrintInventory()
        {
            Console.WriteLine("+---------------------------------+");
            Console.WriteLine("| Id \tName \t Price \t Quantity |");
            Console.WriteLine("+---------------------------------+");
            foreach (Product p in Products)
            {
                Console.WriteLine("| "+p.ToString()+ "\t  |");
            }
            Console.WriteLine("+---------------------------------+");
        }

    }
}
