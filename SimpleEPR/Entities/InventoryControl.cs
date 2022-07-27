using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;


namespace SimpleEPR.Entities
{
    internal abstract class InventoryControl
    {
   
        public static List<Product> Products { get; private set; } = new List<Product>();
        public static int colCount { get; set; }
        public static int rowCount { get; set; }

        static FileInfo existingFile;
        public static string pathInventory = @"C:\Users\Public\Documents\Inventory.xlsx";
        static ExcelPackage Inventory = null;
        static ExcelWorksheet worksheet;
        public  static void ReadInventory()
        {
            if (File.Exists(pathInventory))
            {
                try
                {

                    existingFile = new FileInfo(pathInventory);
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    Inventory = new ExcelPackage(existingFile);

                    worksheet = Inventory.Workbook.Worksheets[0];

                    colCount = worksheet.Dimension.End.Column;
                    rowCount = worksheet.Dimension.End.Row;

                    for (int row = 2; row <= rowCount; row++)
                    {
                        int Id = int.Parse(worksheet.Cells[row, 1].Value.ToString());
                        string Name = worksheet.Cells[row, 2].Value.ToString();
                        double Price = double.Parse(worksheet.Cells[row, 3].Value.ToString());
                        int Quantity = int.Parse(worksheet.Cells[row, 4].Value.ToString());


                        Product p = new Product(Id, Name, Price, Quantity);

                        Products.Add(p);

                    }
                    Inventory.Save();

                }
                catch (IOException e)
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
            else
            {
                Console.WriteLine(@"Inventory file created in path: C:\Users\Public\Documents\Inventory.xlsx");
                
                
                try
                {

                    //existingFile = new FileInfo(pathInventory);
                   ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                   Inventory = new ExcelPackage();

                    worksheet = Inventory.Workbook.Worksheets.Add("Inventory");
                    
                    
                    //colCount = worksheet.Dimension.End.Column;
                   // rowCount = worksheet.Dimension.End.Row;

                    worksheet.Cells[1, 1].Value = "ID";
                    worksheet.Cells[1, 2].Value = "Name";
                    worksheet.Cells[1, 3].Value = "Price";
                    worksheet.Cells[1, 4].Value = "Quantity";
                    Inventory.Save();

                    FileStream obj = File.Create(pathInventory);
                    obj.Close();
                    File.WriteAllBytes(pathInventory, Inventory.GetAsByteArray());
                    

                }
                catch (IOException e)
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

        public static void IncreaseQuantity(int Id ,int Quantity)
        {
            for(int i = 0; i < rowCount-1; i++)
            {
                if ( Products[i].Id == Id)
                {
                    Products[i].Quantity += Quantity;
                    try
                    {
                        existingFile = new FileInfo(@"C:\VsCodeProjects\SimpleERP\SimpleErpFiles\Inventory.xlsx");
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        Inventory = new ExcelPackage(existingFile);
                        worksheet = Inventory.Workbook.Worksheets[0];

                        colCount = worksheet.Dimension.End.Column;
                        rowCount = worksheet.Dimension.End.Row;

                        worksheet.Cells[i+2, 4].Value = Products[i].Quantity;



                        Inventory.Save();

                    }
                    catch (IOException e)
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
                    Console.WriteLine("Modified quantity");
                    return;
                }
               
            }
            Console.WriteLine("Product not found");
            return;
        }

        public static void AddNewItem(int Id, string Name, double Price, int Quantity) 
        {
            foreach(Product p in Products)
            {
                if (p.Name == Name || p.Id == Id)
                {
                    if(p.Name == Name)
                    {
                        Console.WriteLine("This product is already in inventory. Name: "+ Name );
                    }
                    else
                    {
                        Console.WriteLine("This product is already in inventory. Id:" + Id);
                    }
                    
                    return;
                }
                
            }
            try
            {
                existingFile = new FileInfo(@"C:\VsCodeProjects\SimpleERP\SimpleErpFiles\Inventory.xlsx");
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                Inventory = new ExcelPackage(existingFile);
                worksheet = Inventory.Workbook.Worksheets[0];

                colCount = worksheet.Dimension.End.Column;
                rowCount = worksheet.Dimension.End.Row;
                rowCount++;
                Product p = new Product(Id, Name, Price, Quantity);
                Products.Add(p);

                worksheet.Cells[rowCount, 1].Value = Id;
                worksheet.Cells[rowCount, 2].Value = Name;
                worksheet.Cells[rowCount, 3].Value = Price;
                worksheet.Cells[rowCount, 4].Value = Quantity;

              

                Inventory.Save();

            }
            catch (IOException e)
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
   
        public static void DecreaseQuantity(int Id, int Quantity)
        {
            for (int i = 0; i < rowCount-1; i++)
            {
                if (Products[i].Id == Id)
                {
                    int aux = Products[i].Quantity;
                    if ((aux -= Quantity) >= 0)
                    {
                        Products[i].Quantity -= Quantity;
                        try
                        {
                            existingFile = new FileInfo(@"C:\VsCodeProjects\SimpleERP\SimpleErpFiles\Inventory.xlsx");
                            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                            Inventory = new ExcelPackage(existingFile);
                            worksheet = Inventory.Workbook.Worksheets[0];

                            colCount = worksheet.Dimension.End.Column;
                            rowCount = worksheet.Dimension.End.Row;

                            worksheet.Cells[i + 2, 4].Value = Products[i].Quantity;



                            Inventory.Save();

                        }
                        catch (IOException e)
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
                        Console.WriteLine("Modified quantity");
                        return;
                    }
                    Console.WriteLine($"Invalid quantity.There are only {Products[i].Quantity} unit(s) in stock and {Quantity} unit is being taken out.");
                    return;              
                }

            }
            Console.WriteLine("Product not found");
            return;
        }

        public static void Removeitem(int Id)
        {
            for (int i = 0; i < rowCount - 1; i++)
            {
                if (Products[i].Id == Id)
                {
                    Products.Remove(Products[i]);
                    try
                    {
                        existingFile = new FileInfo(@"C:\VsCodeProjects\SimpleERP\SimpleErpFiles\Inventory.xlsx");
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        Inventory = new ExcelPackage(existingFile);
                        worksheet = Inventory.Workbook.Worksheets[0];


                        worksheet.Cells[i + 2, 1].Delete(eShiftTypeDelete.Up);
                        worksheet.Cells[i + 2, 2].Delete(eShiftTypeDelete.Up);
                        worksheet.Cells[i + 2, 3].Delete(eShiftTypeDelete.Up);
                        worksheet.Cells[i + 2, 4].Delete(eShiftTypeDelete.Up);


                        colCount = worksheet.Dimension.End.Column;
                        rowCount = worksheet.Dimension.End.Row;

                        



                        Inventory.Save();

                    }
                    catch (IOException e)
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
                    Console.WriteLine("Product removed");
                    return;
                }


            }
            Console.WriteLine("Product not found");
            return;
        }

        public static void ModifyPrice(int Id, double NewPrice)
        {
            for (int i = 0; i < rowCount - 1; i++)
            {
                if (Products[i].Id == Id)
                {
                    Products[i].Price = NewPrice;
                    try
                    {
                        existingFile = new FileInfo(@"C:\VsCodeProjects\SimpleERP\SimpleErpFiles\Inventory.xlsx");
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        Inventory = new ExcelPackage(existingFile);
                        worksheet = Inventory.Workbook.Worksheets[0];

                        colCount = worksheet.Dimension.End.Column;
                        rowCount = worksheet.Dimension.End.Row;

                        worksheet.Cells[i + 2, 3].Value = Products[i].Price;


                        Inventory.Save();

                    }
                    catch (IOException e)
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
                    Console.WriteLine("Modified Price");
                    return;
                }

            }
            Console.WriteLine("Product not found");
            return;
        }
    }        
}
