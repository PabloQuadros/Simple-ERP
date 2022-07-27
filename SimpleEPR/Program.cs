using System;
using System.IO;
using SimpleEPR.Entities;

namespace SimpleEPR
{
    internal class Program
    {

        static void Main(string[] args)
        {
            Directory.CreateDirectory(@"C:\VsCodeProjects\SimpleERP\SimpleErpFiles");

            InventoryControl.ReadInventory();
            InventoryControl.PrintInventory();
            InventoryControl.DecreaseQuantity(1422, 14);
           


        }    
    }
}
