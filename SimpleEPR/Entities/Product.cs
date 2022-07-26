using System;
using System.Collections.Generic;
using System.Text;

namespace SimpleEPR.Entities
{
    internal class Product
    {
        public int Id { get;  set; }
        public string Name { get;  set; }
        public double Price { get;  set; } 
        public int Quantity { get;  set; }


        public Product(int id, string name, double price, int quantity)
        {
            Id = id;
            Name = name;
            Price = price;
            Quantity = quantity;
        }


        public override string ToString()
        {
            return  Id + "\t" + Name + " \t $ " + Price + "\t "+ Quantity;
        }
    }
}
