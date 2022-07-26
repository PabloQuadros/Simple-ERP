using System;
using System.Collections.Generic;
using System.Text;

namespace SimpleEPR.Entities
{
    internal class Product
    {
        public int Id { get; private set; }
        public string Name { get; private set; }
        public double Price { get; private set; } 
        public int Quantity { get; private set; }


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
