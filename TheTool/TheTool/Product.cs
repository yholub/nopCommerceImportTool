using System;
using System.Collections.Generic;
using System.Text;

namespace TheTool
{
    class Product
    {
        public string Name { get; set; }

        public decimal Price { get; set; }

        public int SKU { get; set; }

        public Attributes Attributes { get; set; } = new Attributes();
    }
}
