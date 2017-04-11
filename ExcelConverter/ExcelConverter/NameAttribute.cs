using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelConverter
{
    public class NameAttribute:Attribute
    {
        public string Name { get; set; }
        public NameAttribute(string name)
        {
            Name = name;
        }
    }
}
