using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace plexos_excel_api_10
{
    internal class PAttribute
    {

        public PAttribute (string classname, string childname, string name, double dvalue)
        {
            ClassName = classname;
            ChildName = childname;
            Name = name;
            Value = dvalue;
        }

        public string Name { get; set; }

        public string ClassName { get; set; }

        public double Value { get; set; }

        public string ChildName { get; set; }

    }
}
