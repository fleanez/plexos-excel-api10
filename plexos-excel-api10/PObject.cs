using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace plexos_excel_api_10
{
    class PObject
    {

        public PObject(string classname, string name)
        {
            Name = name;
            ClassName = classname;
        }

        public string Name { get; }

        public string ClassName { get; }

    }

}
