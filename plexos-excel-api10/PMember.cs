using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace plexos_excel_api_10
{
    internal class PMember
    {
        public PMember(string collection, PObject parent, PObject child)
        {
            Collection = collection;
            Parent = parent;
            Child = child;
        }

        public string Collection { get; set; }

        public PObject Parent { get; set; }

        public PObject Child { get; set; }

    }
}
