using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpuxlClassLibrary
{
    struct DecomposedMatrix
    {
        public object[] Row { get; set; }
        public object[,] Matrix { get; set; }
        public object CellValue { get; set; }
    }
}
