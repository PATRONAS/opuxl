using System.Collections.Generic;

namespace ClientConnector.communication
{
    public class OpuxlMatrix
    {
        public List<MatrixHeader> headers { get; set; }
        public List<List<object>> data { get; set; }
    }
}
