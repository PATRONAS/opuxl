

using System.ComponentModel;

namespace ClientConnector.communication
{
    public class MatrixHeader
    {
        public string text { get; set; }
        [DefaultValue(ExcelType.TEXT)]
        public ExcelType type { get; set; }
    }
}
