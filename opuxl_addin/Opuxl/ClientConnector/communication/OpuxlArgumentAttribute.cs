using ExcelDna.Integration;

namespace ClientConnector.communication
{
    public class OpuxlArgumentAttribute : ExcelArgumentAttribute
    {
        public bool Optional;
        public ExcelType Type;
    }
}
