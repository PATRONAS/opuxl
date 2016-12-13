using ClientConnector.communication;
using System.Collections.Generic;
using System.Text;

namespace ClientConnector
{
    /// <summary>
    /// Wrapper Class which represents the Response from the ServerSocket.
    /// </summary>
    public class ResponsePayload
    {
        public string name { get; set; }
        public OpuxlMatrix matrix { get; set; }
        public string error { get; set; }

        public override string ToString()
        {
            StringBuilder builder = new StringBuilder();
            builder.Append("{ name: ").Append(name).Append(" , matrix: ");

            builder.Append("[");
            if (matrix.data != null && matrix.data.Capacity != 0)
            {
                matrix.data.ForEach(row =>
                {
                    row.ForEach(col =>
                    {
                        builder.Append(col).Append(",");
                    });
                });
            }

            builder.Append("]").Append(", Error: ").Append(error).Append(" }");
            return builder.ToString();
        }
    }
}
