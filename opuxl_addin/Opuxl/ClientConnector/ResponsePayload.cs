using System;
using System.Collections.Generic;
using System.Text;

namespace ClientConnector
{
    /// <summary>
    /// Wrapper Class which represents the Response from the ServerSocket.
    /// </summary>
    public class ResponsePayload
    {
        public string type { get; set; }
        public List<List<object>> data { get; set; }
        public string error { get; set; }

        public override string ToString()
        {
            StringBuilder builder = new StringBuilder();
            builder.Append("{ type: ").Append(type).Append(" , data: ");

            builder.Append("[");
            if (data.Capacity != 0)
            {
                data.ForEach(row =>
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
