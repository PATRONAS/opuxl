using System.Collections.Generic;
using System.Text;

/// <summary>
/// Wrapper class to describe a function, which will be send to the server socket.
/// </summary>
namespace ClientConnector
{
    public class FunctionPayload
    {
        /// <summary>
        /// The Name of the Function
        /// </summary>
        public string name { get; set; }
        /// <summary>
        /// The Arguments of the Function
        /// </summary>
        public List<object> args { get; set; }

        public FunctionPayload()
        {
            args = new List<object>();
        }

        public FunctionPayload(string name) : this()
        {
            this.name = name;
        }

        public FunctionPayload(string name, List<object> args) : this(name)
        {
            this.args = args;
        }


        public override string ToString()
        {
            StringBuilder builder = new StringBuilder();
            builder.Append("{ name: ").Append(name).Append(", ");
            builder.Append("args: [");

            args.ForEach((e) =>
            {
                builder.Append(e).Append(",");
            });

            builder.Append("] }").AppendLine();

            return builder.ToString();
        }
    }
}
