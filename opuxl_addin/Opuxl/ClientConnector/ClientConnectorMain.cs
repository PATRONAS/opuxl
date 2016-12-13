

/// <summary>
/// Simple Test to play around with the socket connection.
/// </summary>
namespace ClientConnector
{
    // TODO: move to unit test
    class ClientConnectorMain
    {
        static int Main(string[] args)
        {
            System.Diagnostics.Debug.WriteLine("Starting ...");
            var server = new SocketConnector();
            ResponsePayload response = server.executeFn(new FunctionPayload(SocketConnector.INITIALIZE));

            if (response.name.Equals(SocketConnector.INITIALIZE))
            {
                System.Diagnostics.Debug.WriteLine("Register the following methods: " + response.ToString());
            }

            return 0;
        }
    }
}
