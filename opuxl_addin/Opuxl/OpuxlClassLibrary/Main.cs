using System.Linq;
using ExcelDna.Integration;
using System.Collections.Generic;
using ClientConnector;
using Newtonsoft.Json.Linq;
using System.Diagnostics;

namespace OpuxlClassLibrary
{
    /// <summary>
    /// The Main Class for the Excel Addin which operates with the Excel Lifecycle Events.
    /// </summary>
    public class Main : IExcelAddIn
    {
        public void AutoClose()
        {
            // on excel sheet close
        }


        void IExcelAddIn.AutoOpen()
        {
            // Create a client connector which will do an initial request to the Opus ServerSocket and retrieve all available methods.
            var socketConnector = new SocketConnector();
            MethodRegistry registry = new MethodRegistry(socketConnector);

            // Execute the INITIALIZE function to retrieve all available functions from the client
            ResponsePayload response = socketConnector.executeFn(new FunctionPayload(SocketConnector.INITIALIZE));

            if (response != null && response.data != null)
            {
                // The inital request returns the available functions.
                // Example Message from Client:
                // { "type":"INITIALIZE",
                //   "data":[["getVersion",[]],["getPortfolio",[["accountSegmentId","The Account Segment Id for which the Portfolio will be fetched."]]],["getCurrencies",[]]],"error":""}
                response.data.ForEach(entry =>
                {
                    // TODO: Create a serializable Wrapper Class to do this.
                    var arguments = new List<ExcelArgumentAttribute>();
                    // First Element is the function name
                    string functionName = (string)entry.ElementAt(0);
                    // and second Element is an array of parameterName/parameterDescription pairs.
                    JArray argsArray = (JArray)entry.ElementAt(1);

                    foreach (JArray arg in argsArray)
                    {
                        // Create an Excel Argument for each Function Parameter.
                        arguments.Add(new ExcelArgumentAttribute
                        {
                            Name = arg.ElementAt(0).ToString(),
                            Description = arg.ElementAt(1).ToString()
                        });
                    }
                    // Create a delegate in our registry for each function
                    registry.addDelegate(functionName, arguments);
                });
            } else
            {
                Debug.WriteLine("Response not set, cant register functions.");
            }
            

            // Register all retrieves functions as delegates which are going to be available in excel.
            registry.createDelegates();
        }

        /// <summary>
        /// Basic Function to display the version of this Excel Plugin
        /// </summary>
        /// <returns>The Excel Plugin Version</returns>
        [ExcelFunction(Name = "Opus.GetOpusInfo", Description = "Returns the Opus Excel Plugin Information")]
        public static string GetOpusInfo()
        {
            return "Opuxl V1.0";
        }
    }
}
