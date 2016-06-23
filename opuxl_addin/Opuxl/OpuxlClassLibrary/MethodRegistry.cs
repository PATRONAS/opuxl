using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;
using System.Diagnostics;
using ClientConnector;

/// <summary>
/// Method Registry interacts with the Excel Sheet and registers methods in it. 
/// The delegates contain a reference to the socketConnector such that we can execute calls to the opus client from within the excel delegate.
/// </summary>
namespace OpuxlClassLibrary
{
    class MethodRegistry
    {
        private SocketConnector connector;

        List<Delegate> Delegates { get; set; }
        List<object> Attributes { get; set; }
        List<List<object>> Arguments { get; set; }

        private MethodRegistry()
        {
            Delegates = new List<Delegate>();
            Attributes = new List<object>();
            Arguments = new List<List<object>>();
        }

        public MethodRegistry(SocketConnector connector) : this()
        {
            this.connector = connector;
        }
 
        /// <summary>
        /// Transforms a list of list of objects into its equivalent 2d array. The Excel Sheet reference needs a 2d Array as its input.
        /// </summary>
        /// <param name="lists"></param>
        /// <returns></returns>
        private static object[,] Create2DArray(List<List<object>> lists)
        {
            int minorLength = lists.ElementAt(0).Count;
            object[,] result = new object[lists.Count, minorLength];

            for (int i = 0; i < lists.Count; i++)
            {
                var array = lists.ElementAt(i);
                if (array.Count != minorLength)
                {
                    // TODO: throw error. arrays are not the same size
                }
                for (int j = 0; j < minorLength; j++)
                {
                    result[i, j] = array.ElementAt(j);
                }
            }

            return result;
        }
 
        // argumentAttributes = new ExcelArgumentAttribute { Name = ..., Description = ... }
        public void addDelegate(string functionName, List<ExcelArgumentAttribute> argumentAttributes)
        {
            dynamic fnDelegate;
            if (argumentAttributes.Count > 0)
            {
                fnDelegate = MakeDelegate(functionName, argumentAttributes);
            }
            else
            {
                fnDelegate = MakeDelegate(functionName);
            }

            var attribute = new ExcelFunctionAttribute
            {
                // getVersion => Opus.GetVersion
                Name = "Opus." + char.ToUpper(functionName[0]) + functionName.Substring(1),
                // TODO Descriptions should be provided by the connector
                Description = "The Function Description"
            };

            registerDelegate(fnDelegate, attribute, argumentAttributes);
        }

        private Func<long, string> MakeDelegate(string delegateName, List<ExcelArgumentAttribute> args)
        {
            // TODO we have to create delegates for a dynamic amount of arguments.
            // Arguments should have expressiv names here, as they will appear in excel as a tooltip.
            Func<long, string> callback = (argument1) =>
            {
                String result;
                FunctionPayload payload = new FunctionPayload(delegateName);
                payload.args.Add(argument1);

                ResponsePayload response;
                try
                {
                    response = connector.executeFn(payload);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex);
                    return ex.Message;
                }

                if (response == null)
                {
                    result = "Function Execution failed.";
                }
                else if (response.error.Length > 0)
                {
                    result = response.error;
                }
                else
                {
                    // Write the result into our excel sheet. This has to be done in an Async Thread.
                    writeMatrixToExcelSheet(response);
                    result = delegateName;
                }

                // This reponse for the currently selected cell
                return result;
            };

            return callback;
        }

        private Func<string> MakeDelegate(string delegateName)
        {

            Func<string> callback = () =>
            {
                String result;
                FunctionPayload payload = new FunctionPayload(delegateName);
                ResponsePayload response;
                try
                {
                    response = connector.executeFn(payload);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex);
                    return ex.Message;
                }

                if (response.error.Length > 0)
                {
                    result = response.error;
                }
                else
                {
                    writeMatrixToExcelSheet(response);
                    result = delegateName;
                }

                // The result is the string, which will be displayed in the currently selected cell.
                return result;
            };

            return callback;
        }

        /// <summary>
        /// Async Task to create a matrix in the currently opened excel sheet.
        /// The Task has to be async such that we can write into adjacent cells as it is not allowed in the 'ui thread'.
        /// </summary>
        /// <param name="response"></param>
        private void writeMatrixToExcelSheet(ResponsePayload response)
        {
            Excel.Application app = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Range range = app.ActiveCell;

            var rows = response.data.Count;
            int cols = 0;
            if (response.data.Count > 0)
            {
                cols = response.data.ElementAt(0).Count;
            }

            object[,] data2d = Create2DArray(response.data);

            // TODO rearrange the matrix (move up by 1 and to the left by 1. selected cell is value [0][0])
            var reference = new ExcelReference(range.Row, range.Row + rows - 1, range.Column - 1, range.Column - 2 + cols);
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                // TODO the setValue method should NOT override the format of the excel cells.
                reference.SetValue(data2d);
            });
        }

        private void registerDelegate(Delegate del, object att, List<ExcelArgumentAttribute> argAtt)
        {
            Delegates.Add(del);
            Attributes.Add(att);
            Arguments.Add((argAtt.Cast<object>().ToList()));
        }

        /// <summary>
        /// Registeres the available Functions with their descriptions into Excel
        /// </summary>
        public void createDelegates()
        {
            ExcelIntegration.RegisterDelegates(
               Delegates,
               Attributes,
               Arguments
            );
        }

    }
}
