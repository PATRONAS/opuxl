using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using System.Diagnostics;
using ClientConnector;
using ClientConnector.communication;
using System.Globalization;

/// <summary>
/// Method Registry interacts with the Excel Sheet and registers methods in it. 
/// The delegates contain a reference to the socketConnector such that we can execute calls to the opus client from within the excel delegate.
/// </summary>
namespace OpuxlClassLibrary
{
    class MethodRegistry
    {
        private SocketConnector connector;
        private bool isCalculating;

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
        private static ContentArea Create2DArray(List<MatrixHeader> headers, List<List<object>> lists)
        {
            // We need to add a row if a header is present
            int headerOffset = (headers != null ? 1 : 0);
            int columnsCount = headers != null ? headers.Count : (lists == null ? 0 : lists.ElementAt(0).Count);
            object[,] data = new object[lists.Count + headerOffset, columnsCount];

            if (headers != null)
            {
                // If headers are present, we create a headers row
                for (int i = 0; i < headers.Count; i++)
                {
                    data[0, i] = headers.ElementAt(i).text;
                }
            }

            for (int i = 0; i < lists.Count; i++)
            {
                var array = lists.ElementAt(i);
                if (array.Count != columnsCount)
                {
                    // TODO: throw error. arrays are not the same size
                }

                for (int j = 0; j < columnsCount; j++)
                {
                    data[i + headerOffset, j] = convertResult(headers != null ? headers.ElementAt(j).type : ExcelType.TEXT, array.ElementAt(j));
                }
            }

            return new ContentArea(data);
        }


        private static object convertResult(ExcelType excelType, object val)
        {
            object result = "";
            if (val == null)
            {
                return result;
            }

            string value = val.ToString();

            switch (excelType)
            {
                case ExcelType.NUMBER:
                    if (!string.IsNullOrEmpty(value))
                    {
                        // We have to use the CultureInfo to correctly parse the decimal point
                        result = double.Parse(value, CultureInfo.InvariantCulture);
                    }
                    break;
                case ExcelType.TEXT:
                    result = value;
                    break;
                case ExcelType.DATETIME:
                    result = value;
                    break;
                case ExcelType.LOGICAL:
                    if (!string.IsNullOrEmpty(value))
                    {
                        result = bool.Parse(value);
                    }
                    break;
            }

            return result;
        }

        public void addDelegate(string functionName, string functionDescription, List<OpuxlArgumentAttribute> argumentAttributes)
        {
            dynamic fnDelegate;

            // TODO Make this generic
            switch (argumentAttributes.Count)
            {
                case 0:
                    fnDelegate = MakeDelegate(functionName);
                    break;
                case 1:
                    fnDelegate = MakeDelegate(functionName, argumentAttributes.ElementAt(0));
                    break;
                case 2:
                    fnDelegate = MakeDelegate(functionName, argumentAttributes.ElementAt(0), argumentAttributes.ElementAt(1));
                    break;
                case 3:
                    fnDelegate = MakeDelegate(functionName, argumentAttributes.ElementAt(0), argumentAttributes.ElementAt(1), argumentAttributes.ElementAt(2));
                    break;
                case 4:
                    fnDelegate = MakeDelegate(functionName, argumentAttributes.ElementAt(0), argumentAttributes.ElementAt(1), argumentAttributes.ElementAt(2), argumentAttributes.ElementAt(3));
                    break;
                case 5:
                    fnDelegate = MakeDelegate(functionName, argumentAttributes.ElementAt(0), argumentAttributes.ElementAt(1), argumentAttributes.ElementAt(2), argumentAttributes.ElementAt(3), argumentAttributes.ElementAt(4));
                    break;
                case 6:
                    fnDelegate = MakeDelegate(functionName, argumentAttributes.ElementAt(0), argumentAttributes.ElementAt(1), argumentAttributes.ElementAt(2), argumentAttributes.ElementAt(3), argumentAttributes.ElementAt(4), argumentAttributes.ElementAt(5));
                    break;
                case 7:
                    fnDelegate = MakeDelegate(functionName, argumentAttributes.ElementAt(0), argumentAttributes.ElementAt(1), argumentAttributes.ElementAt(2), argumentAttributes.ElementAt(3), argumentAttributes.ElementAt(4), argumentAttributes.ElementAt(5), argumentAttributes.ElementAt(6));
                    break;
                case 8:
                    fnDelegate = MakeDelegate(functionName, argumentAttributes.ElementAt(0), argumentAttributes.ElementAt(1), argumentAttributes.ElementAt(2), argumentAttributes.ElementAt(3), argumentAttributes.ElementAt(4), argumentAttributes.ElementAt(5), argumentAttributes.ElementAt(6), argumentAttributes.ElementAt(7));
                    break;
                case 9:
                    fnDelegate = MakeDelegate(functionName, argumentAttributes.ElementAt(0), argumentAttributes.ElementAt(1), argumentAttributes.ElementAt(2), argumentAttributes.ElementAt(3), argumentAttributes.ElementAt(4), argumentAttributes.ElementAt(5), argumentAttributes.ElementAt(6), argumentAttributes.ElementAt(7), argumentAttributes.ElementAt(8));
                    break;
                case 10:
                    fnDelegate = MakeDelegate(functionName, argumentAttributes.ElementAt(0), argumentAttributes.ElementAt(1), argumentAttributes.ElementAt(2), argumentAttributes.ElementAt(3), argumentAttributes.ElementAt(4), argumentAttributes.ElementAt(5), argumentAttributes.ElementAt(6), argumentAttributes.ElementAt(7), argumentAttributes.ElementAt(8), argumentAttributes.ElementAt(9));
                    break;
                default:
                    throw new NotImplementedException("Can not handle an argument count of " + argumentAttributes.Count);
            }

            var attribute = new ExcelFunctionAttribute
            {
                // getVersion => Opus.GetVersion
                Name = "Opus." + char.ToUpper(functionName[0]) + functionName.Substring(1),
                Description = functionDescription
            };

            registerDelegate(fnDelegate, attribute, argumentAttributes);
        }



        private Func<object> MakeDelegate(string delegateName)
        {
            Func<object> callback = () =>
            {
                return executeDelegate(delegateName, new List<object> { });
            };

            return callback;
        }

        private Func<object, object> MakeDelegate(string delegateName, OpuxlArgumentAttribute argument1)
        {
            Func<object, object> callback = (a) =>
            {
                return executeDelegate(delegateName, new List<object> {
                    evaluateParameter(argument1, a)
                });
            };

            return callback;
        }

        private Func<object, object, object> MakeDelegate(string delegateName, OpuxlArgumentAttribute argument1, OpuxlArgumentAttribute argument2)
        {
            Func<object, object, object> callback = (a, b) =>
            {
                return executeDelegate(delegateName, new List<object> {
                    evaluateParameter(argument1, a),
                    evaluateParameter(argument2, b)
                });
            };

            return callback;
        }

        private Func<object, object, object, object> MakeDelegate(string delegateName, OpuxlArgumentAttribute argument1, OpuxlArgumentAttribute argument2, OpuxlArgumentAttribute argument3)
        {
            Func<object, object, object, object> callback = (a, b, c) =>
            {
                return executeDelegate(delegateName, new List<object> {
                   evaluateParameter(argument1, a),
                   evaluateParameter(argument2, b),
                   evaluateParameter(argument3, c)
                });
            };

            return callback;
        }

        private Func<object, object, object, object, object> MakeDelegate(string delegateName, OpuxlArgumentAttribute argument1, OpuxlArgumentAttribute argument2, OpuxlArgumentAttribute argument3, OpuxlArgumentAttribute argument4)
        {
            Func<object, object, object, object, object> callback = (a, b, c, d) =>
            {
                return executeDelegate(delegateName, new List<object> {
                   evaluateParameter(argument1, a),
                   evaluateParameter(argument2, b),
                   evaluateParameter(argument3, c),
                   evaluateParameter(argument4, d)
                });
            };

            return callback;
        }

        private Func<object, object, object, object, object, object> MakeDelegate(string delegateName, OpuxlArgumentAttribute argument1, OpuxlArgumentAttribute argument2, OpuxlArgumentAttribute argument3, OpuxlArgumentAttribute argument4, OpuxlArgumentAttribute argument5)
        {
            Func<object, object, object, object, object, object> callback = (a, b, c, d, e) =>
            {
                return executeDelegate(delegateName, new List<object> {
                   evaluateParameter(argument1, a),
                   evaluateParameter(argument2, b),
                   evaluateParameter(argument3, c),
                   evaluateParameter(argument4, d),
                   evaluateParameter(argument5, e)
                });
            };

            return callback;
        }

        private Func<object, object, object, object, object, object, object> MakeDelegate(string delegateName, OpuxlArgumentAttribute argument1, OpuxlArgumentAttribute argument2, OpuxlArgumentAttribute argument3, OpuxlArgumentAttribute argument4, OpuxlArgumentAttribute argument5, OpuxlArgumentAttribute argument6)
        {
            Func<object, object, object, object, object, object, object> callback = (a, b, c, d, e, f) =>
            {
                return executeDelegate(delegateName, new List<object> {
                   evaluateParameter(argument1, a),
                   evaluateParameter(argument2, b),
                   evaluateParameter(argument3, c),
                   evaluateParameter(argument4, d),
                   evaluateParameter(argument5, e),
                   evaluateParameter(argument6, f)
                });
            };

            return callback;
        }

        private Func<object, object, object, object, object, object, object, object> MakeDelegate(string delegateName, OpuxlArgumentAttribute argument1, OpuxlArgumentAttribute argument2, OpuxlArgumentAttribute argument3, OpuxlArgumentAttribute argument4, OpuxlArgumentAttribute argument5, OpuxlArgumentAttribute argument6, OpuxlArgumentAttribute argument7)
        {
            Func<object, object, object, object, object, object, object, object> callback = (a, b, c, d, e, f, g) =>
             {
                 return executeDelegate(delegateName, new List<object> {
                   evaluateParameter(argument1, a),
                   evaluateParameter(argument2, b),
                   evaluateParameter(argument3, c),
                   evaluateParameter(argument4, d),
                   evaluateParameter(argument5, e),
                   evaluateParameter(argument6, f),
                   evaluateParameter(argument7, g)
                 });
             };

            return callback;
        }

        private Func<object, object, object, object, object, object, object, object, object> MakeDelegate(string delegateName, OpuxlArgumentAttribute argument1, OpuxlArgumentAttribute argument2, OpuxlArgumentAttribute argument3, OpuxlArgumentAttribute argument4, OpuxlArgumentAttribute argument5, OpuxlArgumentAttribute argument6, OpuxlArgumentAttribute argument7, OpuxlArgumentAttribute argument8)
        {
            Func<object, object, object, object, object, object, object, object, object> callback = (a, b, c, d, e, f, g, h) =>
            {
                return executeDelegate(delegateName, new List<object> {
                   evaluateParameter(argument1, a),
                   evaluateParameter(argument2, b),
                   evaluateParameter(argument3, c),
                   evaluateParameter(argument4, d),
                   evaluateParameter(argument5, e),
                   evaluateParameter(argument6, f),
                   evaluateParameter(argument7, g),
                   evaluateParameter(argument8, h)
                 });
            };

            return callback;
        }

        private Func<object, object, object, object, object, object, object, object, object, object> MakeDelegate(string delegateName, OpuxlArgumentAttribute argument1, OpuxlArgumentAttribute argument2, OpuxlArgumentAttribute argument3, OpuxlArgumentAttribute argument4, OpuxlArgumentAttribute argument5, OpuxlArgumentAttribute argument6, OpuxlArgumentAttribute argument7, OpuxlArgumentAttribute argument8, OpuxlArgumentAttribute argument9)
        {
            Func<object, object, object, object, object, object, object, object, object, object> callback = (a, b, c, d, e, f, g, h, i) =>
            {
                return executeDelegate(delegateName, new List<object> {
                   evaluateParameter(argument1, a),
                   evaluateParameter(argument2, b),
                   evaluateParameter(argument3, c),
                   evaluateParameter(argument4, d),
                   evaluateParameter(argument5, e),
                   evaluateParameter(argument6, f),
                   evaluateParameter(argument7, g),
                   evaluateParameter(argument8, h),
                   evaluateParameter(argument9, i)
                 });
            };

            return callback;
        }

        private Func<object, object, object, object, object, object, object, object, object, object, object> MakeDelegate(string delegateName, OpuxlArgumentAttribute argument1, OpuxlArgumentAttribute argument2, OpuxlArgumentAttribute argument3, OpuxlArgumentAttribute argument4, OpuxlArgumentAttribute argument5, OpuxlArgumentAttribute argument6, OpuxlArgumentAttribute argument7, OpuxlArgumentAttribute argument8, OpuxlArgumentAttribute argument9, OpuxlArgumentAttribute argument10)
        {
            Func<object, object, object, object, object, object, object, object, object, object, object> callback = (a, b, c, d, e, f, g, h, i, j) =>
            {
                return executeDelegate(delegateName, new List<object> {
                   evaluateParameter(argument1, a),
                   evaluateParameter(argument2, b),
                   evaluateParameter(argument3, c),
                   evaluateParameter(argument4, d),
                   evaluateParameter(argument5, e),
                   evaluateParameter(argument6, f),
                   evaluateParameter(argument7, g),
                   evaluateParameter(argument8, h),
                   evaluateParameter(argument9, i),
                   evaluateParameter(argument10, j)
                 });
            };

            return callback;
        }

        private object evaluateParameter(OpuxlArgumentAttribute attr, object value)
        {
            object result = value;
            if (attr.Optional && (value is ExcelMissing || value is ExcelEmpty))
            {
                // if the attribute is optional and the value is Missing, then we can just return null as the client knows how to handle that
                result = null;
            }
            return result;
        }

        private object executeDelegate(String delegateName, List<object> args)
        {
            object result;
            // We have to remember whether we are in the fn wizard. As soon as we execute the task, we can't get that information anymore.
            if (ExcelDnaUtil.IsInFunctionWizard() || isCalculating)
            {
                result = delegateName;
            }
            else
            {
                isCalculating = true;
                result = executeTask(delegateName, args);
                isCalculating = false;

            }
            return result;
        }

        private object executeTask(String delegateName, List<object> args)
        {
            FunctionPayload payload = new FunctionPayload(delegateName);

            args.ForEach(arg => payload.args.Add(arg));

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

            object result;
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
                // Execution worked fine, extract the response into our decomposed matrix.
                ContentArea rowMatrix = Create2DArray(response.matrix.headers, response.matrix.data);
                // Rendered the content into the excel sheet
                rowMatrix.display();
                // Set the result of the formula cell to the cell value of the matrix.
                result = rowMatrix.CellValue;
            }

            return result;
        }

        

        private void registerDelegate(Delegate del, object att, List<OpuxlArgumentAttribute> argAtt)
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
