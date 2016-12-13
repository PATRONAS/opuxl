using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
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
        private static char PROPERTY_DIVIDER = '&';
        private SocketConnector connector;
        private bool isCalculating;

        private CustomPropertyUtil propertyUtil = new CustomPropertyUtil();

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
        private static DecomposedMatrix Create2DArray(List<MatrixHeader> headers, List<List<object>> lists)
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

            var rowMatrix = new DecomposedMatrix();
            // Cell Value is the [0,0] value of the matrix
            rowMatrix.CellValue = data[0, 0];
            // We have to remove the first row of the matrix as we will handle it seperatly
            rowMatrix.Matrix = RemoveRowFromMatrix(0, data);
            // And extract the first row without the first element
            rowMatrix.Row = ExtractTailOfRow(data);

            return rowMatrix;
        }

        private static object[] ExtractTailOfRow(object[,] matrix)
        {

            var columnsCount = matrix.GetLength(1);
            object[] result = new object[columnsCount - 1];

            for(int i = 1; i < columnsCount; i++)
            {
                result[i - 1] = matrix[0, i];
            }

            return result;
        }

        private static object[,] RemoveRowFromMatrix(int rowToRemove, object[,] matrix)
        {
            int rowsToKeep = matrix.GetLength(0) - 1;
            object[,] result = new object[rowsToKeep, matrix.GetLength(1)];
            int currentRow = 0;
            for (int i = 0; i < matrix.GetLength(0); i++)
            {
                if (i != rowToRemove)
                {
                    for (int j = 0; j < matrix.GetLength(1); j++)
                    {
                        result[currentRow, j] = matrix[i, j];
                    }
                    currentRow++;
                }
            }
            return result;
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
                // TODO Descriptions should be provided by the connector
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
                // Write the result into our excel sheet. This has to be done in an Async Thread.
                DecomposedMatrix rowMatrix = Create2DArray(response.matrix.headers, response.matrix.data);
                writeMatrixToExcelSheet(rowMatrix);
                result = rowMatrix.CellValue;
            }

            return result;
        }

        /// <summary>
        /// Async Task to create a matrix in the currently opened excel sheet.
        /// The Task has to be async such that we can write into adjacent cells as it is not allowed in the 'ui thread'.
        /// </summary>
        /// <param name="response"></param>
        private void writeMatrixToExcelSheet(DecomposedMatrix rowMatrix)
        {
            Excel.Application app = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Worksheet worksheet = (Excel.Worksheet)app.ActiveWorkbook.ActiveSheet;
            Excel.Range startCell = app.ActiveCell;

            var propKey = "" + startCell.Row + PROPERTY_DIVIDER + startCell.Column;

            /* 
             * We retrieve the existing custom property for the current cell as we are going to closure around it within the async macro.
             * The property contains the current dimensions for the current matrix in the excel sheet. We need to remove this before
             * inserting new data, such that the sheet doesn't contain any stale data.
             */
            int[] existingDimensions = getExistingDimensions(worksheet, propKey);

            // Set a custom propert which has the current cell as its key and the col/rows count of the new matrix as the value.
            // We need that data to "clear" the matrix on the next execution of the method within the cell.
            propertyUtil.SetCustomProp(worksheet, propKey, "" + rowMatrix.Matrix.GetLength(0) + PROPERTY_DIVIDER + rowMatrix.Matrix.GetLength(1));


            // Have to do the update within an async tasks, as UDFs are not allowed to manipulate other cells.
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                if (existingDimensions != null)
                {
                    // If we have existing dimensions for the current formula cell, then we are going to clear the corresponding matrix.
                    int rows = existingDimensions[0];
                    int cols = existingDimensions[1];
                    SetRangeValue(rows, cols, new object[rows, cols]);
                }
                
                SetRangeValue(rowMatrix.Matrix.GetLength(0), rowMatrix.Matrix.GetLength(1), rowMatrix.Matrix);
                SetRowValue(rowMatrix.Row);
            });
        }

        private int[] getExistingDimensions(Excel.Worksheet worksheet, String propertyKey)
        {
            string value = propertyUtil.GetCustomProp(worksheet, propertyKey);
            if (!string.IsNullOrEmpty(value))
            {
                return value.Split(PROPERTY_DIVIDER).Select(x => int.Parse(x)).ToArray();
            }
            return null;
        }

        private void SetRowValue(object[] value)
        {
            Excel.Application app = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Worksheet worksheet = (Excel.Worksheet)app.ActiveWorkbook.ActiveSheet;
            Excel.Range startCell = app.ActiveCell;

            if (startCell.HasFormula)
            {
                // If the current cell has a formula, then we have to move our write operation +1.
                // (this is the case if we use the function wizard)
                startCell = (Excel.Range)worksheet.Cells[startCell.Row + 1, startCell.Column];
            }

            startCell = (Excel.Range)worksheet.Cells[startCell.Row - 1, startCell.Column + 1];

            Excel.Range endCell = (Excel.Range)worksheet.Cells[startCell.Row, startCell.Column + value.Length - 1];
            var writeRange = worksheet.Range[startCell, endCell];
            writeRange.Value2 = value;
        }

        private void SetRangeValue(int rows, int cols, object[,] value)
        {
            Excel.Application app = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Worksheet worksheet = (Excel.Worksheet)app.ActiveWorkbook.ActiveSheet;
            Excel.Range startCell = app.ActiveCell;

            if (startCell.HasFormula)
            {
                // If the current cell has a formula, then we have to move our write operation +1.
                // (this is the case if we use the function wizard)
                startCell = (Excel.Range)worksheet.Cells[startCell.Row + 1, startCell.Column];
            }

            Excel.Range endCell = (Excel.Range)worksheet.Cells[startCell.Row + rows + - 1, startCell.Column + cols - 1];
            var writeRange = worksheet.Range[startCell, endCell];
            writeRange.Value2 = value;
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
