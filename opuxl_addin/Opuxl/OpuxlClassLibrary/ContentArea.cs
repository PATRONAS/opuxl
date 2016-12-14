using ExcelDna.Integration;
using System;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace OpuxlClassLibrary
{
    class ContentArea
    {
        public object[] Row { get; set; }
        public object[,] Matrix { get; set; }
        public object CellValue { get; set; }

        private static char PROPERTY_DIVIDER = '&';
        private CustomPropertyUtil propertyUtil = new CustomPropertyUtil();

        public ContentArea(object[,] data)
        {
            // Cell Value is the [0,0] value of the matrix
            CellValue = data[0, 0];
            // We have to remove the first row of the matrix as we will handle it seperatly
            Matrix = RemoveRowFromMatrix(0, data);
            // And extract the first row without the first element
            Row = ExtractTailOfRow(data);
        }

        /// <summary>
        /// Async Task to create a matrix in the currently opened excel sheet.
        /// The Task has to be async such that we can write into adjacent cells as it is not allowed in the 'ui thread'.
        /// </summary>
        public object display()
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
            propertyUtil.SetCustomProp(worksheet, propKey, "" + Matrix.GetLength(0) + PROPERTY_DIVIDER + Matrix.GetLength(1));


            // Have to do the update within an async tasks, as UDFs are not allowed to manipulate other cells.
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                // If we have existing dimensions for the current formula cell, then we are going to clear the corresponding matrix.
                if (existingDimensions != null)
                {
                    int rows = existingDimensions[0];
                    int cols = existingDimensions[1];
                    // We clear the matrix below the formula row
                    SetRangeValue(rows, cols, new object[rows, cols]);
                    // And then the row behind the starting cell (without the starting cell itself)
                    SetRowValue(new object[cols - 1]);
                }

                // We set the Matrix part (the matrix directly beneath our formula row).
                SetRangeValue(Matrix.GetLength(0), Matrix.GetLength(1), Matrix);
                // And then the row part (the row directly to the right of our formula cell).
                SetRowValue(Row);
            });

            return CellValue;
        }

        /// <summary>
        /// Extracts the existing dimensions for an existing sheet property, if present.
        /// </summary>
        /// <param name="worksheet">The Excel worksheet</param>
        /// <param name="propertyKey">The key of the property</param>
        /// <returns>The dimensions, or null.</returns>
        private int[] getExistingDimensions(Excel.Worksheet worksheet, String propertyKey)
        {
            string value = propertyUtil.GetCustomProp(worksheet, propertyKey);
            if (!string.IsNullOrEmpty(value))
            {
                return value.Split(PROPERTY_DIVIDER).Select(x => int.Parse(x)).ToArray();
            }
            return null;
        }

        /// <summary>
        /// Sets the "row" part of the result matrix, which is the right part next to the starting cell, eg [1, 0] ... [n, 0].
        /// </summary>
        /// <param name="value">The row values</param>
        private void SetRowValue(object[] value)
        {
            Excel.Application app = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Worksheet worksheet = (Excel.Worksheet)app.ActiveWorkbook.ActiveSheet;
            Excel.Range startCell = app.ActiveCell;

            if (startCell.HasFormula)
            {
                // If the current cell has a formula, then we are in the formula wizard and need to move our
                // starting cell row by 1.
                startCell = (Excel.Range)worksheet.Cells[startCell.Row + 1, startCell.Column];
            }

            startCell = (Excel.Range)worksheet.Cells[startCell.Row - 1, startCell.Column + 1];

            Excel.Range endCell = (Excel.Range)worksheet.Cells[startCell.Row, startCell.Column + value.Length - 1];
            var writeRange = worksheet.Range[startCell, endCell];
            writeRange.Value2 = value;
        }

        /// <summary>
        /// Sets the "matrix" part of the result matrix, which is the the complete matrix starting at the second row,
        /// eg [0, 1] ... [n, m]
        /// </summary>
        /// <param name="rows">Count of rows</param>
        /// <param name="cols">Count of cols</param>
        /// <param name="value">The matrix values</param>
        private void SetRangeValue(int rows, int cols, object[,] value)
        {
            Excel.Application app = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Worksheet worksheet = (Excel.Worksheet)app.ActiveWorkbook.ActiveSheet;
            Excel.Range startCell = app.ActiveCell;

            if (startCell.HasFormula)
            {
                // If the current cell has a formula, then we are in the formula wizard and need to move our
                // starting cell row by 1.
                startCell = (Excel.Range)worksheet.Cells[startCell.Row + 1, startCell.Column];
            }

            Excel.Range endCell = (Excel.Range)worksheet.Cells[startCell.Row + rows + -1, startCell.Column + cols - 1];
            var writeRange = worksheet.Range[startCell, endCell];
            writeRange.Value2 = value;
        }

        private static object[] ExtractTailOfRow(object[,] matrix)
        {

            var columnsCount = matrix.GetLength(1);
            object[] result = new object[columnsCount - 1];

            for (int i = 1; i < columnsCount; i++)
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

    }
}
