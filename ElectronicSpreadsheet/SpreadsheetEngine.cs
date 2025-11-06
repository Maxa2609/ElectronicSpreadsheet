using System;
using System.Collections.Generic;
using System.Linq;

namespace ElectronicSpreadsheet
{
    public class Cell
    {
        public string Expression { get; set; } = "";
        public string DisplayValue { get; set; } = "";
        public bool HasError { get; set; } = false;
        public string ErrorMessage { get; set; } = "";
        public object ComputedValue { get; set; }
    }

    public class SpreadsheetEngine
    {
        private Cell[,] _cells;
        private int _rows;
        private int _cols;
        private HashSet<string> _evaluatingCells = new HashSet<string>();

        public SpreadsheetEngine(int rows, int cols)
        {
            _rows = rows;
            _cols = cols;
            _cells = new Cell[rows, cols];

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    _cells[i, j] = new Cell();
                }
            }
        }

        public void Resize(int newRows, int newCols)
        {
            var oldCells = _cells;
            var oldRows = _rows;
            var oldCols = _cols;

            _rows = newRows;
            _cols = newCols;
            _cells = new Cell[newRows, newCols];


            for (int i = 0; i < newRows; i++)
            {
                for (int j = 0; j < newCols; j++)
                {
                    _cells[i, j] = new Cell();
                }
            }

            
            int copyRows = Math.Min(oldRows, newRows);
            int copyCols = Math.Min(oldCols, newCols);

            for (int i = 0; i < copyRows; i++)
            {
                for (int j = 0; j < copyCols; j++)
                {
                    _cells[i, j].Expression = oldCells[i, j].Expression;
                    _cells[i, j].DisplayValue = oldCells[i, j].DisplayValue;
                    _cells[i, j].HasError = oldCells[i, j].HasError;
                    _cells[i, j].ErrorMessage = oldCells[i, j].ErrorMessage;
                    _cells[i, j].ComputedValue = oldCells[i, j].ComputedValue;
                }
            }

            
            for (int i = 0; i < newRows; i++)
            {
                for (int j = 0; j < newCols; j++)
                {
                    if (!string.IsNullOrWhiteSpace(_cells[i, j].Expression))
                    {
                        EvaluateCell(i, j);
                    }
                }
            }
        }

        public Cell GetCell(int row, int col)
        {
            if (row < 0 || row >= _rows || col < 0 || col >= _cols)
                return new Cell { DisplayValue = "#REF!", HasError = true };

            return _cells[row, col];
        }

        public void SetCellExpression(int row, int col, string expression)
        {
            if (row < 0 || row >= _rows || col < 0 || col >= _cols)
                return;

            var cell = _cells[row, col];
            cell.Expression = expression ?? "";

            
            for (int i = 0; i < _rows; i++)
            {
                for (int j = 0; j < _cols; j++)
                {
                    _cells[i, j].ComputedValue = null;
                }
            }

            
            for (int i = 0; i < _rows; i++)
            {
                for (int j = 0; j < _cols; j++)
                {
                    EvaluateCell(i, j);
                }
            }
        }

        private void EvaluateCell(int row, int col)
        {
            var cell = _cells[row, col];
            var cellRef = GetCellReference(row, col);

            if (string.IsNullOrWhiteSpace(cell.Expression))
            {
                cell.DisplayValue = "";
                cell.HasError = false;
                cell.ErrorMessage = "";
                
                return;
            }

            
            if (_evaluatingCells.Contains(cellRef))
            {
                cell.DisplayValue = "#ЦИКЛ!";
                cell.HasError = true;
                cell.ErrorMessage = "Циклічне посилання виявлено";
                cell.ComputedValue = null;
                return;
            }

            try
            {
                _evaluatingCells.Add(cellRef);

                var parser = new ExpressionParser(this, row, col);
                var result = parser.Parse(cell.Expression);

                cell.ComputedValue = result;
                cell.DisplayValue = result?.ToString() ?? "null";
                cell.HasError = false;
                cell.ErrorMessage = "";
            }
            catch (Exception ex)
            {
                cell.HasError = true;
                cell.ErrorMessage = ex.Message;
                cell.ComputedValue = null;

                
                if (ex.Message.Contains("Циклічне посилання виявлено"))
                {
                    cell.DisplayValue = "#ЦИКЛ!";
                }
                else
                {
                    cell.DisplayValue = "#ПОМИЛКА";
                }
            }
            finally
            {
                _evaluatingCells.Remove(cellRef);
            }
        }

        public object GetCellValue(int row, int col)
        {
            if (row < 0 || row >= _rows || col < 0 || col >= _cols)
                throw new Exception("Посилання за межами таблиці");

            var cell = _cells[row, col];

            
            if (cell.HasError)
                throw new Exception($"Клітинка {GetCellReference(row, col)} містить помилку: {cell.ErrorMessage}");

            
            if (cell.ComputedValue == null && !string.IsNullOrWhiteSpace(cell.Expression))
            {
                EvaluateCell(row, col);
            }

            
            if (cell.HasError)
                throw new Exception($"Клітинка {GetCellReference(row, col)} містить помилку: {cell.ErrorMessage}");

            return cell.ComputedValue ?? 0;
        }

        public static string GetCellReference(int row, int col)
        {
            string colName = "";
            int c = col;
            while (c >= 0)
            {
                colName = (char)('A' + (c % 26)) + colName;
                c = c / 26 - 1;
            }
            return colName + (row + 1);
        }

        public static bool ParseCellReference(string reference, out int row, out int col)
        {
            row = -1;
            col = -1;

            if (string.IsNullOrWhiteSpace(reference))
                return false;

            int i = 0;
            string colPart = "";

            
            while (i < reference.Length && char.IsLetter(reference[i]))
            {
                colPart += char.ToUpper(reference[i]);
                i++;
            }

            if (string.IsNullOrEmpty(colPart) || i >= reference.Length)
                return false;

            
            string rowPart = reference.Substring(i);

            if (!int.TryParse(rowPart, out row))
                return false;

            row--; 

        
            col = 0;
            foreach (char c in colPart)
            {
                col = col * 26 + (c - 'A' + 1);
            }
            col--; 

            return true;
        }
    }
}