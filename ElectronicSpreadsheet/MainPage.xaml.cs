using Microsoft.Maui.Controls;
using Microsoft.Maui.Storage;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace ElectronicSpreadsheet
{
    public partial class MainPage : ContentPage
    {
        private SpreadsheetEngine _engine;
        private bool _showValues = true;
        private int _rows = 10;
        private int _cols = 10;
        private string _tableName = "Нова таблиця";
        private Entry _lastFocusedEntry; // Зберігаємо останню активну клітинку

        private int _lastFocusedRow = -1;
        private int _lastFocusedCol = -1;

        public MainPage()
        {
            InitializeComponent();
            _engine = new SpreadsheetEngine(_rows, _cols);
            tableNameEntry.Text = _tableName;
            CreateSpreadsheet();
        }

        private void CreateSpreadsheet()
        {
            spreadsheetGrid.Children.Clear();
            spreadsheetGrid.RowDefinitions.Clear();
            spreadsheetGrid.ColumnDefinitions.Clear();

            // Створення стовпців
            spreadsheetGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = 50 });
            for (int col = 0; col < _cols; col++)
            {
                spreadsheetGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = 120 });
            }

            // Створення рядків
            spreadsheetGrid.RowDefinitions.Add(new RowDefinition { Height = 40 });
            for (int row = 0; row < _rows; row++)
            {
                spreadsheetGrid.RowDefinitions.Add(new RowDefinition { Height = 40 });
            }

            // Заголовки стовпців
            for (int col = 0; col < _cols; col++)
            {
                var header = new Label
                {
                    Text = GetColumnName(col),
                    BackgroundColor = Colors.LightGray,
                    HorizontalTextAlignment = TextAlignment.Center,
                    VerticalTextAlignment = TextAlignment.Center,
                    FontAttributes = FontAttributes.Bold,
                    TextColor = Colors.Black
                };
                spreadsheetGrid.Add(header, col + 1, 0);
            }

            // Заголовки рядків
            for (int row = 0; row < _rows; row++)
            {
                var header = new Label
                {
                    Text = (row + 1).ToString(),
                    BackgroundColor = Colors.LightGray,
                    HorizontalTextAlignment = TextAlignment.Center,
                    VerticalTextAlignment = TextAlignment.Center,
                    FontAttributes = FontAttributes.Bold,
                    TextColor = Colors.Black
                };
                spreadsheetGrid.Add(header, 0, row + 1);
            }

            // Клітинки
            for (int row = 0; row < _rows; row++)
            {
                for (int col = 0; col < _cols; col++)
                {
                    var cell = _engine.GetCell(row, col);
                    var entry = new Entry
                    {
                        Text = _showValues ? cell.DisplayValue : cell.Expression,
                        HorizontalTextAlignment = TextAlignment.Center,
                        BackgroundColor = Colors.White,
                        TextColor = Colors.Black,
                        IsReadOnly = false
                    };

                    int currentRow = row;
                    int currentCol = col;

                    // Обробка фокусу
                    entry.Focused += (s, e) =>
                    {
                        var e2 = (Entry)s;
                        _lastFocusedEntry = e2; // Зберігаємо посилання
                        _lastFocusedRow = currentRow;   // <-- ДОДАНО
                        _lastFocusedCol = currentCol; // <-- ДОДАНО

                        var currentCell = _engine.GetCell(currentRow, currentCol);

                        if (e2.Text == currentCell.DisplayValue)
                        {
                            e2.Text = currentCell.Expression;
                        }
                    };

                    // Обробка втрати фокусу
                    entry.Unfocused += (s, e) =>
                    {
                        var e2 = (Entry)s;
                        UpdateCell(currentRow, currentCol, e2.Text);
                    };

                    // Обробка Enter
                    entry.Completed += (s, e) =>
                    {
                        var e2 = (Entry)s;
                        UpdateCell(currentRow, currentCol, e2.Text);
                        e2.Unfocus();
                    };

                    spreadsheetGrid.Add(entry, col + 1, row + 1);
                }
            }
        }

        private void UpdateCell(int row, int col, string expression)
        {
            _engine.SetCellExpression(row, col, expression);
            RefreshDisplay();
        }

        private void RefreshDisplay()
        {
            for (int row = 0; row < _rows; row++)
            {
                for (int col = 0; col < _cols; col++)
                {
                    var cell = _engine.GetCell(row, col);
                    var entry = GetEntry(row, col);
                    if (entry != null && !entry.IsFocused)
                    {
                        entry.Text = _showValues ? cell.DisplayValue : cell.Expression;
                        entry.BackgroundColor = cell.HasError ? Colors.LightPink : Colors.White;
                    }
                }
            }
        }

        private Entry GetEntry(int row, int col)
        {
            foreach (var child in spreadsheetGrid.Children)
            {
                if (child is Entry entry)
                {
                    int r = Grid.GetRow(entry) - 1;
                    int c = Grid.GetColumn(entry) - 1;
                    if (r == row && c == col)
                        return entry;
                }
            }
            return null;
        }

        private string GetColumnName(int col)
        {
            string name = "";
            while (col >= 0)
            {
                name = (char)('A' + (col % 26)) + name;
                col = col / 26 - 1;
            }
            return name;
        }

        private void OnToggleMode(object sender, EventArgs e)
        {
            _showValues = !_showValues;
            toggleModeButton.Text = _showValues ? "Режим: ЗНАЧЕННЯ" : "Режим: ВИРАЗ";
            RefreshDisplay();
        }

        private void OnAddRow(object sender, EventArgs e)
        {
            _rows++;
            _engine.Resize(_rows, _cols);
            CreateSpreadsheet();
        }

        private void OnRemoveRow(object sender, EventArgs e)
        {
            if (_rows > 1)
            {
                _rows--;
                _engine.Resize(_rows, _cols);
                CreateSpreadsheet();
            }
        }

        private void OnAddColumn(object sender, EventArgs e)
        {
            _cols++;
            _engine.Resize(_rows, _cols);
            CreateSpreadsheet();
        }

        private void OnRemoveColumn(object sender, EventArgs e)
        {
            if (_cols > 1)
            {
                _cols--;
                _engine.Resize(_rows, _cols);
                CreateSpreadsheet();
            }
        }

        private async void OnClear(object sender, EventArgs e)
        {
            bool answer = await DisplayAlert("Підтвердження", "Ви точно хочете очистити таблицю?", "Так", "Ні");
            if (answer)
            {
                _engine = new SpreadsheetEngine(_rows, _cols);
                CreateSpreadsheet();
            }
        }

        private async void OnSave(object sender, EventArgs e)
        {
            try
            {
                var data = new SpreadsheetData
                {
                    Name = _tableName,
                    Rows = _rows,
                    Cols = _cols,
                    Cells = new List<CellData>()
                };

                for (int row = 0; row < _rows; row++)
                {
                    for (int col = 0; col < _cols; col++)
                    {
                        var cell = _engine.GetCell(row, col);
                        if (!string.IsNullOrWhiteSpace(cell.Expression))
                        {
                            data.Cells.Add(new CellData
                            {
                                Row = row,
                                Col = col,
                                Expression = cell.Expression
                            });
                        }
                    }
                }

                string json = JsonSerializer.Serialize(data, new JsonSerializerOptions { WriteIndented = true });
                string fileName = $"{_tableName}.json";

                var filePath = Path.Combine(FileSystem.AppDataDirectory, fileName);
                await File.WriteAllTextAsync(filePath, json);

                await DisplayAlert("Успіх", $"Таблицю '{_tableName}' збережено!", "OK");
            }
            catch (Exception ex)
            {
                await DisplayAlert("Помилка", $"Не вдалося зберегти: {ex.Message}", "OK");
            }
        }

        private async void OnOpen(object sender, EventArgs e)
        {
            try
            {
                var result = await FilePicker.PickAsync(new PickOptions
                {
                    PickerTitle = "Виберіть файл таблиці"
                });

                if (result != null)
                {
                    string json = await File.ReadAllTextAsync(result.FullPath);
                    var data = JsonSerializer.Deserialize<SpreadsheetData>(json);

                    if (data != null)
                    {
                        _tableName = data.Name;
                        _rows = data.Rows;
                        _cols = data.Cols;

                        tableNameEntry.Text = _tableName;

                        _engine = new SpreadsheetEngine(_rows, _cols);

                        foreach (var cellData in data.Cells)
                        {
                            _engine.SetCellExpression(cellData.Row, cellData.Col, cellData.Expression);
                        }

                        CreateSpreadsheet();

                        await DisplayAlert("Успіх", $"Таблицю '{_tableName}' відкрито!", "OK");
                    }
                }
            }
            catch (Exception ex)
            {
                await DisplayAlert("Помилка", $"Не вдалося відкрити: {ex.Message}", "OK");
            }
        }

        private void OnTableNameChanged(object sender, TextChangedEventArgs e)
        {
            _tableName = e.NewTextValue ?? "Нова таблиця";
        }

        private async void OnInsertOperation(object sender, EventArgs e)
        {
            try
            {
                var button = (Button)sender;
                string operation = button.Text;

                // Декодуємо HTML entities
                operation = operation.Replace("&lt;", "<")
                                      .Replace("&gt;", ">")
                                      .Replace("&amp;", "&");

                // 1. Перевіряємо, чи ми взагалі знаємо, яка клітинка активна
                if (_lastFocusedEntry == null || _lastFocusedRow == -1)
                {
                    await DisplayAlert("Підказка", "Спочатку клацніть на клітинку, в яку хочете вставити операцію", "OK");
                    return;
                }

                // 2. Отримуємо клітинку з двигуна
                var cell = _engine.GetCell(_lastFocusedRow, _lastFocusedCol);

                string textToInsertInto;
                int cursorPosition = _lastFocusedEntry.CursorPosition;

                // 3. КЛЮЧОВА ЛОГІКА: Визначаємо, який текст редагувати
                // Якщо текст у полі ("2") дорівнює DisplayValue ("2"), 
                // АЛЕ не дорівнює Expression ("1+1"),
                // це означає, що Unfocused щойно спрацював і перезаписав наш вираз.
                if (_lastFocusedEntry.Text == cell.DisplayValue && _lastFocusedEntry.Text != cell.Expression)
                {
                    // Беремо вираз ("1+1") замість значення ("2")
                    textToInsertInto = cell.Expression;
                    // І ставимо курсор в кінець виразу
                    cursorPosition = (textToInsertInto ?? "").Length;
                }
                else
                {
                    // Unfocused ще не спрацював, або вираз = значенню.
                    // Безпечно беремо текст, який зараз у полі.
                    textToInsertInto = _lastFocusedEntry.Text ?? "";
                }

                // 4. Вставляємо операцію
                _lastFocusedEntry.Text = textToInsertInto.Insert(cursorPosition, operation);
                _lastFocusedEntry.CursorPosition = cursorPosition + operation.Length;

                // 5. Повертаємо фокус на клітинку
                _lastFocusedEntry.Focus();
            }
            catch (Exception ex)
            {
                await DisplayAlert("Помилка", $"Не вдалося вставити операцію: {ex.Message}", "OK");
            }
        }

        private async void OnShowHelp(object sender, EventArgs e)
        {
            await DisplayAlert("Довідка",
                "Електронна таблиця з підтримкою логічних виразів\n\n" +
                "Підтримувані операції:\n" +
                "• Арифметичні: +, -, *, /\n" +
                "• Порівняння: =, <, >, <=, >=, <>\n" +
                "• Логічні: and, or, not\n" +
                "• Функції: max(x,y), min(x,y)\n\n" +
                "Посилання на клітинки: A1, B2, C3...\n\n" +
                "Приклади виразів:\n" +
                "• 5 > 3\n" +
                "• A1 + B1 > 10\n" +
                "• (A1 > 5) and (B1 < 10)\n" +
                "• max(A1, B1) = 100\n" +
                "• not(A1 > B1)\n\n" +
                "Використовуйте кнопки швидкого введення для додавання операцій.",
                "OK");
        }

        private async void OnShowAbout(object sender, EventArgs e)
        {
            await DisplayAlert("Про програму",
                "Електронна таблиця v1.0\n\n" +
                "Лабораторна робота\n" +
                "Варіант 54\n\n" +
                "Програма для роботи з електронними таблицями\n" +
                "з підтримкою логічних виразів та посилань на клітинки.\n\n" +
                "© 2025",
                "OK");
        }
    }

    public class SpreadsheetData
    {
        public string Name { get; set; }
        public int Rows { get; set; }
        public int Cols { get; set; }
        public List<CellData> Cells { get; set; }
    }

    public class CellData
    {
        public int Row { get; set; }
        public int Col { get; set; }
        public string Expression { get; set; }
    }
}