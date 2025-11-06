// Це має бути у вашому новому проекті ElectronicSpreadsheet.Tests
// Не забудьте додати Project Reference на ваш основний проект

using Microsoft.VisualStudio.TestTools.UnitTesting;
using ElectronicSpreadsheet; // Це простір імен з ваших файлів

namespace ElectronicSpreadsheet.Tests
{
    [TestClass]
    public class SpreadsheetEngineTests
    {
        private SpreadsheetEngine _engine;

        // Цей метод [TestInitialize] запускається перед КОЖНИМ тестом,
        // гарантуючи, що ми завжди працюємо з "чистою" таблицею.
        [TestInitialize]
        public void Setup()
        {
            // Створюємо нове поле 10x10 для тестів
            _engine = new SpreadsheetEngine(10, 10);
        }

        /// <summary>
        /// Тест 1: Перевіряє логічні операції та оновлення залежних клітинок.
        /// Сценарій:
        /// 1. A1 = 10, B1 = 5
        /// 2. C1 = A1 > B1 (очікуємо "True")
        /// 3. Змінюємо B1 = 20
        /// 4. Перевіряємо, що C1 автоматично оновилася на "False".
        /// </summary>
        [TestMethod]
        public void SetCellExpression_LogicalOperationWithDependency_UpdatesCorrectly()
        {
            // Arrange (Підготовка)
            _engine.SetCellExpression(0, 0, "10"); // A1
            _engine.SetCellExpression(0, 1, "5");  // B1
            _engine.SetCellExpression(0, 2, "A1 > B1"); // C1

            // Act (Дія 1)
            var cellC1_initial = _engine.GetCell(0, 2);

            // Assert (Перевірка 1)
            Assert.AreEqual("True", cellC1_initial.DisplayValue, "C1 повинна бути 'True' (10 > 5)");
            Assert.IsFalse(cellC1_initial.HasError, "Не повинно бути помилки");

            // Act (Дія 2)
            // Змінюємо залежність. B1 стає 20.
            // Ваш двигун автоматично перераховує всю таблицю.
            _engine.SetCellExpression(0, 1, "20");

            // Assert (Перевірка 2)
            var cellC1_updated = _engine.GetCell(0, 2);
            Assert.AreEqual("False", cellC1_updated.DisplayValue, "C1 повинна оновитися на 'False' (10 > 20)");
            Assert.IsFalse(cellC1_updated.HasError, "Не повинно бути помилки");
        }

        /// <summary>
        /// Тест 2: Перевіряє обробку помилки ділення на нуль.
        /// Сценарій:
        /// 1. A1 = 100
        /// 2. B1 = 0
        /// 3. C1 = A1 / B1
        /// 4. Перевіряємо, що C1 має помилку.
        /// </summary>
        [TestMethod]
        public void SetCellExpression_DivisionByZero_SetsErrorState()
        {
            // Arrange
            _engine.SetCellExpression(0, 0, "100");   // A1
            _engine.SetCellExpression(0, 1, "0");     // B1
            _engine.SetCellExpression(0, 2, "A1 / B1"); // C1

            // Act
            var cellC1 = _engine.GetCell(0, 2);

            // Assert
            Assert.IsTrue(cellC1.HasError, "Клітинка C1 повинна мати прапор помилки");
            Assert.AreEqual("#ПОМИЛКА", cellC1.DisplayValue, "Повинно відображатися '#ПОМИЛКА'");
            // Ми також перевіряємо, що повідомлення про помилку з парсера було коректно передано
            Assert.AreEqual("Ділення на нуль", cellC1.ErrorMessage, "Повідомлення про помилку має бути 'Ділення на нуль'");
        }

        /// <summary>
        /// Тест 3: Перевіряє виявлення циклічного посилання (A1 -> B1 -> A1).
        /// Сценарій:
        /// 1. Встановлюємо A1 = B1
        /// 2. Встановлюємо B1 = A1 (це створює цикл)
        /// 3. Перевіряємо, що обидві клітинки мають помилку циклу.
        /// </summary>
        [TestMethod]
        public void SetCellExpression_CyclicDependency_SetsErrorState()
        {
            // Arrange
            // A1 посилається на B1. На цьому етапі все добре, B1 = 0.
            _engine.SetCellExpression(0, 0, "B1"); // A1
            var cellA1_initial = _engine.GetCell(0, 0);
            Assert.AreEqual("0", cellA1_initial.DisplayValue, "A1 спочатку = 0");

            // Act
            // B1 посилається на A1. Ваш EvaluateCell повинен це виявити.
            _engine.SetCellExpression(0, 1, "A1"); // B1

            // Assert
            // Оскільки ваш двигун перераховує все, обидві клітинки повинні оновитися до стану помилки.
            var cellA1_final = _engine.GetCell(0, 0);
            var cellB1_final = _engine.GetCell(0, 1);

            Assert.IsTrue(cellA1_final.HasError, "A1 повинна мати помилку циклу");
            Assert.AreEqual("#ЦИКЛ!", cellA1_final.DisplayValue, "A1 повинна відображати '#ЦИКЛ!'");

            Assert.IsTrue(cellB1_final.HasError, "B1 повинна мати помилку циклу");
            Assert.AreEqual("#ЦИКЛ!", cellB1_final.DisplayValue, "B1 повинна відображати '#ЦИКЛ!'");
        }
    }
}