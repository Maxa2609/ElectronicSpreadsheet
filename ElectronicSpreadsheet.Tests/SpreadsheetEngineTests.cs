using Microsoft.VisualStudio.TestTools.UnitTesting;
using ElectronicSpreadsheet; 

namespace ElectronicSpreadsheet.Tests
{
    [TestClass]
    public class SpreadsheetEngineTests
    {
        private SpreadsheetEngine _engine;

        [TestInitialize]
        public void Setup()
        {
            
            _engine = new SpreadsheetEngine(10, 10);
        }

        [TestMethod]
        public void SetCellExpression_LogicalOperationWithDependency_UpdatesCorrectly()
        {
            _engine.SetCellExpression(0, 0, "10"); // A1
            _engine.SetCellExpression(0, 1, "5");  // B1
            _engine.SetCellExpression(0, 2, "A1 > B1"); // C1

           
            var cellC1_initial = _engine.GetCell(0, 2);

           
            Assert.AreEqual("True", cellC1_initial.DisplayValue, "C1 повинна бути 'True' (10 > 5)");
            Assert.IsFalse(cellC1_initial.HasError, "Не повинно бути помилки");

            
            _engine.SetCellExpression(0, 1, "20");

            
            var cellC1_updated = _engine.GetCell(0, 2);
            Assert.AreEqual("False", cellC1_updated.DisplayValue, "C1 повинна оновитися на 'False' (10 > 20)");
            Assert.IsFalse(cellC1_updated.HasError, "Не повинно бути помилки");
        }

      
        [TestMethod]
        public void SetCellExpression_DivisionByZero_SetsErrorState()
        {
            
            _engine.SetCellExpression(0, 0, "100");   // A1
            _engine.SetCellExpression(0, 1, "0");     // B1
            _engine.SetCellExpression(0, 2, "A1 / B1"); // C1

            
            var cellC1 = _engine.GetCell(0, 2);

            
            Assert.IsTrue(cellC1.HasError, "Клітинка C1 повинна мати прапор помилки");
            Assert.AreEqual("#ПОМИЛКА", cellC1.DisplayValue, "Повинно відображатися '#ПОМИЛКА'");
            
            Assert.AreEqual("Ділення на нуль", cellC1.ErrorMessage, "Повідомлення про помилку має бути 'Ділення на нуль'");
        }

       
        [TestMethod]
        public void SetCellExpression_CyclicDependency_SetsErrorState()
        {

            _engine.SetCellExpression(0, 0, "B1"); // A1
            var cellA1_initial = _engine.GetCell(0, 0);
            Assert.AreEqual("0", cellA1_initial.DisplayValue, "A1 спочатку = 0");

 
            _engine.SetCellExpression(0, 1, "A1"); // B1

   
            var cellA1_final = _engine.GetCell(0, 0);
            var cellB1_final = _engine.GetCell(0, 1);

            Assert.IsTrue(cellA1_final.HasError, "A1 повинна мати помилку циклу");
            Assert.AreEqual("#ЦИКЛ!", cellA1_final.DisplayValue, "A1 повинна відображати '#ЦИКЛ!'");

            Assert.IsTrue(cellB1_final.HasError, "B1 повинна мати помилку циклу");
            Assert.AreEqual("#ЦИКЛ!", cellB1_final.DisplayValue, "B1 повинна відображати '#ЦИКЛ!'");
        }
    }
}