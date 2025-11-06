using System;
using System.Collections.Generic;

namespace ElectronicSpreadsheet
{
    public class ExpressionParser
    {
        private List<Token> _tokens;
        private int _position;
        private SpreadsheetEngine _engine;
        private int _currentRow;
        private int _currentCol;

        public ExpressionParser(SpreadsheetEngine engine, int row, int col)
        {
            _engine = engine;
            _currentRow = row;
            _currentCol = col;
        }

        public object Parse(string expression)
        {
            var lexer = new ExpressionLexer(expression);
            _tokens = lexer.Tokenize();
            _position = 0;

            var result = ParseOrExpression();

            if (_tokens[_position].Type != TokenType.End)
            {
                throw new Exception($"Неочікуваний токен '{_tokens[_position].Value}' на позиції {_tokens[_position].Position}");
            }

            return result;
        }


        private object ParseOrExpression()
        {
            var left = ParseAndExpression();

            while (_tokens[_position].Type == TokenType.Or)
            {
                _position++;
                var right = ParseAndExpression();
                left = ToBool(left) || ToBool(right);
            }

            return left;
        }

        private object ParseAndExpression()
        {
            var left = ParseNotExpression();

            while (_tokens[_position].Type == TokenType.And)
            {
                _position++;
                var right = ParseNotExpression();
                left = ToBool(left) && ToBool(right);
            }

            return left;
        }

        private object ParseNotExpression()
        {
            if (_tokens[_position].Type == TokenType.Not)
            {
                _position++;
                var operand = ParseNotExpression();
                return !ToBool(operand);
            }

            return ParseComparisonExpression();
        }

        private object ParseComparisonExpression()
        {
            var left = ParseAdditiveExpression();

            if (_tokens[_position].Type == TokenType.Equals ||
                _tokens[_position].Type == TokenType.Less ||
                _tokens[_position].Type == TokenType.Greater ||
                _tokens[_position].Type == TokenType.LessOrEqual ||
                _tokens[_position].Type == TokenType.GreaterOrEqual ||
                _tokens[_position].Type == TokenType.NotEqual)
            {
                var opType = _tokens[_position].Type;
                _position++;
                var right = ParseAdditiveExpression();

                double leftNum = ToNumber(left);
                double rightNum = ToNumber(right);

                switch (opType)
                {
                    case TokenType.Equals:
                        return Math.Abs(leftNum - rightNum) < 1e-10;
                    case TokenType.Less:
                        return leftNum < rightNum;
                    case TokenType.Greater:
                        return leftNum > rightNum;
                    case TokenType.LessOrEqual:
                        return leftNum <= rightNum;
                    case TokenType.GreaterOrEqual:
                        return leftNum >= rightNum;
                    case TokenType.NotEqual:
                        return Math.Abs(leftNum - rightNum) >= 1e-10;
                }
            }

            return left;
        }

        private object ParseAdditiveExpression()
        {
            var left = ParseMultiplicativeExpression();

            while (_tokens[_position].Type == TokenType.Plus ||
                   _tokens[_position].Type == TokenType.Minus)
            {
                var opType = _tokens[_position].Type;
                _position++;
                var right = ParseMultiplicativeExpression();

                double leftNum = ToNumber(left);
                double rightNum = ToNumber(right);

                if (opType == TokenType.Plus)
                    left = leftNum + rightNum;
                else
                    left = leftNum - rightNum;
            }

            return left;
        }

        private object ParseMultiplicativeExpression()
        {
            var left = ParseUnaryExpression();

            while (_tokens[_position].Type == TokenType.Multiply ||
                   _tokens[_position].Type == TokenType.Divide)
            {
                var opType = _tokens[_position].Type;
                _position++;
                var right = ParseUnaryExpression();

                double leftNum = ToNumber(left);
                double rightNum = ToNumber(right);

                if (opType == TokenType.Multiply)
                    left = leftNum * rightNum;
                else
                {
                    if (Math.Abs(rightNum) < 1e-10)
                        throw new Exception("Ділення на нуль");
                    left = leftNum / rightNum;
                }
            }

            return left;
        }

        private object ParseUnaryExpression()
        {
            if (_tokens[_position].Type == TokenType.Plus)
            {
                _position++;
                return ParseUnaryExpression();
            }

            if (_tokens[_position].Type == TokenType.Minus)
            {
                _position++;
                var operand = ParseUnaryExpression();
                return -ToNumber(operand);
            }

            return ParsePrimaryExpression();
        }

        private object ParsePrimaryExpression()
        {
            var token = _tokens[_position];

            if (token.Type == TokenType.Number)
            {
                _position++;
                return double.Parse(token.Value);
            }

            if (token.Type == TokenType.True)
            {
                _position++;
                return true;
            }

            if (token.Type == TokenType.False)
            {
                _position++;
                return false;
            }

            if (token.Type == TokenType.CellReference)
            {
                _position++;
                if (SpreadsheetEngine.ParseCellReference(token.Value, out int row, out int col))
                {
                    return _engine.GetCellValue(row, col);
                }
                throw new Exception($"Некоректне посилання на клітинку '{token.Value}'");
            }

            if (token.Type == TokenType.Max || token.Type == TokenType.Min)
            {
                return ParseFunction();
            }

            if (token.Type == TokenType.LeftParen)
            {
                _position++;
                var result = ParseOrExpression();

                if (_tokens[_position].Type != TokenType.RightParen)
                    throw new Exception($"Очікувалася закриваюча дужка на позиції {_tokens[_position].Position}");

                _position++;
                return result;
            }

            throw new Exception($"Неочікуваний токен '{token.Value}' на позиції {token.Position}");
        }

        private object ParseFunction()
        {
            var funcToken = _tokens[_position];
            _position++;

            if (_tokens[_position].Type != TokenType.LeftParen)
                throw new Exception($"Очікувалася відкриваюча дужка після '{funcToken.Value}'");

            _position++;

            var args = new List<object>();


            args.Add(ParseOrExpression());


            while (_tokens[_position].Type == TokenType.Comma)
            {
                _position++;
                args.Add(ParseOrExpression());
            }

            if (_tokens[_position].Type != TokenType.RightParen)
                throw new Exception($"Очікувалася закриваюча дужка для функції '{funcToken.Value}'");

            _position++;


            if (funcToken.Type == TokenType.Max)
            {
                if (args.Count < 2)
                    throw new Exception("Функція max потребує принаймні 2 аргументи");

                double maxVal = ToNumber(args[0]);
                for (int i = 1; i < args.Count; i++)
                {
                    double val = ToNumber(args[i]);
                    if (val > maxVal)
                        maxVal = val;
                }
                return maxVal;
            }
            else if (funcToken.Type == TokenType.Min)
            {
                if (args.Count < 2)
                    throw new Exception("Функція min потребує принаймні 2 аргументи");

                double minVal = ToNumber(args[0]);
                for (int i = 1; i < args.Count; i++)
                {
                    double val = ToNumber(args[i]);
                    if (val < minVal)
                        minVal = val;
                }
                return minVal;
            }

            throw new Exception($"Невідома функція '{funcToken.Value}'");
        }

        private double ToNumber(object value)
        {
            if (value is double d)
                return d;
            if (value is int i)
                return i;
            if (value is bool b)
                return b ? 1 : 0;
            if (value == null)
                return 0;

            if (double.TryParse(value.ToString(), out double result))
                return result;

            throw new Exception($"Неможливо перетворити '{value}' на число");
        }

        private bool ToBool(object value)
        {
            if (value is bool b)
                return b;
            if (value is double d)
                return Math.Abs(d) >= 1e-10;
            if (value is int i)
                return i != 0;
            if (value == null)
                return false;

            if (bool.TryParse(value.ToString(), out bool result))
                return result;

            return false;
        }
    }
}