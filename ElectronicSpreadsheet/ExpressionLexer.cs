using System;
using System.Collections.Generic;
using System.Text;

namespace ElectronicSpreadsheet
{
    public enum TokenType
    {
        Number,
        Plus,
        Minus,
        Multiply,
        Divide,
        LeftParen,
        RightParen,
        Equals,
        Less,
        Greater,
        LessOrEqual,
        GreaterOrEqual,
        NotEqual,
        And,
        Or,
        Not,
        Max,
        Min,
        CellReference,
        Comma,
        End,
        True,
        False
    }

    public class Token
    {
        public TokenType Type { get; set; }
        public string Value { get; set; }
        public int Position { get; set; }

        public Token(TokenType type, string value, int position)
        {
            Type = type;
            Value = value;
            Position = position;
        }
    }

    public class ExpressionLexer
    {
        private string _input;
        private int _position;

        public ExpressionLexer(string input)
        {
            _input = input ?? "";
            _position = 0;
        }

        public List<Token> Tokenize()
        {
            var tokens = new List<Token>();

            while (_position < _input.Length)
            {
                SkipWhitespace();

                if (_position >= _input.Length)
                    break;

                char current = _input[_position];

                // Числа
                if (char.IsDigit(current))
                {
                    tokens.Add(ReadNumber());
                }
                // Ідентифікатори та ключові слова
                else if (char.IsLetter(current))
                {
                    tokens.Add(ReadIdentifier());
                }
                // Оператори
                else
                {
                    Token token = ReadOperator();
                    if (token != null)
                        tokens.Add(token);
                    else
                        throw new Exception($"Невідомий символ '{current}' на позиції {_position}");
                }
            }

            tokens.Add(new Token(TokenType.End, "", _position));
            return tokens;
        }

        private void SkipWhitespace()
        {
            while (_position < _input.Length && char.IsWhiteSpace(_input[_position]))
            {
                _position++;
            }
        }

        private Token ReadNumber()
        {
            int start = _position;
            StringBuilder number = new StringBuilder();

            while (_position < _input.Length && char.IsDigit(_input[_position]))
            {
                number.Append(_input[_position]);
                _position++;
            }

            return new Token(TokenType.Number, number.ToString(), start);
        }

        private Token ReadIdentifier()
        {
            int start = _position;
            StringBuilder identifier = new StringBuilder();

            while (_position < _input.Length &&
                   (char.IsLetterOrDigit(_input[_position]) || _input[_position] == '_'))
            {
                identifier.Append(_input[_position]);
                _position++;
            }

            string value = identifier.ToString().ToLower();

            // Перевірка ключових слів
            switch (value)
            {
                case "and":
                    return new Token(TokenType.And, value, start);
                case "or":
                    return new Token(TokenType.Or, value, start);
                case "not":
                    return new Token(TokenType.Not, value, start);
                case "max":
                    return new Token(TokenType.Max, value, start);
                case "min":
                    return new Token(TokenType.Min, value, start);
                case "true":
                    return new Token(TokenType.True, value, start);
                case "false":
                    return new Token(TokenType.False, value, start);
                default:
                    // Перевірка, чи це посилання на клітинку
                    if (IsCellReference(identifier.ToString()))
                        return new Token(TokenType.CellReference, identifier.ToString(), start);

                    throw new Exception($"Невідомий ідентифікатор '{identifier}' на позиції {start}");
            }
        }

        private bool IsCellReference(string text)
        {
            if (string.IsNullOrEmpty(text))
                return false;

            int i = 0;
            // Перевіряємо літери
            while (i < text.Length && char.IsLetter(text[i]))
                i++;

            if (i == 0 || i == text.Length)
                return false;

            // Перевіряємо цифри
            while (i < text.Length && char.IsDigit(text[i]))
                i++;

            return i == text.Length;
        }

        private Token ReadOperator()
        {
            int start = _position;
            char current = _input[_position];

            switch (current)
            {
                case '+':
                    _position++;
                    return new Token(TokenType.Plus, "+", start);

                case '-':
                    _position++;
                    return new Token(TokenType.Minus, "-", start);

                case '*':
                    _position++;
                    return new Token(TokenType.Multiply, "*", start);

                case '/':
                    _position++;
                    return new Token(TokenType.Divide, "/", start);

                case '(':
                    _position++;
                    return new Token(TokenType.LeftParen, "(", start);

                case ')':
                    _position++;
                    return new Token(TokenType.RightParen, ")", start);

                case ',':
                    _position++;
                    return new Token(TokenType.Comma, ",", start);

                case '=':
                    _position++;
                    return new Token(TokenType.Equals, "=", start);

                case '<':
                    _position++;
                    if (_position < _input.Length)
                    {
                        if (_input[_position] == '=')
                        {
                            _position++;
                            return new Token(TokenType.LessOrEqual, "<=", start);
                        }
                        else if (_input[_position] == '>')
                        {
                            _position++;
                            return new Token(TokenType.NotEqual, "<>", start);
                        }
                    }
                    return new Token(TokenType.Less, "<", start);

                case '>':
                    _position++;
                    if (_position < _input.Length && _input[_position] == '=')
                    {
                        _position++;
                        return new Token(TokenType.GreaterOrEqual, ">=", start);
                    }
                    return new Token(TokenType.Greater, ">", start);

                default:
                    return null;
            }
        }
    }
}