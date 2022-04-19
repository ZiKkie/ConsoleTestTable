//Программа вычисления значений таблицы в консольном приложении
using System.Data;
using System.Text.RegularExpressions;
using static System.Console;
class Program
{
    //Тестовые значения
    //const string TestDimension = "3\t4";
    //const string TestValues = "12\t=C2\t3\t'Sample\t=A1+B1*C1/5\t=A2*B1\t=B3-C3\t'Spread\t'Test\t=4-3\t5\t'Sheet";
    //const string TestValues = "12\t=C\t3\t'Sampl\t=A1+B1*C1/5\t=A2*B1\t=B3-C3\t'Spread\t'Test\t=4-3\t5\t'Sheet";
    //const string TestValues = "12	=C2	3	'Sample	=A1+B1*C1/5	=A2*B1	=B3-C3	'Spread	'Test	=4-3	5	'Sheet";
    //const string TestDimension = "2\t2";
    //const string TestValues = "asd\t=a 1\t'sd5sfs\t4";
    //const string TestValues = "2\t=a1*+a2\ta1\t=1a";
    //const string TestValues = "=B1\t=A2\t=A1\t4";
    //const string TestValues = "=B1\t\t=A1\t4";
    //const string TestValues = "=a2\t\t'asd\t";
    //const string TestValues = "=-1\t-1\t=+1\t=2-1";
    //const string TestValues = "=b1+b1\t2\t\t";

    static void Main()
    {
        WriteLine("Введите размерность таблицы:");
        Table table;
        while (true)
        {
            string? dim = ReadLine(); //TestDimension; //Здесь можно подставлять тестовые данные
            var (row, col) = ParseDimension(dim);
            if (row != 0 && col != 0)
            {
                table = new(row, col);
                break;
            }
            else
                WriteLine("Неверный формат, повторно введите размерность:");
        }
        WriteLine("Введите данные ячеек таблицы:");
        while (true)
        {
            string? val = ReadLine(); //TestValues; //Здесь можно подставлять тестовые данные
            if (table.SetValues(val))
                break;
            else
                WriteLine("Неверный формат, повторно введите данные ячеек:");
        }
        table.Process();
        table.Print();
        ReadLine();
    }
    //Валидация размерности таблицы
    static (int row, int col) ParseDimension(string? input)
    {
        if (!string.IsNullOrEmpty(input))
        {
            string[]? words = input.Split('\t');
            if (words != null && words.Length == 2)
            {
                int.TryParse(words[0], out int row);
                int.TryParse(words[1], out int col);
                if (Enumerable.Range(1, 9).Contains(row) && Enumerable.Range(1, 26).Contains(col))
                    return (row, col);
            }
        }
        return (0, 0);
    }
}

class Table
{
    public Table(int rows, int cols)
    {
        _rows = rows;
        _cols = cols;
        _values = new Cell[_rows, _cols];
    }
    private readonly int _rows;
    private readonly int _cols;
    private Cell[,] _values;

    class Cell
    {
        public Cell(string value)
        {
            Value = value;
        }
        public string? Value; //Значение
        public bool Evaluated; //Защита от повторного вызова
    }
    //Внесение данных
    public bool SetValues(string? input)
    {
        if (!string.IsNullOrEmpty(input))
        {
            string[]? words = input.Split('\t');
            if (words != null && words.Length == _cols * _rows)
            {
                int k = 0;
                for (int i = 0; i < _rows; i++)
                {
                    for (int j = 0; j < _cols; j++)
                    {
                        _values[i, j] = new Cell(words[k]);
                        k++;
                    }
                }
                return true;
            }
        }
        return false;
    }
    //Обработка данных
    public void Process()
    {
        for (int i = 0; i < _rows; i++)
        {
            for (int j = 0; j < _cols; j++)
            {
                _values[i, j].Value = ComputeCell(i, j);
            }
        }
    }
    //Рекурсивная функция
    private string ComputeCell(int row, int col)
    {
        ref string? cell = ref _values[row, col].Value;

        if (_values[row, col].Evaluated)
            return cell ?? string.Empty;

        //Проверка на простое число
        if (int.TryParse(cell, out int value))
            return value.ToString();
        else if (!string.IsNullOrEmpty(cell))
        {
            //Обработка первого знака строки
            char first = cell[0];
            if (first.Equals('\''))
            {
                cell = cell.Substring(1);
            }
            else if (first.Equals('='))
            {
                cell = cell.Substring(1);
                //Удаление пробелов и разбиение строки на элементы
                cell = Regex.Replace(cell, @"\s+", "");
                string[] terms = cell.Split('+', '-', '*', '/');
                //Обработка элементов строки
                foreach (string term in terms)
                {
                    //Обработка буквенной ссылки
                    if (!string.IsNullOrEmpty(term) && !int.TryParse(term, out _))
                    {
                        //Валидация идентификатора
                        int iRow;
                        int iCol;
                        if (term.Length == 2 && char.IsLetter(term[0]) && char.IsDigit(term[1]))
                        {
                            iRow = (int)char.GetNumericValue(term[1]);
                            iCol = char.ToUpper(term[0]) - 64;
                        }
                        else
                            return cell = "#Неверный формат идентификатора или выражения";

                        //Рекурсивный вызов
                        string sVal = ComputeCell(iRow - 1, iCol - 1);

                        //Допускается присваивание одиночной или пустой строки, если она не содержит ошибок
                        if ((!int.TryParse(sVal, out _) && terms.Length != 1)
                            || (!string.IsNullOrEmpty(sVal) && sVal[0].Equals('#')))
                                return cell = "#Арифметическая операция со строкой или полем с ошибкой";
                        //Замена идентификатора на значение
                        cell = cell.Replace(term, sVal);
                    }
                }
                //Вычисление ячейки
                try
                {
                    if (terms.Length != 1)
                    {
                        DataTable dt = new DataTable();
                        var result = dt.Compute(cell, "");
                        //Округление на случай дробного числа и проверка на null результата вычисления
                        if (!(result is DBNull))
                            cell = Convert.ToInt32(result).ToString();
                    }
                }catch
                {
                    cell = "#Ошибка при вычислении";
                }
            }
            else if (!first.Equals('#'))
            {
                cell = "#Неверный формат строки";
            }
        }
        //Установка флага и возвращение объекта
        _values[row, col].Evaluated = true;
        return cell ?? string.Empty;
    }
    //Печать таблицы
    public void Print()
    {
        for (int i = 0; i < _rows; i++)
        {
            for (int j = 0; j < _cols; j++)
            {
                Write(_values[i, j].Value + "\t");
            }
            WriteLine();
        }
    }
}