using System;
using System.Runtime.InteropServices;
using static System.Net.Mime.MediaTypeNames;

class UserInput
{
    public static string StringInput(string text = "Введите строку")
    {
        string? user_input = "";

        bool b = false;

        while (b != true)
        {
            Console.Write(text);
            user_input = Console.ReadLine();
            if (user_input == "" || user_input == null) 
            {
                Console.WriteLine("Строка не дожна быть пустой!");
                b = false;
            }
            else
            {
                b = true;
            }

        }
        return user_input;
    }

    // Для ввода целых чисел
    public static int intInput(bool isPositive = false, string text = "Введите целое число: ")
    {
        string user_input = "";
        bool b = false;

        while (b != true)
        {
            Console.Write(text);
            user_input = Console.ReadLine();
            b = int.TryParse(user_input, out int result2);
            if (b && isPositive)
            {
                int i = int.Parse(user_input);
                if (i < 0)
                {
                    b = false;
                    Console.WriteLine("Число должно быть положительным!");
                }
            }
            else if (!b) { Console.WriteLine("Введенное значение не является целым числом!"); }
        }
        return int.Parse(user_input);
    }

    // Для ввода дробных чисел
    public static double doubleInput(bool isPositive = false, string text = "Введите целое число: ")
    {
        string user_input = "";
        bool b = false;

        while (b != true)
        {
            Console.Write(text);
            user_input = Console.ReadLine();
            b = double.TryParse(user_input, out double result2);
            if (b && isPositive)
            {
                double i = double.Parse(user_input);
                if (i < 0)
                {
                    b = false;
                    Console.WriteLine("Число должно быть положительным!");
                }
            }
            else if (!b) { Console.WriteLine("Введенное значение не является целым числом!"); }
        }
        return double.Parse(user_input);
    }

    public static DateTime DateTimeInput(string text = "Введите дату: ")
    {
        string? user_input = "";
        bool b = false;

        while (b != true)
        {
            Console.Write(text);
            user_input = Console.ReadLine();

            b = DateTime.TryParse(user_input, out DateTime result2);
            if (!b) { Console.WriteLine("Введенное значение не является датой!"); }
        }
        return DateTime.Parse(user_input);
    }

    // Для ввода 1 символа
    public static char charInput()
    {
        Console.Write("Введите символ: ");
        string user_input = Console.ReadLine();
        while (user_input.Length > 1 || user_input.Length == 0)
        {
            Console.Write("Введите 1 символ!: ");
            user_input = Console.ReadLine();
        }
        return user_input[0];

    }

    // Для ввода массива заданной длинны и случайными числами
    public static int[] randomArrInput()
    {
        int arr_size = intInput(true, "Введите размер массива: ");
        int[] arr = new int[arr_size];
        Random rand = new Random();
        for (int i = 0; i < arr_size; i++)
        {
            arr[i] = rand.Next(-arr_size, arr_size);
        }

        return arr;
    }

    // Для ввода массива
    public static double[] ArrInput(int size, string text = "Введите элементы массива через пробел: ")
    {
        Console.Write(text);
        string s = Console.ReadLine();

        string[] split_s = s.Split(' ');
        double[] arr_d = new double[size];


        bool b = true;

        while (b)
        {
            b = false;
            if (split_s.Length != size)
            {
                Console.WriteLine($"Количество элементов в строке должно быть равно {size}");
                b = true;
            }

            for (int i = 0; i < size; i++)
            {
                if (double.TryParse(split_s[i], out double result))
                {
                    arr_d[i] = result;
                }
                else
                {
                    Console.WriteLine("В строке должны быть только числа типа double или пробелы");
                    b = true;
                }
            }
        }


        return arr_d;
    }


}