/// <include file='Docs/UserInput.xml' path='Docs/members[@name="userinput"]/UserInput/*'/>
class UserInput
{
    /// <include file='Docs/UserInput.xml' path='Docs/members[@name="userinput"]/StringInput/*'/>
    public static string StringInput(string prompt = "Введите строку: ")
    {
        string? userInput;

        do
        {
            Console.Write(prompt);
            userInput = Console.ReadLine();

            if (string.IsNullOrEmpty(userInput))
            {
                Console.WriteLine("Строка не должна быть пустой!");
            }
            else
            {
                return userInput;
            }
        } while (string.IsNullOrEmpty(userInput));

        return userInput;
    }

    /// <include file='Docs/UserInput.xml' path='Docs/members[@name="userinput"]/IntInput/*'/>
    public static int IntInput(bool isPositive = false, string prompt = "Введите целое число: ")
    {
        string? userInput;
        bool isValidInput = false;
        int result = 0;

        while (!isValidInput)
        {
            Console.Write(prompt);
            userInput = Console.ReadLine();

            if (int.TryParse(userInput, out result))
            {
                if (isPositive && result < 0)
                {
                    Console.WriteLine("Число должно быть положительным!");
                }
                else
                {
                    isValidInput = true;
                }
            }
            else
            {
                Console.WriteLine("Введенное значение не является целым числом!");
            }
        }
        return result;
    }

    /// <include file='Docs/UserInput.xml' path='Docs/members[@name="userinput"]/DateTimeInput/*'/>
    public static DateTime DateTimeInput(string prompt = "Введите дату: ")
    {
        DateTime result;
        bool isValid;

        do
        {
            Console.Write(prompt);
            string? userInput = Console.ReadLine();

            isValid = DateTime.TryParse(userInput, out result);

            if (!isValid)
            {
                Console.WriteLine("Введенное значение не является датой!");
            }
        } while (!isValid);

        return result;
    }
}