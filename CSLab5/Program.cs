namespace CSLab5
{
    class Program
    {
        public static void Main()
        {
            Console.WriteLine("Протокол дописывать в старый файл?");
            Console.Write("1 - да: ");
            string? input = Console.ReadLine();
            string protocolFile = "protocol.txt";
            if (input != "1" || !File.Exists(protocolFile)) 
            {
                bool allGood = false;
                if (!File.Exists(protocolFile))
                    Console.WriteLine("Стандартного файла не существует.");

                while (!allGood)
                {
                    protocolFile = UserInput.StringInput("Введите название нового файла: ");

                    if (!protocolFile.EndsWith(".txt"))
                    {
                        Console.WriteLine("Файл должен иметь расширение txt!");
                        continue;
                    }
                   
                    try
                    {
                        File.Create(protocolFile).Close();
                        allGood = true;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Произошла ошибка: {ex}");
                    }        
                }
            }
            Console.Clear();
            Protocoler protocoler = new(protocolFile);

            protocoler.WriteLine("Программа запущенна");
            protocoler.Save();
        
            DataBase dataBase = new("LR5-var3.xls");
        
            bool userContinue = true;
            while (userContinue) 
            {
                Console.WriteLine("Главное меню");
                Console.WriteLine($"Текущий файл: {dataBase.File}");
                Console.WriteLine("Функции:");
                Console.WriteLine("1 - Чтение базы данных из excel файла");
                Console.WriteLine("2 - Просмотр базы данных");
                Console.WriteLine("3 - Удаление элемента по ключу");
                Console.WriteLine("4 - Корректировка элемента по ключу");
                Console.WriteLine("5 - Добавление элемента");

                Console.WriteLine("6 - Запрос для получения суммарной выручки" +
                    " за указанный период от указанного экспоната");

                Console.WriteLine("7 - Запрос для получение суммарной выручки" +
                    " от экспонатов казанной эры, за указанный промежуток времени");

                Console.WriteLine(
                    "8 - Запрос на полчение информации о песетителях, посетивших указанный экспонат, " +
                    "из указанного города, за указанный промежуток времени"
                    );

                Console.WriteLine("9 - Запрос на получение информации о посетителях указанного возраста," +
                    " посетивших указанный экспонат");

                Console.WriteLine("10 - Выход из программы");
                Console.WriteLine("11 - Очистить консоль");
                Console.WriteLine("12 - Сохранение базы данных");
                Console.WriteLine();

                Console.Write("Введите номер функции: ");

                switch (Console.ReadLine())
                {
                    case "1":
                        Console.WriteLine("функция 1");

                        protocoler.WriteLine("Вызвана функция 1");
                        protocoler.Save();

                        string file = UserInput.StringInput("Введите путь к файлу: ");
                        try
                        {
                            dataBase = new DataBase(file);
                            Console.WriteLine("База данных успешно прочитана");
                        }
                        catch (Exception ex) 
                        {
                            Console.WriteLine(ex.Message);

                            protocoler.WriteLine($"В функции 1 произошда ошибка: {ex.Message}");
                            protocoler.Save();

                        }
                        Console.WriteLine();
                        break;

                    case "2":
                        Console.WriteLine("База данных: ");
                        Console.WriteLine(dataBase);

                        protocoler.WriteLine("Вызвана функция 2");
                        protocoler.Save();
                        break;

                    case "3":
                        Console.WriteLine("Функция 3");
                        Console.WriteLine();

                        protocoler.WriteLine("Вызвана функция 3");
                        protocoler.Save();

                        int table = UserInput.IntInput(true, "Введите id таблицы: ");

                        int id = UserInput.IntInput(true, "Введите id элемента: ");

                        try
                        {
                            dataBase.DeleteObjectById(table, id);
                            Console.WriteLine("Элемент успешно удален из базы данных");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);

                            protocoler.WriteLine($"В функции 3 произошда ошибка: {ex.Message}");
                            protocoler.Save();

                        }
                        Console.WriteLine();
                        break;
                    
                    case "4":
                        Console.WriteLine("Функция 4");
                        Console.WriteLine();

                        protocoler.WriteLine("Вызвана функция 4");
                        protocoler.Save();

                        table = UserInput.IntInput(true, "Введите id таблицы: ");

                        id = UserInput.IntInput(true, "Введите id элемента: ");

                        string attr = UserInput.StringInput("Введите название атрибута: ");

                        string val = UserInput.StringInput("Введите новое значение атрибута: ");

                        try
                        {
                            dataBase.UpdateObjectbyId(table, id, attr, val);
                            Console.WriteLine("Элемент успешно изменен в базе данных");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);

                            protocoler.WriteLine($"В функции 4 произошда ошибка: {ex.Message}");
                            protocoler.Save();

                        }
                        Console.WriteLine();
                        break;

                    case "5":
                        Console.WriteLine("Функция 5");
                        Console.WriteLine();

                        protocoler.WriteLine("Вызвана функция 5");
                        protocoler.Save();

                        Console.WriteLine("1 - экспонат");
                        Console.WriteLine("2 - посетителя");
                        Console.WriteLine("3 - билет");

                        Console.Write("Выберите, что хотите добавить: ");
                        string? s = Console.ReadLine();

                        switch(s)
                        {
                            case "1":
                                Console.WriteLine("Добаление экспоната");

                                string name = UserInput.StringInput("Введите название экспоната: ");

                                string era1 = UserInput.StringInput("Введите эпозху: ");

                                try
                                {
                                    dataBase.AddExhibit(name, era1);
                                    Console.WriteLine("Экспонат успешно добавлен");
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"{ex.Message}");
                                    protocoler.WriteLine($"В функции 5 произошда ошибка: {ex.Message}");
                                    protocoler.Save();
                                }
                            
                                break;

                            case "2":
                                Console.WriteLine("Добаление пометителя");

                                name = UserInput.StringInput("Введите полное имя посетителя: ");
                                int age1 = UserInput.IntInput(true, "Введите возраст: ");
                                string city1 = UserInput.StringInput("Введите город: "); 
                            
                                try 
                                { 
                                    dataBase.AddVisitor(name, age1, city1);
                                    Console.WriteLine("Посетитель успешно добавлен");
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"{ex.Message}");

                                    protocoler.WriteLine($"В функции 5 произошда ошибка: {ex.Message}");
                                    protocoler.Save();
                                }
                                break;

                            case "3":
                                Console.WriteLine("Добаление билета");

                                int id1 = UserInput.IntInput(true, "Введите id экспоната: ");
                                int id2 = UserInput.IntInput(true, "Введите id посетителя: ");
                                int price = UserInput.IntInput(true, "Введите цену билета: ");

                                try
                                {
                                    dataBase.AddTicket(id1, id2, DateTime.Now, price);
                                    Console.WriteLine("Билет успешно добавлен");

                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"{ex.Message}");

                                    protocoler.WriteLine($"В функции 5 произошда ошибка: {ex.Message}");
                                    protocoler.Save();
                                }
                                break;

                            default:
                                Console.WriteLine("Введенное значение не является параметром!");
                                break;
                        }

                        Console.WriteLine();
                        break;

                    case "6":
                        Console.WriteLine("Функция 6");
                        Console.WriteLine();

                        protocoler.WriteLine("Вызвана функция 6");
                        protocoler.Save();

                        int idExhibit = UserInput.IntInput(true, "Введите id экспоната: ");

                        DateTime begin = UserInput.DateTimeInput("Введите дату 1: ");
                        DateTime end = UserInput.DateTimeInput("Введите дату 2: ");

                        int a = dataBase.Request1(idExhibit, begin, end);
                        Console.WriteLine($"Суммарная выручка = {a}");

                        Console.WriteLine();
                        break;

                    case "7":
                        Console.WriteLine("Функция 7");
                        Console.WriteLine();

                        protocoler.WriteLine("Вызвана функция 7");
                        protocoler.Save();

                        string era = UserInput.StringInput("Введите название эпохи: ");

                        begin = UserInput.DateTimeInput("Введите дату 1: ");
                        end = UserInput.DateTimeInput("Введите дату 2: ");

                        int s1 = dataBase.Request2(era, begin, end);
                        Console.WriteLine($"Суммарная выручка = {s1}");

                        Console.WriteLine();
                        break;

                    case "8":
                        Console.WriteLine("Функция 8");
                        Console.WriteLine();

                        protocoler.WriteLine("Вызвана функция 8");
                        protocoler.Save();

                        idExhibit = UserInput.IntInput(true, "Введите id экспоната: ");

                        string city = UserInput.StringInput("Введите название города: ");

                        begin = UserInput.DateTimeInput("Введите дату 1: ");
                        end = UserInput.DateTimeInput("Введите дату 2: ");

                        var visitors = dataBase.Request3(idExhibit, city, begin, end);

                        Console.WriteLine("Результат: ");
                        foreach ( var visitor in visitors)
                        {
                            Console.WriteLine(visitor);
                        }

                        Console.WriteLine();
                        break;

                    case "9":
                        Console.WriteLine("Функция 9");
                        Console.WriteLine();

                        protocoler.WriteLine("Вызвана функция 9");
                        protocoler.Save();

                        int age = UserInput.IntInput(true, "Введите возраст: ");

                        era = UserInput.StringInput("Введите эпоху: ");

                        var k = dataBase.Request4(age, era);

                        Console.WriteLine("Результат: ");
                        foreach (var visitor in k)
                        {
                            Console.WriteLine(visitor);
                        }

                        Console.WriteLine();
                        break;

                    case "10":
                        protocoler.WriteLine("Программа завершена");
                        protocoler.Close();
                        return;

                    case "11":
                        Console.Clear();

                        protocoler.WriteLine("Консоль очищена");
                        protocoler.Save();
                        break;

                    case "12":
                        dataBase.Save();
                        break;

                    default:
                        Console.WriteLine("Введенное значение не является номером функции!");
                        Console.WriteLine();
                        break;
                }
            }
        }
    }
}
