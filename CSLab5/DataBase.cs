using Aspose.Cells;
using System;

namespace CSLab5
{
    class DataBase
    {
        private string _file;

        private List<Exhibit> _exhibits;
        private List<Visitor> _visitors;
        private List<Ticket> _tickets;

        public string File
        {
            get { return _file; }
        }

        public DataBase(string file)
        {
            _file = file;

            if (!System.IO.File.Exists(file))
                throw new Exception("Файла с заданным путем не существет!");

            if (!file.EndsWith(".xls"))
                throw new Exception("Тип файла должен быть xls!");

            Workbook wb = new(file);

            _exhibits = [];

            Worksheet ws = wb.Worksheets[0];

            for (int i = 1; i <= ws.Cells.MaxDataRow; i++)
            {
                Row row = ws.Cells.Rows[i];

                _exhibits.Add(new Exhibit(
                    row[0].IntValue,
                    row[1].StringValue,
                    row[2].StringValue
                    ));

            }

            _visitors = [];

            ws = wb.Worksheets[1];

            for (int i = 1; i <= ws.Cells.MaxDataRow; i++)
            {
                Row row = ws.Cells.Rows[i];

                _visitors.Add(new Visitor(
                    row[0].IntValue,
                    row[1].StringValue,
                    row[2].IntValue,
                    row[3].StringValue
                    ));

            }

            _tickets = [];

            ws = wb.Worksheets[2];

            for (int i = 1; i <= ws.Cells.MaxDataRow; i++)
            {
                Row row = ws.Cells.Rows[i];

                _tickets.Add(new Ticket(
                    row[0].IntValue,
                    row[1].IntValue,
                    row[2].IntValue,
                    row[3].DateTimeValue,
                    row[4].IntValue
                    ));

            }
        }

        public void DelExhibitById(int id)
        {
            var tic = from Ticket t in _tickets
                      where t.IdExhibit == id
                      select t;

            foreach (var t in tic)
            {
                this.DelTickettById(t.Id);
            }

            var exs = from Exhibit e in _exhibits
                      where id == e.Id
                      select e;
            if (!exs.Any())
                throw new Exception($"экспоната с id = {id} не существеут");

            _exhibits.Remove(_exhibits.First());

            exs = from Exhibit e in _exhibits
                  where id < e.Id
                  select e;

            foreach (Exhibit e in exs)
            {
                e.Id--;
            }

        }

        public void DelVisitorById(int id)
        {
            var tic = from Ticket t in _tickets
                      where t.IdVisitor == id
                      select t;

            foreach (var t in tic)
            {
                this.DelTickettById(t.Id);
            }

            var exs = from Visitor e in _visitors
                      where id == e.Id
                      select e;
            if (!exs.Any())
                throw new Exception($"посетителя с id = {id} не существеут");

            _exhibits.Remove(_exhibits.First());

            exs = from Visitor e in _visitors
                  where id < e.Id
                  select e;

            foreach (Visitor e in exs)
            {
                e.Id--;
            }

        }

        public void DelTickettById(int id)
        {
            var exs = from Ticket e in _tickets
                      where id == e.Id
                      select e;
            if (!exs.Any())
                throw new Exception($"билета с id = {id} не существеут");

            _exhibits.Remove(_exhibits.First());

            exs = from Ticket e in _tickets
                  where id < e.Id
                  select e;

            foreach (Ticket e in exs)
            {
                e.Id--;
            }

        }

        public void DelElById(int idTable, int id)
        {
            if (idTable < 0 || idTable > 2)
                throw new Exception($"таблицы с id '{idTable}' не существеут");

            switch (id)
            {
                case 0:
                    DelExhibitById(id);
                    break;
                case 1:
                    DelVisitorById(id);
                    break;
                case 2:
                    DelTickettById(id);
                    break;
            }

            Save();
        }

        public void UpdateExhibitById(int id, string AttributeName, string newValue)
        {
            var exs = from Exhibit e in _exhibits
                      where id == e.Id
                      select e;
            if (!exs.Any())
                throw new Exception($"экспоната с id = {id} не существеут");

            Exhibit ex = exs.First();

            switch (AttributeName)
            {
                case "name":
                    ex.Name = newValue;
                    break;
                case "era":
                    ex.Era = newValue;
                    break;
                default:
                    throw new Exception($"атрибут с названием = {AttributeName} не существеут");
            }

        }

        public void UpdateVisitorById(int id, string AttributeName, string newValue)
        {
            var exs = from Visitor e in _visitors
                      where id == e.Id
                      select e;
            if (!exs.Any())
                throw new Exception($"посетителя с id = {id} не существеут");

            Visitor ex = exs.First();

            switch (AttributeName)
            {
                case "name":
                    ex.Name = newValue;
                    break;
                case "age":
                    ex.Age = Convert.ToInt32(newValue);
                    break;
                case "city":
                    ex.City = newValue;
                    break;
                default:
                    throw new Exception($"атрибут с названием = {AttributeName} не существеут");
            }

        }

        public void UpdateTicketById(int id, string AttributeName, string newValue)
        {
            var exs = from Ticket e in _tickets
                      where id == e.Id
                      select e;
            if (!exs.Any())
                throw new Exception($"экспоната с id = {id} не существеут");

            Ticket ex = exs.First();

            switch (AttributeName)
            {
                case "idExhibit":
                    ex.IdExhibit = Convert.ToInt32(newValue);
                    break;
                case "idVisitor":
                    ex.IdVisitor = Convert.ToInt32(newValue);
                    break;
                case "price":
                    ex.Price = Convert.ToInt32(newValue);
                    break;
                case "time":
                    ex.Time = Convert.ToDateTime(newValue);
                    break;
                default:
                    throw new Exception($"атрибут с названием = {AttributeName} не существеут");
            }

        }

        public void UpdateElbyId(int idTable, int id, string attributeName, string newValue)
        {
            if (idTable < 0 || idTable > 2)
                throw new Exception($"таблицы с id '{idTable}' не существеут");

            switch (idTable)
            {
                case 0:
                    UpdateExhibitById(id, attributeName, newValue);
                    break;
                case 1:
                    UpdateVisitorById(id, attributeName, newValue);
                    break;
                case 2:
                    UpdateTicketById(id, attributeName, newValue);
                    break;
            }

            Save();
        }

        public void AddExhibit(string name, string era)
        {
            _exhibits.Add(new Exhibit(
                _exhibits[-1].Id + 1,
                name,
                era
                ));
            Save();
        }

        public void AddVisitor(string name, int age, string city)
        {
            _visitors.Add(new Visitor(
                _visitors[-1].Id + 1,
                name,
                age,
                city
                ));
            Save();

        }

        public void AddTicket(int idExhibit, int idVisitor, DateTime time, int price)
        {
            _tickets.Add(new Ticket(
                _tickets[-1].Id + 1,
                idExhibit,
                idVisitor,
                time,
                price
                ));
            Save();

        }

        public int Request1(int idExhibit, DateTime? begin = null, DateTime? end = null)
        {
            // Запрос для получения суммарной выручки за данный период от одного экспоната
            if (begin == null)
                begin = new DateTime(1970, 1, 1);
            if (end == null)
                end = DateTime.Today;

            var prices = from Ticket r in _tickets
                         where
                            r.IdExhibit == idExhibit
                            && r.Time >= begin
                            && r.Time <= end
                         select r.Price;

            return prices.Sum();
        }

        public int Request2(string era, DateTime? begin = null, DateTime? end = null)
        {
            // Запрос для получение суммарной выручки от экспонатов казанной эры, за указанный промежуток времени
            if (begin == null)
                begin = new DateTime(1970, 1, 1);
            if (end == null)
                end = DateTime.Today;

            var prices = from Ticket r1 in _tickets
                         join Exhibit r2 in _exhibits on r1.IdExhibit equals r2.Id
                         where
                            r2.Era == era
                            && r1.Time >= begin
                            && r1.Time <= end
                         select r1.Price;

            return prices.Sum();
        }

        public IEnumerable<object> Request3(int idExhibit, string city, DateTime? begin = null, DateTime? end = null)
        {
            // Запрос на полчение информации о песетителях, посетивших данный экспонат,
            // из указанного города
            // за указанный промежуток
            if (begin == null)
                begin = new DateTime(1970, 1, 1);
            if (end == null)
                end = DateTime.Today;

            var people = from Ticket rT in _tickets
                         join Exhibit rEx in _exhibits on rT.IdExhibit equals rEx.Id
                         join Visitor rV in _visitors on rT.IdVisitor equals rV.Id
                         where
                            rEx.Id == idExhibit
                            && rV.City == city
                            && rT.Time >= begin
                            && rT.Time <= end
                         select new
                         {
                             idTicket = rT.Id,
                             name = rV.Name,
                             age = rV.Age,
                             price = rT.Price
                         };
            return people;
        }

        public IEnumerable<object> Request4(int age, string era)
        {
            // запрос на получение id, имен, времени посещения экспонатов данной эпохи, посетителями данного возраста
            var people = from Ticket rT in _tickets
                         join Exhibit rEx in _exhibits on rT.IdExhibit equals rEx.Id
                         join Visitor rV in _visitors on rT.IdVisitor equals rV.Id
                         where
                            rV.Age == age
                            && rEx.Era == era
                         select new
                         {
                             name = rV.Name,
                             idTicket = rT.Id,
                             date = rT.Time
                         };
            return people;
        }

        public void Save()
        {

            Workbook wb = new Workbook(_file);

            Worksheet ws = wb.Worksheets[0];

            Style s1 = ws.Cells.Rows[1][0].GetDisplayStyle();
            Style s2 = ws.Cells.Rows[1][1].GetDisplayStyle();
            Style s3 = ws.Cells.Rows[1][2].GetDisplayStyle();

            ws.Cells.DeleteRows(1, ws.Cells.Rows.Count - 1);


            foreach (Exhibit e in _exhibits)
            {
                ws.Cells.InsertRow(ws.Cells.MaxDataRow + 1);

                ws.Cells.Rows[^1][0].SetStyle(s1);
                ws.Cells.Rows[^1][0].PutValue(e.Id);

                ws.Cells.Rows[^1][1].SetStyle(s2);
                ws.Cells.Rows[^1][1].PutValue(e.Name);

                ws.Cells.Rows[^1][2].SetStyle(s3);
                ws.Cells.Rows[^1][2].PutValue(e.Era);

            }

            ws = wb.Worksheets[1];
            s1 = ws.Cells.Rows[1][0].GetDisplayStyle();
            s2 = ws.Cells.Rows[1][1].GetDisplayStyle();
            s3 = ws.Cells.Rows[1][2].GetDisplayStyle();
            Style s4 = ws.Cells.Rows[1][3].GetDisplayStyle();
            ws.Cells.DeleteRows(1, ws.Cells.Rows.Count - 1);

            foreach (Visitor e in _visitors)
            {
                ws.Cells.InsertRow(ws.Cells.MaxDataRow + 1);

                ws.Cells.Rows[^1][0].SetStyle(s1);
                ws.Cells.Rows[^1][0].PutValue(e.Id);

                ws.Cells.Rows[^1][1].SetStyle(s2);
                ws.Cells.Rows[^1][1].PutValue(e.Name);

                ws.Cells.Rows[^1][2].SetStyle(s3);
                ws.Cells.Rows[^1][2].PutValue(e.Age);

                ws.Cells.Rows[^1][3].SetStyle(s4);
                ws.Cells.Rows[^1][3].PutValue(e.City);
            }

            ws = wb.Worksheets[2];

            s1 = ws.Cells.Rows[1][0].GetDisplayStyle();
            s2 = ws.Cells.Rows[1][1].GetDisplayStyle();
            s3 = ws.Cells.Rows[1][2].GetDisplayStyle();
            s4 = ws.Cells.Rows[1][3].GetDisplayStyle();
            Style s5 = ws.Cells.Rows[1][4].GetDisplayStyle();

            ws.Cells.DeleteRows(1, ws.Cells.Rows.Count - 1);

            foreach (Ticket e in _tickets)
            {
                ws.Cells.InsertRow(ws.Cells.MaxDataRow + 1);

                ws.Cells.Rows[^1][0].SetStyle(s1);
                ws.Cells.Rows[^1][0].PutValue(e.Id);

                ws.Cells.Rows[^1][1].SetStyle(s2);
                ws.Cells.Rows[^1][1].PutValue(e.IdExhibit);

                ws.Cells.Rows[^1][2].SetStyle(s3);
                ws.Cells.Rows[^1][2].PutValue(e.IdVisitor);

                ws.Cells.Rows[^1][3].SetStyle(s4);
                ws.Cells.Rows[^1][3].PutValue(e.Time);

                ws.Cells.Rows[^1][4].SetStyle(s5);
                ws.Cells.Rows[^1][4].PutValue(e.Price);
            }

            wb.Save("save" + _file);

        }

        public override string ToString()
        {
            string s = "";

            int[] maxLength = new int[3];

            var k = from Exhibit e in _exhibits
                    select e.Id.ToString().Length;
            maxLength[0] = k.Max();

            k = from Exhibit e in _exhibits
                select e.Name.Length;
            maxLength[1] = k.Max();

            k = from Exhibit e in _exhibits
                select e.Era.Length;
            maxLength[2] = k.Max();

            for (int j = 0; j < maxLength.Sum() + 3 * 2 + 3; j++)
                s += "-";
            s += "\n";

            s += "|";
            for (int j = 0; j < (maxLength.Sum() + 3 * 2 + 3 - "Экспонаты".Length) / 2; j++)
                s += " ";

            s += "Экспонаты";

            for (int j = 0; j < (maxLength.Sum() + 3 * 2 + 3 - "Экспонаты".Length) / 2 - 1; j++)
                s += " ";
            s += "|";
            s += "\n";

            for (int j = 0; j < maxLength.Sum() + 3 * 2 + 3; j++)
                s += "-";
            s += "\n";

            foreach (Exhibit e in _exhibits)
            {
                s += "|";
                string space = "";
                for (int l = 0; l < maxLength[0] - e.Id.ToString().Length; l++)
                    space += " ";
                s += $"{space}{e.Id} |";

                space = "";
                for (int l = 0; l < maxLength[1] - e.Name.Length; l++)
                    space += " ";
                s += $"{space}{e.Name} |";

                space = "";
                for (int l = 0; l < maxLength[2] - e.Era.Length; l++)
                    space += " ";
                s += $"{space}{e.Era} |";
                s += "\n";
            }
            for (int j = 0; j < maxLength.Sum() + 3 * 2 + 3; j++)
                s += "-";
            s += "\n";


            maxLength = new int[4];

            k = from Visitor e in _visitors
                select e.Id.ToString().Length;
            maxLength[0] = k.Max();

            k = from Visitor e in _visitors
                select e.Name.Length;
            maxLength[1] = k.Max();

            k = from Visitor e in _visitors
                select e.Age.ToString().Length;
            maxLength[2] = k.Max();

            k = from Visitor e in _visitors
                select e.City.Length;
            maxLength[3] = k.Max();

            for (int j = 0; j < maxLength.Sum() + 3 * 2 + 3; j++)
                s += "-";
            s += "\n";

            s += "|";
            for (int j = 0; j < (maxLength.Sum() + 3 * 2 + 3 - "Посетители".Length) / 2; j++)
                s += " ";

            s += "Посетители";

            for (int j = 0; j < (maxLength.Sum() + 3 * 2 + 3 - "Посетители".Length) / 2 - 1; j++)
                s += " ";
            s += "|";
            s += "\n";

            for (int j = 0; j < maxLength.Sum() + 3 * 2 + 3; j++)
                s += "-";
            s += "\n";

            foreach (Visitor e in _visitors)
            {
                s += "|";
                string space = "";
                for (int l = 0; l < maxLength[0] - e.Id.ToString().Length; l++)
                    space += " ";
                s += $"{space}{e.Id} |";

                space = "";
                for (int l = 0; l < maxLength[1] - e.Name.Length; l++)
                    space += " ";
                s += $"{space}{e.Name} |";

                space = "";
                for (int l = 0; l < maxLength[2] - e.Age.ToString().Length; l++)
                    space += " ";
                s += $"{space}{e.Age} |";

                space = "";
                for (int l = 0; l < maxLength[3] - e.City.Length; l++)
                    space += " ";
                s += $"{space}{e.City} |";
                s += "\n";
            }
            for (int j = 0; j < maxLength.Sum() + 3 * 2 + 3; j++)
                s += "-";
            s += "\n";


            maxLength = new int[5];

            k = from Ticket e in _tickets
                select e.Id.ToString().Length;
            maxLength[0] = k.Max();

            k = from Ticket e in _tickets
                select e.IdExhibit.ToString().Length;
            maxLength[1] = k.Max();

            k = from Ticket e in _tickets
                select e.IdVisitor.ToString().Length;
            maxLength[2] = k.Max();

            k = from Ticket e in _tickets
                select e.Time.ToShortDateString().Length;
            maxLength[3] = k.Max();

            k = from Ticket e in _tickets
                select e.Price.ToString().Length;
            maxLength[4] = k.Max();

            for (int j = 0; j < maxLength.Sum() + 3 * 2 + 3; j++)
                s += "-";
            s += "\n";

            s += "|";
            for (int j = 0; j < (maxLength.Sum() + 3 * 2 + 3 - "Билеты".Length) / 2; j++)
                s += " ";

            s += "Билеты";

            for (int j = 0; j < (maxLength.Sum() + 3 * 2 + 3 - "Билеты".Length) / 2 - 1; j++)
                s += " ";
            s += "|";
            s += "\n";

            for (int j = 0; j < maxLength.Sum() + 3 * 2 + 3; j++)
                s += "-";
            s += "\n";

            foreach (Ticket e in _tickets)
            {
                s += "|";
                string space = "";
                for (int l = 0; l < maxLength[0] - e.Id.ToString().Length; l++)
                    space += " ";
                s += $"{space}{e.Id} |";

                space = "";
                for (int l = 0; l < maxLength[1] - e.IdExhibit.ToString().Length; l++)
                    space += " ";
                s += $"{space}{e.IdExhibit} |";

                space = "";
                for (int l = 0; l < maxLength[2] - e.IdVisitor.ToString().Length; l++)
                    space += " ";
                s += $"{space}{e.IdVisitor} |";

                space = "";
                for (int l = 0; l < maxLength[3] - e.Time.ToShortDateString().Length; l++)
                    space += " ";
                s += $"{space}{e.Time.ToShortDateString()} |";

                space = "";
                for (int l = 0; l < maxLength[2] - e.Price.ToString().Length; l++)
                    space += " ";
                s += $"{space}{e.Price} |";
                s += "\n";
            }
            for (int j = 0; j < maxLength.Sum() + 3 * 2 + 3; j++)
                s += "-";
            s += "\n";


            return s;
        }
    }
}