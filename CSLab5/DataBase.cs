﻿using Aspose.Cells;
using System;

namespace CSLab5
{
    /// <include file='Docs/DataBase.xml' 
    /// path='Docs/members[@name="database"]/DataBase/*'/>
    class DataBase
    {
        private readonly string _file;
        private readonly List<Exhibit> _exhibits;
        private readonly List<Visitor> _visitors;
        private readonly List<Ticket> _tickets;

        /// <include file='Docs/DataBase.xml' 
        /// path='Docs/members[@name="database"]/File/*'/>
        public string File
        {
            get { return _file; }
        }

        /// <include file='Docs/DataBase.xml' 
        /// path='Docs/members[@name="database"]/Constructor/*'/>
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

        /// <include file='Docs/DataBase.xml' 
        /// path='Docs/members[@name="database"]/DeleteExhibitById/*'/>
        public void DeleteExhibitById(int id)
        {
            var tickets = from Ticket t in _tickets
                      where t.IdExhibit == id
                      select t;

            foreach (var t in tickets)
            {
                DeleteTicketById(t.Id);
            }

            var exhibits = from Exhibit e in _exhibits
                      where id == e.Id
                      select e;
            if (!exhibits.Any())
                throw new Exception($"экспоната с id = {id} не существеут");

            _exhibits.Remove(_exhibits.First());

            exhibits = from Exhibit e in _exhibits
                  where id < e.Id
                  select e;

            foreach (Exhibit e in exhibits)
            {
                e.Id--;
            }

        }

        /// <include file='Docs/DataBase.xml' 
        /// path='Docs/members[@name="database"]/DeleteVisitorById/*'/>
        public void DeleteVisitorById(int id)
        {
            var tickets = from Ticket t in _tickets
                      where t.IdVisitor == id
                      select t;

            foreach (var t in tickets)
            {
                DeleteTicketById(t.Id);
            }

            var visitors = from Visitor e in _visitors
                      where id == e.Id
                      select e;
            if (!visitors.Any())
                throw new Exception($"посетителя с id = {id} не существеут");

            _exhibits.Remove(_exhibits.First());

            visitors = from Visitor e in _visitors
                  where id < e.Id
                  select e;

            foreach (Visitor e in visitors)
            {
                e.Id--;
            }

        }

        /// <include file='Docs/DataBase.xml' 
        /// path='Docs/members[@name="database"]/DeleteTickettById/*'/>
        public void DeleteTicketById(int id)
        {
            var tickets = from Ticket e in _tickets
                      where id == e.Id
                      select e;

            if (!tickets.Any())
                throw new Exception($"билета с id = {id} не существеут");

            _exhibits.Remove(_exhibits.First());

            tickets = from Ticket e in _tickets
                  where id < e.Id
                  select e;

            foreach (Ticket e in tickets)
            {
                e.Id--;
            }

        }

        /// <include file='Docs/DataBase.xml' 
        /// path='Docs/members[@name="database"]/DeleteObjectById/*'/>
        public void DeleteObjectById(int idTable, int id)
        {
            switch (id)
            {
                case 0:
                    DeleteExhibitById(id);
                    break;
                case 1:
                    DeleteVisitorById(id);
                    break;
                case 2:
                    DeleteTicketById(id);
                    break;
                default:
                    throw new Exception($"таблицы с id '{idTable}' не существеут");
            }

            Save();
        }

        /// <include file='Docs/DataBase.xml' 
        /// path='Docs/members[@name="database"]/UpdateExhibitById/*'/>
        public void UpdateExhibitById(int id, string attributeName, string newValue)
        {
            var exhibits = from Exhibit e in _exhibits
                      where id == e.Id
                      select e;
            
            if (!exhibits.Any())
                throw new Exception($"экспоната с id = {id} не существеут");

            Exhibit ex = exhibits.First();

            switch (attributeName)
            {
                case "name":
                    ex.Name = newValue;
                    break;
                case "era":
                    ex.Era = newValue;
                    break;
                default:
                    throw new Exception($"атрибут с названием = {attributeName} не существуют");
            }

        }

        /// <include file='Docs/DataBase.xml' 
        /// path='Docs/members[@name="database"]/UpdateVisitorById/*'/>
        public void UpdateVisitorById(int id, string attributeName, string newValue)
        {
            var visitors = from Visitor e in _visitors
                      where id == e.Id
                      select e;
            
            if (!visitors.Any())
                throw new Exception($"посетителя с id = {id} не существеут");

            Visitor ex = visitors.First();

            switch (attributeName)
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
                    throw new Exception($"атрибут с названием = {attributeName} не существеут");
            }

        }

        /// <include file='Docs/DataBase.xml' 
        /// path='Docs/members[@name="database"]/UpdateTicketById/*'/>
        public void UpdateTicketById(int id, string attributeName, string newValue)
        {
            var tickets = from Ticket e in _tickets
                      where id == e.Id
                      select e;

            if (!tickets.Any())
                throw new Exception($"экспоната с id = {id} не существеут");

            Ticket ex = tickets.First();

            switch (attributeName)
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
                    throw new Exception($"атрибут с названием = {attributeName} не существеут");
            }

        }

        /// <include file='Docs/DataBase.xml' 
        /// path='Docs/members[@name="database"]/UpdateObjectbyId/*'/>
        public void UpdateObjectbyId(int idTable, int id, 
            string attributeName, string newValue)
        {
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
                default:
                    throw new Exception($"таблицы с id '{idTable}' не существеут");
            }

            Save();
        }

        /// <include file='Docs/DataBase.xml' 
        /// path='Docs/members[@name="database"]/AddExhibit/*'/>
        public void AddExhibit(string name, string era)
        {
            _exhibits.Add(new Exhibit(
                _exhibits[-1].Id + 1,
                name,
                era
                ));

            Save();
        }

        /// <include file='Docs/DataBase.xml' 
        /// path='Docs/members[@name="database"]/AddVisitor/*'/>
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

        /// <include file='Docs/DataBase.xml' 
        /// path='Docs/members[@name="database"]/AddTicket/*'/>
        public void AddTicket(int idExhibit, int idVisitor, 
            DateTime time, int price)
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

        /// <include file='Docs/DataBase.xml' 
        /// path='Docs/members[@name="database"]/Request1/*'/>
        public int Request1(int idExhibit, 
            DateTime? begin = null, DateTime? end = null)
        {
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

        /// <include file='Docs/DataBase.xml' 
        /// path='Docs/members[@name="database"]/Request2/*'/>
        public int Request2(string era, 
            DateTime? begin = null, DateTime? end = null)
        {
            if (begin == null)
                begin = new DateTime(1970, 1, 1);
            if (end == null)
                end = DateTime.Today;

            var prices = from Ticket rT in _tickets
                         join Exhibit rEx in _exhibits on rT.IdExhibit equals rEx.Id
                         where
                            rEx.Era == era
                            && rT.Time >= begin
                            && rT.Time <= end
                         select rT.Price;

            return prices.Sum();
        }

        /// <include file='Docs/DataBase.xml' 
        /// path='Docs/members[@name="database"]/Request3/*'/>
        public IEnumerable<object> Request3(int idExhibit, string city, 
            DateTime? begin = null, DateTime? end = null)
        {
            if (begin == null)
                begin = new DateTime(1970, 1, 1);
            if (end == null)
                end = DateTime.Today;

            return from Ticket rT in _tickets
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
        }

        /// <include file='Docs/DataBase.xml' 
        /// path='Docs/members[@name="database"]/Request4/*'/>
        public IEnumerable<object> Request4(int age, string era)
        {
            return from Ticket rT in _tickets
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
        }

        /// <include file='Docs/DataBase.xml' 
        /// path='Docs/members[@name="database"]/Save/*'/>
        public void Save()
        {

            Workbook wb = new(_file);

            Worksheet ws = wb.Worksheets[0];

            Style fisrtColumnStyle = ws.Cells.Rows[1][0].GetDisplayStyle();
            Style secondColumnStyle = ws.Cells.Rows[1][1].GetDisplayStyle();
            Style thirdColumStyle = ws.Cells.Rows[1][2].GetDisplayStyle();

            ws.Cells.DeleteRows(1, ws.Cells.Rows.Count - 1);


            foreach (Exhibit e in _exhibits)
            {
                ws.Cells.InsertRow(ws.Cells.MaxDataRow + 1);

                ws.Cells.Rows[^1][0].SetStyle(fisrtColumnStyle);
                ws.Cells.Rows[^1][0].PutValue(e.Id);

                ws.Cells.Rows[^1][1].SetStyle(secondColumnStyle);
                ws.Cells.Rows[^1][1].PutValue(e.Name);

                ws.Cells.Rows[^1][2].SetStyle(thirdColumStyle);
                ws.Cells.Rows[^1][2].PutValue(e.Era);

            }

            ws = wb.Worksheets[1];
            fisrtColumnStyle = ws.Cells.Rows[1][0].GetDisplayStyle();
            secondColumnStyle = ws.Cells.Rows[1][1].GetDisplayStyle();
            thirdColumStyle = ws.Cells.Rows[1][2].GetDisplayStyle();
            Style fourthColumStyle = ws.Cells.Rows[1][3].GetDisplayStyle();
            ws.Cells.DeleteRows(1, ws.Cells.Rows.Count - 1);

            foreach (Visitor e in _visitors)
            {
                ws.Cells.InsertRow(ws.Cells.MaxDataRow + 1);

                ws.Cells.Rows[^1][0].SetStyle(fisrtColumnStyle);
                ws.Cells.Rows[^1][0].PutValue(e.Id);

                ws.Cells.Rows[^1][1].SetStyle(secondColumnStyle);
                ws.Cells.Rows[^1][1].PutValue(e.Name);

                ws.Cells.Rows[^1][2].SetStyle(thirdColumStyle);
                ws.Cells.Rows[^1][2].PutValue(e.Age);

                ws.Cells.Rows[^1][3].SetStyle(fourthColumStyle);
                ws.Cells.Rows[^1][3].PutValue(e.City);
            }

            ws = wb.Worksheets[2];

            fisrtColumnStyle = ws.Cells.Rows[1][0].GetDisplayStyle();
            secondColumnStyle = ws.Cells.Rows[1][1].GetDisplayStyle();
            thirdColumStyle = ws.Cells.Rows[1][2].GetDisplayStyle();
            fourthColumStyle = ws.Cells.Rows[1][3].GetDisplayStyle();
            Style fifthColumStyle = ws.Cells.Rows[1][4].GetDisplayStyle();

            ws.Cells.DeleteRows(1, ws.Cells.Rows.Count - 1);

            foreach (Ticket e in _tickets)
            {
                ws.Cells.InsertRow(ws.Cells.MaxDataRow + 1);

                ws.Cells.Rows[^1][0].SetStyle(fisrtColumnStyle);
                ws.Cells.Rows[^1][0].PutValue(e.Id);

                ws.Cells.Rows[^1][1].SetStyle(secondColumnStyle);
                ws.Cells.Rows[^1][1].PutValue(e.IdExhibit);

                ws.Cells.Rows[^1][2].SetStyle(thirdColumStyle);
                ws.Cells.Rows[^1][2].PutValue(e.IdVisitor);

                ws.Cells.Rows[^1][3].SetStyle(fourthColumStyle);
                ws.Cells.Rows[^1][3].PutValue(e.Time);

                ws.Cells.Rows[^1][4].SetStyle(fifthColumStyle);
                ws.Cells.Rows[^1][4].PutValue(e.Price);
            }

            wb.Save("save" + _file);

        }

        /// <include file='Docs/DataBase.xml' 
        /// path='Docs/members[@name="database"]/ToString/*'/>
        public override string ToString()
        {
            string output = "";

            int[] maxLength = new int[3];

            var lengths = from Exhibit e in _exhibits
                    select e.Id.ToString().Length;

            maxLength[0] = lengths.Max();

            lengths = from Exhibit e in _exhibits
                select e.Name.Length;
            maxLength[1] = lengths.Max();

            lengths = from Exhibit e in _exhibits
                select e.Era.Length;
            maxLength[2] = lengths.Max();

            for (int j = 0; j < maxLength.Sum() + 3 * 2 + 3; j++)
                output += "-";
            output += "\n";

            output += "|";
            for (int j = 0; j < (maxLength.Sum() + 3 * 2 + 3 - "Экспонаты".Length) / 2; j++)
                output += " ";

            output += "Экспонаты";

            for (int j = 0; j < (maxLength.Sum() + 3 * 2 + 3 - "Экспонаты".Length) / 2 - 1; j++)
                output += " ";
            output += "|";
            output += "\n";

            for (int j = 0; j < maxLength.Sum() + 3 * 2 + 3; j++)
                output += "-";
            output += "\n";

            foreach (Exhibit e in _exhibits)
            {
                output += "|";
                string space = "";
                for (int l = 0; l < maxLength[0] - e.Id.ToString().Length; l++)
                    space += " ";
                output += $"{space}{e.Id} |";

                space = "";
                for (int l = 0; l < maxLength[1] - e.Name.Length; l++)
                    space += " ";
                output += $"{space}{e.Name} |";

                space = "";
                for (int l = 0; l < maxLength[2] - e.Era.Length; l++)
                    space += " ";
                output += $"{space}{e.Era} |";
                output += "\n";
            }
            for (int j = 0; j < maxLength.Sum() + 3 * 2 + 3; j++)
                output += "-";
            output += "\n";


            maxLength = new int[4];

            lengths = from Visitor e in _visitors
                select e.Id.ToString().Length;
            maxLength[0] = lengths.Max();

            lengths = from Visitor e in _visitors
                select e.Name.Length;
            maxLength[1] = lengths.Max();

            lengths = from Visitor e in _visitors
                select e.Age.ToString().Length;
            maxLength[2] = lengths.Max();

            lengths = from Visitor e in _visitors
                select e.City.Length;
            maxLength[3] = lengths.Max();

            for (int j = 0; j < maxLength.Sum() + 3 * 2 + 3; j++)
                output += "-";
            output += "\n";

            output += "|";
            for (int j = 0; j < (maxLength.Sum() + 3 * 2 + 3 - "Посетители".Length) / 2; j++)
                output += " ";

            output += "Посетители";

            for (int j = 0; j < (maxLength.Sum() + 3 * 2 + 3 - "Посетители".Length) / 2 - 1; j++)
                output += " ";
            output += "|";
            output += "\n";

            for (int j = 0; j < maxLength.Sum() + 3 * 2 + 3; j++)
                output += "-";
            output += "\n";

            foreach (Visitor e in _visitors)
            {
                output += "|";
                string space = "";
                for (int l = 0; l < maxLength[0] - e.Id.ToString().Length; l++)
                    space += " ";
                output += $"{space}{e.Id} |";

                space = "";
                for (int l = 0; l < maxLength[1] - e.Name.Length; l++)
                    space += " ";
                output += $"{space}{e.Name} |";

                space = "";
                for (int l = 0; l < maxLength[2] - e.Age.ToString().Length; l++)
                    space += " ";
                output += $"{space}{e.Age} |";

                space = "";
                for (int l = 0; l < maxLength[3] - e.City.Length; l++)
                    space += " ";
                output += $"{space}{e.City} |";
                output += "\n";
            }

            for (int j = 0; j < maxLength.Sum() + 3 * 2 + 3; j++)
                output += "-";
            output += "\n";

            maxLength = new int[5];

            lengths = from Ticket e in _tickets
                select e.Id.ToString().Length;
            maxLength[0] = lengths.Max();

            lengths = from Ticket e in _tickets
                select e.IdExhibit.ToString().Length;
            maxLength[1] = lengths.Max();

            lengths = from Ticket e in _tickets
                select e.IdVisitor.ToString().Length;
            maxLength[2] = lengths.Max();

            lengths = from Ticket e in _tickets
                select e.Time.ToShortDateString().Length;
            maxLength[3] = lengths.Max();

            lengths = from Ticket e in _tickets
                select e.Price.ToString().Length;
            maxLength[4] = lengths.Max();

            for (int j = 0; j < maxLength.Sum() + 3 * 2 + 3; j++)
                output += "-";
            output += "\n";

            output += "|";
            for (int j = 0; j < (maxLength.Sum() + 3 * 2 + 3 - "Билеты".Length) / 2; j++)
                output += " ";

            output += "Билеты";

            for (int j = 0; j < (maxLength.Sum() + 3 * 2 + 3 - "Билеты".Length) / 2 - 1; j++)
                output += " ";
            output += "|";
            output += "\n";

            for (int j = 0; j < maxLength.Sum() + 3 * 2 + 3; j++)
                output += "-";
            output += "\n";

            foreach (Ticket e in _tickets)
            {
                output += "|";
                string space = "";
                for (int l = 0; l < maxLength[0] - e.Id.ToString().Length; l++)
                    space += " ";
                output += $"{space}{e.Id} |";

                space = "";
                for (int l = 0; l < maxLength[1] - e.IdExhibit.ToString().Length; l++)
                    space += " ";
                output += $"{space}{e.IdExhibit} |";

                space = "";
                for (int l = 0; l < maxLength[2] - e.IdVisitor.ToString().Length; l++)
                    space += " ";
                output += $"{space}{e.IdVisitor} |";

                space = "";
                for (int l = 0; l < maxLength[3] - e.Time.ToShortDateString().Length; l++)
                    space += " ";
                output += $"{space}{e.Time.ToShortDateString()} |";

                space = "";
                for (int l = 0; l < maxLength[2] - e.Price.ToString().Length; l++)
                    space += " ";
                output += $"{space}{e.Price} |";
                output += "\n";
            }

            for (int j = 0; j < maxLength.Sum() + 3 * 2 + 3; j++)
                output += "-";
            output += "\n";

            return output;
        }
    }
}