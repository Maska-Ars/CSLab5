using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Net.Http.Headers;
using System.Runtime.CompilerServices;
using System.Runtime.ExceptionServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Microsoft.VisualBasic;
using static System.Runtime.InteropServices.JavaScript.JSType;

class DataBase
{
    private string file;

    private List<Exhibit> exhibits;
    private List<Visitor> visitors;
    private List<Ticket> tickets;

    public DataBase(string file)
    {
        this.file = file;

        if (!File.Exists(file))
            throw new Exception("Файла с заданным путем не существет!");

        if (!file.EndsWith(".xls"))
            throw new Exception("Тип файла должен быть xls!");

        Workbook wb = new Workbook(file);

        exhibits = new List<Exhibit>();

        Worksheet ws = wb.Worksheets[0];

        for (int i = 1; i <= ws.Cells.MaxDataRow; i++)
        {
            Row row = ws.Cells.Rows[i];

            exhibits.Add(new Exhibit(
                row[0].IntValue,
                row[1].StringValue,
                row[2].StringValue
                ));

        }

        visitors = new List<Visitor>();

        ws = wb.Worksheets[1];

        for (int i = 1; i <= ws.Cells.MaxDataRow; i++)
        {
            Row row = ws.Cells.Rows[i];

            visitors.Add(new Visitor(
                row[0].IntValue,
                row[1].StringValue,
                row[2].IntValue,
                row[3].StringValue
                ));

        }

        tickets = new List<Ticket>();

        ws = wb.Worksheets[2];

        for (int i = 1; i <= ws.Cells.MaxDataRow; i++)
        {
            Row row = ws.Cells.Rows[i];

            tickets.Add(new Ticket(
                row[0].IntValue,
                row[1].IntValue,
                row[2].IntValue,
                row[3].DateTimeValue,
                row[4].IntValue
                ));

        }


    }

    public string GetFile() => this.file;

    public void DelExhibitById(int id)
    {
        var tic = from Ticket t in tickets
                  where t.GetIdExhibit() == id
                  select t;

        foreach (var t in tic)
        {
            this.DelTickettById(t.GetId());
        }

        var exs = from Exhibit e in exhibits
                  where id == e.GetId()
                  select e;
        if (exs.Count() == 0)
            throw new Exception($"экспоната с id = {id} не существеут");

        exhibits.Remove(exhibits.First());

        exs = from Exhibit e in exhibits
              where id < e.GetId()
              select e;

        foreach (Exhibit e in exs)
        {
            e.SetId(e.GetId() - 1);
        }

    }

    public void DelVisitorById(int id)
    {
        var tic = from Ticket t in tickets
                  where t.GetIdVisitor() == id
                  select t;

        foreach (var t in tic)
        {
            this.DelTickettById(t.GetId());
        }

        var exs = from Visitor e in visitors
                  where id == e.GetId()
                  select e;
        if (exs.Count() == 0)
            throw new Exception($"посетителя с id = {id} не существеут");

        exhibits.Remove(exhibits.First());

        exs = from Visitor e in visitors
              where id < e.GetId()
              select e;

        foreach (Visitor e in exs)
        {
            e.SetId(e.GetId() - 1);
        }

    }

    public void DelTickettById(int id)
    {
        var exs = from Ticket e in tickets
                  where id == e.GetId()
                  select e;
        if (exs.Count() == 0)
            throw new Exception($"билета с id = {id} не существеут");

        exhibits.Remove(exhibits.First());

        exs = from Ticket e in tickets
              where id < e.GetId()
              select e;

        foreach (Ticket e in exs)
        {
            e.SetId(e.GetId() - 1);
        }

    }

    public void DelElById(int idTable, int id)
    {
        if (idTable < 0 || idTable > 2)
            throw new Exception($"таблицы с id '{idTable}' не существеут");

        switch(id)
        {
            case 0:
                this.DelExhibitById(id);
                break;
            case 1:
                this.DelVisitorById(id);
                break;
            case 2:
                this.DelTickettById(id);
                break;
        }

        this.Save();
    }

    public void UpdateExhibitById(int id, string AttributeName, string newValue)
    {
        var exs = from Exhibit e in exhibits
                  where id == e.GetId()
                  select e;
        if (exs.Count() == 0)
            throw new Exception($"экспоната с id = {id} не существеут");

        Exhibit ex = exs.First();

        switch(AttributeName)
        {
            case "name":
                ex.SetName(newValue);
                break;
            case "era":
                ex.SetEra(newValue);
                break;
            default:
                throw new Exception($"атрибут с названием = {AttributeName} не существеут");
        }

    }

    public void UpdateVisitorById(int id, string AttributeName, string newValue)
    {
        var exs = from Visitor e in visitors
                  where id == e.GetId()
                  select e;
        if (exs.Count() == 0)
            throw new Exception($"посетителя с id = {id} не существеут");

        Visitor ex = exs.First();

        switch (AttributeName)
        {
            case "name":
                ex.SetName(newValue);
                break;
            case "age":
                ex.SetAge(Convert.ToInt32(newValue));
                break;
            case "city":
                ex.SetCity(newValue);
                break;
            default:
                throw new Exception($"атрибут с названием = {AttributeName} не существеут");
        }

    }

    public void UpdateTicketById(int id, string AttributeName, string newValue)
    {
        var exs = from Ticket e in tickets
                  where id == e.GetId()
                  select e;
        if (exs.Count() == 0)
            throw new Exception($"экспоната с id = {id} не существеут");

        Ticket ex = exs.First();

        switch (AttributeName)
        {
            case "idExhibit":
                ex.SetIdExhibit(Convert.ToInt32(newValue));
                break;
            case "idVisitor":
                ex.SetIdVisitor(Convert.ToInt32(newValue));
                break;
            case "price":
                ex.SetPrice(Convert.ToInt32(newValue));
                break;
            case "time":
                ex.SetTime(Convert.ToDateTime(newValue));
                break;
            default:
                throw new Exception($"атрибут с названием = {AttributeName} не существеут");
        }

    }

    public void UpdateElbyId(int idTable, int id, string attributeName, string newValue)
    {
        if (idTable < 0 || idTable > 2)
            throw new Exception($"таблицы с id '{idTable}' не существеут");

        switch(idTable)
        {
            case 0:
                this.UpdateExhibitById(id, attributeName, newValue);
                break;
            case 1:
                this.UpdateVisitorById(id, attributeName, newValue);
                break;
            case 2:
                this.UpdateTicketById(id, attributeName, newValue);
                break;
        }

        this.Save();
    }

    public void AddExhibit(string name, string era)
    {
        exhibits.Add(new Exhibit(
            exhibits[-1].GetId() + 1,
            name,
            era
            ));
        this.Save();
    }

    public void AddVisitor(string name, int age, string city)
    {
        visitors.Add(new Visitor(
            visitors[-1].GetId() + 1,
            name,
            age,
            city
            ));
        this.Save();

    }

    public void AddTicket(int idExhibit, int idVisitor, DateTime time, int price)
    {
        tickets.Add(new Ticket(
            tickets[-1].GetId() + 1,
            idExhibit,
            idVisitor,
            time,
            price
            ));
        this.Save();

    }

    public int Request1(int idExhibit, DateTime? begin = null, DateTime? end = null)
    {
        // Запрос для получения суммарной выручки за данный период от одного экспоната
        if (begin == null)
            begin = new DateTime(1970, 1, 1);
        if (end == null)
            end = DateTime.Today;

        var prices = from Ticket r in tickets
                     where
                        r.GetIdExhibit() == idExhibit
                        && r.GetTime() >= begin
                        && r.GetTime() <= end
                     select r.GetPrice();

        return prices.Sum();
    }

    public int Request2(string era, DateTime? begin = null, DateTime? end = null)
    {
        // Запрос для получение суммарной выручки от экспонатов казанной эры, за указанный промежуток времени
        if (begin == null)
            begin = new DateTime(1970, 1, 1);
        if (end == null)
            end = DateTime.Today;

        var prices = from Ticket r1 in tickets
                     join Exhibit r2 in exhibits on r1.GetIdExhibit() equals r2.GetId()
                     where
                        r2.GetEra() == era
                        && r1.GetTime() >= begin
                        && r1.GetTime() <= end
                     select r1.GetPrice();

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

        var people = from Ticket rT in tickets
                     join Exhibit rEx in exhibits on rT.GetIdExhibit() equals rEx.GetId()
                     join Visitor rV in visitors on rT.GetIdVisitor() equals rV.GetId()
                     where
                        rEx.GetId() == idExhibit
                        && rV.GetCity() == city
                        && rT.GetTime() >= begin
                        && rT.GetTime() <= end
                     select new {
                         idTicket = rT.GetId(),
                         name = rV.GetName(),
                         age = rV.GetAge(),
                         price = rT.GetPrice()
                     };
        return people;
    }

    public IEnumerable<object> Request4(int age, string era)
    {
        // запрос на получение id, имен, времени посещения экспонатов данной эпохи, посетителями данного возраста
        var people = from Ticket rT in tickets
                     join Exhibit rEx in exhibits on rT.GetIdExhibit() equals rEx.GetId()
                     join Visitor rV in visitors on rT.GetIdVisitor() equals rV.GetId()
                     where
                        rV.GetAge() == age
                        && rEx.GetEra() == era
                     select new
                     {
                         name = rV.GetName(),
                         idTicket = rT.GetId(),
                         date = rT.GetTime()
                     };
        return people;
    }

    public void Save()
    {

        Workbook wb = new Workbook(file);

        Worksheet ws = wb.Worksheets[0];

        Style s1 = ws.Cells.Rows[1][0].GetDisplayStyle();
        Style s2 = ws.Cells.Rows[1][1].GetDisplayStyle();
        Style s3 = ws.Cells.Rows[1][2].GetDisplayStyle();

        ws.Cells.DeleteRows(1, ws.Cells.Rows.Count-1);


        foreach (Exhibit e in exhibits)
        {
            ws.Cells.InsertRow(ws.Cells.MaxDataRow + 1);

            ws.Cells.Rows[^1][0].SetStyle(s1);
            ws.Cells.Rows[^1][0].PutValue(e.GetId());

            ws.Cells.Rows[^1][1].SetStyle(s2);
            ws.Cells.Rows[^1][1].PutValue(e.GetName());

            ws.Cells.Rows[^1][2].SetStyle(s3);
            ws.Cells.Rows[^1][2].PutValue(e.GetEra());

        }

        ws = wb.Worksheets[1];
        s1 = ws.Cells.Rows[1][0].GetDisplayStyle();
        s2 = ws.Cells.Rows[1][1].GetDisplayStyle();
        s3 = ws.Cells.Rows[1][2].GetDisplayStyle();
        Style s4 = ws.Cells.Rows[1][3].GetDisplayStyle();
        ws.Cells.DeleteRows(1, ws.Cells.Rows.Count - 1);

        foreach (Visitor e in visitors)
        {
            ws.Cells.InsertRow(ws.Cells.MaxDataRow + 1);

            ws.Cells.Rows[^1][0].SetStyle(s1);
            ws.Cells.Rows[^1][0].PutValue(e.GetId());

            ws.Cells.Rows[^1][1].SetStyle(s2);
            ws.Cells.Rows[^1][1].PutValue(e.GetName());

            ws.Cells.Rows[^1][2].SetStyle(s3);
            ws.Cells.Rows[^1][2].PutValue(e.GetAge());

            ws.Cells.Rows[^1][3].SetStyle(s4);
            ws.Cells.Rows[^1][3].PutValue(e.GetCity());
        }

        ws = wb.Worksheets[2];

        s1 = ws.Cells.Rows[1][0].GetDisplayStyle();
        s2 = ws.Cells.Rows[1][1].GetDisplayStyle();
        s3 = ws.Cells.Rows[1][2].GetDisplayStyle();
        s4 = ws.Cells.Rows[1][3].GetDisplayStyle();
        Style s5 = ws.Cells.Rows[1][4].GetDisplayStyle();

        ws.Cells.DeleteRows(1, ws.Cells.Rows.Count - 1);

        foreach (Ticket e in tickets)
        {
            ws.Cells.InsertRow(ws.Cells.MaxDataRow + 1);

            ws.Cells.Rows[^1][0].SetStyle(s1);
            ws.Cells.Rows[^1][0].PutValue(e.GetId());

            ws.Cells.Rows[^1][1].SetStyle(s2);
            ws.Cells.Rows[^1][1].PutValue(e.GetIdExhibit());

            ws.Cells.Rows[^1][2].SetStyle(s3);
            ws.Cells.Rows[^1][2].PutValue(e.GetIdVisitor());

            ws.Cells.Rows[^1][3].SetStyle(s4);
            ws.Cells.Rows[^1][3].PutValue(e.GetTime());

            ws.Cells.Rows[^1][4].SetStyle(s5);
            ws.Cells.Rows[^1][4].PutValue(e.GetPrice());
        }

        wb.Save("save"+file);

    }

    public override string ToString()
    {
        string s = "";
        
        int[] maxLength = new int[3];

        var k = from Exhibit e in exhibits
                select e.GetId().ToString().Length;
        maxLength[0] = k.Max();

        k = from Exhibit e in exhibits
            select e.GetName().Length;
        maxLength[1] = k.Max();

        k = from Exhibit e in exhibits
            select e.GetEra().Length;
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

        foreach (Exhibit e in exhibits)
        {
            s += "|";
            string space = "";
            for (int l = 0; l < maxLength[0] - e.GetId().ToString().Length; l++)
                space += " ";
            s += $"{space}{e.GetId().ToString()} |";

            space = "";
            for (int l = 0; l < maxLength[1] - e.GetName().Length; l++)
                space += " ";
            s += $"{space}{e.GetName()} |";

            space = "";
            for (int l = 0; l < maxLength[2] - e.GetEra().Length; l++)
                space += " ";
            s += $"{space}{e.GetEra()} |";
            s += "\n";
        }
        for (int j = 0; j < maxLength.Sum() + 3 * 2 + 3; j++)
            s += "-";
        s += "\n";


        maxLength = new int[4];

        k = from Visitor e in visitors
            select e.GetId().ToString().Length;
        maxLength[0] = k.Max();

        k = from Visitor e in visitors
            select e.GetName().Length;
        maxLength[1] = k.Max();

        k = from Visitor e in visitors
            select e.GetAge().ToString().Length;
        maxLength[2] = k.Max();

        k = from Visitor e in visitors
            select e.GetCity().Length;
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

        foreach (Visitor e in visitors)
        {
            s += "|";
            string space = "";
            for (int l = 0; l < maxLength[0] - e.GetId().ToString().Length; l++)
                space += " ";
            s += $"{space}{e.GetId().ToString()} |";

            space = "";
            for (int l = 0; l < maxLength[1] - e.GetName().Length; l++)
                space += " ";
            s += $"{space}{e.GetName()} |";

            space = "";
            for (int l = 0; l < maxLength[2] - e.GetAge().ToString().Length; l++)
                space += " ";
            s += $"{space}{e.GetAge()} |";

            space = "";
            for (int l = 0; l < maxLength[3] - e.GetCity().Length; l++)
                space += " ";
            s += $"{space}{e.GetCity()} |";
            s += "\n";
        }
        for (int j = 0; j < maxLength.Sum() + 3 * 2 + 3; j++)
            s += "-";
        s += "\n";


        maxLength = new int[5];

        k = from Ticket e in tickets
            select e.GetId().ToString().Length;
        maxLength[0] = k.Max();

        k = from Ticket e in tickets
            select e.GetIdExhibit().ToString().Length;
        maxLength[1] = k.Max();

        k = from Ticket e in tickets
            select e.GetIdVisitor().ToString().Length;
        maxLength[2] = k.Max();

        k = from Ticket e in tickets
            select e.GetTime().ToShortDateString().Length;
        maxLength[3] = k.Max();

        k = from Ticket e in tickets
            select e.GetPrice().ToString().Length;
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

        foreach (Ticket e in tickets)
        {
            s += "|";
            string space = "";
            for (int l = 0; l < maxLength[0] - e.GetId().ToString().Length; l++)
                space += " ";
            s += $"{space}{e.GetId().ToString()} |";

            space = "";
            for (int l = 0; l < maxLength[1] - e.GetIdExhibit().ToString().Length; l++)
                space += " ";
            s += $"{space}{e.GetIdExhibit()} |";

            space = "";
            for (int l = 0; l < maxLength[2] - e.GetIdVisitor().ToString().Length; l++)
                space += " ";
            s += $"{space}{e.GetIdVisitor()} |";

            space = "";
            for (int l = 0; l < maxLength[3] - e.GetTime().ToShortDateString().Length; l++)
                space += " ";
            s += $"{space}{e.GetTime().ToShortDateString()} |";

            space = "";
            for (int l = 0; l < maxLength[2] - e.GetPrice().ToString().Length; l++)
                space += " ";
            s += $"{space}{e.GetPrice()} |";
            s += "\n";
        }
        for (int j = 0; j < maxLength.Sum() + 3 * 2 + 3; j++)
            s += "-";
        s += "\n";


        return s;
    }
}