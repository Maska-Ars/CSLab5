using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Runtime.CompilerServices;
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
    private Workbook wb;

    public DataBase(string file)
    {
        this.file = file;

        if (!File.Exists(file))
            throw new Exception("Файла с заданным путем не существет!");

        if (!file.EndsWith(".xls"))
            throw new Exception("Тип файла должен быть xls!");

        this.wb = new Workbook(file);
    }

    public void DelElById(string tableName, int id)
    {
        var worksheets = from w in wb.Worksheets
                         where w.Name == tableName
                         select w;

        if (worksheets.Count() == 0)
            throw new Exception($"таблицы с названием '{tableName}' не существеут");

        Worksheet ws = worksheets.First();

        var rows = from Row r in ws.Cells.Rows
                   where r[0].IsNumericValue && r[0].IntValue == id
                   select r;

        if (rows.Count() == 0)
            throw new Exception($"объекта с id = {id} не существеут");

        Row row = rows.First();

        rows = from Row r in ws.Cells.Rows
               where r.Index > row.Index
               select r;

        foreach (Row r in rows)
            r[0].Value = r[0].IntValue - 1;

        ws.Cells.DeleteRow(row.Index);

        wb.Save(file);
    }

    public void UpdateElbyId(string tableName, int id, string attributeName, string newValue)
    {
        var worksheets = from w in wb.Worksheets
                         where w.Name == tableName
                         select w;

        Worksheet ws = worksheets.First();

        var columns = from column in ws.Cells.Columns
                      where ws.Cells[0, column.Index].StringValue == attributeName
                      select column;

        int col = columns.First().Index;

        var rows = from Row r in ws.Cells.Rows
                   where r[0].IsNumericValue && r[0].IntValue == id
                   select r;

        Row row = rows.First();

        if (row[col].Type == CellValueType.IsNumeric)
            row[col].Value = Int32.Parse(newValue);
        else if (row[col].Type == CellValueType.IsDateTime)
            row[col].Value = DateTime.Parse(newValue);
        else if (row[col].Type == CellValueType.IsString)
            row[col].Value = newValue;

        wb.Save(file);
    }

    public void AddExhibit(string name, string era)
    {
        Worksheet ws = wb.Worksheets[0];
        ws.Cells.InsertRow(ws.Cells.MaxDataRow + 1);
        Row row = ws.Cells.Rows[ws.Cells.Rows.Count - 1];

        row[0].Value = row.Index;
        row[1].Value = name;
        row[2].Value = era;

        wb.Save(file);
    }

    public void AddVisitor(string name, int age, string city)
    {
        Worksheet ws = wb.Worksheets[1];
        ws.Cells.InsertRow(ws.Cells.MaxDataRow + 1);
        Row row = ws.Cells.Rows[ws.Cells.Rows.Count - 1];

        row[0].Value = row.Index;
        row[1].Value = name;
        row[2].Value = age;
        row[3].Value = city;

        wb.Save(file);
    }

    public void AddTicket(int idExhibit, int idVisitor, DateTime time, int price)
    {
        Worksheet ws = wb.Worksheets[0];
        ws.Cells.InsertRow(ws.Cells.MaxDataRow + 1);
        Row row = ws.Cells.Rows[ws.Cells.Rows.Count - 1];

        var RowIdEx = from Row r in wb.Worksheets[0].Cells.Rows
                      where r[0].IntValue == idExhibit
                      select r;

        if (RowIdEx.Count() == 0)
            throw new Exception($"экспоната с id = {idExhibit} не существеут");

        var RowIdVis = from Row r in wb.Worksheets[1].Cells.Rows
                       where r[0].IntValue == idVisitor
                       select r;

        if (RowIdVis.Count() == 0)
            throw new Exception($"посетителя с id = {idVisitor} не существеут");


        row[0].Value = row.Index;
        row[1].Value = idExhibit;
        row[2].Value = idVisitor;
        row[3].Value = time;
        row[4].Value = price;

        wb.Save(file);
    }

    public int Request1(int idExhibit, DateTime? begin = null, DateTime? end = null)
    {
        // Запрос для получения суммарной выручки за данный период от одного экспоната
        if (begin == null)
            begin = new DateTime(1970, 1, 1);
        if (end == null)
            end = DateTime.Today;

        var prices = from Row r in wb.Worksheets[2].Cells.Rows
                     where
                        r[1].Type == CellValueType.IsNumeric
                        && r[1].IntValue == idExhibit
                        && r[3].DateTimeValue >= begin
                        && r[3].DateTimeValue <= end
                     select r[4].IntValue;

        return prices.Sum();
    }

    public int Request2(string era, DateTime? begin = null, DateTime? end = null)
    {
        // Запрос для получение суммарной выручки от экспонатов казанной эры, за указанный промежуток времени
        if (begin == null)
            begin = new DateTime(1970, 1, 1);
        if (end == null)
            end = DateTime.Today;

        var prices = from Row r1 in wb.Worksheets[2].Cells.Rows
                     join Row r2 in wb.Worksheets[0].Cells.Rows on r1[1].Value equals r2[0].Value
                     where
                        r2[2].StringValue == era
                        && r1[3].DateTimeValue >= begin
                        && r1[3].DateTimeValue <= end
                     select r1[4].IntValue;

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

        var people = from Row rT in wb.Worksheets[2].Cells.Rows
                     join Row rEx in wb.Worksheets[0].Cells.Rows on rT[1].Value equals rEx[0].Value
                     join Row rV in wb.Worksheets[1].Cells.Rows on rT[2].Value equals rV[0].Value
                     where
                        rEx[0].Type == CellValueType.IsNumeric
                        && rEx[0].IntValue == idExhibit
                        && rV[3].StringValue == city
                        && rT[1].DateTimeValue >= begin
                        && rT[3].DateTimeValue <= end
                     select new {
                         id = rT[0].IntValue,
                         name = rV[1].StringValue,
                         age = rV[2].IntValue,
                         price = rT[4].IntValue
                     };
        return people;
    }

    public IEnumerable<object> Request4(int age, string era)
    {
        // запрос на получение id, имен, времени посещения экспонатов данной эпохи, посетителями данного возраста
        var people = from Row rT in wb.Worksheets[2].Cells.Rows
                     join Row rEx in wb.Worksheets[0].Cells.Rows on rT[1].Value equals rEx[0].Value
                     join Row rV in wb.Worksheets[1].Cells.Rows on rT[2].Value equals rV[0].Value
                     where
                        rV[2].Type == CellValueType.IsNumeric
                        && rV[2].IntValue == age
                        && rEx[2].StringValue == era
                     select new
                     {
                         name = rV[1].StringValue,
                         idTicket = rT[0].IntValue,
                         date = rT[3].DateTimeValue
                     };
        return people;
    }

    public string GetFile()
    {
        return this.file;
    }

    public override string ToString()
    {
        string s = "";
        WorksheetCollection col = wb.Worksheets;

        for (int i = 0; i < 3; i++)
        {
            Worksheet ws = col[i];

            int rows = ws.Cells.MaxDataRow;
            int cols = ws.Cells.MaxDataColumn;

            int[] maxLength = new int[cols+1];
            for (int j = 0; j <= cols; j++)
            {
                var k = from Row r in ws.Cells.Rows
                        select r[j].StringValue.Length;
                maxLength[j] = k.Max();
            }

            for (int j = 0; j < maxLength.Sum() + cols * 2 + 3; j++)
                s += "-";
            s += "\n";

            s += "|";
            for (int j = 0; j < (maxLength.Sum() + cols * 2 + 3 - ws.Name.Length) / 2; j++)
                s += " ";

            s += ws.Name;

            for (int j = 0; j < (maxLength.Sum() + cols * 2 + 3 - ws.Name.Length) / 2 - 1; j++)
                s += " ";
            s += "|";
            s += "\n";

            for (int j = 0; j < maxLength.Sum() + cols * 2 + 3; j++)
                s += "-";
            s += "\n";

            for (int j = 0; j <= rows; j++)
            {
                s += "|";
                for (int k = 0; k <= cols; k++)
                {
                    string space = "";
                    for (int l = 0; l < maxLength[k] - ws.Cells[j, k].StringValue.Length; l++)
                        space += " ";

                    if (ws.Cells[j, k].Type == CellValueType.IsDateTime)
                        s += $"{space}{ws.Cells[j, k].DateTimeValue.ToShortDateString()} |";
                    else
                        s += $"{space}{ws.Cells[j, k].StringValue} |";
                }
                s += "\n";
            }
            for (int j = 0; j < maxLength.Sum() + cols * 2 + 3; j++)
                s += "-";
            s += "\n";

        }

        return s;
    }
}