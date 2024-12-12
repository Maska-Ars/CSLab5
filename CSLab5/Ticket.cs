using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

class Ticket
{
    private int id;
    private int idExhibit;
    private int idVisitor;
    private DateTime time;
    private int price;

    public Ticket(int id, int idExhibit, int idVisitor, DateTime time, int price)
    {
        this.id = id;
        this.idExhibit = idExhibit;
        this.idVisitor = idVisitor;
        this.time = time;
        this.price = price;
    }

    public int GetId() => id;
    public void SetId(int id) => this.id = id;

    public int GetIdExhibit() => idExhibit;
    public void SetIdExhibit(int idExhibit) => this.idExhibit = idExhibit;

    public int GetIdVisitor() => idVisitor;
    public void SetIdVisitor(int idVisitor) => this.idVisitor = idVisitor;

    public DateTime GetTime() => time;
    public void SetTime(DateTime time) => this.time = time;

    public int GetPrice() => price;
    public void SetPrice(int price) => this.price = price;

    public override string ToString()
    {
        return $"id = {id}, idEx = {idExhibit}, idV = {idVisitor}, time = {time}, price = {price}";
    }
}