using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

class Ticket
{
    public int id { get; set; }
    public int idExhibit { get; set; }
    public int idVisitor { get; set; }
    public DateTime date { get; set; }
    public int price { get; set; }

    public Ticket(int id,int idExhibit,int idVisitor, DateTime date, int price)
    { 
        this.id = id;
        this.idExhibit = idExhibit;
        this.idVisitor = idVisitor;
        this.date = date;
        this.price = price;
    }

    public override string ToString()
    {
        return $"{this.id}, {this.idExhibit}, {this.idVisitor}, {this.date.ToString("d")}, {this.price}";
    }
}