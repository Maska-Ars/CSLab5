using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
class Visitor
{
    public int id { get; set; }
    public string name { get; set; }
    public int age { get; set; }
    public string city { get; set; }

    public Visitor(int id, string name, int age, string city) 
    { 
        this.id = id;
        this.name = name;
        this.age = age;
        this.city = city;
    }

    public override string ToString()
    {
        return $"{this.id}, {this.name}, {this.age}, {this.city}";
    }

}