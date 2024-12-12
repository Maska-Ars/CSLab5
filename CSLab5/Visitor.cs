using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

class Visitor
{
    private int id;
    private string name;
    private int age;
    private string city;

    public Visitor(int id, string name, int age, string city)
    {
        this.id = id;
        this.name = name;
        this.age = age;
        this.city = city;
    }

    public int GetId() => id;
    public void SetId(int id) => this.id = id;

    public string GetName() => name;
    public void SetName(string Name) => this.name = Name;

    public int GetAge() => age;
    public void SetAge(int age) => this.age = age;

    public string GetCity() => city;
    public void SetCity(string city) => this.city = city;

    public override string ToString()
    {
        return $"id = {id}, name = {name}, age = {age}, city = {city}";
    }
}
