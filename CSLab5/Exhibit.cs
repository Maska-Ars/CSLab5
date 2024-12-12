using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

class Exhibit
{
    private int id;
    private string name;
    private string era;

    public Exhibit(int id, string name, string era)
    {
        this.id = id;
        this.name = name;
        this.era = era;
    }

    public int GetId() => id;
    public void SetId(int id) => this.id = id;

    public string GetName() => name;
    public void SetName(string name) => this.name = name;

    public string GetEra() => era;
    public void SetEra(string era) => this.era = era;

    public override string ToString()
    {
        return $"id = {id}, name = {name}, era = {era}";
    }

}