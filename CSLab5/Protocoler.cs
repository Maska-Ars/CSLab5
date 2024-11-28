using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

class Protocoler
{
    private string file;
    private StreamWriter sw;

    public Protocoler(string file = "Protocol.txt") 
    {
        if (!File.Exists(file))
            throw new Exception("Файла с заданным путем не существет!");

        if (!file.EndsWith(".txt"))
            throw new Exception("Тип файла должен быть txt!");

        this.file = file;
        this.sw = new StreamWriter(file, true);
    }

    public void WriteLine(string s)
    {
        sw.WriteLine($"{DateTime.Now} - {s}");
    }

    public void Save()
    {
        sw.Close();
        sw = new StreamWriter(file, true);
    }

    public void Close() 
    {
        sw.Close();
    }
}