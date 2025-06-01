namespace CSLab5 
{
    class Protocoler
    {
        private string _file;
        private StreamWriter _sw;

        public Protocoler(string file = "Protocol.txt")
        {
            if (!File.Exists(file))
                throw new Exception("Файла с заданным путем не существет!");

            if (!file.EndsWith(".txt"))
                throw new Exception("Тип файла должен быть txt!");

            _file = file;
            _sw = new StreamWriter(file, true);
        }

        public void WriteLine(string s)
        {
            _sw.WriteLine($"{DateTime.Now} - {s}");
        }

        public void Save()
        {
            _sw.Close();
            _sw = new StreamWriter(_file, true);
        }

        public void Close()
        {
            _sw.Close();
        }
    }
}