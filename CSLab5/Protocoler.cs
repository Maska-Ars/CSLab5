namespace CSLab5 
{
    /// <include file='Docs/Protocoler.xml' 
    /// path='Docs/members[@name="protocoler"]/Protocoler/*'/>
    class Protocoler
    {
        private readonly string _file;
        private StreamWriter _sw;

        /// <include file='Docs/Protocoler.xml' 
        /// path='Docs/members[@name="protocoler"]/Constructor/*'/>
        public Protocoler(string file = "Protocol.txt")
        {
            if (!File.Exists(file))
                throw new Exception("Файла с заданным путем не существет!");

            if (!file.EndsWith(".txt"))
                throw new Exception("Тип файла должен быть txt!");

            _file = file;
            _sw = new StreamWriter(file, true);
        }

        /// <include file='Docs/Protocoler.xml' 
        /// path='Docs/members[@name="protocoler"]/WriteLine/*'/>
        public void WriteLine(string s)
        {
            _sw.WriteLine($"{DateTime.Now} - {s}");
        }

        /// <include file='Docs/Protocoler.xml' 
        /// path='Docs/members[@name="protocoler"]/Save/*'/>
        public void Save()
        {
            _sw.Close();
            _sw = new StreamWriter(_file, true);
        }

        /// <include file='Docs/Protocoler.xml' 
        /// path='Docs/members[@name="protocoler"]/Close/*'/>
        public void Close() => _sw.Close();
    }
}