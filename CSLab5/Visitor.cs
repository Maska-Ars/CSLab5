namespace CSLab5
{
    class Visitor(int id, string name, int age, string city)
    {
        private int _id = id;
        private string _name = name;
        private int _age = age;
        private string _city = city;

        public int Id
        {
            get { return _id; }
            set { _id = value; }
        }

        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }

        public int Age
        {
            get { return _age; }
            set { _age = value; }
        }

        public string City
        {
            get { return _city; }
            set { _city = value; }
        }

        public override string ToString()
        {
            return $"id = {_id}, name = {_name}, age = {_age}, city = {_city}";
        }
    }

}