namespace CSLab5
{
    /// <include file='Docs/Visitor.xml' path='Docs/members[@name="visitor"]/Visitor/*'/>
    class Visitor(int id, string name, int age, string city)
    {
        private int _id = id;
        private string _name = name;
        private int _age = age;
        private string _city = city;

        /// <include file='Docs/Visitor.xml' path='Docs/members[@name="visitor"]/Id/*'/>
        public int Id
        {
            get { return _id; }
            set { _id = value; }
        }

        /// <include file='Docs/Visitor.xml' path='Docs/members[@name="visitor"]/Name/*'/>
        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }

        /// <include file='Docs/Visitor.xml' path='Docs/members[@name="visitor"]/Age/*'/>
        public int Age
        {
            get { return _age; }
            set { _age = value; }
        }

        /// <include file='Docs/Visitor.xml' path='Docs/members[@name="visitor"]/City/*'/>
        public string City
        {
            get { return _city; }
            set { _city = value; }
        }

        /// <include file='Docs/Visitor.xml' path='Docs/members[@name="visitor"]/ToString/*'/>
        public override string ToString() => $"id = {_id}, name = {_name}, age = {_age}, city = {_city}";
    }

}