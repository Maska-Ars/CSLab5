namespace CSLab5
{
    class Exhibit(int id, string name, string era)
    {
        private int _id = id;
        private string _name = name;
        private string _era = era;

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

        public string Era
        {
            get { return _era; }
            set { _era = value; }
        }

        public override string ToString()
        {
            return $"id = {_id}, name = {_name}, era = {_era}";
        }

    }
}