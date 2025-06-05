using Aspose.Cells.Charts;

namespace CSLab5
{
    /// <include file='Docs/Exhibit.xml' 
    /// path='Docs/members[@name="exhibit"]/Exhibit/*'/>
    class Exhibit(int id, string name, string era)
    {
        private int _id = id;
        private string _name = name;
        private string _era = era;

        /// <include file='Docs/Exhibit.xml' 
        /// path='Docs/members[@name="exhibit"]/Id/*'/>
        public int Id
        {
            get { return _id; }
            set { _id = value; }
        }

        /// <include file='Docs/Exhibit.xml' 
        /// path='Docs/members[@name="exhibit"]/Name/*'/>
        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }

        /// <include file='Docs/Exhibit.xml' 
        /// path='Docs/members[@name="exhibit"]/Era/*'/>
        public string Era
        {
            get { return _era; }
            set { _era = value; }
        }

        /// <include file='Docs/Exhibit.xml' 
        /// path='Docs/members[@name="exhibit"]/ToString/*'/>
        public override string ToString()
        {
            return $"id = {_id}, name = {_name}, era = {_era}";
        }
    }
}