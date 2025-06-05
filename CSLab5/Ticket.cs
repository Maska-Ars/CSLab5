namespace CSLab5
{
    /// <include file='Docs/Ticket.xml' 
    /// path='Docs/members[@name="ticket"]/Ticket/*'/>
    class Ticket(int id, int idExhibit, int idVisitor, DateTime time, int price)
    {
        private int _id = id;
        private int _idExhibit = idExhibit;
        private int _idVisitor = idVisitor;
        private DateTime _time = time;
        private int _price = price;

        /// <include file='Docs/Ticket.xml' 
        /// path='Docs/members[@name="ticket"]/Id/*'/>
        public int Id
        {
            get { return _id; }
            set { _id = value; }
        }

        /// <include file='Docs/Ticket.xml' 
        /// path='Docs/members[@name="ticket"]/IdExhibit/*'/>
        public int IdExhibit
        {
            get { return _idExhibit; }
            set { _idExhibit = value; }
        }

        /// <include file='Docs/Ticket.xml' 
        /// path='Docs/members[@name="ticket"]/IdVisitor/*'/>
        public int IdVisitor
        {
            get { return _idVisitor; }
            set { _idVisitor = value; }
        }

        /// <include file='Docs/Ticket.xml' 
        /// path='Docs/members[@name="ticket"]/Time/*'/>
        public DateTime Time
        {
            get { return _time; }
            set { _time = value; }
        }

        /// <include file='Docs/Ticket.xml' 
        /// path='Docs/members[@name="ticket"]/Price/*'/>
        public int Price
        {
            get { return _price; }
            set { _price = value; }
        }

        /// <include file='Docs/Ticket.xml' 
        /// path='Docs/members[@name="ticket"]/ToString/*'/>
        public override string ToString()
        {
            return $"id = {_id}, " +
                $"idEx = {_idExhibit}, " +
                $"idV = {_idVisitor}, " +
                $"time = {_time}, " +
                $"price = {_price}";
        }
    }
}