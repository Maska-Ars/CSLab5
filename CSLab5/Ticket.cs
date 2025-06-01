namespace CSLab5
{
    class Ticket(int id, int idExhibit, int idVisitor, DateTime time, int price)
    {
        private int _id = id;
        private int _idExhibit = idExhibit;
        private int _idVisitor = idVisitor;
        private DateTime _time = time;
        private int _price = price;

        public int Id
        {
            get { return _id; }
            set { _id = value; }
        }

        public int IdExhibit
        {
            get { return _idExhibit; }
            set { _idExhibit = value; }
        }

        public int IdVisitor
        {
            get { return _idVisitor; }
            set { _idVisitor = value; }
        }

        public DateTime Time
        {
            get { return _time; }
            set { _time = value; }
        }

        public int Price
        {
            get { return _price; }
            set { _price = value; }
        }

        public override string ToString()
        {
            return $"id = {_id}, idEx = {_idExhibit}, idV = {_idVisitor}, time = {_time}, price = {_price}";
        }
    }
}