namespace Toxy
{
    public class ToxyAttribute
    {
        public ToxyAttribute()
        { 
            
        }
        public ToxyAttribute(string name, string value)
        {
            this.Name = name;
            this.Value = value;
        }
        public string Name { get; set; }
        public string Value { get; set; }
        public override string ToString()
        {
            if (Name == null)
                return string.Empty;

            if(Value != null)
                return string.Format("{0}=\"{1}\"", Name, Value);
            else
                return string.Format("{0}", Name);
        }
    }


    public class ToxyDom
    {
        private ToxyNode root = null;
        public ToxyNode Root
        {
            get { return root; }
            internal set { root = value; }
        }
    }
}
