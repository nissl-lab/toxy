using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Toxy
{
    public class ToxyProperty
    {
        public string Name { get; set; }
        public object Value { get; set; }
    }
    public class ToxyMetadata:IEnumerable<ToxyProperty>
    {
        Dictionary<string, ToxyProperty> list = new Dictionary<string, ToxyProperty>();
        public ToxyProperty Add(string name, object value)
        {
            ToxyProperty prop = new ToxyProperty() { Name = name, Value = value };
            if (list.ContainsKey(name.ToLower()))
                list[name.ToLower()] = prop;
            else
                list.Add(name.ToLower(), prop);
            return prop;
        }
        public void Remove(string name)
        {
            list.Remove(name.ToLower());
        }
        public string[] GetNames()
        {
            return list.Keys.ToArray();
        }
        public void Set(ToxyProperty property)
        {
            if(list.ContainsKey(property.Name))
                list[property.Name] = property;
            else
                list.Add(property.Name, property);
        }
        public ToxyProperty Get(string name)
        {
            return list[name.ToLower()];
        }
        public int Count
        {
            get
            {
                return list.Count;
            }
        }
        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            foreach (var prop in list.Values)
            {
                sb.AppendLine(string.Format("{0}:{1}", prop.Name, prop.Value.ToString()));
            }
            return sb.ToString();
        }


        public IEnumerator<ToxyProperty> GetEnumerator()
        {
            return list.Values.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return list.Values.GetEnumerator();
        }
    }   
}
