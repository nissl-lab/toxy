using System;
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
        private readonly Dictionary<string, ToxyProperty> list = new Dictionary<string, ToxyProperty>(StringComparer.OrdinalIgnoreCase);
        public ToxyProperty Add(string name, object value)
        {
            ToxyProperty prop = new ToxyProperty() { Name = name, Value = value };
            
            // Adds the Item if it does not exist
            list[name] = prop;
            return prop;
        }
        public void Remove(string name)
        {
            list.Remove(name);
        }
        public string[] GetNames()
        {
            return list.Keys.Select(x => x.ToLower()).ToArray();
        }
        public void Set(ToxyProperty property)
        {
            // Adds the Item if it does not exist or Sets it
            list[property.Name] = property;
        }
        public ToxyProperty Get(string name)
        {
            return list[name];
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
