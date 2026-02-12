using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace Toxy
{
    public class ToxySpreadsheet:ICloneable
    {
        public ToxySpreadsheet()
        {
            this.Tables = new List<ToxyTable>();
        }
        public string Name { get; set; }
        public List<ToxyTable> Tables { get; set; }

        public int Length 
        {
            get {
                return this.Tables.Count;
            }
        }
        public ToxyTable[] Slice(int start, int length)
        {
            if (start<0 || this.Tables.Count==0|| this.Tables.Max(t => t.SheetIndex) < start+length-1)
                throw new ArgumentOutOfRangeException();

            return this.Tables.Where(t => t.SheetIndex >= start && t.SheetIndex < start + length).ToArray();
        }
        public ToxyTable this[string name]
        {
            get {
                if (string.IsNullOrEmpty(name))
                {
                    throw new ArgumentNullException("sheet name cannot be empty or null");
                }
                for (int i = 0; i < this.Tables.Count; i++)
                { 
                    if(this.Tables[i].Name == name)
                        return this.Tables[i];
                }
                return null;
            }
        }

        public DataSet ToDataSet()
        {
            DataSet ds = new DataSet();
            ds.DataSetName = this.Name;
            foreach (var table in this.Tables)
            {
                ds.Tables.Add(table.ToDataTable());
            }
            return ds;
        }

        public object Clone()
        {
            ToxySpreadsheet newss = new ToxySpreadsheet();
            newss.Name = this.Name;
            for (int i = 0; i < this.Tables.Count; i++)
            {
                newss.Tables.Add(this.Tables[i].Clone() as ToxyTable);
            }
            return newss;
        }
    }
}
