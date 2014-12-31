using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace Toxy
{




    public class ToxySpreadsheet
    {
        public ToxySpreadsheet()
        {
            this.Tables = new List<ToxyTable>();
        }
        public string Name { get; set; }
        public List<ToxyTable> Tables { get; set; }

        public ToxyTable this[string name]
        {
            get {
                if (string.IsNullOrEmpty(name))
                {
                    throw new ArgumentNullException("sheet name cannot be empty or null");
                }
                for (int i = 0; i < this.Tables.Count; i++)
                { 
                    if(this.Tables[i].Name==name)
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
    }
}
