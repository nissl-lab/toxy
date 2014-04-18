using System;
using System.Collections.Generic;
using System.Text;

namespace Toxy
{
    public class ToxyAttribute
    {
        public string Name { get; set; }
        public string Value { get; set; }
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
