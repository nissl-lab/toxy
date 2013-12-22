using System;
using System.Collections.Generic;
using System.Text;

namespace Toxy.Entities
{
    public class ToxyAttribute
    {
        public string Name { get; set; }
        public string Value { get; set; }
    }
    public class ToxyNode
    {
        public List<ToxyAttribute> Attributes { get; set; }
        public string Text { get; set; }
        public ToxyNode Parent { get; set; }
        public ToxyNode Nodes { get; set; }

        public string TagName { get; set; }
        public string Type { get; set; }

        public ToxyNode SingleSelect(string selector) 
        {
            throw new NotImplementedException();
        }
        public List<ToxyNode> SelectNodes(string selector)
        {
            throw new NotImplementedException();
        }

        public string ToText()
        {
            throw new NotImplementedException();
        }
        public string ToHtml()
        {
            throw new NotImplementedException();
        }

    }

    public class ToxyDom
    {
        private ToxyNode root = null;
        public ToxyNode Root
        {
            get { return root; }
        }
    }
}
