using System.Collections.Generic;
using System.Text;

namespace Toxy
{
    public class ToxyNode
    {
        /// <summary>
        /// node name, if it's a textnode, this will be "#text"
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// for text node use only
        /// </summary>
        public string Text { get; set; }
        private List<ToxyAttribute> attributes;
        public List<ToxyAttribute> Attributes { get { return attributes ?? (attributes = new List<ToxyAttribute>()); } set { attributes = value; } }

        private List<ToxyNode> childrenNodes;
        public List<ToxyNode> ChildrenNodes
        {
            get { return childrenNodes ?? (childrenNodes = new List<ToxyNode>()); }
            set { childrenNodes = value; }
        }

        internal HtmlAgilityPack.HtmlNode HtmlNode { get; set; }

        internal static ToxyNode TransformHtmlNodeToToxyNode(HtmlAgilityPack.HtmlNode htmlNode)
        {
            if (htmlNode == null)
            {
                return null;
            }
            ToxyNode toxyNode = new ToxyNode();
            toxyNode.Name = htmlNode.Name;
            foreach (var item in htmlNode.Attributes)
            {
                toxyNode.Attributes.Add(new ToxyAttribute { Name = item.Name, Value = item.Value });
            }
            toxyNode.HtmlNode = htmlNode;
            return toxyNode;
        }

        public ToxyNode SingleSelect(string selector)
        {
            //todo:need to copy ChildNodes to the seleted node
            return ToxyNode.TransformHtmlNodeToToxyNode(this.HtmlNode.SelectSingleNode(selector));
        }
        public List<ToxyNode> SelectNodes(string selector)
        {
            HtmlAgilityPack.HtmlNodeCollection htmlNodeList = this.HtmlNode.SelectNodes(selector);
            if (htmlNodeList == null)
            {
                return null;
            }

            List<ToxyNode> toxyNodeList = new List<ToxyNode>();
            foreach (var item in htmlNodeList)
            {
                //todo:need to copy ChildNodes to the seleted node
                toxyNodeList.Add(ToxyNode.TransformHtmlNodeToToxyNode(item));
            }
            return toxyNodeList;
        }
        /// <summary>
        /// node string with name and attributes
        /// </summary>
        public string NodeString
        {
            get
            {
                StringBuilder sb = new StringBuilder();
                sb.Append("<");
                sb.Append(this.Name);
                for (int i = 0; i < this.Attributes.Count; i++)
                {
                    sb.Append(" ");
                    sb.Append(this.Attributes[i].Name);
                    if (this.Attributes[i].Value!=null)
                    {
                        sb.Append("=\"");
                        sb.Append(this.Attributes[i].Value);
                        sb.Append("\"");
                    }
                }
                sb.Append(">");
                return sb.ToString();
            }
        }
        public string InnerText
        {
            get{return HtmlNode.InnerText;}
        }
    }
}
