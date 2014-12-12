using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;

namespace Toxy.Parsers
{
    public class XMLDomParser:IDomParser
    {
        public XMLDomParser(ParserContext context) 
        {
            this.Context = context;
        }

        public ToxyDom Parse()
        {
            if (!File.Exists(Context.Path))
                throw new FileNotFoundException("File " + Context.Path + " is not found");

            XmlDocument doc = new XmlDocument();
            doc.Load(Context.Path);

            ToxyNode rootNode = ConvertToToxyNode(doc.DocumentElement);
            ToxyDom dom = new ToxyDom();
            dom.Root = rootNode;
            AppendChildren(rootNode, doc.DocumentElement);
            return dom;
        }
        void AppendChildren(ToxyNode tnode, XmlNode ele)
        {
            if (ele.ChildNodes.Count == 0)
                return;

            foreach (XmlNode child in ele.ChildNodes)
            { 
                ToxyNode x = ConvertToToxyNode(child);
                tnode.ChildrenNodes.Add(x);
                AppendChildren(x, child);
            }
        }
        ToxyNode ConvertToToxyNode(XmlNode ele)
        {
            ToxyNode tnode = new ToxyNode();
            tnode.Name = ele.Name;
            if (ele.Name == "#text")
            {
                tnode.Text = ele.InnerText;
                return tnode;
            }
            if (ele.Attributes != null)
            {
                foreach (XmlAttribute attr in ele.Attributes)
                    tnode.Attributes.Add(new ToxyAttribute(attr.Name, attr.Value));
            }
            return tnode;
        }

        public ParserContext Context
        {
            get;
            set;
        }
    }
}
