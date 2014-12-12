using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Toxy.Parsers
{
    public class HtmlDomParser : IDomParser
    {
        public ParserContext Context { get; set; }

        public HtmlDomParser(ParserContext context)
        {
            this.Context = context;
        }

        /// <summary>
        /// Parse HTML document
        /// Note:Context.Path must be absolute path,not relative path
        /// </summary>
        /// <returns></returns>
        public ToxyDom Parse()
        {
            if (!File.Exists(Context.Path))
                throw new FileNotFoundException("File " + Context.Path + " is not found");

            HtmlWeb hw = new HtmlWeb();
            HtmlDocument htmlDoc = hw.Load(Context.Path);
            HtmlNode docNode = htmlDoc.DocumentNode;
            ToxyNode root = ToxyNode.TransformHtmlNodeToToxyNode(docNode);

            Queue<KeyValuePair<HtmlNode, ToxyNode>> nodeQueue = new Queue<KeyValuePair<HtmlNode, ToxyNode>>();
            nodeQueue.Enqueue(new KeyValuePair<HtmlNode, ToxyNode>(docNode, root));
            while (nodeQueue.Count > 0)
            {
                KeyValuePair<HtmlNode, ToxyNode> pair = nodeQueue.Dequeue();
                HtmlNode htmlParentNode = pair.Key;
                ToxyNode toxyParentNode = pair.Value;
                foreach (HtmlNode htmlChildNode in htmlParentNode.ChildNodes)
                {
                    ToxyNode toxyChildNode = ToxyNode.TransformHtmlNodeToToxyNode(htmlChildNode);
                    if (htmlChildNode.Name == "#text")
                        toxyChildNode.Text = htmlChildNode.InnerText;
                    toxyParentNode.ChildrenNodes.Add(toxyChildNode);
                    nodeQueue.Enqueue(new KeyValuePair<HtmlNode, ToxyNode>(htmlChildNode, toxyChildNode));
                }
            }

            return new ToxyDom() { Root = root };
        }


    }
}
