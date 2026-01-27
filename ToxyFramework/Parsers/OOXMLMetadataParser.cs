using NPOI;
using NPOI.OpenXml4Net.OPC;

namespace Toxy.Parsers
{
    public class OOXMLMetadataParser : IMetadataParser
    {
        public OOXMLMetadataParser(ParserContext context)
        {
            this.Context = context;
        }

        public ParserContext Context
        {
            get;
            set;
        }

        public ToxyMetadata Parse()
        {
            if (!System.IO.File.Exists(Context.Path))
                throw new System.IO.FileNotFoundException("File " + Context.Path + " is not found");

            ToxyMetadata metadata = new ToxyMetadata();
            OPCPackage pack = null;
            try
            {
                pack = OPCPackage.Open(Context.Path, PackageAccess.READ);
                POIXMLProperties props = new POIXMLProperties(pack);
                if (props.CoreProperties != null)
                {
                    if (props.CoreProperties.Title != null)
                    {
                        metadata.Add("Title", props.CoreProperties.Title);
                    }
                    else if (props.CoreProperties.Identifier != null)
                    {
                        metadata.Add("Identifier", props.CoreProperties.Identifier);
                    }
                    else if (props.CoreProperties.Keywords != null)
                    {
                        metadata.Add("Keywords", props.CoreProperties.Keywords);
                    }
                    else if (props.CoreProperties.Revision != null)
                    {
                        metadata.Add("Revision", props.CoreProperties.Revision);
                    }
                    else if (props.CoreProperties.Subject != null)
                    {
                        metadata.Add("Subject", props.CoreProperties.Subject);
                    }
                    else if (props.CoreProperties.Modified != null)
                    {
                        metadata.Add("Modified", props.CoreProperties.Modified);
                    }
                    else if (props.CoreProperties.LastPrinted != null)
                    {
                        metadata.Add("LastPrinted", props.CoreProperties.LastPrinted);
                    }
                    else if (props.CoreProperties.Created != null)
                    {
                        metadata.Add("Created", props.CoreProperties.Created);
                    }
                    else if (props.CoreProperties.Creator != null)
                    {
                        metadata.Add("Creator", props.CoreProperties.Creator);
                    }
                    else if (props.CoreProperties.Description != null)
                    {
                        metadata.Add("Description", props.CoreProperties.Description);
                    }
                }
                if (props.ExtendedProperties != null && props.ExtendedProperties.props != null)
                {
                    var extProps = props.ExtendedProperties.props.GetProperties();
                    if (extProps.Application != null)
                    {
                        metadata.Add("Application", extProps.Application);
                    }
                    if (extProps.AppVersion != null)
                    {
                        metadata.Add("AppVersion", extProps.AppVersion);
                    }
                    if (extProps.Characters > 0)
                    {
                        metadata.Add("Characters", extProps.Characters);
                    }
                    if (extProps.CharactersWithSpaces > 0)
                    {
                        metadata.Add("CharactersWithSpaces", extProps.CharactersWithSpaces);
                    }
                    if (extProps.Company != null)
                    {
                        metadata.Add("Company", extProps.Company);
                    }
                    if (extProps.Lines > 0)
                    {
                        metadata.Add("Lines", extProps.Lines);
                    }
                    if (extProps.Manager != null)
                    {
                        metadata.Add("Manager", extProps.Manager);
                    }
                    if (extProps.Notes > 0)
                    {
                        metadata.Add("Notes", extProps.Notes);
                    }
                    if (extProps.Pages > 0)
                    {
                        metadata.Add("Pages", extProps.Pages);
                    }
                    if (extProps.Paragraphs > 0)
                    {
                        metadata.Add("Paragraphs", extProps.Paragraphs);
                    }
                    if (extProps.Words > 0)
                    {
                        metadata.Add("Words", extProps.Words);
                    }
                    if (extProps.TotalTime > 0)
                    {
                        metadata.Add("TotalTime", extProps.TotalTime);
                    }
                }
            }
            finally
            {
                pack?.Close();
            }
            return metadata;
        }
    }
}
