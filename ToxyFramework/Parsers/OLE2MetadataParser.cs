using NPOI.HPSF;
using NPOI.POIFS.FileSystem;
using System.IO;

namespace Toxy.Parsers
{
    public class OLE2MetadataParser : IMetadataParser
    {
        public OLE2MetadataParser(ParserContext context)
        {
            this.Context = context;
        }
        public ToxyMetadata Parse()
        {
            if (!System.IO.File.Exists(Context.Path))
                throw new System.IO.FileNotFoundException("File " + Context.Path + " is not found");

            ToxyMetadata metadata = new ToxyMetadata();
            using (Stream stream = File.OpenRead(Context.Path))
            {
                POIFSFileSystem poifs = new NPOI.POIFS.FileSystem.POIFSFileSystem(stream);
                DirectoryNode root = poifs.Root;

                SummaryInformation si = null;
                DocumentSummaryInformation dsi = null;
                if (root.HasEntry(SummaryInformation.DEFAULT_STREAM_NAME))
                {
                    PropertySet ps = GetPropertySet(root, SummaryInformation.DEFAULT_STREAM_NAME);
                    if (ps is SummaryInformation)
                    {
                        si = ps as SummaryInformation;

                        if (si.Author != null)
                            metadata.Add("Author", si.Author);
                        if (si.ApplicationName != null)
                            metadata.Add("ApplicationName", si.ApplicationName);
                        if (si.ClassID != null)
                            metadata.Add("ClassID", si.ClassID.ToString());
                        if (si.CharCount != 0)
                            metadata.Add("CharCount", si.CharCount);
                        if (si.ByteOrder != 0)
                            metadata.Add("ByteOrder", si.ByteOrder);
                        if (si.CreateDateTime != null)
                            metadata.Add("CreateDateTime", si.CreateDateTime);
                        if (si.LastAuthor != null)
                            metadata.Add("LastAuthor", si.LastAuthor);
                        if (si.Keywords != null)
                            metadata.Add("Keywords", si.Keywords);
                        if (si.LastPrinted != null)
                            metadata.Add("LastPrinted", si.LastPrinted);
                        if (si.LastSaveDateTime != null)
                            metadata.Add("LastSaveDateTime", si.LastSaveDateTime);
                        if (si.PageCount != 0)
                            metadata.Add("PageCount", si.PageCount);
                        if (si.WordCount != 0)
                            metadata.Add("WordCount", si.WordCount);
                        if (si.Comments != null)
                            metadata.Add("Comments", si.Comments);
                        if (si.EditTime != 0)
                            metadata.Add("EditTime", si.EditTime);
                        if (si.RevNumber != null)
                            metadata.Add("RevNumber", si.RevNumber);
                        if (si.Security != 0)
                            metadata.Add("Security", si.Security);
                        if (si.Subject != null)
                            metadata.Add("Subject", si.Subject);
                        if (si.Title != null)
                            metadata.Add("Title", si.Title);
                        if (si.Template != null)
                            metadata.Add("Template", si.Template);
                    }
                }

                if (root.HasEntry(DocumentSummaryInformation.DEFAULT_STREAM_NAME))
                {
                    PropertySet ps = GetPropertySet(root, DocumentSummaryInformation.DEFAULT_STREAM_NAME);
                    if (ps is DocumentSummaryInformation)
                    {
                        dsi = ps as DocumentSummaryInformation;
                        if (dsi.ByteCount > 0)
                        {
                            metadata.Add("ByteCount", dsi.ByteCount);
                        }
                        else if (dsi.Company != null)
                        {
                            metadata.Add("Company", dsi.Company);
                        }
                        else if (dsi.Format > 0)
                        {
                            metadata.Add("Format", dsi.Format);
                        }
                        else if (dsi.LineCount > 0)
                        {
                            metadata.Add("LineCount", dsi.LineCount);
                        }
                        else if (dsi.LinksDirty)
                        {
                            metadata.Add("LinksDirty", true);
                        }
                        else if (dsi.Manager != null)
                        {
                            metadata.Add("Manager", dsi.Manager);
                        }
                        else if (dsi.NoteCount > 0)
                        {
                            metadata.Add("NoteCount", dsi.NoteCount);
                        }
                        else if (dsi.Scale)
                        {
                            metadata.Add("Scale", dsi.Scale);
                        }
                        else if (dsi.HiddenCount > 0)
                        {
                            metadata.Add("HiddenCount", dsi.Company);
                        }
                        else if (dsi.MMClipCount > 0)
                        {
                            metadata.Add("MMClipCount", dsi.MMClipCount);
                        }
                        else if (dsi.ParCount > 0)
                        {
                            metadata.Add("ParCount", dsi.ParCount);
                        }
                    }
                }
            }
            return metadata;
        }

        PropertySet GetPropertySet(DirectoryNode dir, string name)
        {
            DocumentInputStream dis = dir.CreateDocumentInputStream(name);
            PropertySet set = PropertySetFactory.Create(dis);
            return set;
        }
        public ParserContext Context
        {
            get;
            set;
        }
    }
}
