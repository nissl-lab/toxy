using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TagLib.IFD;
using TagLib.Image;

namespace Toxy.Parsers
{
    public class ImageMetadataParser:IMetadataParser
    {
        public ImageMetadataParser(ParserContext context)
        {
            this.Context = context;
        }
        public ToxyMetadata Parse()
        {
            if (!System.IO.File.Exists(Context.Path))
                throw new System.IO.FileNotFoundException("File " + Context.Path + " is not found");

            ToxyMetadata metadatas = new ToxyMetadata();
            TagLib.File file = TagLib.Image.File.Create(Context.Path);
            if (file.Properties.PhotoHeight != 0)
                metadatas.Add("PhotoHeight", file.Properties.PhotoHeight);
            if (file.Properties.PhotoQuality != 0)
                metadatas.Add("PhotoQuality", file.Properties.PhotoQuality);
            if (file.Properties.PhotoWidth != 0)
                metadatas.Add("PhotoWidth", file.Properties.PhotoWidth);

            CombinedImageTag tags = file.Tag as CombinedImageTag;
            if (tags.Altitude != null)
                metadatas.Add("Altitude", (double)tags.Altitude);
            if(tags.Latitude!=null)
                metadatas.Add("Latitude", (double)tags.Latitude);
            if(tags.Model!=null)
                metadatas.Add("Model", tags.Model);
            if (tags.Make != null)
                metadatas.Add("Make", tags.Make);
            if (tags.Orientation != ImageOrientation.None)
                metadatas.Add("Make", tags.Orientation);
            if(tags.Longitude!=null)
                metadatas.Add("Longitude", (double)tags.Longitude);
            if(tags.Keywords.Length>0)
                metadatas.Add("Keywords", string.Join(",", tags.Keywords));
            if(tags.ISOSpeedRatings!=null)
                metadatas.Add("ISOSpeedRatings", (uint)tags.ISOSpeedRatings);
            if (tags.Creator != null)
                metadatas.Add("Creator", tags.Creator);
            if (!string.IsNullOrWhiteSpace(tags.Comment))
                metadatas.Add("Comment", tags.Comment);
            if (tags.Rating != null)
                metadatas.Add("Rating", (uint)tags.Rating);
            if (tags.Software != null)
                metadatas.Add("Software", tags.Software);
            if (tags.FNumber != null)
                metadatas.Add("FNumber", (double)tags.FNumber);
            if (tags.ExposureTime != null)
                metadatas.Add("ExposureTime", (double)tags.ExposureTime);
            if (tags.FocalLength != null)
                metadatas.Add("FocalLength", (double)tags.FocalLength);
            if (tags.FocalLengthIn35mmFilm != null)
                metadatas.Add("FocalLengthIn35mmFilm", (uint)tags.FocalLengthIn35mmFilm);
            if (tags.DateTime != null)
                metadatas.Add("DateTime", (DateTime)tags.DateTime);

            if ((file.TagTypes & TagLib.TagTypes.XMP) == TagLib.TagTypes.XMP)
            {
                TagLib.Xmp.XmpTag tagXmp = file.GetTag(TagLib.TagTypes.XMP) as TagLib.Xmp.XmpTag;
                foreach (TagLib.Xmp.XmpNode node in tagXmp.NodeTree.Children)
                {
                    if(!string.IsNullOrWhiteSpace(node.Value))
                    metadatas.Add(node.Name, node.Value);
                }
            }
            if ((file.TagTypes & TagLib.TagTypes.GifComment) == TagLib.TagTypes.GifComment)
            {
                TagLib.Gif.GifCommentTag tagGif = file.GetTag(TagLib.TagTypes.GifComment) as TagLib.Gif.GifCommentTag;
                metadatas.Add("GifComment", tagGif.Comment);
            }

            return metadatas;
        }

        public ParserContext Context
        {
            get;
            set;
        }
    }
}
