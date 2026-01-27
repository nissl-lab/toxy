namespace Toxy.Parsers
{
    public class AudioMetadataParser:IMetadataParser
    {
        public AudioMetadataParser(ParserContext context)
        {
            this.Context = context;
        }
        public ToxyMetadata Parse()
        {
            if (!System.IO.File.Exists(Context.Path))
                throw new System.IO.FileNotFoundException("File " + Context.Path + " is not found");
            ToxyMetadata metadatas = new ToxyMetadata();
            TagLib.File file = TagLib.Audible.File.Create(Context.Path);
            if (file.Properties.AudioSampleRate != 0)
                metadatas.Add("AudioSampleRate", file.Properties.AudioSampleRate);
            if (file.Properties.AudioChannels != 0)
                metadatas.Add("AudioChannels", file.Properties.AudioChannels);
            if (file.Properties.AudioBitrate != 0)
                metadatas.Add("AudioBitrate", file.Properties.AudioBitrate);
            if (file.Properties.BitsPerSample != 0)
                metadatas.Add("BitsPerSample", file.Properties.BitsPerSample);
            if (file.Properties.Description != null)
                metadatas.Add("Description", file.Properties.Description);
            if (file.Length>0)
                metadatas.Add("FileLength", file.Length);


            TagLib.Tag tag1 = file.Tag;
            if (tag1 != null)
            {
                if (tag1.Album != null)
                    metadatas.Add("Album", tag1.Album);
                if (tag1.AlbumArtists.Length > 0)
                    metadatas.Add("AlbumArtists", tag1.JoinedAlbumArtists);
                if (tag1.BeatsPerMinute != 0)
                    metadatas.Add("BeatsPerMinute", tag1.BeatsPerMinute);
                if (tag1.Composers.Length > 0)
                    metadatas.Add("Composers", tag1.JoinedComposers);
                if (tag1.Genres.Length > 0)
                    metadatas.Add("Genres", tag1.JoinedGenres);
                if (tag1.Performers.Length > 0)
                    metadatas.Add("Performers", tag1.JoinedPerformers);
                if (tag1.Comment != null)
                    metadatas.Add("Comment", tag1.Comment);
                if (tag1.Conductor != null)
                    metadatas.Add("Conductor", tag1.Conductor);
                if (tag1.Copyright != null)
                    metadatas.Add("Copyright", tag1.Copyright);
                if (tag1.Disc != 0)
                    metadatas.Add("Disc", tag1.Disc);
                if (tag1.DiscCount != 0)
                    metadatas.Add("DiscCount", tag1.DiscCount);
                if (tag1.Lyrics != null)
                    metadatas.Add("Lyrics", tag1.Lyrics);
                if (tag1.Title != null)
                    metadatas.Add("Title", tag1.Title);
                if (tag1.Track != 0)
                    metadatas.Add("Track", tag1.Track);
                if (tag1.TrackCount != 0)
                    metadatas.Add("TrackCount", tag1.TrackCount);
                if (tag1.Year != 0)
                    metadatas.Add("Year", tag1.Year);
                if (tag1.AmazonId != null)
                    metadatas.Add("AmazonId", tag1.AmazonId);
                if (tag1.MusicIpId != null)
                    metadatas.Add("MusicIpId", tag1.MusicIpId);
            }
            
            return metadatas;
        }

        public ParserContext Context
        {
            get;set;
        }
    }
}
