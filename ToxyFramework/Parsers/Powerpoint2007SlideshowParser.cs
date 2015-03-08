using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Toxy.Parsers
{
    public class Powerpoint2007SlideshowParser: ISlideshowParser
    {
        public Powerpoint2007SlideshowParser(ParserContext context)
        {
            this.Context = context;
        }
        public ToxySlideshow Parse()
        {
            if (!File.Exists(Context.Path))
                throw new FileNotFoundException("File " + Context.Path + " is not found");

            ToxySlideshow ss = new ToxySlideshow();

            using (PresentationDocument ppt = PresentationDocument.Open(Context.Path, false))
            {
                // Get the relationship ID of the first slide.
                PresentationPart part = ppt.PresentationPart;
                DocumentFormat.OpenXml.OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;
                for (int index = 0; index < slideIds.Count; index++)
                {
                    string relId = (slideIds[index] as SlideId).RelationshipId;
                    relId = (slideIds[index] as SlideId).RelationshipId;

                    // Get the slide part from the relationship ID.
                    SlidePart slide = (SlidePart)part.GetPartById(relId);
                    var tslide= Parse(slide);
                    ss.Slides.Add(tslide);
                }
            }
            return ss;
        }
        ToxySlide Parse(SlidePart slidePart)
        {
            if (slidePart == null)
            {
                throw new ArgumentNullException("slidePart");
            }
            ToxySlide slide =null;
            if (slidePart.Slide != null)
            {
                slide=new ToxySlide();
                // Iterate through all the paragraphs in the slide.
                foreach (DocumentFormat.OpenXml.Drawing.Paragraph paragraph in
                    slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>())
                {
                    foreach (DocumentFormat.OpenXml.Drawing.Text text in
                        paragraph.Descendants<DocumentFormat.OpenXml.Drawing.Text>())
                    {
                        slide.Texts.Add(text.Text);
                    }
                }
            }
            return slide;
        }

        public ToxySlide Parse(int slideIndex)
        {
            if (!File.Exists(Context.Path))
                throw new FileNotFoundException("File " + Context.Path + " is not found");

            using (PresentationDocument ppt = PresentationDocument.Open(Context.Path, false))
            {
                // Get the relationship ID of the first slide.
                PresentationPart part = ppt.PresentationPart;
                DocumentFormat.OpenXml.OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;
                if (slideIds.Count - 1 < slideIndex)
                {
                    throw new ArgumentOutOfRangeException(string.Format("This file only contains {0} slide(s).", slideIds.Count));
                }
                string relId = (slideIds[slideIndex] as SlideId).RelationshipId;
                relId = (slideIds[slideIndex] as SlideId).RelationshipId;

                // Get the slide part from the relationship ID.
                SlidePart slide = (SlidePart)part.GetPartById(relId);
                var tslide = Parse(slide);
                return tslide;
            }
        }

        public ParserContext Context
        {
            get;
            set;
        }
    }
}
