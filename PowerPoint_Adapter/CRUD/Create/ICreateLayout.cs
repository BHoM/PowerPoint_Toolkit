using BH.oM.PowerPoint;
using BH.oM.PowerPoint.Layout;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Text;

namespace BH.Adapter.PowerPoint
{
    public partial class PowerPointAdapter : BHoMAdapter
    {
        public void ICreateLayout(PresentationPart presentationPart, ISlideLayout layout)
        {
            BH.Engine.Base.Compute.StartSuppressRecordingEvents(suppressNotes: true);
            CreateLayout(presentationPart, layout as dynamic);
            BH.Engine.Base.Compute.StopSuppressRecordingEvents();
        }

        private void CreateLayout(PresentationPart presentationPart, Blank blank)
        {
            SlideCreate create = new SlideCreate() { SlideNumber = -1, SlideMasterName = "Office Theme", LayoutName = "Blank" };
            CreateSlide(presentationPart, create);
        }

        private void CreateLayout(PresentationPart presentationPart, Comparison comparison)
        {
            SlideCreate create = new SlideCreate() { SlideNumber = -1, SlideMasterName = "Office Theme", LayoutName = "Comparison" };
            int slideNumber = CreateSlide(presentationPart, create);
            SlidePart slidePart = GetSlide(presentationPart, slideNumber - 1);

            IUpdateSlide(slidePart, comparison.TitleElement.IToUpdate(slideNumber, "Title 1"));
            IUpdateSlide(slidePart, comparison.LeftSubtitleElement.IToUpdate(slideNumber, "Text Placeholder 2"));
            IUpdateSlide(slidePart, comparison.RightSubtitleElement.IToUpdate(slideNumber, "Text Placeholder 4"));
            IUpdateSlide(slidePart, comparison.LeftContentElement.IToUpdate(slideNumber, "Content Placeholder 3"));
            IUpdateSlide(slidePart, comparison.RightContentElement.IToUpdate(slideNumber, "Content Placeholder 5"));
        }

        private void CreateLayout(PresentationPart presentationPart, ContentWithCaption contentWithCaption)
        {
            SlideCreate create = new SlideCreate() { SlideNumber = -1, SlideMasterName = "Office Theme", LayoutName = "Content with Caption" };
            int slideNumber = CreateSlide(presentationPart, create);
            SlidePart slidePart = GetSlide(presentationPart, slideNumber - 1);

            IUpdateSlide(slidePart, contentWithCaption.TitleElement.IToUpdate(slideNumber, "Title 1"));
            IUpdateSlide(slidePart, contentWithCaption.CaptionElement.IToUpdate(slideNumber, "Text Placeholder 3"));
            IUpdateSlide(slidePart, contentWithCaption.ContentElement.IToUpdate(slideNumber, "Content Placeholder 2"));
        }

        private void CreateLayout(PresentationPart presentationPart, PictureWithCaption pictureWithCaption)
        {
            SlideCreate create = new SlideCreate() { SlideNumber = -1, SlideMasterName = "Office Theme", LayoutName = "Picture with Caption" };
            int slideNumber = CreateSlide(presentationPart, create);
            SlidePart slidePart = GetSlide(presentationPart, slideNumber - 1);

            IUpdateSlide(slidePart, pictureWithCaption.TitleElement.IToUpdate(slideNumber, "Title 1"));
            IUpdateSlide(slidePart, pictureWithCaption.CaptionElement.IToUpdate(slideNumber, "Text Placeholder 3"));
            IUpdateSlide(slidePart, pictureWithCaption.ImageElement.IToUpdate(slideNumber, "Picture Placeholder 2"));
        }

        private void CreateLayout(PresentationPart presentationPart, SectionHeader sectionHeader)
        {
            SlideCreate create = new SlideCreate() { SlideNumber = -1, SlideMasterName = "Office Theme", LayoutName = "Section Header" };
            int slideNumber = CreateSlide(presentationPart, create);
            SlidePart slidePart = GetSlide(presentationPart, slideNumber - 1);

            IUpdateSlide(slidePart, sectionHeader.TitleElement.IToUpdate(slideNumber, "Title 1"));
            IUpdateSlide(slidePart, sectionHeader.SubtitleElement.IToUpdate(slideNumber, "Text Placeholder 2"));
        }

        private void CreateLayout(PresentationPart presentationPart, Title title)
        {
            SlideCreate create = new SlideCreate() { SlideNumber = -1, SlideMasterName = "Office Theme", LayoutName = "Title Slide" };
            int slideNumber = CreateSlide(presentationPart, create);
            SlidePart slidePart = GetSlide(presentationPart, slideNumber - 1);

            IUpdateSlide(slidePart, title.TitleElement.IToUpdate(slideNumber, "Title 1"));
            IUpdateSlide(slidePart, title.SubtitleElement.IToUpdate(slideNumber, "Subtitle 2"));
        }

        private void CreateLayout(PresentationPart presentationPart, TitleOnly titleOnly)
        {
            SlideCreate create = new SlideCreate() { SlideNumber = -1, SlideMasterName = "Office Theme", LayoutName = "Title Only" };
            int slideNumber = CreateSlide(presentationPart, create);
            SlidePart slidePart = GetSlide(presentationPart, slideNumber - 1);

            IUpdateSlide(slidePart, titleOnly.TitleElement.IToUpdate(slideNumber, "Title 1"));
        }

        private void CreateLayout(PresentationPart presentationPart, TitleWithContent titleWithContent)
        {
            SlideCreate create = new SlideCreate() { SlideNumber = -1, SlideMasterName = "Office Theme", LayoutName = "Title and Content" };
            int slideNumber = CreateSlide(presentationPart, create);
            SlidePart slidePart = GetSlide(presentationPart, slideNumber - 1);

            IUpdateSlide(slidePart, titleWithContent.TitleElement.IToUpdate(slideNumber, "Title 1"));
            IUpdateSlide(slidePart, titleWithContent.ContentElement.IToUpdate(slideNumber, "Content Placeholder 2"));
        }

        private void CreateLayout(PresentationPart presentationPart, TitleWithVerticalText titleWithVerticalText)
        {
            SlideCreate create = new SlideCreate() { SlideNumber = -1, SlideMasterName = "Office Theme", LayoutName = "Title and Vertical Text" };
            int slideNumber = CreateSlide(presentationPart, create);
            SlidePart slidePart = GetSlide(presentationPart, slideNumber - 1);

            IUpdateSlide(slidePart, titleWithVerticalText.TitleElement.IToUpdate(slideNumber, "Title 1"));
            IUpdateSlide(slidePart, titleWithVerticalText.TextElement.IToUpdate(slideNumber, "Vertical Text Placeholder 2"));
        }

        private void CreateLayout(PresentationPart presentationPart, TwoContent twoContent)
        {
            SlideCreate create = new SlideCreate() { SlideNumber = -1, SlideMasterName = "Office Theme", LayoutName = "Two Content" };
            int slideNumber = CreateSlide(presentationPart, create);
            SlidePart slidePart = GetSlide(presentationPart, slideNumber - 1);

            IUpdateSlide(slidePart, twoContent.TitleElement.IToUpdate(slideNumber, "Title 1"));
            IUpdateSlide(slidePart, twoContent.LeftContentElement.IToUpdate(slideNumber, "Content Placeholder 2"));
            IUpdateSlide(slidePart, twoContent.RightContentElement.IToUpdate(slideNumber, "Content Placeholder 3"));
        }

        private void CreateLayout(PresentationPart presentationPart, VerticalTitleAndText verticalTitleAndText)
        {
            SlideCreate create = new SlideCreate() { SlideNumber = -1, SlideMasterName = "Office Theme", LayoutName = "Vertical Title and Text" };
            int slideNumber = CreateSlide(presentationPart, create);
            SlidePart slidePart = GetSlide(presentationPart, slideNumber - 1);

            IUpdateSlide(slidePart, verticalTitleAndText.TitleElement.IToUpdate(slideNumber, "Vertical Title 1"));
            IUpdateSlide(slidePart, verticalTitleAndText.TextElement.IToUpdate(slideNumber, "Vertical Text Placeholder 2"));
        }

        private void CreateLayout(PresentationPart presentationPart, ISlideLayout layout)
        {
            BH.Engine.Base.Compute.RecordError($"Layouts of type {layout.GetType().FullName} are not supported for use in the PowerPointAdapter.");
        }
    }
}
