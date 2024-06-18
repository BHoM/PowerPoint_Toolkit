using BH.Engine.Base;
using BH.oM.PowerPoint;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BH.Adapter.PowerPoint
{
    public partial class PowerPointAdapter: BHoMAdapter
    {
        public void ICreateSlide(PresentationPart presentationPart, ISlideCreate create)
        {
            CreateSlide(presentationPart, create as dynamic);
        }

        private void CreateSlide(PresentationPart presentationPart, SlideCreate create)
        {
            SlideMasterPart slideMasterPart = presentationPart.SlideMasterParts.FirstOrDefault();

            if (slideMasterPart == null)
            {
                BH.Engine.Base.Compute.RecordError("There was no slide master in the presentation to get the layout from.");
                return;
            }

            SlideLayoutPart slideLayoutPart = slideMasterPart.SlideLayoutParts.SingleOrDefault(sl => sl.SlideLayout.CommonSlideData.Name.Value.Equals(create.LayoutName, StringComparison.OrdinalIgnoreCase));

            if (slideLayoutPart == null)
            {
                BH.Engine.Base.Compute.RecordError($"The slide layout ({create.LayoutName}) could not be found in the master.");
                return;
            }

            // Create a new slide and add to presentation.
            Slide slide = new Slide();
            SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();
            slide.Save(slidePart);

            // Copy the slide layout to the created slide.
            slidePart.AddPart(slideLayoutPart);
            slidePart.Slide.CommonSlideData = (CommonSlideData)slideLayoutPart.SlideLayout.CommonSlideData.Clone();

            // Remove extra images placed on top of layout images, as reference issues occur from cloning slide layouts with images in the layout.
            foreach (Picture picture in slidePart.Slide.CommonSlideData.ShapeTree.Descendants<Picture>())
                slidePart.Slide.CommonSlideData.ShapeTree.RemoveChild(picture);

            // Insert the slide at the position given.
            string id = slideMasterPart.GetIdOfPart(slideLayoutPart);
            presentationPart.SetSlideID(slidePart, create.SlideNumber - 1);
        }

        private void CreateSlide(PresentationPart presentationPart, ISlideCreate create)
        {
            BH.Engine.Base.Compute.RecordError($"Objects of type {create.GetType().FullName} are not currently supported for creating slides.");
        }
    }
}
