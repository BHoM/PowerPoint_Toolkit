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
        public void CreateSlide(PresentationPart presentationPart, SlideCreate command)
        {
            SlideMasterPart slideMasterPart;

            if (command.SlideMasterName.IsNullOrEmpty())
            {
                BH.Engine.Base.Compute.RecordNote("The slide master name was empty, using the first master.");
                slideMasterPart = presentationPart.SlideMasterParts.FirstOrDefault();
            }
            else
                slideMasterPart = presentationPart.SlideMasterParts.SingleOrDefault(mp => mp.SlideMaster.CommonSlideData.Name.Value.Equals(command.SlideMasterName, StringComparison.OrdinalIgnoreCase));

            if (slideMasterPart == null)
            {
                BH.Engine.Base.Compute.RecordError("There was no slide master in the presentation to get the layout from.");
                return;
            }

            SlideLayoutPart slideLayoutPart = slideMasterPart.SlideLayoutParts.SingleOrDefault(sl => sl.SlideLayout.CommonSlideData.Name.Value.Equals(command.LayoutName, StringComparison.OrdinalIgnoreCase));

            if (slideLayoutPart == null)
            {
                BH.Engine.Base.Compute.RecordError($"The slide layout ({command.LayoutName}) could not be found in the master.");
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
            presentationPart.SetSlideID(slidePart, command.SlideIndex);
        }
    }
}
