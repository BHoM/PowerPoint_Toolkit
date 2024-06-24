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
        public int ICreateSlide(PresentationPart presentationPart, ISlideCreate create)
        {
            return CreateSlide(presentationPart, create as dynamic);
        }

        private int CreateSlide(PresentationPart presentationPart, SlideCreate create)
        {
            SlideMasterPart slideMasterPart;

            // Get the slide master, and retreive the layout from the master.
            if (create.SlideMasterName.IsNullOrEmpty())
                slideMasterPart = presentationPart.SlideMasterParts.FirstOrDefault();
            else
                slideMasterPart = presentationPart.SlideMasterParts.SingleOrDefault(sm => sm.ThemePart.Theme.Name.Value.Equals(create.SlideMasterName, StringComparison.OrdinalIgnoreCase));

            if (slideMasterPart == null)
            {
                BH.Engine.Base.Compute.RecordError($"There was no slide master with name '{create.SlideMasterName}'.");
                return -1;
            }

            SlideLayoutPart slideLayoutPart = slideMasterPart.SlideLayoutParts.SingleOrDefault(sl => sl.SlideLayout.CommonSlideData.Name.Value.Equals(create.LayoutName, StringComparison.OrdinalIgnoreCase));

            if (slideLayoutPart == null)
            {
                BH.Engine.Base.Compute.RecordError($"The slide layout ({create.LayoutName}) could not be found in the master.");
                return -1;
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
            return presentationPart.SetSlideID(slidePart, create.SlideNumber - 1);
        }

        private int CreateSlide(PresentationPart presentationPart, ISlideCreate create)
        {
            BH.Engine.Base.Compute.RecordError($"Objects of type {create.GetType().FullName} are not currently supported for use in the PowerPointAdapter.");
            return -1;
        }
    }
}
