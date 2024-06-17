using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BH.Adapter.PowerPoint
{
    public static partial class Compute
    {
        public static void SetSlideID(this PresentationPart presentationPart, SlidePart slidePart, int slideIndex)
        {
            SlideIdList slideIDList = presentationPart.Presentation.SlideIdList;

            // Create a new slide ID list if there isn't one already.
            if (slideIDList == null)
            {
                slideIDList = new SlideIdList();
                presentationPart.Presentation.SlideIdList = slideIDList;
            }

            uint newID = slideIDList.ChildElements.Count() == 0 ? 256 : slideIDList.GetMaxSlideID() + 1;

            if (slideIndex < 0 || slideIndex >= slideIDList.Count())
            {
                BH.Engine.Base.Compute.RecordNote($"The slide index ({slideIndex}) was outside the range of slides ({slideIDList.Count()}). Appending the slide to the end of the presentation.");
                SlideId slideID = new SlideId() { Id = newID, RelationshipId = presentationPart.GetIdOfPart(slidePart) };
                slideIDList.AppendChild(slideID);
            }
            else
            {
                SlideId nextSlideID = (SlideId)slideIDList.ChildElements[slideIndex];
                SlideId slideID = new SlideId() { Id = newID, RelationshipId = presentationPart.GetIdOfPart(slidePart) };
                slideIDList.InsertBefore(slideID, nextSlideID);
            }
        }

        public static uint GetMaxSlideID(this SlideIdList slideIDList)
        {
            uint maxSlideID = 0;
            if (slideIDList.ChildElements.Count() > 0)
                maxSlideID = slideIDList.ChildElements
                    .Cast<SlideId>()
                    .Max(x => x.Id.Value);
            return maxSlideID;
        }
    }
}
