/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2024, the respective contributors. All rights reserved.
 *
 * Each contributor holds copyright over their respective contributions.
 * The project versioning (Git) records all such contribution source information.
 *                                           
 *                                                                              
 * The BHoM is free software: you can redistribute it and/or modify         
 * it under the terms of the GNU Lesser General Public License as published by  
 * the Free Software Foundation, either version 3.0 of the License, or          
 * (at your option) any later version.                                          
 *                                                                              
 * The BHoM is distributed in the hope that it will be useful,              
 * but WITHOUT ANY WARRANTY; without even the implied warranty of               
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the                 
 * GNU Lesser General Public License for more details.                          
 *                                                                            
 * You should have received a copy of the GNU Lesser General Public License     
 * along with this code. If not, see <https://www.gnu.org/licenses/lgpl-3.0.html>.      
 */

using BH.oM.Adapter;
using BH.oM.PowerPoint;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BH.Adapter.PowerPoint
{
    public partial class PowerPointAdapter : BHoMAdapter
    {
        /***************************************************/
        /**** Private Methods                           ****/
        /***************************************************/

        private int DeleteSlides(PresentationPart presentationPart, DeleteSlides deleteSlide)
        {
            if (deleteSlide.SlideNumbers.Count == 0)
                return 0;

            return DeleteSlideFromPresentation(presentationPart, deleteSlide.SlideNumbers);
        }

        /***************************************************/

        // Delete the specified slide from the presentation.
        private int DeleteSlideFromPresentation(PresentationPart presentationPart, List<int> slideNumbers)
        {
            //Method modified version taken from https://learn.microsoft.com/en-us/office/open-xml/presentation/how-to-delete-a-slide-from-a-presentation?tabs=cs-0%2Ccs-1%2Ccs-2%2Ccs-3%2Ccs-4%2Ccs-5%2Ccs-6%2Ccs-7%2Ccs-8%2Ccs-9%2Ccs
            // Use the CountSlides sample to get the number of slides in the presentation.
            int slidesCount = presentationPart.SlideParts.Count();

            // Get the presentation from the presentation part.
            Presentation presentation = presentationPart?.Presentation;

            // Get the list of slide IDs in the presentation.
            SlideIdList slideIdList = presentation?.SlideIdList;

            List<SlideId> slideIds = new List<SlideId>();
            List<string> slideRelIds = new List<string>();

            foreach (int slideNum in slideNumbers)
            {
                int slideIndex = slideNum - 1;

                if (slideIndex < 0 || slideIndex >= slidesCount)
                {
                    BH.Engine.Base.Compute.RecordError($"Presentation does not contain slide with number {slideNum}. Slide not deleted.");
                    continue;
                }

                // Get the slide ID of the specified slide
                SlideId slideId = slideIdList?.ChildElements[slideIndex] as SlideId;

                // Get the relationship ID of the slide.
                string slideRelId = slideId?.RelationshipId;

                // If there's no relationship ID, there's no slide to delete.
                if (slideRelId != null)
                {
                    slideIds.Add(slideId);
                    slideRelIds.Add(slideRelId);
                }
            }

            for (int i = 0; i < slideIds.Count; i++)
            {

                SlideId slideId = slideIds[i];
                string slideRelId = slideRelIds[i];

                // Remove the slide from the slide list.
                slideIdList?.RemoveChild(slideId);

                //
                // Remove references to the slide from all custom shows.
                if (presentation?.CustomShowList != null)
                {
                    // Iterate through the list of custom shows.
                    foreach (var customShow in presentation.CustomShowList.Elements<CustomShow>())
                    {
                        if (customShow.SlideList != null)
                        {
                            // Declare a link list of slide list entries.
                            LinkedList<SlideListEntry> slideListEntries = new LinkedList<SlideListEntry>();
                            foreach (SlideListEntry slideListEntry in customShow.SlideList.Elements())
                            {
                                // Find the slide reference to remove from the custom show.
                                if (slideListEntry.Id != null && slideListEntry.Id == slideRelId)
                                {
                                    slideListEntries.AddLast(slideListEntry);
                                }
                            }

                            // Remove all references to the slide from the custom show.
                            foreach (SlideListEntry slideListEntry in slideListEntries)
                            {
                                customShow.SlideList.RemoveChild(slideListEntry);
                            }
                        }
                    }
                }
            }

            return slideIds.Count;
        }


        /***************************************************/
    }
}


