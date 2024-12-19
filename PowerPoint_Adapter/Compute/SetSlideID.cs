/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2025, the respective contributors. All rights reserved.
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
        public static int SetSlideID(this PresentationPart presentationPart, SlidePart slidePart, int slideIndex)
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
                // Dont show the note if the slide number is one above the number of slides, as appending should be expected behaviour.
                if (slideIndex > slideIDList.Count())
                    BH.Engine.Base.Compute.RecordNote($"The slide number given ({slideIndex + 1}) was outside the range of slides ({slideIDList.Count()}). Appending the slide to the end of the presentation.");

                SlideId slideID = new SlideId() { Id = newID, RelationshipId = presentationPart.GetIdOfPart(slidePart) };
                slideIDList.AppendChild(slideID);
                return slideIDList.Count();
            }
            else
            {
                SlideId nextSlideID = (SlideId)slideIDList.ChildElements[slideIndex];
                SlideId slideID = new SlideId() { Id = newID, RelationshipId = presentationPart.GetIdOfPart(slidePart) };
                slideIDList.InsertBefore(slideID, nextSlideID);
                return slideIndex + 1;
            }
        }

        public static uint GetMaxSlideID(this SlideIdList slideIDList)
        {
            uint maxSlideID = 0;
            if (slideIDList.ChildElements.Count() > 0)
                maxSlideID = slideIDList.ChildElements
                    .OfType<SlideId>()
                    .Max(x => x.Id.Value);
            return maxSlideID;
        }
    }
}

