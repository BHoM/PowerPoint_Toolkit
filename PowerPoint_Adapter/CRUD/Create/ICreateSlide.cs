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
            SlideMasterPart slideMasterPart;

            // Get the slide master, and retreive the layout from the master.
            if (create.SlideMasterName.IsNullOrEmpty())
                slideMasterPart = presentationPart.SlideMasterParts.FirstOrDefault();
            else
                slideMasterPart = presentationPart.SlideMasterParts.SingleOrDefault(sm => sm.ThemePart.Theme.Name.Value.Equals(create.SlideMasterName, StringComparison.OrdinalIgnoreCase));

            if (slideMasterPart == null)
            {
                BH.Engine.Base.Compute.RecordError($"There was no slide master with name '{create.SlideMasterName}'.");
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

            // Try to insert the slide at the position given, and get the slide number.
            int slideNumber = presentationPart.SetSlideID(slidePart, create.SlideNumber - 1);

            List<ISlideUpdate> slideUpdates = create.SlideUpdates ?? new List<ISlideUpdate>();

            // If the slide number of the part is not the same as the slide number in the create, then set the updates to that number instead.
            // Run all slide updates
            foreach (ISlideUpdate update in slideUpdates)
            {
                if (slideNumber != update.SlideNumber)
                    update.SlideNumber = slideNumber;
                IUpdateSlide(slidePart, update);
            }
        }

        private void CreateSlide(PresentationPart presentationPart, ISlideCreate create)
        {
            BH.Engine.Base.Compute.RecordError($"Objects of type {create.GetType().FullName} are not currently supported for use in the PowerPointAdapter.");
            return;
        }
    }
}
