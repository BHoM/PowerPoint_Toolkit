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
using BH.oM.Base;
using BH.oM.PowerPoint;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BH.Adapter.PowerPoint
{
    public partial class PowerPointAdapter : BHoMAdapter
    {
        /***************************************************/
        /**** Adapter overload method                   ****/
        /***************************************************/

        protected override IEnumerable<IBHoMObject> IRead(Type type, IList ids, ActionConfig actionConfig = null)
        {
            if (type == typeof(SlideMasterInfo) || type == typeof(SlideLayoutInfo))
                return ReadMasterTemplateInfo();
            else if (type == typeof(SlideInfo))
                return ReadSlideInfo();
            else
                BH.Engine.Base.Compute.RecordError($"The type {type.FullName} is not supported for pulling from presentations.");

            return new List<IBHoMObject>();
        }

        /***************************************************/

        protected List<SlideInfo> ReadSlideInfo()
        {
            List<SlideInfo> objects = new List<SlideInfo>();

            using (MemoryStream memoryStream = GetTemplateMemoryStream())
            using (PresentationDocument presentationDoc = PresentationDocument.Open(memoryStream, true))
            {
                SlideIdList slideIdList = presentationDoc.PresentationPart?.Presentation.SlideIdList ?? new SlideIdList();
                int slideNumber = 1;

                foreach (SlideId slideId in slideIdList)
                {
                    SlideInfo info = new SlideInfo() { SlideNumber = slideNumber };
                    SlidePart slidePart = (SlidePart)presentationDoc.PresentationPart.GetPartById(slideId.RelationshipId);
                    
                    if (slidePart == null)
                    {
                        slideNumber++;
                        continue;
                    }

                    info.ElementNames = slidePart.Slide.CommonSlideData?.ShapeTree?.Elements<Shape>().Select(shape => shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name.Value).ToList() ?? new List<string>();
                    info.ElementNames.RemoveAll(x => x == null);

                    objects.Add(info);
                    slideNumber++;
                }
            }

            return objects;
        }

        /***************************************************/

        protected List<SlideMasterInfo> ReadMasterTemplateInfo()
        {
            List<SlideMasterInfo> objects = new List<SlideMasterInfo>();

            using (MemoryStream memoryStream = GetTemplateMemoryStream())
            using (PresentationDocument presentationDoc = PresentationDocument.Open(memoryStream, true))
            {
                IEnumerable<SlideMasterPart> slideMasterParts = presentationDoc.PresentationPart?.SlideMasterParts ?? new List<SlideMasterPart>();

                foreach (SlideMasterPart slideMasterPart in slideMasterParts)
                {
                    SlideMasterInfo info = new SlideMasterInfo();
                    info.Name = slideMasterPart.ThemePart.Theme.Name;

                    IEnumerable<SlideLayoutPart> slideLayoutParts = slideMasterPart.SlideLayoutParts;

                    foreach (SlideLayoutPart slideLayoutPart in slideLayoutParts)
                    {
                        SlideLayoutInfo templateInfo = new SlideLayoutInfo();
                        templateInfo.Name = slideLayoutPart.SlideLayout.CommonSlideData.Name;

                        // Get all the shape names from the layout, if at any point the properties are null, return a new List<string>();
                        templateInfo.ElementNames = slideLayoutPart.SlideLayout.CommonSlideData.ShapeTree?.Elements<Shape>().Select(shape => shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value).ToList() ?? new List<string>();

                        // Remove any elements that have no name (if that is possible, but better to make sure)
                        templateInfo.ElementNames.RemoveAll(x => x == null);

                        info.Templates.Add(templateInfo);
                    }

                    objects.Add(info);
                }
            }

            return objects;
        }

        /***************************************************/
    }
}