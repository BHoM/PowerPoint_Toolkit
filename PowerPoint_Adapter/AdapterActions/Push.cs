/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2023, the respective contributors. All rights reserved.
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

using BH.Engine.Adapter;
using BH.oM.Adapter;
using BH.oM.Base;
using BH.oM.Data.Collections;
using BH.oM.PowerPoint;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace BH.Adapter.PowerPoint
{
    public partial class PowerPointAdapter : BHoMAdapter
    {
        /***************************************************/
        /**** Public Overrides                          ****/
        /***************************************************/

        public override List<object> Push(IEnumerable<object> objects, string tag = "", PushType pushType = PushType.AdapterDefault, ActionConfig actionConfig = null)
        {
            if (objects == null || !objects.Any())
            {
                BH.Engine.Base.Compute.RecordError("No objects were provided for Push action.");
                return new List<object>();
            }
            objects = objects.Where(x => x != null).ToList();

            // If unset, set the pushType to AdapterSettings' value (base AdapterSettings default is FullCRUD).
            if (pushType == PushType.AdapterDefault)
                pushType = PushType.UpdateOnly;

            // Copy the content of the template into a MemoryStream
            MemoryStream memoryStream = null;
            if (m_TemplateFileSettings != null)
                memoryStream = OpenTemplateFile(m_TemplateFileSettings.GetFullFileName());
            else if (m_TemplateStream != null)
            {
                memoryStream = new MemoryStream();
                m_TemplateStream.CopyTo(memoryStream);
            }

            if (memoryStream == null)
            {
                BH.Engine.Base.Compute.RecordError("The content of the template was not extracted successfully.");
                return new List<object>();
            }

            // Open the presentation
            PresentationDocument presentationDoc = null;
            try
            {
                presentationDoc = PresentationDocument.Open(memoryStream, true);
            }
            catch (Exception e)
            {
                memoryStream.Close();
                BH.Engine.Base.Compute.RecordError("Could not open the file: " + e.Message);
                return new List<object>();
            }
            
            // Update the slides
            PresentationPart presentationPart = presentationDoc.PresentationPart;
            Presentation presentation = presentationPart.Presentation;
            foreach (ISlideUpdate update in objects.OfType<ISlideUpdate>())
            {
                SlidePart slidePart = GetSlide(presentationPart, update.SlideNumber - 1);
                if (slidePart != null)
                    IUpdateSlide(slidePart, update);
            }

            // Save the output 
            try
            {
                if (m_OutputFileSettings != null)
                    presentationDoc.SaveAs(m_OutputFileSettings.GetFullFileName());
                else if (m_OutputStream != null)
                {
                    presentationDoc.Clone(m_OutputStream);
                    m_OutputStream.Position = 0;
                }
                    
            }
            catch (Exception e)
            {
                BH.Engine.Base.Compute.RecordError("Could not save the changes: " + e.Message);
            }

            // Release all content from memory
            presentationDoc.Close();
            memoryStream.Close();
            
            return objects.ToList();
        }

        /***************************************************/
        /**** Private Methods                           ****/
        /***************************************************/

        private MemoryStream OpenTemplateFile(string filePath)
        {
            // Make sure the file exists
            if (!File.Exists(filePath))
            {
                BH.Engine.Base.Compute.RecordError($"There is no presentation with the file path {filePath}");
                return null;
            }

            // Copy the template file to the memory stream
            MemoryStream memoryStream = new MemoryStream();
            try
            {
                FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                fileStream.CopyTo(memoryStream);
                fileStream.Close();
            }
            catch (Exception e)
            {
                BH.Engine.Base.Compute.RecordError("Could not open the file: " + e.Message);
            }

            return memoryStream;
        }


        /***************************************************/

        private SlidePart GetSlide(PresentationPart presentationPart, int index)
        {
            var slideIds = presentationPart.Presentation.SlideIdList.ChildElements;
            if (index > slideIds.Count)
            {
                BH.Engine.Base.Compute.RecordError($"The slide index is too high. There are only {slideIds.Count} in the presentation.");
                return null;
            }

            SlidePart slidePart = presentationPart.GetPartById((slideIds[index] as SlideId).RelationshipId) as SlidePart;
            if (slidePart == null)
                BH.Engine.Base.Compute.RecordError($"The slide cannot be found.");

            return slidePart;
        }

        /***************************************************/
    }
}



