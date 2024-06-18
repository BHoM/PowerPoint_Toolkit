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

using BH.Engine.Adapter;
using BH.Engine.Base;
using BH.oM.Adapter;
using BH.oM.Base;
using BH.oM.Data.Collections;
using BH.oM.PowerPoint;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
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

            objects = objects.Where(x => x != null);

            // Filter out objects based on the push type given
            switch (pushType)
            {
                case PushType.UpdateOnly:
                    objects = objects.Where(x => typeof(ISlideUpdate).IsAssignableFrom(x.GetType()));
                    break;
                case PushType.CreateNonExisting:
                case PushType.CreateOnly:
                    objects = objects.Where(x => typeof(ISlideCreate).IsAssignableFrom(x.GetType()));
                    break;
            }

            using (MemoryStream memoryStream = GetTemplateMemoryStream())
            using (PresentationDocument presentationDoc = PresentationDocument.Open(memoryStream, true))
            {
                // Ensure the streams are not null
                if (memoryStream == null)
                {
                    BH.Engine.Base.Compute.RecordError("The content of the template was not extracted successfully.");
                    return new List<object>();
                }
                else if (presentationDoc == null)
                {
                    BH.Engine.Base.Compute.RecordError("The presentation document could not be opened correctly.");
                    return new List<object>();
                }
            
                // Get the presentation part and the presentation.
                PresentationPart presentationPart = presentationDoc.PresentationPart;
                Presentation presentation = presentationPart.Presentation;

                // Update/create slides based upon given actions.
                foreach (object action in objects)
                {
                    switch (action)
                    {
                        case ISlideUpdate update:
                            SlidePart slidePart = GetSlide(presentationPart, update.SlideNumber - 1);
                            if (slidePart != null)
                                IUpdateSlide(slidePart, update);
                            break;
                        case ISlideCreate create:
                            ICreateSlide(presentationPart, create);
                            break;
                    }
                }

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(presentationDoc);

                if (errors.Any())
                    BH.Engine.Base.Compute.RecordWarning($"There are some ({errors.Count()}) validation errors in the presentation caused by the some of the changes made in this push. The presentation may still be recoverable in PowerPoint, though some elements may have been affected.");

                // Save the output
                try
                {
                    if (m_OutputFileSettings != null)
                        presentationDoc.Clone(m_OutputFileSettings.GetFullFileName()).Dispose();
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
            }

            return objects.ToList();
        }

        /***************************************************/
        /**** Private Methods                           ****/
        /***************************************************/

        private MemoryStream GetTemplateMemoryStream()
        {
            if (m_TemplateFileSettings != null)
                return OpenTemplateFile(m_TemplateFileSettings.GetFullFileName());
            else if (m_TemplateStream != null)
            {
                MemoryStream memoryStream = new MemoryStream();
                m_TemplateStream.CopyTo(memoryStream);
                return memoryStream;
            }
            else
                return null;
        }

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
                using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    fileStream.CopyTo(memoryStream);
            }
            catch (Exception e)
            {
                BH.Engine.Base.Compute.RecordError(e, "An error occurred while opening the template file.");
                return null;
            }

            return memoryStream;
        }

        /***************************************************/

        private SlidePart GetSlide(PresentationPart presentationPart, int index)
        {
            var slideIds = presentationPart.Presentation.SlideIdList.ChildElements;
            if (index > slideIds.Count)
            {
                BH.Engine.Base.Compute.RecordError($"The slide index is too high. There are only {slideIds.Count} slides in the presentation.");
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