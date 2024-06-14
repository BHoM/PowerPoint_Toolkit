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
using DocumentFormat.OpenXml;
using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using Drawing = DocumentFormat.OpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using BH.Engine.Base;
using BH.Engine.Geometry;
using BH.oM.Geometry;

namespace BH.Adapter.PowerPoint
{
    public partial class PowerPointAdapter : BHoMAdapter
    {
       
        /***************************************************/
        /**** Private Methods                           ****/
        /***************************************************/

        private void UpdateSlide(SlidePart slidePart, ImageUpdate update)
        {

            // Get the image element matching the name provided in update
            NonVisualDrawingProperties matchingProperty = slidePart.Slide.Descendants<NonVisualDrawingProperties>()
                .Where(x => x.Name.Value == update.ElementName)
                .FirstOrDefault();

            if (matchingProperty == null)
            {
                BH.Engine.Base.Compute.RecordError("Could not find the element with the name " + update.ElementName);
                return;
            }

            Picture picture = matchingProperty.Parent?.Parent as Picture;
            if (picture == null)
            {
                BH.Engine.Base.Compute.RecordError("The element with the name " + update.ElementName + " is not an image.");
                return;
            }

            // Read the image file
            FileStream stream;
            try
            {
                stream = File.OpenRead(update.ImageFilePath);
            }
            catch (Exception e)
            {
                BH.Engine.Base.Compute.RecordError("The image could not be opened: " + e.Message);
                return;
            }

            // Add the image to the PowerPoint
            string imageExtension = System.IO.Path.GetExtension(update.ImageFilePath).ToLower();
            ImagePartType imageType = ImagePartType.Jpeg;
            switch (System.IO.Path.GetExtension(update.ImageFilePath))
            {
                case "bmp":
                    imageType = ImagePartType.Bmp;
                    break;
                case "png":
                    imageType = ImagePartType.Png;
                    break;
                case "gif":
                    imageType = ImagePartType.Gif;
                    break;
                case "svg":
                    imageType = ImagePartType.Svg;
                    break;
            }

            ImagePart imagePart = slidePart.AddImagePart(imageType);
            imagePart.FeedData(stream);
            stream.Close();

            // Link the image element to the new image file
            Drawing.Blip blip = picture.BlipFill?.Blip;
            if (blip == null)
            {
                BH.Engine.Base.Compute.RecordError("Could not replace the image in element " + update.ElementName);
                return;
            }
            blip.Embed = slidePart.GetIdOfPart(imagePart);
        }

        /***************************************************/

        private void UpdateSlide(SlidePart slidePart, ImageUpdateStream update)
        {

            if (update.ImageStream == null)
            {
                BH.Engine.Base.Compute.RecordError("Null stream provided. Unable to update image.");
                return;
            }

            // Get the image element matching the name provided in update
            NonVisualDrawingProperties matchingProperty = slidePart.Slide.Descendants<NonVisualDrawingProperties>()
                .Where(x => x.Name.Value == update.ElementName)
                .FirstOrDefault();

            if (matchingProperty == null)
            {
                BH.Engine.Base.Compute.RecordError("Could not find the element with the name " + update.ElementName);
                return;
            }

            Picture picture = matchingProperty.Parent?.Parent as Picture;
            if (picture == null)
            {
                BH.Engine.Base.Compute.RecordError("The element with the name " + update.ElementName + " is not an image.");
                return;
            }

            // Add the image to the PowerPoint
            ImagePartType imageType = ImagePartType.Jpeg;
            switch (update.ImageType.ToLower())
            {
                case "bmp":
                    imageType = ImagePartType.Bmp;
                    break;
                case "png":
                    imageType = ImagePartType.Png;
                    break;
                case "gif":
                    imageType = ImagePartType.Gif;
                    break;
                case "svg":
                    imageType = ImagePartType.Svg;
                    break;
            }

            ImagePart imagePart = slidePart.AddImagePart(imageType);
            imagePart.FeedData(update.ImageStream);

            // Link the image element to the new image file
            Drawing.Blip blip = picture.BlipFill?.Blip;
            if (blip == null)
            {
                BH.Engine.Base.Compute.RecordError("Could not replace the image in element " + update.ElementName);
                return;
            }
            blip.Embed = slidePart.GetIdOfPart(imagePart);
        }

        /***************************************************/

    }
}


