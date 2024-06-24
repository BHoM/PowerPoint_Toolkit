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
            OpenXmlElement element = GetElementByName(slidePart, update.ElementName);

            if (element == null)
            {
                BH.Engine.Base.Compute.RecordError($"Could not find an element with the name {update.ElementName}");
                return;
            }

            Picture picture;

            switch (element)
            {
                case Picture oldPicture:
                    picture = oldPicture;
                    break;
                case Shape oldShape:
                    picture = ConvertShapeToPicture(oldShape);
                    slidePart.Slide.CommonSlideData.ShapeTree.ReplaceChild(picture, oldShape);
                    break;
                default:
                    BH.Engine.Base.Compute.RecordError($"The element with name '{update.ElementName}' must be either a Shape or a Picture to be updated with an image.");
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

            // Read the image file

            ImagePart imagePart = slidePart.AddImagePart(imageType);

            try
            {
                using (FileStream stream = File.OpenRead(update.ImageFilePath))
                    imagePart.FeedData(stream);
            }
            catch (Exception ex)
            {
                BH.Engine.Base.Compute.RecordError(ex, "An error occurred while copying the image into the presentation.");
                return;
            }

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

            OpenXmlElement element = GetElementByName(slidePart, update.ElementName);

            if (element == null)
            {
                BH.Engine.Base.Compute.RecordError($"Could not find an element with the name {update.ElementName}");
                return;
            }

            Picture picture;

            switch (element)
            {
                case Picture oldPicture:
                    picture = oldPicture;
                    break;
                case Shape oldShape:
                    picture = ConvertShapeToPicture(oldShape);
                    slidePart.Slide.CommonSlideData.ShapeTree.ReplaceChild(picture, oldShape);
                    break;
                default:
                    BH.Engine.Base.Compute.RecordError($"The element with name '{update.ElementName}' must be either a Shape or a Picture to be updated with an image.");
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

        private Picture ConvertShapeToPicture(Shape oldShape)
        {
            ShapeProperties shapeProperties = (ShapeProperties)oldShape.Descendants<ShapeProperties>().Single().CloneNode(true);
            NonVisualDrawingProperties drawingProperties = (NonVisualDrawingProperties)oldShape.Descendants<NonVisualDrawingProperties>().Single().CloneNode(true);

            // If the shape doesn't have a custom or preset geometry (for some reason) it does not display the image, so create a rectangular presetgeometry if it doesn't exist already.
            if (shapeProperties.Descendants<Drawing.CustomGeometry>().SingleOrDefault() == null)
                _ = shapeProperties.Descendants<Drawing.PresetGeometry>().SingleOrDefault() ?? shapeProperties.AppendChild(new Drawing.PresetGeometry() { Preset=Drawing.ShapeTypeValues.Rectangle });

            Picture picture = new Picture
            (
                new NonVisualPictureProperties
                (
                    drawingProperties,
                    new NonVisualPictureDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()
                ),
                new BlipFill
                (
                    new Drawing.Blip(),
                    new Drawing.Stretch()
                ),
                shapeProperties
            );

            return picture;
        }

    }
}


