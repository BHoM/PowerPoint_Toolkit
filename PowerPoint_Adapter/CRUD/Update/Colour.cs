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
using DocumentFormat.OpenXml.VariantTypes;

namespace BH.Adapter.PowerPoint
{
    public partial class PowerPointAdapter : BHoMAdapter
    {
       
        /***************************************************/
        /**** Private Methods                           ****/
        /***************************************************/

        private void UpdateSlide(SlidePart slidePart, ShapeColourUpdate update)
        {
            // Get the shape element matching the name provided in update
            OpenXmlElement element = GetElementByName(slidePart, update.ElementName);

            if (element  == null)
            {
                BH.Engine.Base.Compute.RecordError($"Could not find an element with the name {update.ElementName}");
                return;
            }

            Shape shape = element as Shape;

            if (shape == null)
            {
                BH.Engine.Base.Compute.RecordError("The element with the name " + update.ElementName + " is not a shape.");
                return;
            }

            // Replace the colours
            if (!string.IsNullOrEmpty(update.EdgeColour))
                SetOutlineColour(shape.ShapeProperties, update.EdgeColour);

            if (!string.IsNullOrEmpty(update.FillColour))
                SetFillColour(shape.ShapeProperties, update.FillColour);
        }

        /***************************************************/
        /**** Private Methods - Helpers                 ****/
        /***************************************************/

        private void SetOutlineColour(OpenXmlElement element, string hexColour, double opacity = -1)
        {
            Drawing.Outline outline = element.GetFirstChild<Drawing.Outline>()?? element.AppendChild(new Drawing.Outline());

            SetFillColour(outline, hexColour, opacity);
        }

        private void SetFillColour(OpenXmlElement element, string hexColour, double opacity = -1)
        {
            //Remove all instances of NoFill
            element.RemoveAllChildren<Drawing.NoFill>();

            //Try get an existing SolidFill element out
            Drawing.SolidFill fill = element.GetFirstChild<Drawing.SolidFill>()?? element.AppendChild(new Drawing.SolidFill());

            //If existing, make sure any SchemeColor is removed - to be replaced by explicit RGB color
            fill.SchemeColor?.Remove();

            Drawing.RgbColorModelHex rgb = fill.GetFirstChild<Drawing.RgbColorModelHex>()?? fill.AppendChild(new Drawing.RgbColorModelHex());
            rgb.Val = hexColour.TrimStart('#');

            //Check if opacity value is to be assigned
            if (opacity >= 0)
            {
                if (opacity > 1)
                {
                    BH.Engine.Base.Compute.RecordWarning("Opacity value above 1 provided. Opacity of 1 means full opacity (no transparency). Value of full opacity assumed.");
                    opacity = 1;
                }
                rgb.Append(new Drawing.Alpha() { Val = (int)Math.Round(opacity * 100000) });
            }
        }

        /***************************************************/
    }
}