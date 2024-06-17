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

        private void UpdateSlide(SlidePart slidePart, SimpleTextUpdate update)
        {
            // Get the shape element matching the name provided in update
            NonVisualDrawingProperties matchingProperty = slidePart.Slide.Descendants<NonVisualDrawingProperties>()
                .Where(x => x.Name.Value == update.ElementName)
                .FirstOrDefault();

            if (matchingProperty == null)
            {
                BH.Engine.Base.Compute.RecordError("Could not find the element with the name " + update.ElementName);
                return;
            }
                
            Shape shape = matchingProperty.Parent?.Parent as Shape;
            if (shape == null)
            {
                BH.Engine.Base.Compute.RecordError("The element with the name " + update.ElementName + " is not a shape.");
                return;
            }

            // Replace the text
            var paragraph = shape.Descendants<Drawing.Paragraph>().FirstOrDefault();
            var runs = paragraph.Descendants<Drawing.Run>().ToList();

            if (runs.Count == 0)
            {
                paragraph.AddChild(new Drawing.Run(new Drawing.Text(update.Text)));
            }
            else if (runs.Count == 1)
            {
                Drawing.Text text = runs.First().Text;
                if (text != null)
                    text.Text = update.Text;
                else
                    runs.First().Text = new Drawing.Text(update.Text);
            }
            else
            { 
                BH.Engine.Base.Compute.RecordError("The element contains more than one line of text. Please use MultiLineTextUpdate for this.");
                return;
            }
           
        }

        /***************************************************/


        private void UpdateSlide(SlidePart slidePart, MultiLineTextUpdate update)
        {
            // Get the shape element matching the name provided in update
            NonVisualDrawingProperties matchingProperty = slidePart.Slide.Descendants<NonVisualDrawingProperties>()
                .Where(x => x.Name.Value == update.ElementName)
                .FirstOrDefault();

            if (matchingProperty == null)
            {
                BH.Engine.Base.Compute.RecordError("Could not find the element with the name " + update.ElementName);
                return;
            }

            Shape shape = matchingProperty.Parent?.Parent as Shape;
            if (shape == null)
            {
                BH.Engine.Base.Compute.RecordError("The element with the name " + update.ElementName + " is not a shape.");
                return;
            }

            // Replace the text
            var paragraph = shape.Descendants<Drawing.Paragraph>().FirstOrDefault();
            var runs = paragraph.Descendants<Drawing.Run>().ToList();

            int textCount = update.Text.Count;
            int runCount = runs.Count;

            string fullText = "";
            for (int i = 0; i < update.Text.Count - 1; i++)
            {
                fullText += update.Text[i] + Environment.NewLine;
            }
            fullText += update.Text[update.Text.Count - 1];

            if (runs.Count == 0)
            {
                paragraph.AddChild(new Drawing.Run(new Drawing.Text(fullText)));
            }
            else
            {
                Drawing.Text text = runs[0].Text;
                if (text != null)
                    text.Text = fullText;
                else
                    runs.First().Text = new Drawing.Text(fullText);

                Drawing.RunProperties runProps = runs[0].RunProperties;
                if (runProps != null)
                    runProps.SpellingError = null;  //Make sure no spelling error underlines are left from template
            }

            for (int i = 1; i < runCount; i++)
            {
                runs[i].Remove();
            }

            if (!string.IsNullOrEmpty(update.Colour))
            {
                Drawing.RunProperties rp = shape.Descendants<Drawing.RunProperties>().FirstOrDefault();

                if (rp == null)
                {
                    rp = new Drawing.RunProperties();
                    shape.AddChild(rp);
                }
                SetFillColour(rp, update.Colour);
            }
        }

        /***************************************************/

    }
}


