/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2026, the respective contributors. All rights reserved.
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
            OpenXmlElement element = GetElementByName(slidePart, update.ElementName);

            if (element == null)
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
            OpenXmlElement element = GetElementByName(slidePart, update.ElementName);

            if (element == null)
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

            // Create text body and internal properties if they do not exist.
            TextBody textBody = shape.TextBody ?? shape.AppendChild(new TextBody(new Drawing.BodyProperties(), new Drawing.ListStyle()));

            List<Drawing.Paragraph> paragraphs = textBody.Elements<Drawing.Paragraph>().ToList();

            //create a new paragraph in the text body with the text for that paragraph in a run if it does not already exist, otherwise change 

            bool updateColour = !string.IsNullOrEmpty(update.Colour);
            List<Drawing.Paragraph> newParagraphs = new List<Drawing.Paragraph>();

            Drawing.RunProperties lastRunProperties = null;
            Drawing.ParagraphProperties lastParagraphProperties = null;

            for (int paragraphIndex = 0; paragraphIndex < update.Text.Count; paragraphIndex++)
            {
                Drawing.Paragraph paragraph = (Drawing.Paragraph)paragraphs.ElementAtOrDefault(paragraphIndex)?.CloneNode(true) ?? new Drawing.Paragraph();

                if (paragraph.ParagraphProperties == null && update.UseLastParagraphProperties)
                    paragraph.AddChild(lastParagraphProperties ?? new Drawing.ParagraphProperties());

                Drawing.Run run = (Drawing.Run)paragraph.GetFirstChild<Drawing.Run>()?.CloneNode(true) ?? paragraph.AppendChild(new Drawing.Run());

                if (run.RunProperties != null)
                    run.RunProperties.SpellingError = null;
                else
                    // If using last run properties, use the last run properties (create new if null), otherwise create new run properties.
                    run.AddChild(update.UseLastParagraphProperties ? lastRunProperties ?? new Drawing.RunProperties() : new Drawing.RunProperties());

                if (updateColour)
                    SetFillColour(run.RunProperties, update.Colour);

                lastRunProperties = (Drawing.RunProperties)run.RunProperties.CloneNode(true);
                lastParagraphProperties = (Drawing.ParagraphProperties)paragraph.ParagraphProperties?.CloneNode(true);

                paragraph.RemoveAllChildren<Drawing.Run>();
                paragraph.AddChild(run);

                Drawing.Text text = run.Text ?? run.AppendChild(new Drawing.Text());
                text.Text = update.Text[paragraphIndex];

                newParagraphs.Add(paragraph);
            }

            textBody.RemoveAllChildren<Drawing.Paragraph>();
            textBody.Append(newParagraphs);
        }

        /***************************************************/

    }
}




