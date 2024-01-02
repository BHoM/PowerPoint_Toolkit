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

namespace BH.Adapter.PowerPoint
{
    public partial class PowerPointAdapter : BHoMAdapter
    {
        /***************************************************/
        /**** Interface Methods                         ****/
        /***************************************************/

        protected void IUpdateSlide(SlidePart slidePart, ISlideUpdate update)
        {
            if (update == null)
                BH.Engine.Base.Compute.RecordError("No action was found for an update of type " + update.GetType().Name);
            else
                UpdateSlide(slidePart, update as dynamic);
        }


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

        private void UpdateSlide(SlidePart slidePart, ChartUpdate update)
        {
            // Get the chart element matching the name provided in update
            NonVisualDrawingProperties matchingProperty = slidePart.Slide.Descendants<NonVisualDrawingProperties>()
                .Where(x => x.Name.Value == update.ElementName)
                .FirstOrDefault();

            if (matchingProperty == null)
            {
                BH.Engine.Base.Compute.RecordError("Could not find the element with the name " + update.ElementName);
                return;
            }

            GraphicFrame frame = matchingProperty.Parent?.Parent as GraphicFrame;
            if (frame == null)
            {
                BH.Engine.Base.Compute.RecordError("The element with the name " + update.ElementName + " is not a chart.");
                return;
            }
            Drawing.Charts.ChartReference reference = frame.Descendants<Drawing.Charts.ChartReference>().FirstOrDefault();

            ChartPart chartPart = null;
            try
            {
                chartPart = slidePart.GetPartById(reference.Id) as ChartPart;
            }
            catch { }
            if (chartPart == null)
            {
                BH.Engine.Base.Compute.RecordError("Cannot find the reference to the chart " + update.ElementName + ".");
                return;
            }

            // Check that the update data is the correct size
            if (update.Data.Count != update.Series.Count)
            {
                BH.Engine.Base.Compute.RecordError("The number of data lists must be equal to the number of series provided in the update.");
                return;
            }
            if (update.Data.Any(x => x.Count != update.Categories.Count))
            {
                BH.Engine.Base.Compute.RecordError("Each list of data must contain a number of values equal to the number of categories provided in the update.");
                return;
            }

            // Update the title if 'update.Title' is provided
            Drawing.Charts.Chart chart = chartPart.ChartSpace.Descendants<Drawing.Charts.Chart>().FirstOrDefault();
            if (update.Title?.Length > 0)
            {
                Drawing.Text title = chart.Elements<Drawing.Charts.Title>().FirstOrDefault().Descendants<Drawing.Text>().FirstOrDefault();
                if (title != null)
                    title.Text = update.Title;
            }

            // Get access to the embedded spreadsheet
            //UpdateEmbeddedSpreadsheet(chartPart, update); // Not needed anymore since the chart will remove its relation to the spreadsheet below
            chartPart.ChartSpace.Elements<Drawing.Charts.ExternalData>().FirstOrDefault()?.Remove();
            chartPart.DeletePart(chartPart.EmbeddedPackagePart);

            // Remove the existing chart series
            var chartSeries = chart.Descendants<Drawing.Charts.SeriesText>().Select(x => x.Parent).ToList();

            var shapeProperties = chartSeries.Select(x => x.ChildElements.OfType<Drawing.Charts.ChartShapeProperties>().FirstOrDefault()).ToList();
            shapeProperties.ForEach(x => x.Remove());

            var seriesParent = chartSeries.First().Parent;
            chartSeries.ForEach(x => x.Remove());

            // Create the template for the chart series
            var seriesTemplate = chartSeries.First().DeepClone();
            seriesTemplate.ReplaceChild(
                new Drawing.Charts.CategoryAxisData(
                    new OpenXmlElement[] {
                        new Drawing.Charts.StringLiteral(
                            update.Categories.Select((x, i) => new Drawing.Charts.StringPoint(
                                new OpenXmlElement[] { new Drawing.Charts.NumericValue(x) }) { Index = new UInt32Value((uint)i) }
                            )
                        ) { PointCount = new Drawing.Charts.PointCount { Val = new UInt32Value((uint)update.Categories.Count) } }
                    }
                ),
                seriesTemplate.ChildElements.OfType<Drawing.Charts.CategoryAxisData>().FirstOrDefault()
            );

            // Add the new series
            for (int i = 0; i < update.Series.Count; i++)
            {
                var serie = seriesTemplate.DeepClone();

                serie.ReplaceChild(
                    new Drawing.Charts.Index { Val = new UInt32Value((uint)i)},
                    serie.ChildElements.OfType<Drawing.Charts.Index>().First()
                );

                serie.ReplaceChild(
                    new Drawing.Charts.Order { Val = new UInt32Value((uint)i) },
                    serie.ChildElements.OfType<Drawing.Charts.Order>().First()
                );

                serie.ReplaceChild(
                    new Drawing.Charts.SeriesText(
                        new Drawing.Charts.NumericValue(update.Series[i])
                    ),
                    serie.ChildElements.OfType<Drawing.Charts.SeriesText>().FirstOrDefault()
                );

                serie.ReplaceChild(
                    new Drawing.Charts.Values(
                        new OpenXmlElement[] {
                            new Drawing.Charts.NumberLiteral(
                                update.Data[i].Select((x, j) => new Drawing.Charts.NumericPoint(
                                    new OpenXmlElement[] { new Drawing.Charts.NumericValue(x.ToString()) }) { Index = new UInt32Value((uint)j) }
                                )
                            ) { PointCount = new Drawing.Charts.PointCount { Val = new UInt32Value((uint)(update.Data[i].Count)) } }
                        }
                    ),
                    serie.ChildElements.OfType<Drawing.Charts.Values>().FirstOrDefault()
                );

                serie.AppendChild(shapeProperties[i % shapeProperties.Count].DeepClone());

                seriesParent.AppendChild(serie);
            }
        }

        /***************************************************/

        private void CopyStream(Stream input, Stream output)
        {
            byte[] buffer = new byte[32768];
            while (true)
            {
                int read = input.Read(buffer, 0, buffer.Length);
                if (read <= 0)
                    return;
                output.Write(buffer, 0, read);
            }
        }

        /***************************************************/
        /**** Fallback Methods                          ****/
        /***************************************************/

        private void UpdateSlide(SlidePart slidePart, ISlideUpdate update)
        {
            BH.Engine.Base.Compute.RecordError("No action was found for an update of type " + update.GetType().Name);
        }


        /***************************************************/
        /**** Helper Methods                            ****/
        /***************************************************/

        // Not used anymore but kept for reference as it shows how to edit a data table in an internal spreadsheet
        private void UpdateEmbeddedSpreadsheet(ChartPart chartPart, ChartUpdate update) 
        {
            try
            {
                EmbeddedPackagePart epp = chartPart.EmbeddedPackagePart;
                Stream stream = epp.GetStream();
                SpreadsheetDocument spreadsheetDoc = SpreadsheetDocument.Open(stream, true);

                Spreadsheet.SharedStringTable sharedStringTable = spreadsheetDoc.GetAllParts().OfType<SharedStringTablePart>().FirstOrDefault()?.SharedStringTable;
                WorksheetPart worksheetPart = spreadsheetDoc.GetAllParts().OfType<WorksheetPart>().FirstOrDefault();
                Spreadsheet.Worksheet worksheet = worksheetPart?.Worksheet;
                Spreadsheet.SheetData sheetData = worksheet.Elements<Spreadsheet.SheetData>().FirstOrDefault();

                // Clean the existing content
                List<Spreadsheet.Row> rows = sheetData.Elements<Spreadsheet.Row>().ToList();
                foreach (Spreadsheet.Row row in rows)
                {
                    foreach (Spreadsheet.Cell cell in row.Elements<Spreadsheet.Cell>())
                    {
                        if (cell.DataType?.Value != Spreadsheet.CellValues.SharedString)
                            cell.CellValue.Text = "";
                    }
                }

                // Updating the sheet's content
                for (int r = 0; r < update.Data.Count; r++)
                {
                    Spreadsheet.Row row = rows.FirstOrDefault(x => x.RowIndex == r + 1);
                    if (row == null)
                        row = new Spreadsheet.Row { RowIndex = (uint)(r + 1) };

                    List<object> values = r == 0 ?
                        Enumerable.Concat<object>(new List<object> { " " }, update.Series.ToList<object>()).ToList() 
                        : Enumerable.Concat<object>(new List<object> { update.Categories[r - 1] }, update.Data[r-1].Cast<object>()).ToList();
                    List<Spreadsheet.Cell> cells = row.Elements<Spreadsheet.Cell>().ToList();
                    for (int c = 0; c < values.Count; c++)
                    {
                        string cellReference = BH.Engine.Excel.Query.ColumnName(c) + (r + 1);
                        Spreadsheet.Cell cell = cells.FirstOrDefault(x => x.CellReference.Value == cellReference);
                        if (cell == null)
                        {
                            row.AddChild(new Spreadsheet.Cell()
                            {
                                CellReference = cellReference,
                                CellValue = new Spreadsheet.CellValue(values[c] as dynamic)
                            });
                        }
                        else if (cell.DataType?.Value == Spreadsheet.CellValues.SharedString)
                        {
                            int sharedIndex = -1;
                            if (int.TryParse(cell.CellValue?.Text, out sharedIndex))
                            {
                                Spreadsheet.SharedStringItem item = sharedStringTable.ElementAt(sharedIndex) as Spreadsheet.SharedStringItem;
                                if (item != null)
                                    item.SetPropertyValue("Text", new Spreadsheet.Text { Text = values[c].ToString(), Space = new EnumValue<SpaceProcessingModeValues>(SpaceProcessingModeValues.Preserve) });
                            }
                        }
                        else
                            cell.CellValue = new Spreadsheet.CellValue(values[c] as dynamic);
                    }
                }

                // Updating the sheet's table
                Spreadsheet.TablePart tablePart = worksheet.Descendants<Spreadsheet.TablePart>().FirstOrDefault();
                if (tablePart != null)
                {
                    int nbColumns = update.Series.Count;
                    Spreadsheet.Table table = (worksheetPart.GetPartById(tablePart.Id) as TableDefinitionPart)?.Table;
                    table.Reference = "A1:" + BH.Engine.Excel.Query.ColumnName(nbColumns - 1) + update.Data.Count;
                    table.TableColumns.Count.Value = (uint)nbColumns;

                    List<Spreadsheet.TableColumn> columns = table.TableColumns.Elements<Spreadsheet.TableColumn>().ToList();
                    for (int c = 0; c < columns.Count; c++)
                        columns[c].Name = c == 0 ? " " : update.Series[c-1].ToString();
                }

                // Saving the spreadsheet
                spreadsheetDoc.Close();
                stream.Close();
            }
            catch (Exception e)
            {
                BH.Engine.Base.Compute.RecordError("Failed to update the chart data: " + e.Message);
                return;
            }
        }

        /***************************************************/

    }
}


