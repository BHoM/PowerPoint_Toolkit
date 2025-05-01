/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2025, the respective contributors. All rights reserved.
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

using BH.oM.PowerPoint;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using BH.Engine.Base;

namespace BH.Adapter.PowerPoint
{
    public partial class PowerPointAdapter
    {
        private void UpdateSlide(SlidePart slidePart, TableUpdate update)
        {
            if (update.Contents == null || update.Contents.IsNullOrEmpty())
            {
                BH.Engine.Base.Compute.RecordError("The table update has no contents to update the table with.");
                return;
            }

            if (update.Contents[0].IsNullOrEmpty())
            {
                BH.Engine.Base.Compute.RecordError("The table update has no contents to update the table with.");
                return;
            }

            int rowCount = update.Contents.Count;
            int columnCount = update.Contents[0].Count;

            if (update.Contents.Any(x => columnCount != x.Count))
            {
                BH.Engine.Base.Compute.RecordError("The length of all of the rows in Content must be equal.");
                return;
            }

            if (update.HeaderRow.Count > 0 && update.HeaderRow.Count != columnCount)
            {
                BH.Engine.Base.Compute.RecordError($"The length of the header row ({update.HeaderRow.Count}) must be equal to the length of the content rows ({columnCount}).");
                return;
            }

            if (update.HeaderRow.Count > 0)
                rowCount += 1;

            OpenXmlElement element = GetElementByName(slidePart, update.ElementName);
            GraphicFrame frame; //tables are contained within graphic frames

            switch (element) //if the element is a graphic frame already then just use that, but if it's a placeholder then there is nothing contained so need to create a new table from scratch
            {
                case GraphicFrame graphicFrame:
                    frame = graphicFrame;
                    break;
                case Shape shape:
                    frame = ConvertToGraphicFrame(shape, rowCount, columnCount);
                    slidePart.Slide.CommonSlideData.ShapeTree.ReplaceChild(frame, shape);
                    break;
                default:
                    BH.Engine.Base.Compute.RecordError($"The element with name '{update.ElementName}' on slide {update.SlideNumber} must be either a Shape, a Table, or a Table placeholder to be updated with a table.");
                    return;
            }

            D.Table table = frame.Graphic?.GraphicData?.GetFirstChild<D.Table>();

            if (table == null)
            {
                BH.Engine.Base.Compute.RecordError($"The element with the name `{update.ElementName}` on slide {update.SlideNumber} must contain a Table.");
                return;
            }

            IEnumerable<D.GridColumn> columns = table.TableGrid.Descendants<D.GridColumn>();
            IEnumerable<D.TableRow> rows = table.Descendants<D.TableRow>();

            //compare to the column and row counts in the data to update.
            if (columns.Count() != columnCount || rows.Count() != rowCount)
            {
                BH.Engine.Base.Compute.RecordError($"The shape of the data in Content must be the same shape as the table that is being updated.\nContent has {rowCount} rows and {columnCount} columns, whereas the table has {rows.Count()} rows and {columns.Count()} columns.");
                return;
            }

            //update the text in each row/column - This assumes that each table cell will have only one paragraph and run originally and does not modify any other properties.

            int r = 0;
            int c = 0;
            bool headerRow = update.HeaderRow.Count != 0;
            foreach (D.TableRow row in rows)
            {
                foreach (D.TableCell cell in row.Descendants<D.TableCell>())
                {
                    try
                    {
                        D.Paragraph par = cell.TextBody.Elements<D.Paragraph>().SingleOrDefault();
                        if (par == null)
                        {
                            par = new D.Paragraph();
                            par.AddChild(new D.Run(new D.RunProperties(), new D.Text()));
                            cell.TextBody.AddChild(par);
                        }

                        D.Run run = par.Elements<D.Run>().SingleOrDefault();
                        int? fontSize = null;
                        if (run == null)
                        {
                            run = new D.Run(new D.RunProperties(), new D.Text());
                            par.AddChild(run);
                        }
                        else
                            fontSize = run.RunProperties.FontSize / 100;

                        if (update.UpdatedTextFontSize <= 0)
                            fontSize = fontSize ?? 20;
                        else
                            fontSize = update.UpdatedTextFontSize;

                        if (headerRow)
                            run.Text.Text = update.HeaderRow[c];
                        else 
                            run.Text.Text = update.Contents[r][c];

                        run.RunProperties.FontSize = fontSize * 100;
                    }
                    catch (InvalidOperationException ex)
                    {
                        BH.Engine.Base.Compute.RecordWarning(ex, $"An error occurred while trying to update cell in row {r} column {c}, due to there being more than one paragraph or run in the template cell. This cell has not been updated, but others might have been. Occurred in element `{update.ElementName}` on slide {update.SlideNumber}.");
                    }
                    catch (Exception ex)
                    {
                        BH.Engine.Base.Compute.RecordError(ex, $"An error occurred while trying to update cell in row {r} column {c}. Occurred in element `{update.ElementName}` on slide {update.SlideNumber}.");
                        return;
                    }
                    c++;
                }
                r++;
                c = 0;
            }
        }

        private static GraphicFrame ConvertToGraphicFrame(Shape oldShape, int rowCount, int columnCount)
        {
            ShapeProperties shapeProperties = (ShapeProperties)oldShape.Descendants<ShapeProperties>().Single();
            NonVisualDrawingProperties drawingProps = (NonVisualDrawingProperties)oldShape.Descendants<NonVisualDrawingProperties>().Single().CloneNode(true);

            GraphicFrame gf =  new GraphicFrame(
                new NonVisualGraphicFrameProperties(
                    drawingProps,
                    new NonVisualGraphicFrameDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()
                    ),
                new Transform() { Offset = (D.Offset)shapeProperties.Transform2D.Offset.CloneNode(true), Extents = (D.Extents)shapeProperties.Transform2D.Extents.CloneNode(true) }
                );

            (long width, long height) = GetSizeOfShape(shapeProperties);

            gf.Append(new D.Graphic(new D.GraphicData(ConstructNewTable(rowCount, columnCount, width, height)) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/table" }));

            return gf;
        }

        private static (long, long) GetSizeOfShape(ShapeProperties shapeProps)
        {
            long width = shapeProps.Transform2D.Extents.Cx;
            long height = shapeProps.Transform2D.Extents.Cy;
            return (width, height);
        }

        private static D.Table ConstructNewTable(int rowCount, int columnCount, long shapeWidth, long shapeHeight)
        {
            D.Table table = new D.Table();
            D.TableProperties tableProperties = new D.TableProperties();
            D.TableGrid tGrid = new D.TableGrid();

            for (int i = 0; i < columnCount; i++)
            {
                tGrid.Append(new D.GridColumn() { Width = shapeWidth / columnCount });
            }

            table.Append(tableProperties);
            table.Append(tGrid);

            for (int row = 0; row < rowCount; row++)
            {
                D.TableRow tRow = new D.TableRow() { Height = shapeHeight / rowCount };
                for (int col = 0; col < columnCount; col++)
                {
                    tRow.Append(ConstructNewTableCell());
                }
                table.Append(tRow);
            }

            return table;
        }

        private static D.TableCell ConstructNewTableCell()
        {
            D.TableCell cell = new D.TableCell();
            D.TableCellProperties properties = new D.TableCellProperties();
            D.TextBody textBody = new D.TextBody();
            D.BodyProperties bodyProperties = new D.BodyProperties();
            D.ListStyle listStyle = new D.ListStyle();

            D.Paragraph par = new D.Paragraph();
            D.Run run = new D.Run();
            D.RunProperties runProps = new D.RunProperties();
            D.Text text = new D.Text(); //.Text property is what will be modified later
            run.Append(runProps);
            run.Append(text);
            D.EndParagraphRunProperties endProps = new D.EndParagraphRunProperties();

            par.Append(run);
            par.Append(endProps);
            textBody.Append(bodyProperties);
            textBody.Append(listStyle);
            textBody.Append(par);

            cell.Append(textBody);
            cell.Append(properties);

            return cell;
        }
    }
}
