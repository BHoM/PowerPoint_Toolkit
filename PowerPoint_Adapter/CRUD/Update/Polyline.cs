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

        private void UpdateSlide(SlidePart slidePart, MultiPolylineUpdate update)
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

            Shape shape = matchingProperty.Parent?.Parent as Shape;
            if (shape == null)
            {
                BH.Engine.Base.Compute.RecordError("The element with the name " + update.ElementName + " is not a shape.");
                return;
            }
            long left = shape.ShapeProperties.Transform2D.Offset.X;
            long top = shape.ShapeProperties.Transform2D.Offset.Y;
            long width = shape.ShapeProperties.Transform2D.Extents.Cx;
            long height = shape.ShapeProperties.Transform2D.Extents.Cy;

            BoundingBox totalBox = update.Shapes.Select(x => x.Path.Bounds()).ToList().Bounds();

            //Drawing.PathList pathList = shape.ShapeProperties.Descendants<Drawing.CustomGeometry>().First().PathList;
            //var currentPaths = pathList.ChildElements.ToList();
            //currentPaths.ForEach(x => x?.Remove());

            double bhWidth = totalBox.Max.X - totalBox.Min.X;
            double bhHeight = totalBox.Max.Y - totalBox.Min.Y;




            double scaleX;
            double scaleY;

            if (update.KeepShapeAspectRatio)
            {
                double scale = Math.Min(width / bhWidth, height / bhHeight);
                scaleX = scale;
                scaleY = scale;
            }
            else
            {
                scaleX = width / bhWidth;
                scaleY = height / bhHeight;
            }

            long offsetX = (long)Math.Round(-totalBox.Min.X * scaleX);
            long offsetY = (long)Math.Round(-totalBox.Min.Y * scaleY);


            var shapeOwner = shape.Parent;
            shape.Remove();
            var initialOutline = shape.ShapeProperties.GetFirstChild<Drawing.Outline>();
            if (initialOutline != null)
                initialOutline.Remove();

            for (int i = 0; i < update.Shapes.Count; i++)
            {
                Shape newShape = shape.DeepClone();
                Drawing.PathList pathList = newShape.ShapeProperties.Descendants<Drawing.CustomGeometry>().First().PathList;
                var currentPaths = pathList.ChildElements.ToList();
                currentPaths.ForEach(x => x?.Remove());
                PolylineData polylineData = update.Shapes[i];
                Polyline polyline = polylineData.Path;
                Drawing.Path path = new Drawing.Path();
                path.Append(new Drawing.MoveTo(ShapePoint(polyline.ControlPoints[0], scaleX, scaleY, offsetX, offsetY)));
                for (int j = 1; j < polyline.ControlPoints.Count; j++)
                {
                    path.Append(new Drawing.LineTo(ShapePoint(polyline.ControlPoints[j], scaleX, scaleY, offsetX, offsetY)));
                }

                if (!string.IsNullOrEmpty(polylineData.FillColour))
                {
                    Drawing.RgbColorModelHex rgb = new Drawing.RgbColorModelHex() { Val = polylineData.FillColour.TrimStart('#') };
                    Drawing.Alpha alpha = new Drawing.Alpha();
                    alpha.Val = (int)Math.Round(polylineData.FillOpacity * 100000);
                    rgb.Append(alpha);
                    foreach (var noFill in newShape.ShapeProperties.Elements<Drawing.NoFill>())
                    {
                        noFill.Remove();
                    }

                    var fill = newShape.ShapeProperties.GetFirstChild<Drawing.SolidFill>();
                    if (fill != null)
                    {
                        if (fill.SchemeColor != null)
                            fill.SchemeColor.Remove();

                        fill.Append(rgb);
                    }
                    else
                    {
                        fill = new Drawing.SolidFill(rgb);

                        newShape.ShapeProperties.AddChild(fill);
                    }
                }

                var outline = newShape.ShapeProperties.GetFirstChild<Drawing.Outline>();
                if (outline == null)
                {
                    outline = new Drawing.Outline();
                    newShape.ShapeProperties.AddChild(outline);
                }
                else
                {
                    outline.Remove();
                    outline = outline.DeepClone();
                    newShape.ShapeProperties.AddChild(outline);
                }

                outline.Width = (int)Math.Round(polylineData.Thickness * 12700);
                if (!string.IsNullOrEmpty(polylineData.EdgeColour))
                {
                    Drawing.RgbColorModelHex rgb = new Drawing.RgbColorModelHex() { Val = polylineData.EdgeColour.TrimStart('#') };
                    foreach (var noFill in outline.Elements<Drawing.NoFill>())
                    {
                        noFill.Remove();
                    }

                    var fill = outline.GetFirstChild<Drawing.SolidFill>();
                    if (fill != null)
                    {
                        if (fill.SchemeColor != null)
                            fill.SchemeColor.Remove();

                        fill.Append(rgb);
                    }
                    else
                    {
                        fill = new Drawing.SolidFill(rgb);

                        outline.AddChild(fill);
                    }
                }

                if (polylineData.IsDashed)
                {
                    var dashed = outline.GetFirstChild<Drawing.PresetDash>();
                    if (dashed == null)
                    {
                        dashed = new Drawing.PresetDash();
                        outline.AddChild(dashed);
                    }

                    dashed.Val = Drawing.PresetLineDashValues.Dash;

                }

                pathList.Append(path);
                shapeOwner.Append(newShape);

            }


        }

        /***************************************************/

        private Drawing.Point ShapePoint(Point bhPoint, double scaleX, double scaleY, long offsetX, long offsetY)
        {
            return new Drawing.Point { X = (Math.Round(bhPoint.X * scaleX) + offsetX).ToString(), Y = (Math.Round(bhPoint.Y * scaleY) + offsetY).ToString() };
        }

        /***************************************************/


    }
}


