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

using BH.oM.Base;
using BH.oM.Geometry;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace BH.oM.PowerPoint
{
    [Description("Allows to change the geometry of a shape element.")]
    public class MultiPolylineUpdate : BHoMObject, ISlideUpdate
    {
        [Description("Number of the slide where the update needs to happen.")]
        public virtual int SlideNumber { get; set; } = 0;

        [Description("Name of the shape element that needs to be updated.")]
        public virtual string ElementName { get; set; } = "";

        [Description("Shape to be updated.")]
        public virtual List<PolylineData> Shapes { get; set; } = new List<PolylineData>();

        [Description("If true, the provided shapes are centred in the template box, if false, the shapes are drawn from the top left corner.")]
        public virtual bool CentreShapes { get; set; } = true;

        [Description("If true, the shape aspect ratio is kept, and the shapes are made to fit the extents of the template shape. If false, the shapes are atempted to fill up the template shape as much as possible which can lead to change in aspect ratio of the provided shapes.")]
        public virtual bool KeepShapeAspectRatio { get; set; } = true;

        [Description("If true, the provided shapes that share all properties in terms of colours will be added to the same shape object. If false, all PolylineData obejcts will be added to separate shape objects.")]
        public virtual bool GroupPolylinesWithSameProperties { get; set; } = false;

        [Description("Scale factor to be applied to the figure.")]
        public virtual double Scale { get; set; } = 1.0;
    }
}


