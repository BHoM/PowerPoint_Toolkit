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
 * along with this code. If not, see https://www.gnu.org/licenses/lgpl-3.0.html.      
 */

using BH.oM.Base;
using BH.oM.Geometry;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace BH.oM.PowerPoint
{
    [Description("Data corresponding to polylinear data and style properties.")]
    public class PolylineData : BHoMObject
    {
        [Description("Shape to be updated.")]
        public virtual Polyline Path { get; set; } = null;

        [Description("Thickness of the path.")]
        public virtual double Thickness { get; set; } = 1;

        [Description("Colour of the edge of the path. No colour change will be made if left empty.")]
        public virtual string EdgeColour { get; set; } = "";

        [Description("Fill colour for the shape. If left empty, no fill will be applied.")]
        public virtual string FillColour { get; set; } = "";

        [Description("Opacity of the fill where 1 means full opacity and 0 means full transperency.")]
        public virtual double FillOpacity { get; set; } = 1.0;

        [Description("Toggles if the line should be dashed or not.")]
        public virtual bool IsDashed { get; set; } = false;
    }
}

