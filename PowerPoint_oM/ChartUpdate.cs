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

using BH.oM.Base;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace BH.oM.PowerPoint
{
    [Description("Allows to update the content of chart element.")]
    public class ChartUpdate : BHoMObject, ISlideUpdate
    {
        [Description("Number of the slide where the update needs to happen.")]
        public virtual int SlideNumber { get; set; } = 0;

        [Description("Name of the chart element that needs to be updated.")]
        public virtual string ElementName { get; set; } = "";

        [Description("New title for the chart. If left empty, the existing title will not be replaced.")]
        public virtual string Title { get; set; } = "";

        [Description("Names of the series.")]
        public virtual List<string> Series { get; set; } = new List<string>();

        [Description("Names of the categories.")]
        public virtual List<string> Categories { get; set; } = new List<string>();

        [Description("Numerical values for the chart data. There must be a list per serie and each list's length must be equal to the number of categories.")]
        public virtual List<List<double>> Data { get; set; } = new List<List<double>>();
    }
}

