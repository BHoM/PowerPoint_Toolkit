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
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace BH.oM.PowerPoint
{
    public class TableUpdate : BHoMObject, ISlideUpdate
    {
        [Description("Number of the slide where the update needs to happen.")]
        public virtual int SlideNumber { get; set; } = -1;

        [Description("Name of the table that needs to be updated.")]
        public virtual string ElementName { get; set; } = "";

        [Description("Content to be placed into the table. Outer list indexes correspond to row numbers, and inner list indexes correspond to the column numbers. each row must be the same length.")]
        public virtual List<List<string>> Contents { get; set; } = new List<List<string>>();

        [Description("The font size of any text in the table.")]
        public virtual int UpdatedTextFontSize { get; set; } = 20;
    }
}
