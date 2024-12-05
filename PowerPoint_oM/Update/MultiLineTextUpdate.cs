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
    [Description("Allows to update the text content of a shape")]
    public class MultiLineTextUpdate : BHoMObject, ISlideUpdate
    {
        [Description("Number of the slide where the update needs to happen.")]
        public virtual int SlideNumber { get; set; } = 0;

        [Description("Name of the text element that needs to be updated.")]
        public virtual string ElementName { get; set; } = "";

        [Description("New text for the element. Each item in the list represents a new row.")]
        public virtual List<string> Text { get; set; } = new List<string>();

        [Description("Colour of the text element that needs to be updated.")]
        public virtual string Colour { get; set; } = "";

        [Description("Whether to use the properties of the previous paragraph for subsequent paragraphs where there is no text to update. If set to false, new paragraphs will have default properties with no indentation. Default false.")]
        public virtual bool UseLastParagraphProperties { get; set; } = false;
    }
}

