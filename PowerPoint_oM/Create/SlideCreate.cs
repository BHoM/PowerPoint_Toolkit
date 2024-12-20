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
    [Description("Use with a push action to create a new slide from a slide layout in the first slide master template, at the position provided.")]
    public class SlideCreate : BHoMObject, ISlideCreate, IImmutable
    {
        [Description("The name of the slide master to get the layout from. If this is blank, the first slide master in the list will be used instead.")]
        public virtual string SlideMasterName { get; } = "";

        [Description("The name of the layout to use from the slide master.")]
        public virtual string LayoutName { get; } = "";

        [Description("The location to place the slide in the presentation, starting from 1. -1 to append the slide to the end of the presentation.")]
        public virtual int SlideNumber { get; } = -1;

        [Description("The slide updates to be applied to the created slide. Any set slide numbers will b")]
        public virtual List<ISlideUpdate> SlideUpdates { get; } = new List<ISlideUpdate>();

        public SlideCreate(string slideMasterName = "", string layoutName = "", int slideNumber = -1, List<ISlideUpdate> slideUpdates = null)
        {
            SlideMasterName = slideMasterName;
            LayoutName = layoutName;
            SlideNumber = slideNumber;
            SlideUpdates = slideUpdates;
        }
    }
}

