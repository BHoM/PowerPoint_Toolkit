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

using BH.oM.Base.Attributes;
using BH.oM.PowerPoint;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace BH.Engine.Adapters.PowerPoint
{
    public static partial class Create
    {
        [Description("Create a command that tells the PowerPointAdapter to create a slide using the slide layout provided in the slide master at the position given and update the slide with the given updates.")]
        [Input("slideMasterName", "The name of the slide master within the presentation to use. Input an empty string to use the first slide master in the list.")]
        [Input("slideLayoutName", "The name of the slide layout within the slide master to use.")]
        [Input("slideNumber", "The position that the slide should be in when created. If the slide should be appended to the end of the presentation, set to -1 or a number larger than the number of slides in the presentation.")]
        [Input("slideUpdates", "A list of slide updates to apply to the created slide. The slide number for each update should be ignored as these are overwritten when creating a slide.")]
        [Output("slideCreate", "The resultant SlideCreate command.")]
        public static SlideCreate SlideCreate(string slideMasterName, string slideLayoutName, int slideNumber = -1, List<ISlideUpdate> slideUpdates = null)
        {
            if (slideUpdates == null)
                slideUpdates = new List<ISlideUpdate>();

            foreach (ISlideUpdate update in slideUpdates)
                update.SlideNumber = slideNumber;

            SlideCreate slideCreate = new SlideCreate(slideMasterName, slideLayoutName, slideNumber, slideUpdates);
            return slideCreate;
        }
    }
}


