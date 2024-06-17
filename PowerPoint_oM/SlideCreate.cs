using BH.oM.Base;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace BH.oM.PowerPoint
{
    [Description("Use with a push action to create a new slide from a slide layout in the slide master templates, at the position provided.")]
    public class SlideCreate : BHoMObject
    {
        [Description("The name of the slide master to try getting the layout from.")]
        public virtual string SlideMasterName { get; set; } = "";

        [Description("The name of the layout to use from the slide master.")]
        public virtual string LayoutName { get; set; } = "";

        [Description("The location to place the slide in the presentation, starting from 0. Use -1 to append the slide to the end of the presentation.")]
        public virtual int SlideIndex { get; set; } = -1;
    }
}
