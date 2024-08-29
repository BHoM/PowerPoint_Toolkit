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
