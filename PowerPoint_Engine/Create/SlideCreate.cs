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
