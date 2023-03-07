using BH.oM.Base;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace BH.oM.PowerPoint
{
    [Description("Allows to update the text content of a shape")]
    public class SimpleTextUpdate : BHoMObject, ISlideUpdate
    {
        [Description("Number of the slide where the update needs to happen.")]
        public virtual int SlideNumber { get; set; } = 0;

        [Description("Name of the text element that needs to be updated.")]
        public virtual string ElementName { get; set; } = "";

        [Description("New text for the element.")]
        public virtual string Text { get; set; } = "";
    }
}
