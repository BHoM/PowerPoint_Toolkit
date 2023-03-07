using BH.oM.Base;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace BH.oM.PowerPoint
{
    public interface ISlideUpdate : IBHoMObject
    {
        [Description("Number of the slide where the update needs to happen.")]
        int SlideNumber { get; set; }

        [Description("Name of the element that needs to be updated.")]
        string ElementName { get; set; }
    }
}
