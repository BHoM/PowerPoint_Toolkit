using BH.oM.Base;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace BH.oM.PowerPoint
{
    [Description("Allows to replace the image content of an image element.")]
    public class ImageUpdate : BHoMObject, ISlideUpdate
    {
        [Description("Number of the slide where the update needs to happen.")]
        public virtual int SlideNumber { get; set; } = 0;

        [Description("Name of the image element that needs to be updated.")]
        public virtual string ElementName { get; set; } = "";

        [Description("File path of the new image.")]
        public virtual string ImageFilePath { get; set; } = "";
    }
}
