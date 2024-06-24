using BH.oM.Base;
using System;
using System.Collections.Generic;
using System.Text;

namespace BH.oM.PowerPoint.Layout.SlideParts
{
    public class ImageElement : BHoMObject, IShapeElement
    {
        public virtual string ImagePath { get; set; } = "";
    }
}