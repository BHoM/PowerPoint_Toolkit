using BH.oM.Base;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace BH.oM.PowerPoint.Layout.SlideParts
{
    public class SimpleTextElement : BHoMObject, IShapeElement
    {
        public virtual string Text { get; set; } = "";
    }
}