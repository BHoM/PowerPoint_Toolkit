using BH.oM.Base;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace BH.oM.PowerPoint.Layout.SlideParts
{
    public class MultiLineTextElement : BHoMObject, IShapeElement
    {
        public virtual List<string> Text { get; set; } = new List<string>();

        public virtual Color Colour { get; set; } = Color.Black;
    }
}