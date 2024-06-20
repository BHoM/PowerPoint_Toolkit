using BH.oM.Base;
using System;
using System.Collections.Generic;
using System.Text;

namespace BH.oM.PowerPoint
{
    public class SlideLayoutInfo : BHoMObject
    {
        public override string Name { get; set; } = "";

        public List<string> ElementNames { get; set; } = new List<string>();
    }
}
