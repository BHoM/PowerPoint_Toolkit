using BH.oM.Base;
using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text;

namespace BH.oM.PowerPoint
{
    public class SlideMasterInfo : BHoMObject
    {
        public override string Name { get; set; } = "";

        public virtual List<SlideLayoutInfo> Templates { get; set; } = new List<SlideLayoutInfo>();
    }
}
