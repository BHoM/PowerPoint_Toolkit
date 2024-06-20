using BH.oM.Base;
using System;
using System.Collections.Generic;
using System.Text;

namespace BH.oM.PowerPoint
{
    public class SlideInfo : BHoMObject
    {
        public int SlideNumber { get; set; }

        public List<string> ElementNames { get; set; } = new List<string>();
    }
}
