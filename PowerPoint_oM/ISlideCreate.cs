using BH.oM.Base;
using System;
using System.Collections.Generic;
using System.Text;

namespace BH.oM.PowerPoint
{
    public interface ISlideCreate : IBHoMObject
    {
        int SlideIndex { get; set; }

        string LayoutName { get; set; }
    }
}
