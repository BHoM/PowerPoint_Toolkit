using BH.oM.Base;
using BH.oM.PowerPoint.Layout.SlideParts;
using System;
using System.Collections.Generic;
using System.Text;

namespace BH.oM.PowerPoint.Layout
{
    public class SectionHeader : BHoMObject, ISlideLayout
    {
        public virtual SimpleTextElement TitleElement { get; set; } = new SimpleTextElement();

        public virtual SimpleTextElement SubtitleElement { get; set; } = new SimpleTextElement();
    }
}