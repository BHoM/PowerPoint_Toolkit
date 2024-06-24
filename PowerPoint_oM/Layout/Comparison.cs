using BH.oM.Base;
using BH.oM.PowerPoint.Layout.SlideParts;
using System;
using System.Collections.Generic;
using System.Text;

namespace BH.oM.PowerPoint.Layout
{
    public class Comparison : BHoMObject, ISlideLayout
    {
        public virtual SimpleTextElement TitleElement { get; set; } = new SimpleTextElement();

        public virtual SimpleTextElement LeftSubtitleElement { get; set; } = new SimpleTextElement();

        public virtual IShapeElement LeftContentElement { get; set; } = new MultiLineTextElement();

        public virtual SimpleTextElement RightSubtitleElement { get; set; } = new SimpleTextElement();

        public virtual IShapeElement RightContentElement { get; set; } = new MultiLineTextElement();
    }
}
