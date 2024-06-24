using BH.oM.Base;
using BH.oM.PowerPoint.Layout.SlideParts;
using System;
using System.Collections.Generic;
using System.Text;

namespace BH.oM.PowerPoint.Layout
{
    public class PictureWithCaption : BHoMObject, ISlideLayout
    {
        public virtual SimpleTextElement TitleElement { get; set; } = new SimpleTextElement();

        public virtual MultiLineTextElement CaptionElement { get; set; } = new MultiLineTextElement();

        public virtual ImageElement ImageElement { get; set; } = new ImageElement();
    }
}
