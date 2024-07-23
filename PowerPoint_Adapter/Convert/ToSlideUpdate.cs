using BH.oM.PowerPoint;
using BH.oM.PowerPoint.Layout.SlideParts;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace BH.Adapter.PowerPoint
{
    public static partial class Convert
    {
        public static ISlideUpdate IToSlideUpdate(this IShapeElement element, int slideNumber, string elementName)
        {
            return ToSlideUpdate(element as dynamic, slideNumber, elementName);
        }

        private static ImageUpdate ToSlideUpdate(ImageElement element, int slideNumber, string elementName)
        {
            return new ImageUpdate() { SlideNumber = slideNumber, ElementName=elementName, ImageFilePath = element.ImagePath };
        }

        private static MultiLineTextUpdate ToSlideUpdate(MultiLineTextElement element, int slideNumber, string elementName)
        {
            return new MultiLineTextUpdate() { SlideNumber = slideNumber, ElementName = elementName, Text = element.Text, Colour = $"#{element.Colour.R:X2}{element.Colour.G:X2}{element.Colour.B:X2}" };
        }

        private static SimpleTextUpdate ToSlideUpdate(SimpleTextElement element, int slideNumber, string elementName)
        {
            return new SimpleTextUpdate() { SlideNumber = slideNumber, ElementName = elementName, Text = element.Text };
        }

        private static ISlideUpdate ToSlideUpdate(IShapeElement element, int slideNumber, string elementName)
        {
            BH.Engine.Base.Compute.RecordError($"The shape element type {element.GetType().FullName} is not supported for the PowerPointAdapter");
            return null;
        }
    }
}
