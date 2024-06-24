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
        public static ISlideUpdate IToUpdate(this IShapeElement element, int slideNumber, string elementName)
        {
            return ToUpdate(element as dynamic, slideNumber, elementName);
        }

        private static ImageUpdate ToUpdate(ImageElement element, int slideNumber, string elementName)
        {
            return new ImageUpdate() { SlideNumber = slideNumber, ElementName=elementName, ImageFilePath = element.ImagePath };
        }

        private static MultiLineTextUpdate ToUpdate(MultiLineTextElement element, int slideNumber, string elementName)
        {
            return new MultiLineTextUpdate() { SlideNumber = slideNumber, ElementName = elementName, Text = element.Text, Colour = $"#{element.Colour.R:X2}{element.Colour.G:X2}{element.Colour.B:X2}" };
        }

        private static SimpleTextUpdate ToUpdate(SimpleTextElement element, int slideNumber, string elementName)
        {
            return new SimpleTextUpdate() { SlideNumber = slideNumber, ElementName = elementName, Text = element.Text };
        }

        private static ISlideUpdate ToUpdate(IShapeElement element, int slideNumber, string elementName)
        {
            BH.Engine.Base.Compute.RecordError($"The shape element type {element.GetType().FullName} is not supported for the PowerPointAdapter");
            return null;
        }
    }
}
