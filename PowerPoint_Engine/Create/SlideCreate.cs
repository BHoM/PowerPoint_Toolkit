using BH.oM.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BH.Engine.Adapters.PowerPoint
{
    public static partial class Create
    {
        public static SlideCreate SlideCreate(int slideNumber, string slideMasterName, string slideLayoutName, List<ISlideUpdate> slideUpdates)
        {
            foreach (ISlideUpdate update in slideUpdates)
            {
                update.SlideNumber = slideNumber;
            }

            SlideCreate slideCreate = new SlideCreate(slideMasterName, slideLayoutName, slideNumber, slideUpdates);
            return slideCreate;
        }
    }
}
