using BH.oM.Base;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace BH.oM.PowerPoint
{
    public class TableUpdate : BHoMObject, ISlideUpdate
    {
        [Description("Number of the slide where the update needs to happen.")]
        public virtual int SlideNumber { get; set; } = -1;

        [Description("Name of the table that needs to be updated.")]
        public virtual string ElementName { get; set; }

        [Description("Content to be placed into the table. Outer list indexes correspond to row numbers, and inner list indexes correspond to the column numbers. each row must be the same length.")]
        public virtual List<List<string>> Contents { get; set; }

        [Description("The font size of any text in the table.")]
        public virtual int UpdatedTextFontSize { get; set; } = 20;
    }
}
