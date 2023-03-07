using BH.oM.Adapters.Excel;
using BH.oM.Base;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace BH.oM.PowerPoint
{
    [Description("Allows to update the content of chart element.")]
    public class ChartUpdate : BHoMObject, ISlideUpdate
    {
        [Description("Number of the slide where the update needs to happen.")]
        public virtual int SlideNumber { get; set; } = 0;

        [Description("Name of the chart element that needs to be updated.")]
        public virtual string ElementName { get; set; } = "";

        [Description("New title for the chart. If left empty, the existing title will not be replaced.")]
        public virtual string Title { get; set; } = "";

        [Description("Names of the series.")]
        public virtual List<string> Series { get; set; } = new List<string>();

        [Description("Names of the categories.")]
        public virtual List<string> Categories { get; set; } = new List<string>();

        [Description("Numerical values for the chart data. There must be a list per serie and each list's length must be equal to the number of categories.")]
        public virtual List<List<double>> Data { get; set; } = new List<List<double>>();
    }
}
