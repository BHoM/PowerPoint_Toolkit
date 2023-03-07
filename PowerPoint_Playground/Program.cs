using BH.Adapter.PowerPoint;
using BH.oM.Adapter;
using BH.oM.Adapters.Excel;
using BH.oM.PowerPoint;
using System;
using System.Collections.Generic;

namespace PowerPoint_Playground
{
    class Program
    {
        static void Main(string[] args)
        {
            FileSettings templateSettings = new FileSettings { Directory = @"C:\Users\adecler\Documents", FileName = "TemplateTest.pptx" };
            FileSettings outputSettings = new FileSettings { Directory = @"C:\Users\adecler\Documents", FileName = "test.pptx" };
            PowerPointAdapter adapter = new PowerPointAdapter(templateSettings, outputSettings);

            adapter.Push(new List<object> {
                new SimpleTextUpdate { SlideNumber = 3, ElementName = "TemplateTitle", Text = "New slide title" },
                new SimpleTextUpdate { SlideNumber = 3, ElementName = "TemplateText_01", Text = "New text" },
                new ImageUpdate { SlideNumber = 3, ElementName = "TemplatePicture_01", ImageFilePath = @"C:\Users\adecler\OneDrive - BuroHappold\Pictures\Vega_Toolkit01.JPG" },
                new ImageUpdate { SlideNumber = 3, ElementName = "TemplatePicture_02", ImageFilePath = @"C:\Users\adecler\OneDrive - BuroHappold\Pictures\InstallerCreation.jpg" },
                new ChartUpdate {
                    SlideNumber = 3, 
                    ElementName = "TemplateChart_01", 
                    Title = "New Chart Title",
                    Series = new List<string> { "S1", "S2", "S3", "S4" },
                    Categories = { "C1", "C2", "C3", "C4", "C5" },
                    Data = new List<List<double>> {
                        new List<double> { 30.1, 50.2, 40.1, 20.1, 15.3 },
                        new List<double> { 50.1, 40.2, 20.1, 30.3, 10.6 },
                        new List<double> { 51.1, 41.2, 21.1, 31.3, 11.6 },
                        new List<double> { 52.1, 42.2, 22.1, 32.3, 12.6 }
                    }
                }
            });
        }
    }
}
