/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2023, the respective contributors. All rights reserved.
 *
 * Each contributor holds copyright over their respective contributions.
 * The project versioning (Git) records all such contribution source information.
 *                                           
 *                                                                              
 * The BHoM is free software: you can redistribute it and/or modify         
 * it under the terms of the GNU Lesser General Public License as published by  
 * the Free Software Foundation, either version 3.0 of the License, or          
 * (at your option) any later version.                                          
 *                                                                              
 * The BHoM is distributed in the hope that it will be useful,              
 * but WITHOUT ANY WARRANTY; without even the implied warranty of               
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the                 
 * GNU Lesser General Public License for more details.                          
 *                                                                            
 * You should have received a copy of the GNU Lesser General Public License     
 * along with this code. If not, see <https://www.gnu.org/licenses/lgpl-3.0.html>.      
 */

using BH.Adapter;
using BH.oM.Base.Attributes;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BH.Adapter.PowerPoint
{
    public partial class PowerPointAdapter : BHoMAdapter
    {
        /***************************************************/
        /**** Constructors                              ****/
        /***************************************************/

        [Description("Adapter to create a new PowerPoint file based on an existing template.")]
        [Input("templateFileSettings", "Defines the location of the template PowerPoint file.")]
        [Output("outputFileSettings", "Defines the location of the new PowerPoint file.")]
        public PowerPointAdapter(BH.oM.Adapter.FileSettings templateFileSettings, BH.oM.Adapter.FileSettings outputFileSettings)
        {
            if (templateFileSettings == null)
            {
                BH.Engine.Base.Compute.RecordError("Please set the File Settings for the template file to enable the PowerPoint Adapter to work correctly.");
                return;
            }

            if (!Path.HasExtension(templateFileSettings.FileName) || Path.GetExtension(templateFileSettings.FileName) != ".pptx")
            {
                BH.Engine.Base.Compute.RecordError("PowerPoint adapter supports only .pptx files.");
                return;
            }

            m_TemplateFileSettings = templateFileSettings;

            m_OutputFileSettings = outputFileSettings;
            if (!Directory.Exists(m_OutputFileSettings.Directory))
                Directory.CreateDirectory(m_OutputFileSettings.Directory);
        }

        /***************************************************/

        [Description("Adapter to create a new PowerPoint file based on an existing template.")]
        [Input("templateStream", "Defines the content of the template PowerPoint file.")]
        [Output("outputStream", "Defines the content of the new PowerPoint file.")]
        public PowerPointAdapter(Stream templateStream, Stream outputStream)
        {
            if (templateStream == null)
            {
                BH.Engine.Base.Compute.RecordError("Please set the Stream for the template to enable the PowerPoint Adapter to work correctly.");
                return;
            }

            if (outputStream == null)
            {
                BH.Engine.Base.Compute.RecordError("Please set the Stream for the output to enable the PowerPoint Adapter to work correctly.");
                return;
            }

            m_TemplateStream = templateStream;
            m_OutputStream = outputStream;
        }


        /***************************************************/
        /**** Private Fields                            ****/
        /***************************************************/

        private BH.oM.Adapter.FileSettings m_TemplateFileSettings = null;

        private BH.oM.Adapter.FileSettings m_OutputFileSettings = null;

        private Stream m_TemplateStream = null;

        private Stream m_OutputStream = null;

        /***************************************************/
    }
}

