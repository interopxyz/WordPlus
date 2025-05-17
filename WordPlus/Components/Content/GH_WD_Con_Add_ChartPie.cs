using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

using Sd = System.Drawing;

namespace WordPlus.Components.Content
{
    public class GH_WD_Con_ChartPie : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_WD_Con_ChartPie class.
        /// </summary>
        public GH_WD_Con_ChartPie()
          : base("Pie Chart", "ChartPie",
              "Construct a Pie Chart Content Object from a list of numbers",
              Constants.ShortName, Constants.SubContent)
        {
        }

        public override GH_Exposure Exposure => GH_Exposure.senary;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddNumberParameter("Numbers", "N", "A list of numbers, each branch is a series and each list item a category", GH_ParamAccess.list);
            pManager[0].Optional = false;
            pManager.AddGenericParameter("Title", "T", "The chart title as a paragraph, fragment, or string. (Note charts do not respect font formatting)", GH_ParamAccess.item);
            pManager[1].Optional = true;
            pManager.AddGenericParameter("Series Labels", "S", "An optional list of labels per item as paragraphs, fragments, or strings. These are the values that will show up on the legend. (Note charts do not respect font formatting)", GH_ParamAccess.list);
            pManager[2].Optional = true;
            pManager.AddIntegerParameter("Legend Location", "L", "The location of the chart legend", GH_ParamAccess.item, 0);
            pManager[3].Optional = true;

            Param_Integer paramA = (Param_Integer)pManager[3];
            foreach (WdContent.LegendLocations value in Enum.GetValues(typeof(WdContent.LegendLocations)))
            {
                paramA.AddNamedValue((int)value + " | " + value.ToString(), (int)value);
            }
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Content.Name, Constants.Content.NickName, Constants.Content.Output, GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo gooA = null;
            WdParagraph title = new WdParagraph("Pie Chart"); ;
            if (DA.GetData(1, ref gooA)) if (!gooA.TryGetParagraph(out title)) ;

            List<double> dataSet = new List<double>();
            if (!DA.GetDataList(0, dataSet)) return;

            List<IGH_Goo> goos = new List<IGH_Goo>();
            DA.GetDataList(2, goos);
            List<WdParagraph> series = new List<WdParagraph>();
            foreach (IGH_Goo goo in goos) if (goo.TryGetParagraph(out WdParagraph paragraph)) series.Add(paragraph);

            List<Sd.Color> colors = new List<Sd.Color>();
            //DA.GetDataList(3, colors);

            int legend = 0;
            DA.GetData(3, ref legend);

            WdContent content = WdContent.CreateChartPieContent(title, dataSet, series, colors, (WdContent.LegendLocations)legend);

            DA.SetData(0, content);
        }

        /// <summary>
        /// Provides an Icon for the component.
        /// </summary>
        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                //You can add image files to your project resources and access them like this:
                // return Resources.IconForThisComponent;
                return Properties.Resources.WD_Con_Chart_Pie;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("61E80D2B-1A75-437E-A0A3-A78CD1D971AA"); }
        }
    }
}