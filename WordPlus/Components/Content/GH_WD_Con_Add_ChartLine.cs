using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

using Sd = System.Drawing;

namespace WordPlus.Components
{
    public class GH_WD_Con_ChartLine : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the MyComponent1 class.
        /// </summary>
        public GH_WD_Con_ChartLine()
          : base("Word Line Chart", "WD Ln Cht",
              "Construct a Word Line Chart Content Object from a list of numbers",
              Constants.ShortName, Constants.SubContent)
        {
        }

        public override GH_Exposure Exposure => GH_Exposure.senary;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddNumberParameter("Numbers", "N", "A datatree of numbers, each branch is a series and each list item a category (Note charts do not respect font formatting)", GH_ParamAccess.tree);
            pManager[0].Optional = false;
            pManager.AddGenericParameter("Title", "T", "The chart title as a paragraph, fragment, or string. (Note charts do not respect font formatting)", GH_ParamAccess.item);
            pManager[1].Optional = true;
            pManager.AddGenericParameter("Axis Labels", "A", "An optional list of labels per list item as paragraphs, fragments, or strings. These are the values that will be plotted along the Axis opposite the values (Note charts do not respect font formatting)", GH_ParamAccess.list);
            pManager[2].Optional = true;
            pManager.AddGenericParameter("Series Labels", "S", "An optional list of labels per branch as paragraphs, fragments, or strings. These are the values that will show up on the legend. (Note charts do not respect font formatting)", GH_ParamAccess.list);
            pManager[3].Optional = true;
            pManager.AddColourParameter("Series Colors", "C", "An optional list of colors per branch.", GH_ParamAccess.list);
            pManager[4].Optional = true;
            pManager.AddIntegerParameter("Legend Location", "L", "The location of the chart legend", GH_ParamAccess.item, 0);
            pManager[5].Optional = true;

            Param_Integer paramA = (Param_Integer)pManager[5];
            foreach (Content.LegendLocations value in Enum.GetValues(typeof(Content.LegendLocations)))
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
            Paragraph title = new Paragraph("Line Chart"); ;
            if (DA.GetData(1, ref gooA)) if (!gooA.TryGetParagraph(out title)) ;

            List<List<double>> dataSet = new List<List<double>>();
            if (!DA.GetDataTree(0, out GH_Structure<GH_Number> ghData)) return;

            foreach (List<GH_Number> data in ghData.Branches)
            {
                List<double> values = new List<double>();
                foreach (GH_Number value in data)
                {
                    values.Add(value.Value);
                }
                dataSet.Add(values);
            }

            List<IGH_Goo> goosA = new List<IGH_Goo>();
            DA.GetDataList(2, goosA);
            List<Paragraph> axis = new List<Paragraph>();
            foreach (IGH_Goo goo in goosA) if (goo.TryGetParagraph(out Paragraph paragraph)) axis.Add(paragraph);

            List<IGH_Goo> goosB = new List<IGH_Goo>();
            DA.GetDataList(3, goosB);
            List<Paragraph> series = new List<Paragraph>();
            foreach (IGH_Goo goo in goosB) if (goo.TryGetParagraph(out Paragraph paragraph)) series.Add(paragraph);

            List<Sd.Color> colors = new List<Sd.Color>();
            DA.GetDataList(4, colors);

            int legend = 0;
            DA.GetData(5, ref legend);

            Content content = Content.CreateChartLineContent(title, dataSet, axis, series, colors, (Content.LegendLocations)legend);

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
                return Properties.Resources.WD_Con_Chart_Line;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("A3FFD8D2-3F74-431A-9FBB-48FB91A9F823"); }
        }
    }
}