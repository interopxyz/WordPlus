using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace WordPlus.Components.Modify
{
    public class GH_WD_Con_Mod_Align : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_WD_Con_Mod_Align class.
        /// </summary>
        public GH_WD_Con_Mod_Align()
          : base("Content Align", "ConAlign",
              "Modify Content Alignment if applicable",
              Constants.ShortName, Constants.Modify)
        {
        }

        public override GH_Exposure Exposure => GH_Exposure.secondary;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Content.Name, Constants.Content.NickName, Constants.Content.Input, GH_ParamAccess.item);
            pManager[0].Optional = false;
            pManager.AddIntegerParameter("Horizontal Alignment", "H", "Optionally override the horizontal alignment", GH_ParamAccess.item);
            pManager[1].Optional = true;
            //pManager.AddIntegerParameter("Vertical Alignment", "V", "Optionally override the vertical alignment", GH_ParamAccess.item);
            //pManager[2].Optional = true;

            Param_Integer paramA = (Param_Integer)pManager[1];
            foreach (Font.HorizontalAlignments value in Enum.GetValues(typeof(Font.HorizontalAlignments)))
            {
                paramA.AddNamedValue(value.ToString(), (int)value);
            }

            //Param_Integer paramB = (Param_Integer)pManager[2];
            //foreach (Font.VerticalAlignments value in Enum.GetValues(typeof(Font.VerticalAlignments)))
            //{
            //    paramB.AddNamedValue(value.ToString(), (int)value);
            //}
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Content.Name, Constants.Content.NickName, Constants.Content.Input, GH_ParamAccess.item);
            pManager[0].Optional = false;
            pManager.AddIntegerParameter("Horizontal Alignment", "H", "Get the horizontal alignment", GH_ParamAccess.item);
            //pManager.AddIntegerParameter("Vertical Alignment", "V", "Get the vertical alignment", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo gooA = null;
            if (!DA.GetData(0, ref gooA)) return;

            if (!gooA.TryGetContent(out WdContent content))
            {
                this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Con input must be a Content Object, Text Fragment Object, or a string");
                return;
            }

            int horizontal = 0;
            if (DA.GetData(1, ref horizontal)) content.Font.HorizontalAlignment = (Font.HorizontalAlignments)horizontal;

            int vertical = 0;
            //if (DA.GetData(2, ref vertical)) content.Font.VerticalAlignment = (Font.VerticalAlignments)vertical;

            DA.SetData(0, content);
            DA.SetData(1, content.Font.HorizontalAlignment);
            //DA.SetData(2, content.Font.VerticalAlignment);
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
                return Properties.Resources.WD_Con_Mod_Align;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("D6CBDDA0-E311-41AB-ABB1-751D254E97AF"); }
        }
    }
}