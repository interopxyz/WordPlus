using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Rhino.Geometry;

namespace WordPlus.Components
{
    public class GH_WD_Con_BreakPage : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_WB_Con_PageBreak class.
        /// </summary>
        public GH_WD_Con_BreakPage()
          : base("Page Break Contents", "PgBrk",
              "Construct a Page Break Content Object",
              Constants.ShortName, Constants.SubContent)
        {
        }

        public override GH_Exposure Exposure => GH_Exposure.quinary;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
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
            WdContent content = WdContent.CreatePageBreakContent();

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
                return Properties.Resources.WD_Con_Break;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("592E7181-D42F-4DF5-8318-8EC555551F7A"); }
        }
    }
}