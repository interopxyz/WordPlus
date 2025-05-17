using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace WordPlus.Components.Document
{
    public class GH_WD_Doc_Stream : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_WD_Doc_Stream class.
        /// </summary>
        public GH_WD_Doc_Stream()
          : base("Write Document", "Write Doc",
              "Write a Document to a memory stream",
              Constants.ShortName, Constants.SubDocument)
        {
        }

        public override GH_Exposure Exposure => GH_Exposure.septenary;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Document.Name, Constants.Document.NickName, Constants.Document.Output, GH_ParamAccess.item);
            pManager[0].Optional = false;

            pManager.AddBooleanParameter(Constants.Activate.Name, Constants.Activate.NickName, Constants.Activate.Input, GH_ParamAccess.item);
            pManager[1].Optional = false;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("Memory Stream", "S", "The document as a memory stream", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            bool activate = false;
            if (DA.GetData(1, ref activate))
            {
                if (activate)
                {
                    IGH_Goo gooA = null;
                    if (!DA.GetData(0, ref gooA)) return;

                    if (!gooA.CastTo<WdDocument>(out WdDocument document))
                    {
                        this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Doc input must be a Word Document Object");
                        return;
                    }
                    document = new WdDocument(document);

                    DA.SetData(0,document.StreamDocument());
                }
            }
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
                return Properties.Resources.WD_Doc_Stream_Write;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("09BEAC86-082A-4756-B46B-9A82EE978ADE"); }
        }
    }
}