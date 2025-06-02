using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace WordPlus.Components.Document
{
    public class GH_WD_Doc_Read : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_WD_Doc_Read class.
        /// </summary>
        public GH_WD_Doc_Read()
          : base("Read Word Document", "WD Read Doc",
              "Read a Word Document from a memory stream",
              Constants.ShortName, Constants.SubDocument)
        {
        }

        public override GH_Exposure Exposure => GH_Exposure.septenary;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("Memory Stream String", "S", "ASCII string of a memory stream", GH_ParamAccess.item);
            pManager[0].Optional = true;
            pManager.AddBooleanParameter(Constants.Activate.Name, Constants.Activate.NickName, Constants.Activate.Input, GH_ParamAccess.item);
            pManager[1].Optional = false;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Document.Name, Constants.Document.NickName, Constants.Document.Output, GH_ParamAccess.item);
            pManager[0].Optional = false;
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
                    string streamText = string.Empty;
                    if (DA.GetData(0, ref streamText))
                    {
                        WdDocument document = WdDocument.ReadDocument(streamText);
                        DA.SetData(0, document);
                    }
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
                return Properties.Resources.WD_Doc_Stream_Read;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("dc0b0226-3d8f-427d-b07f-d5dff5b0c98a"); }
        }
    }
}