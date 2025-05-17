using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace WordPlus.Components.Document
{
    public class GH_WD_Doc_Merge : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_WD_Doc_Merge class.
        /// </summary>
        public GH_WD_Doc_Merge()
          : base("Merge Documents", "WD Merge",
              "Sequentially merge Documents into a single document and return document contents",
              Constants.ShortName, Constants.SubDocument)
        {
        }

        public override GH_Exposure Exposure => GH_Exposure.secondary;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Document.Name, Constants.Document.NickName, Constants.Document.Inputs, GH_ParamAccess.list);
            pManager[0].Optional = false;
            pManager.AddTextParameter("Name", "N", "Document Name", GH_ParamAccess.item);
            pManager[1].Optional = true;
            pManager.AddBooleanParameter("Page Breaks", "B", "Create at new page at the start of each document", GH_ParamAccess.item, false);
            pManager[2].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Document.Name, Constants.Document.NickName, Constants.Document.Output, GH_ParamAccess.item);
            pManager.AddGenericParameter(Constants.Content.Name, Constants.Content.NickName, Constants.Content.Output, GH_ParamAccess.list);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            List<IGH_Goo> goos = new List<IGH_Goo>();
            List<WdDocument> documents = new List<WdDocument>();
            if (!DA.GetDataList(0, goos)) return;
            foreach (IGH_Goo goo in goos) if (goo.CastTo<WdDocument>(out WdDocument document)) documents.Add(document);

            string name = "MergedDocument";
            DA.GetData(1, ref name);

            bool pageBreak = false;
            DA.GetData(2, ref pageBreak);

            WdDocument output = new WdDocument(documents, name, pageBreak);

            DA.SetData(0, output);
            DA.SetDataList(1, output.GetContents());
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
                return Properties.Resources.WD_Doc_Merge;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("8475B37A-15B8-42E7-ADCF-733CB56ADEEE"); }
        }
    }
}