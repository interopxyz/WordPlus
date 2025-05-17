using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace WordPlus.Components
{
    public class GH_WD_Doc_AddContents : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_WB_Doc_AddContents class.
        /// </summary>
        public GH_WD_Doc_AddContents()
          : base("Document Contents", "WD Content",
              "Sequentially adds contents to a document and return document contnets",
              Constants.ShortName, Constants.SubDocument)
        {
        }

        public override GH_Exposure Exposure => GH_Exposure.secondary;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Document.Name, Constants.Document.NickName, Constants.Document.Input, GH_ParamAccess.item);
            pManager[0].Optional = true;
            pManager.AddGenericParameter(Constants.Content.Name, Constants.Content.NickName, "A list of Word Content Objects to add to the Word Document", GH_ParamAccess.list);
            pManager[1].Optional = false;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Document.Name, Constants.Document.NickName, Constants.Document.Output, GH_ParamAccess.item);
            pManager.AddGenericParameter(Constants.Content.Name, Constants.Content.NickName, Constants.Content.Outputs, GH_ParamAccess.list);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            WdDocument document = new WdDocument();
            IGH_Goo gooA = null;
            if (DA.GetData(0, ref gooA))
            {
                if (!gooA.CastTo<WdDocument>(out document))
                {
                    this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Doc input must be a Word Document Object");
                    return;
                }
                document = new WdDocument(document);
            }

            List<IGH_Goo> goos = new List<IGH_Goo>();
            if (!DA.GetDataList(1, goos)) return;
            foreach(IGH_Goo goo in goos) if (goo.TryGetContent(out WdContent content)) document.AddContent(content);

            DA.SetData(0, document);
            DA.SetDataList(1, document.GetContents());
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
                return Properties.Resources.WD_Doc_Add_Content;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("D5F294D6-494A-4225-9C19-BDBC34EB9073"); }
        }
    }
}