using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace WordPlus.Components.Document
{
    public class GH_WD_Doc_AddFooter : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_WD_Doc_AddFooter class.
        /// </summary>
        public GH_WD_Doc_AddFooter()
          : base("Footer Contents", "WD Footer",
              "Sequentially adds contents to a document footer and return footer contents",
              Constants.ShortName, Constants.SubDocument)
        {
        }

        public override GH_Exposure Exposure => GH_Exposure.tertiary;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Document.Name, Constants.Document.NickName, Constants.Document.Input, GH_ParamAccess.item);
            pManager[0].Optional = false;
            pManager.AddGenericParameter(Constants.Content.Name, Constants.Content.NickName, "A list of Word Content Objects to add to the Word Document", GH_ParamAccess.list);
            pManager[1].Optional = false;
            pManager.AddIntegerParameter("Type", "T", "Header Type", GH_ParamAccess.item, 1);
            pManager[2].Optional = true;

            Param_Integer paramA = (Param_Integer)pManager[2];
            foreach (WdDocument.HeaderFooterTypes value in Enum.GetValues(typeof(WdDocument.HeaderFooterTypes)))
            {
                paramA.AddNamedValue((int)value + " | " + value.ToString(), (int)value);
            }

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
            IGH_Goo gooA = null;
            if (!DA.GetData(0, ref gooA)) return;

            if (!gooA.CastTo<WdDocument>(out WdDocument document))
            {
                this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Doc input must be a Word Document Object");
                return;
            }
            document = new WdDocument(document);

            int type = 1;
            DA.GetData(2, ref type);
            document.FooterType = (WdDocument.HeaderFooterTypes)type;

            List<IGH_Goo> goos = new List<IGH_Goo>();
            if (!DA.GetDataList(1, goos)) return;
            foreach (IGH_Goo goo in goos) if (goo.TryGetContent(out WdContent content)) document.Footer.Add(content);

            DA.SetData(0, document);
            DA.SetDataList(1, document.Footer.GetContents());
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
                return Properties.Resources.WD_Doc_Footer;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("036B6C17-4A0D-43F6-B3F6-45948BEFC03B"); }
        }
    }
}