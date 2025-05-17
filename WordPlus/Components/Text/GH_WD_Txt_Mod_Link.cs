using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace WordPlus.Components.Text
{
    public class GH_WD_Txt_Mod_Link : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_WD_Txt_Mod_Link class.
        /// </summary>
        public GH_WD_Txt_Mod_Link()
          : base("Content Link", "ConLink",
              "Modify Content Link if applicable",
              Constants.ShortName, Constants.Format)
        {
        }

        public override GH_Exposure Exposure => GH_Exposure.secondary;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Paragraph.Name, Constants.Paragraph.NickName, Constants.Paragraph.Input, GH_ParamAccess.item);
            pManager[0].Optional = false;
            pManager.AddTextParameter("Hyperlink", "L", "The optional hyperlink for the text", GH_ParamAccess.item);
            pManager[1].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Paragraph.Name, Constants.Paragraph.NickName, Constants.Paragraph.Output, GH_ParamAccess.item);
            pManager.AddTextParameter("Hyperlink", "L", "The optional hyperlink for the text", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo gooA = null;
            if (!DA.GetData(0, ref gooA)) return;

            WdParagraph paragraph = null;
            bool isParagraph = false;
            WdFragment fragment = null;

            if (gooA.TryCastToParagraph(out paragraph))
            {
                isParagraph = true;
            }
            else if (!gooA.TryGetFragment(out fragment))
            {
                this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Txt input must be a Text Paragraph Object, Text Fragment Object or a string");
                return;
            }

            string link = string.Empty;
            if (DA.GetData(1, ref link))
            {
                if (isParagraph)
                {
                    foreach (WdFragment f in paragraph.Fragments) f.Hyperlink = link;
                }
                else
                {
                    fragment.Hyperlink = link;
                }
            }

            if (isParagraph)
            {
                DA.SetData(0, paragraph);
                DA.SetData(1, paragraph.Fragments[0].Hyperlink);
            }
            else
            {
                DA.SetData(0, fragment);
                DA.SetData(1, fragment.Hyperlink);
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
                return Properties.Resources.WD_Page_Txt_Link;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("B5399F24-7EF9-41CE-893D-832F8BFECB40"); }
        }
    }
}