using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace WordPlus.Components.Text
{
    public class GH_WD_Txt_Mod_TextSplit : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_WD_Txt_Mod_GetFragments class.
        /// </summary>
        public GH_WD_Txt_Mod_TextSplit()
          : base("Split Paragraph", "SplitPara",
              "Splits a Text Paragraph Object into it's Text Fragment Objects and Text Strings",
              Constants.ShortName, Constants.Format)
        {
        }
        public override GH_Exposure Exposure => GH_Exposure.quinary;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Paragraph.Name, Constants.Paragraph.NickName, Constants.Paragraph.Input, GH_ParamAccess.item);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Fragment.Name, Constants.Fragment.NickName, Constants.Fragment.Input, GH_ParamAccess.list);
            pManager.AddGenericParameter("Text", "T", "The text content string", GH_ParamAccess.list);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo gooA = null;
            if (!DA.GetData(0, ref gooA)) return;

            if (!gooA.TryGetParagraph(out WdParagraph paragraph))
            {
                this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Txt input must be a Text Paragraph Object, Text Fragment Object or a string");
                return;
            }

            List<WdFragment> fragments = new List<WdFragment>();
            List<string> texts = new List<string>();
            foreach (WdFragment fragment in paragraph.Fragments)
            {
                fragments.Add(new WdFragment(fragment));
                texts.Add(fragment.Text);
            }

            DA.SetDataList(0, fragments);
            DA.SetDataList(1, texts);
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
                return Properties.Resources.WD_Txt_Split;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("CE1F54DB-7400-4FC6-B6B8-4062028A3A92"); }
        }
    }
}