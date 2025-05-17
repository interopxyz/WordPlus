using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

using Sd = System.Drawing;

namespace WordPlus.Components.Content
{
    public class GH_WD_Txt_Mod_FontSetting : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_WD_Con_Mod_Font class.
        /// </summary>
        public GH_WD_Txt_Mod_FontSetting()
          : base("Content Font", "ConFont",
              "Modify Content Font if applicable",
              Constants.ShortName, Constants.Format)
        {
        }

        public override GH_Exposure Exposure => GH_Exposure.primary;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Paragraph.Name, Constants.Paragraph.NickName, Constants.Paragraph.Input, GH_ParamAccess.item);
            pManager[0].Optional = false;
            pManager.AddTextParameter("Family Name", "F", "Optionally set Font Family name", GH_ParamAccess.item);
            pManager[1].Optional = true;
            pManager.AddNumberParameter("Size", "S", "Optionaly set Text Size", GH_ParamAccess.item);
            pManager[2].Optional = true;
            pManager.AddColourParameter("Color", "C", "Optionaly set Text Color", GH_ParamAccess.item);
            pManager[3].Optional = true;
            pManager.AddIntegerParameter("Highlight", "H", "Optionaly set Text Highlight Color", GH_ParamAccess.item);
            pManager[4].Optional = true;

            Param_Integer paramA = (Param_Integer)pManager[4];
            foreach (Font.HighlightColors value in Enum.GetValues(typeof(Font.HighlightColors)))
            {
                paramA.AddNamedValue(value.ToString(), (int)value);
            }
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Paragraph.Name, Constants.Paragraph.NickName, Constants.Paragraph.Input, GH_ParamAccess.item);
            pManager.AddTextParameter("Family Name", "F", "Get Font family name", GH_ParamAccess.item);
            pManager.AddNumberParameter("Size", "S", "Get Text size", GH_ParamAccess.item);
            pManager.AddColourParameter("Color", "C", "Get Text color", GH_ParamAccess.item);
            pManager.AddIntegerParameter("Highlight", "H", "Get Text Highlight Index", GH_ParamAccess.item);
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

            string family = "Arial";
            if (DA.GetData(1, ref family))
            {
                if (isParagraph)
                {
                    foreach (WdFragment f in paragraph.Fragments) f.Font.Family = family;
                }
                else
                {
                    fragment.Font.Family = family;
                }
            }

            double size = 8.0;
            if (DA.GetData(2, ref size))
            {
                if (isParagraph)
                {
                    foreach (WdFragment f in paragraph.Fragments) f.Font.Size = size;
                }
                else
                {
                    fragment.Font.Size = size;
                }
            }

            Sd.Color color = Sd.Color.Black;
            if (DA.GetData(3, ref color))
            {
                if (isParagraph)
                {
                    foreach (WdFragment f in paragraph.Fragments) f.Font.Color = color;
                }
                else
                {
                    fragment.Font.Color = color;
                }
            }

            int highlight = 0;
            if (DA.GetData(4, ref highlight))
            {
                if (isParagraph)
                {
                    foreach (WdFragment f in paragraph.Fragments) f.Font.Highlight = (Font.HighlightColors)highlight;
                }
                else
                {
                    fragment.Font.Highlight = (Font.HighlightColors)highlight;
                }
                    }

            if (isParagraph)
            {
                DA.SetData(0, paragraph);
                DA.SetData(1, paragraph.Fragments[0].Font.Family);
                DA.SetData(2, paragraph.Fragments[0].Font.Size);
                DA.SetData(3, paragraph.Fragments[0].Font.Color);
                DA.SetData(4, paragraph.Fragments[0].Font.Highlight);
            }
            else
            {
                DA.SetData(0, fragment);
                DA.SetData(1, fragment.Font.Family);
                DA.SetData(2, fragment.Font.Size);
                DA.SetData(3, fragment.Font.Color);
                DA.SetData(4, fragment.Font.Highlight);
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
                return Properties.Resources.WD_Doc_Modify_Font;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("B15CE411-0084-4F9B-B4FC-EC4AE8048DEE"); }
        }
    }
}