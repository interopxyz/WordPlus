using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace WordPlus.Components.Content
{
    public class GH_WD_Txt_Mod_FontStyle : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_WD_Con_Mod_CharStyle class.
        /// </summary>
        public GH_WD_Txt_Mod_FontStyle()
          : base("Content Char Style", "ConCharStyle",
              "Modify Content Character Styles if applicable",
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
            pManager.AddBooleanParameter("Bold", "B", "Optionally set the Bold status of the text", GH_ParamAccess.item);
            pManager[1].Optional = true;
            pManager.AddBooleanParameter("Italic", "I", "Optionally set the Italic status of the text", GH_ParamAccess.item);
            pManager[2].Optional = true;
            pManager.AddBooleanParameter("Underlined", "U", "Optionally set the Underlined status of the text", GH_ParamAccess.item);
            pManager[3].Optional = true;
            pManager.AddBooleanParameter("Strikethrough", "X", "Optionally set the Strikethrough status of the text", GH_ParamAccess.item);
            pManager[4].Optional = true;
            pManager.AddIntegerParameter("Sub/Super Script", "S", "Optionally set the Subscript, Superscript, or Default status of the text", GH_ParamAccess.item);
            pManager[5].Optional = true;

            Param_Integer paramA = (Param_Integer)pManager[5];
            paramA.AddNamedValue("0 | Default", 0);
            paramA.AddNamedValue("1 | Superscript", 1);
            paramA.AddNamedValue("2 | Subscript", 2);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Paragraph.Name, Constants.Paragraph.NickName, Constants.Paragraph.Output, GH_ParamAccess.item);
            pManager.AddBooleanParameter("Bold", "B", "Get the Bold status of the text", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Italic", "I", "Get the Italic status of the text", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Underlined", "U", "Get the Underlined status of the text", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Strikethrough", "X", "Get the Strikethrough status of the text", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Subscript", "S", "Get the Subscript status of the text", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Superscript", "^", "Get the Superscript status of the text", GH_ParamAccess.item);
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

            bool bold = false;
            if (DA.GetData(1, ref bold))
            {
                if (isParagraph)
                {
                    foreach (WdFragment f in paragraph.Fragments) f.Font.Bold = bold;
                }
                else
                {
                    fragment.Font.Bold = bold;
                }
            }

            bool italic = false;
            if (DA.GetData(2, ref italic))

            {
                if (isParagraph)
                {
                    foreach (WdFragment f in paragraph.Fragments) f.Font.Italic = italic;
                }
                else
                {
                    fragment.Font.Italic = italic;
                }
            }

            bool underlined = false;
            if (DA.GetData(3, ref underlined))
            {
                if (isParagraph)
                {
                    foreach (WdFragment f in paragraph.Fragments) f.Font.Underlined = underlined;
                }
                else
                {
                    fragment.Font.Underlined = underlined;
                }
            }

            bool strikethrough = false;
            if (DA.GetData(4, ref strikethrough))
            {
                if (isParagraph)
                {
                    foreach (WdFragment f in paragraph.Fragments) f.Font.Strikethrough = strikethrough;
                }
                else
                {
                    fragment.Font.Strikethrough = strikethrough;
                }
            }

            int style = 0;
            if (DA.GetData(5, ref style))

                if (isParagraph)
                {
                    if (style == 1) foreach (WdFragment f in paragraph.Fragments) f.Font.Superscript = true;
                    if (style == 2) foreach (WdFragment f in paragraph.Fragments) f.Font.Subscript = true;
                }
                else
                {
                    if (style == 1) fragment.Font.Superscript = true;
                    if (style == 2) fragment.Font.Subscript = true;
                }

            if (isParagraph)
            {
                DA.SetData(0, paragraph);
                DA.SetData(1, paragraph.Fragments[0].Font.Bold);
                DA.SetData(2, paragraph.Fragments[0].Font.Italic);
                DA.SetData(3, paragraph.Fragments[0].Font.Underlined);
                DA.SetData(4, paragraph.Fragments[0].Font.Strikethrough);
                DA.SetData(5, paragraph.Fragments[0].Font.Subscript);
                DA.SetData(6, paragraph.Fragments[0].Font.Superscript);
            }
            else
            {
                DA.SetData(0, fragment);
                DA.SetData(1, fragment.Font.Bold);
                DA.SetData(2, fragment.Font.Italic);
                DA.SetData(3, fragment.Font.Underlined);
                DA.SetData(4, fragment.Font.Strikethrough);
                DA.SetData(5, fragment.Font.Subscript);
                DA.SetData(6, fragment.Font.Superscript);
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
                return Properties.Resources.WD_Doc_Modify_FontStyle;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("6999EFDA-8459-42ED-A0CC-B159652E06DF"); }
        }
    }
}