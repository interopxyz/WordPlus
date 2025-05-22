using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace WordPlus.Components
{
    public class GH_WD_Txt_Mod_FontPresets : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_WD_Con_Mod_FontPresets class.
        /// </summary>
        public GH_WD_Txt_Mod_FontPresets()
          : base("Word Content Font Presets", "WD ConPreset",
              "Apply a preset Word Font Style",
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
            pManager.AddIntegerParameter("Formatting", "F", "Optional predefined text formatting", GH_ParamAccess.item);
            pManager[1].Optional = true;

            Param_Integer paramA = (Param_Integer)pManager[1];
            foreach (Font.Presets value in Enum.GetValues(typeof(Font.Presets)))
            {
                paramA.AddNamedValue(value.ToString(), (int)value);
            }
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Fragment.Name, Constants.Fragment.NickName, Constants.Fragment.Input, GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo gooA = null;
            if (!DA.GetData(0, ref gooA)) return;

            Paragraph paragraph = null;
            bool isParagraph = false;
            Fragment fragment = null;

            if (gooA.TryCastToParagraph(out paragraph))
            {
                isParagraph = true;
            }
            else if (!gooA.TryGetFragment(out fragment))
            {
                this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Txt input must be a Text Paragraph Object, Text Fragment Object or a string");
                return;
            }

            int type = 0;
            if(DA.GetData(1, ref type))
            {
                if (isParagraph)
                {
                    foreach (Fragment f in paragraph.Fragments) f.Font = Fonts.GetPreset((Font.Presets)type);
                }
                else
                {
                    if (type > 0) fragment.Font = Fonts.GetPreset((Font.Presets)type);
                }
            }

            if (isParagraph)
            {
                DA.SetData(0, paragraph);
            }
            else
            {
                DA.SetData(0, fragment);
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
                return Properties.Resources.WD_Doc_Modify_FontPresets;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("1682F93F-35BE-429F-A91E-918C37134AF9"); }
        }
    }
}