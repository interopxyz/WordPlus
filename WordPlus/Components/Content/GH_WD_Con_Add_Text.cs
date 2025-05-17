using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace WordPlus.Components
{
    public class GH_WD_Con_Text : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_WB_Con_Text class.
        /// </summary>
        public GH_WD_Con_Text()
          : base("Text Contents", "Txt",
              "Construct a Text Content Object",
              Constants.ShortName, Constants.SubContent)
        {
        }

        public override GH_Exposure Exposure => GH_Exposure.primary;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Paragraph.Name,Constants.Paragraph.NickName,Constants.Paragraph.Input, GH_ParamAccess.item);
            pManager[0].Optional = false;
            pManager.AddIntegerParameter("Formatting", "F", "Optional predefined text formatting", GH_ParamAccess.item, 0);
            pManager[1].Optional = true;

            Param_Integer paramA = (Param_Integer)pManager[1];
            foreach (Font.Presets value in Enum.GetValues(typeof(Font.Presets)))
            {
                paramA.AddNamedValue((int)value + " | " + value.ToString(), (int)value);
            }
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Content.Name, Constants.Content.NickName, Constants.Content.Output, GH_ParamAccess.item);
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
                this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Txt input must be a Text Fragment Object or a string");
                return;
            }

            int preset = 0;
            if (DA.GetData(1, ref preset)) if(preset>0) foreach(WdFragment fragment in paragraph.Fragments) fragment.Font = Fonts.GetPreset((Font.Presets)preset);

            WdContent content = WdContent.CreateTextContent(paragraph);

            DA.SetData(0, content);
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
                return Properties.Resources.WD_Con_Text;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("C8C8CEE6-1327-403A-B033-D6D0129E43BD"); }
        }
    }
}