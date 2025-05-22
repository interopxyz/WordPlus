using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

using WD = OfficeIMO.Word;

namespace WordPlus.Components
{
    public class GH_WD_Con_List : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_WB_Con_List class.
        /// </summary>
        public GH_WD_Con_List()
          : base("Word List Contents", "WD Lst",
              "Construct a Word Text List Content Object",
              Constants.ShortName, Constants.SubContent)
        {
        }

        public override GH_Exposure Exposure => GH_Exposure.primary;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Paragraph.Name+"s", Constants.Paragraph.NickName, Constants.Paragraph.Input, GH_ParamAccess.list);
            pManager[0].Optional = false;
            pManager.AddIntegerParameter("Level","L","Sets the indentation level", GH_ParamAccess.list);
            pManager[1].Optional = true;
            pManager.AddIntegerParameter("Style", "S", "Bullet Style", GH_ParamAccess.item, 0);
            pManager[2].Optional = true;
            pManager.AddIntegerParameter("Formatting", "F", "Optional predefined text formatting", GH_ParamAccess.item, 0);
            pManager[3].Optional = true;

            Param_Integer paramA = (Param_Integer)pManager[2];
            foreach (Content.BulletPoints value in Enum.GetValues(typeof(Content.BulletPoints)))
            {
                paramA.AddNamedValue((int)value + " | " + value.ToString(), (int)value);
            }

            Param_Integer paramB = (Param_Integer)pManager[3];
            foreach (Font.Presets value in Enum.GetValues(typeof(Font.Presets)))
            {
                paramB.AddNamedValue((int)value + " | " + value.ToString(), (int)value);
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
            List<IGH_Goo> goos = new List<IGH_Goo>();
            if(!DA.GetDataList(0, goos))return;

            int preset = 0;
            DA.GetData(3, ref preset);
            bool hasPreset = preset > 0; 

            List<Paragraph> paragraphs = new List<Paragraph>();
            foreach (IGH_Goo goo in goos)
            {
                if (goo.TryGetParagraph(out Paragraph paragraph))
                {
                    if(hasPreset) foreach (Fragment fragment in paragraph.Fragments) fragment.Font = Fonts.GetPreset((Font.Presets)preset);
                    paragraphs.Add(paragraph);
                }
            }

            List<int> levels = new List<int>();
            DA.GetDataList(1, levels);

            int style = 0;
            DA.GetData(2, ref style);

            Content content = Content.CreateListContent(paragraphs, levels,(Content.BulletPoints)style);

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
                return Properties.Resources.WD_Con_List;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("BF486C80-5DE4-4752-A94C-47D2BF9F80F3"); }
        }
    }
}