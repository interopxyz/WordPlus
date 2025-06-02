using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace WordPlus.Components.Text
{
    public class GH_WD_Txt_Mod_Indent : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_WD_Txt_Mod_Indent class.
        /// </summary>
        public GH_WD_Txt_Mod_Indent()
          : base("Word Content Indentation", "WD Indent",
              "Modify Word Content Indentation if applicable",
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
            pManager.AddIntegerParameter("Units", "U", "The Units for the document Width and Height" + Environment.NewLine + "(Note: Pixels are at 96dpi)", GH_ParamAccess.item, 1);
            pManager[1].Optional = true;
            pManager.AddNumberParameter("Indent Before", "B", "Indentation Before", GH_ParamAccess.item);
            pManager[2].Optional = true;
            pManager.AddNumberParameter("Indent After", "A", "Indentation After", GH_ParamAccess.item);
            pManager[3].Optional = true;
            pManager.AddNumberParameter("Indent First Line", "F", "Indentation First Line", GH_ParamAccess.item);
            pManager[4].Optional = true;
            pManager.AddNumberParameter("Indent Hanging", "H", "Indentation Hanging", GH_ParamAccess.item);
            pManager[5].Optional = true;

            Param_Integer paramA = (Param_Integer)pManager[1];
            foreach (WdDocument.Units value in Enum.GetValues(typeof(WdDocument.Units)))
            {
                paramA.AddNamedValue((int)value + " | " + value.ToString(), (int)value);
            }
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Paragraph.Name, Constants.Paragraph.NickName, Constants.Paragraph.Input, GH_ParamAccess.item);
            pManager.AddNumberParameter("Indent Before", "B", "Indentation Before", GH_ParamAccess.item);
            pManager.AddNumberParameter("Indent After", "A", "Indentation After", GH_ParamAccess.item);
            pManager.AddNumberParameter("Indent First Line", "F", "Indentation First Line", GH_ParamAccess.item);
            pManager.AddNumberParameter("Indent Hanging", "H", "Indentation Hanging", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo gooA = null;
            if (!DA.GetData(0, ref gooA)) return;

            if (!gooA.TryGetParagraph(out Paragraph paragraph))
            {
                this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Txt input must be a Text Paragraph Object, Text Fragment Object or a string");
                return;
            }
            int type = 1;
            DA.GetData(1, ref type);

            double before = -1;
            double after = -1;
            double first = -1;
            double hanging = -1;

            switch ((WdDocument.Units)type)
            {
                default:
                    if (DA.GetData(2, ref before)) paragraph.IndentBefore = (int)before.InchToTwip();
                    if (DA.GetData(3, ref after)) paragraph.IndentAfter = (int)after.InchToTwip();
                    if (DA.GetData(4, ref first)) paragraph.IndentFirstLine = (int)first.InchToTwip();
                    if (DA.GetData(5, ref hanging)) paragraph.IndentHanging = (int)hanging.InchToTwip();
                    DA.SetData(1, ((double)paragraph.IndentBefore).TwipToInch());
                    DA.SetData(2, ((double)paragraph.IndentAfter).TwipToInch());
                    DA.SetData(3, ((double)paragraph.IndentFirstLine).TwipToInch());
                    DA.SetData(4, ((double)paragraph.IndentHanging).TwipToInch());
                    break;
                case WdDocument.Units.Pixels:
                    if (DA.GetData(2, ref before)) paragraph.IndentBefore = (int)before.PixelsToTwip();
                    if (DA.GetData(3, ref after)) paragraph.IndentAfter = (int)after.PixelsToTwip();
                    if (DA.GetData(4, ref first)) paragraph.IndentFirstLine = (int)first.PixelsToTwip();
                    if (DA.GetData(5, ref hanging)) paragraph.IndentHanging = (int)hanging.PixelsToTwip();
                    DA.SetData(1, ((double)paragraph.IndentBefore).TwipToPixels());
                    DA.SetData(2, ((double)paragraph.IndentAfter).TwipToPixels());
                    DA.SetData(3, ((double)paragraph.IndentFirstLine).TwipToPixels());
                    DA.SetData(4, ((double)paragraph.IndentHanging).TwipToPixels());
                    break;
                case WdDocument.Units.Centimeters:
                    if (DA.GetData(2, ref before)) paragraph.IndentBefore = (int)before.CentimeterToTwip();
                    if (DA.GetData(3, ref after)) paragraph.IndentAfter = (int)after.CentimeterToTwip();
                    if (DA.GetData(4, ref first)) paragraph.IndentFirstLine = (int)first.CentimeterToTwip();
                    if (DA.GetData(5, ref hanging)) paragraph.IndentHanging = (int)hanging.CentimeterToTwip();
                    DA.SetData(1, ((double)paragraph.IndentBefore).TwipToCentimeter());
                    DA.SetData(2, ((double)paragraph.IndentAfter).TwipToCentimeter());
                    DA.SetData(3, ((double)paragraph.IndentFirstLine).TwipToCentimeter());
                    DA.SetData(4, ((double)paragraph.IndentHanging).TwipToCentimeter());
                    break;
                case WdDocument.Units.Millimeters:
                    if (DA.GetData(2, ref before)) paragraph.IndentBefore = (int)before.MillimeterToTwip();
                    if (DA.GetData(3, ref after)) paragraph.IndentAfter = (int)after.MillimeterToTwip();
                    if (DA.GetData(4, ref first)) paragraph.IndentFirstLine = (int)first.MillimeterToTwip();
                    if (DA.GetData(5, ref hanging)) paragraph.IndentHanging = (int)hanging.MillimeterToTwip();
                    DA.SetData(1, ((double)paragraph.IndentBefore).TwipToMillimeter());
                    DA.SetData(2, ((double)paragraph.IndentAfter).TwipToMillimeter());
                    DA.SetData(3, ((double)paragraph.IndentFirstLine).TwipToMillimeter());
                    DA.SetData(4, ((double)paragraph.IndentHanging).TwipToMillimeter());
                    break;
            }

            DA.SetData(0, paragraph);
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
                return Properties.Resources.WD_Doc_Modify_Indentation;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("7C719479-1F84-43A2-B684-22395C1462DD"); }
        }
    }
}