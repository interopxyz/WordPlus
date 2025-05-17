using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace WordPlus.Components.Content
{
    public class GH_WD_Con_Mod_Size : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_WD_Con_Mod_Size class.
        /// </summary>
        public GH_WD_Con_Mod_Size()
          : base("Content Size", "ConSize",
              "Modify Content Size if applicable",
              Constants.ShortName, Constants.Modify)
        {
        }

        public override GH_Exposure Exposure => GH_Exposure.secondary;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Content.Name, Constants.Content.NickName, Constants.Content.Input, GH_ParamAccess.item);
            pManager[0].Optional = false;
            pManager.AddIntegerParameter("Units", "U", "The Units for the document Width and Height" + Environment.NewLine + "(Note: Pixels are at 96dpi)", GH_ParamAccess.item, 1);
            pManager[1].Optional = true;
            pManager.AddNumberParameter("Width", "W", "Optionally override Content Width if applicable", GH_ParamAccess.item);
            pManager[2].Optional = true;
            pManager.AddNumberParameter("Height", "H", "Optionally override Content Height if applicable", GH_ParamAccess.item);
            pManager[3].Optional = true;

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
            pManager.AddGenericParameter(Constants.Content.Name, Constants.Content.NickName, Constants.Content.Input, GH_ParamAccess.item);
            pManager.AddNumberParameter("Width", "W", "The Content Width. (0 = autofit)", GH_ParamAccess.item);
            pManager.AddNumberParameter("Height", "H", "The Content Height. (0 = autofit)", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo gooA = null;
            if (!DA.GetData(0, ref gooA)) return;

            if (!gooA.TryGetContent(out WdContent content))
            {
                this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Con input must be a Content Object, Text Fragment Object or a string");
                return;
            }

            switch (content.ContentType)
            {
                default:
                    this.AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, content.ContentType.ToString().SplitAtCapitols() + " content does not support resizing");
                    break;
                case WdContent.ContentTypes.Table:
                    this.AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "Table content does not support changing height");
                    break;
                case WdContent.ContentTypes.ChartArea:
                case WdContent.ContentTypes.ChartBar:
                case WdContent.ContentTypes.ChartColumn:
                case WdContent.ContentTypes.ChartLine:
                case WdContent.ContentTypes.ChartPie:
                case WdContent.ContentTypes.Image:
                    break;
            }
            int units = 0;
            DA.GetData(1, ref units);

            double width = 0;
            bool hasWidth = DA.GetData(2, ref width);

            double height = 0;
            bool hasHeight = DA.GetData(3, ref height) ;

            switch ((WdDocument.Units)units)
            {
                default://Inches
                    if (hasWidth) content.Width = (int)width.InchToTwip();
                    if (hasHeight) content.Height = (int)height.InchToTwip();
                    break;
                case WdDocument.Units.Millimeters:
                    if (hasWidth) content.Width = (int)width.MillimeterToTwip();
                    if (hasHeight) content.Height = (int)height.MillimeterToTwip();
                    break;
                case WdDocument.Units.Centimeters:
                    if (hasWidth) content.Width = (int)width.CentimeterToTwip();
                    if (hasHeight) content.Height = (int)height.CentimeterToTwip();
                    break;
                case WdDocument.Units.Pixels:
                    if (hasWidth) content.Width = (int)width.PixelsToTwip();
                    if (hasHeight) content.Height = (int)height.PixelsToTwip();
                    break;
            }

            switch ((WdDocument.Units)units)
            {
                default://Inches
                    DA.SetData(1, content.Width.TwipToInch());
                    DA.SetData(2, content.Height.TwipToInch());
                    break;
                case WdDocument.Units.Millimeters:
                    DA.SetData(1, content.Width.TwipToMillimeter());
                    DA.SetData(2, content.Height.TwipToMillimeter());
                    break;
                case WdDocument.Units.Centimeters:
                    DA.SetData(1, content.Width.TwipToCentimeter());
                    DA.SetData(2, content.Height.TwipToCentimeter());
                    break;
                case WdDocument.Units.Pixels:
                    DA.SetData(1, content.Width.TwipToPixels());
                    DA.SetData(2, content.Height.TwipToPixels());
                    break;
            }

            DA.SetData(0,content);

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
                return Properties.Resources.WD_Con_Mod_ReSize;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("CDCBB4E3-B28A-4064-9EBD-A66076436B65"); }
        }
    }
}