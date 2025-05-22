using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Rhino.Geometry;

namespace WordPlus.Components
{
    public class GH_WD_Doc_PageConstruct : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_WB_Doc_Construct class.
        /// </summary>
        public GH_WD_Doc_PageConstruct()
          : base("Construct Word Document", "WD Doc",
              "Construct a Word Document",
              Constants.ShortName, Constants.SubDocument)
        {
        }

        public override GH_Exposure Exposure => GH_Exposure.primary;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddIntegerParameter("Units", "U", "The Units for the document Width and Height" + Environment.NewLine + "(Note: Pixels are at 96dpi)", GH_ParamAccess.item, 1);
            pManager[0].Optional = true;
            pManager.AddNumberParameter("Width", "W", "The Width of the Document Pages in the specified Units.", GH_ParamAccess.item, 8.5);
            pManager[1].Optional = true;
            pManager.AddNumberParameter("Height", "H", "The Height of the Document Pages in the specified Units", GH_ParamAccess.item, 11);
            pManager[2].Optional = true;
            pManager.AddIntegerParameter("Margin", "M", "Select a preset margin setting", GH_ParamAccess.item, 0);
            pManager[3].Optional = true;

            Param_Integer paramA = (Param_Integer)pManager[0];
            foreach (WdDocument.Units value in Enum.GetValues(typeof(WdDocument.Units)))
            {
                paramA.AddNamedValue((int)value + " | " + value.ToString(), (int)value);
            }

            Param_Integer paramC = (Param_Integer)pManager[3];
            foreach (Page.Margins value in Enum.GetValues(typeof(Page.Margins)))
            {
                paramC.AddNamedValue((int)value + " | " + value.ToString(), (int)value);
            }

        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Document.Name, Constants.Document.NickName, Constants.Document.Output, GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            WdDocument document = new WdDocument();

            int units = 0;
            DA.GetData(0, ref units);

            double width = 8.5;
            DA.GetData(1, ref width);

            double height = 11;
            DA.GetData(2, ref height);

            switch ((WdDocument.Units)units)
            {
                default://Inches
                    document.Page.PageWidth = (int)width.InchToTwip();
                    document.Page.PageHeight = (int)height.InchToTwip();
                    break;
                case WdDocument.Units.Millimeters:
                    document.Page.PageWidth = (int)width.MillimeterToTwip();
                    document.Page.PageHeight = (int)height.MillimeterToTwip();
                    break;
                case WdDocument.Units.Centimeters:
                    document.Page.PageWidth = (int)width.CentimeterToTwip();
                    document.Page.PageHeight = (int)height.CentimeterToTwip();
                    break;
                case WdDocument.Units.Pixels:
                    document.Page.PageWidth = (int)width.PixelsToTwip();
                    document.Page.PageHeight = (int)height.PixelsToTwip();
                    break;
            }

            int margin = 0;
            DA.GetData(3, ref margin);
            document.Page.SetMargin((Page.Margins)margin);

            DA.SetData(0, document);
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
                return Properties.Resources.WD_Doc_New;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("96BCE992-BE4D-4B3A-8D33-5C2EF1B8C352"); }
        }
    }
}