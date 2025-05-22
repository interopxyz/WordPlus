using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace WordPlus.Components
{
    public class GH_WD_Doc_SetMargins : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_WD_Doc_SetMargins class.
        /// </summary>
        public GH_WD_Doc_SetMargins()
          : base("Word Document Margins", "WD Margins",
              "Sets the number of text columns in a Word document",
              Constants.ShortName, Constants.SubDocument)
        {
        }

        public override GH_Exposure Exposure => GH_Exposure.tertiary;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Document.Name, Constants.Document.NickName, Constants.Document.Input, GH_ParamAccess.item);
            pManager[0].Optional = false;
            pManager.AddIntegerParameter("Units", "U", "The Units for the document Margins", GH_ParamAccess.item,1);
            pManager[1].Optional = true;
            pManager.AddNumberParameter("Top Margin", "T", "The top margin of the page", GH_ParamAccess.item);
            pManager[2].Optional = true;
            pManager.AddNumberParameter("Bottom Margin", "B", "The bottom margin of the page", GH_ParamAccess.item);
            pManager[3].Optional = true;
            pManager.AddNumberParameter("Left Margin", "L", "The left margin of the page", GH_ParamAccess.item);
            pManager[4].Optional = true;
            pManager.AddNumberParameter("Right Margin", "R", "The right margin of the page", GH_ParamAccess.item);
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
            pManager.AddGenericParameter(Constants.Document.Name, Constants.Document.NickName, Constants.Document.Output, GH_ParamAccess.item);
            pManager.AddNumberParameter("Top Margin", "T", "The top margin of the page", GH_ParamAccess.item);
            pManager.AddNumberParameter("Bottom Margin", "B", "The bottom margin of the page", GH_ParamAccess.item);
            pManager.AddNumberParameter("Left Margin", "L", "The left margin of the page", GH_ParamAccess.item);
            pManager.AddNumberParameter("Right Margin", "R", "The right margin of the page", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo gooA = null;
            if (!DA.GetData(0, ref gooA)) return;

            if (!gooA.CastTo<WdDocument>(out WdDocument document))
            {
                this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Doc input must be a Word Document Object");
                return;
            }
            document = new WdDocument(document);

            int units = 0;
            DA.GetData(1, ref units);

            double top = 0;
            double bottom = 0;
            double left = 0;
            double right = 0;

            switch((WdDocument.Units)units)
            {
                default:
                    if (DA.GetData(2, ref top)) document.Page.MarginTop = (int)Extension.InchToTwip( top);
                    if (DA.GetData(3, ref bottom)) document.Page.MarginBottom = (int)Extension.InchToTwip(bottom);
                    if (DA.GetData(4, ref left)) document.Page.MarginLeft = (int)Extension.InchToTwip(left);
                    if (DA.GetData(5, ref right)) document.Page.MarginRight = (int)Extension.InchToTwip(right);
                    break;
                case WdDocument.Units.Millimeters:
                    if (DA.GetData(2, ref top)) document.Page.MarginTop = (int)Extension.MillimeterToTwip(top);
                    if (DA.GetData(3, ref bottom)) document.Page.MarginBottom = (int)Extension.MillimeterToTwip(bottom);
                    if (DA.GetData(4, ref left)) document.Page.MarginLeft = (int)Extension.MillimeterToTwip(left);
                    if (DA.GetData(5, ref right)) document.Page.MarginRight = (int)Extension.MillimeterToTwip(right);
                    break;
                case WdDocument.Units.Centimeters:
                    break;
                    if (DA.GetData(2, ref top)) document.Page.MarginTop = (int)Extension.CentimeterToTwip(top);
                    if (DA.GetData(3, ref bottom)) document.Page.MarginBottom = (int)Extension.CentimeterToTwip(bottom);
                    if (DA.GetData(4, ref left)) document.Page.MarginLeft = (int)Extension.CentimeterToTwip(left);
                    if (DA.GetData(5, ref right)) document.Page.MarginRight = (int)Extension.CentimeterToTwip(right);
                case WdDocument.Units.Pixels:
                    if (DA.GetData(2, ref top)) document.Page.MarginTop = (int)Extension.PixelsToTwip(top);
                    if (DA.GetData(3, ref bottom)) document.Page.MarginBottom = (int)Extension.PixelsToTwip(bottom);
                    if (DA.GetData(4, ref left)) document.Page.MarginLeft = (int)Extension.PixelsToTwip(left);
                    if (DA.GetData(5, ref right)) document.Page.MarginRight = (int)Extension.PixelsToTwip(right);
                    break;
            }

            DA.SetData(0, document);
            switch ((WdDocument.Units)units)
            {
                default:
                    DA.SetData(1, Extension.TwipToInch(document.Page.MarginTop));
                    DA.SetData(2, Extension.TwipToInch(document.Page.MarginBottom));
                    DA.SetData(3, Extension.TwipToInch(document.Page.MarginLeft));
                    DA.SetData(4, Extension.TwipToInch(document.Page.MarginRight));
                    break;
                case WdDocument.Units.Millimeters:
                    DA.SetData(1, Extension.TwipToMillimeter(document.Page.MarginTop));
                    DA.SetData(2, Extension.TwipToMillimeter(document.Page.MarginBottom));
                    DA.SetData(3, Extension.TwipToMillimeter(document.Page.MarginLeft));
                    DA.SetData(4, Extension.TwipToMillimeter(document.Page.MarginRight));
                    break;
                case WdDocument.Units.Centimeters:
                    DA.SetData(1, Extension.TwipToCentimeter(document.Page.MarginTop));
                    DA.SetData(2, Extension.TwipToCentimeter(document.Page.MarginBottom));
                    DA.SetData(3, Extension.TwipToCentimeter(document.Page.MarginLeft));
                    DA.SetData(4, Extension.TwipToCentimeter(document.Page.MarginRight));
                    break;
                case WdDocument.Units.Pixels:
                    DA.SetData(1, Extension.TwipToPixels(document.Page.MarginTop));
                    DA.SetData(2, Extension.TwipToPixels(document.Page.MarginBottom));
                    DA.SetData(3, Extension.TwipToPixels(document.Page.MarginLeft));
                    DA.SetData(4, Extension.TwipToPixels(document.Page.MarginRight));
                    break;
            }

            DA.SetData(1, document.Page.Columns);
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
                return Properties.Resources.WD_Doc_Margins;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("5474309B-6276-445A-964E-1EDDDE6D82A6"); }
        }
    }
}