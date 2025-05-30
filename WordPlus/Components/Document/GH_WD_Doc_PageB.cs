﻿using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Rhino.Geometry;

namespace WordPlus.Components
{
    public class GH_WD_Doc_PageB : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_WD_Doc_PageB class.
        /// </summary>
        public GH_WD_Doc_PageB()
          : base("Iso B Word Pages", "WD IsoB",
              "Create a Word Document from a standard Iso B Paper Size",
              Constants.ShortName, Constants.SubDocument)
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddIntegerParameter("Size", "S", "A standard Iso B paper size", GH_ParamAccess.item, 0);
            pManager[0].Optional = true;
            pManager.AddIntegerParameter("Orientation", "O", "Set the portrait or landscape orientation of the page", GH_ParamAccess.item, 0);
            pManager[1].Optional = true;
            pManager.AddIntegerParameter("Margin", "M", "Select a preset margin setting", GH_ParamAccess.item);
            pManager[2].Optional = true;

            Param_Integer paramA = (Param_Integer)pManager[0];
            foreach (Page.SizesIsoB value in Enum.GetValues(typeof(Page.SizesIsoB)))
            {
                paramA.AddNamedValue((int)value + " | " + value.ToString(), (int)value);
            }

            Param_Integer paramB = (Param_Integer)pManager[1];
            paramB.AddNamedValue("Portrait", 0);
            paramB.AddNamedValue("Landscape", 1);

            Param_Integer paramC = (Param_Integer)pManager[2];
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
            int type = 0;
            DA.GetData(0, ref type);

            WdDocument document = new WdDocument();
            document.Page = Page.IsoB((Page.SizesIsoB)type);

            int orientation = 0;
            DA.GetData(1, ref orientation);
            if (orientation != 0) document.Page.Orientation = Page.Orientations.Landscape;

            int margin = 0;
            if(DA.GetData(2, ref margin)) document.Page.SetMargin((Page.Margins)margin);

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
                return Properties.Resources.WD_Doc_Iso_B;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("35221F66-89D7-4DE8-999C-49F00577D67D"); }
        }
    }
}