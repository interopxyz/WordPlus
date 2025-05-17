using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace WordPlus.Components.Document
{
    public class GH_WD_Doc_SetPassword : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_WD_Doc_SetPassword class.
        /// </summary>
        public GH_WD_Doc_SetPassword()
          : base("Document Password", "WD Pass",
              "Sets a password on a document. (Note: Make sure to save this password as this action is irreversable on the output doucment.",
              Constants.ShortName, Constants.SubDocument)
        {
        }

        public override GH_Exposure Exposure => GH_Exposure.quarternary;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Document.Name, Constants.Document.NickName, Constants.Document.Input, GH_ParamAccess.item);
            pManager[0].Optional = false;
            pManager.AddTextParameter("Password", "P", "The document password. (Note: This document will only open with this password.", GH_ParamAccess.item);
            pManager[1].Optional = true;
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
            IGH_Goo gooA = null;
            if (!DA.GetData(0, ref gooA)) return;

            if (!gooA.CastTo<WdDocument>(out WdDocument document))
            {
                this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Doc input must be a Word Document Object");
                return;
            }
            document = new WdDocument(document);

            string password = string.Empty;
            if (DA.GetData(1, ref password))
            {
                if(password.Contains(" "))
                {
                    this.AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Password cannot contain spaces");
                    return;
                }
                else
                {
                    document.Password = password;
                }
            }

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
                return Properties.Resources.WD_Doc_Password;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("C51DA233-D1D2-4246-B47E-3FA2049D186A"); }
        }
    }
}