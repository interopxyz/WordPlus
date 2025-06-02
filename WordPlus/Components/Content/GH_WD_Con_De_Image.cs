using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace WordPlus.Components
{
    public class GH_WD_Con_De_Image : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_WD_Con_De_Image class.
        /// </summary>
        public GH_WD_Con_De_Image()
          : base("Deconstruct Word Image Contents", "WD De Img",
              "DeConstruct a Word Image Content Object",
              Constants.ShortName, Constants.SubContent)
        {
        }

        public override GH_Exposure Exposure => GH_Exposure.septenary;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Image " + Constants.Content.Name, Constants.Content.NickName, "Image " + Constants.Content.Input, GH_ParamAccess.item);
            pManager[0].Optional = false;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Bitmap Image", "I", "A System Drawing Bitmap", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo gooA = null;
            if (DA.GetData(0, ref gooA))
            {
                if (!gooA.CastTo<Content>(out Content content))
                {
                    this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Con input must be a Image Content Object");
                    return;
                }
                else
                {
                    if (content.ContentType == Content.ContentTypes.Image)
                    {
                        DA.SetData(0, content.Image);
                    }
                    else
                    {
                        this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Con input must be a Image Content Object");
                        return;
                    }
                }
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
                return Properties.Resources.WD_Con_Image_De;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("7691757C-9550-4972-B332-EFB3614E7630"); }
        }
    }
}