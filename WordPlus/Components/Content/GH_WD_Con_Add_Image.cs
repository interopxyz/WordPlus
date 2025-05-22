using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

using Sd = System.Drawing;

namespace WordPlus.Components
{
    public class GH_WD_Con_Image : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_WB_Con_Image class.
        /// </summary>
        public GH_WD_Con_Image()
          : base("Word Image Content", "WD Img",
              "Construct a Word Image Content Object",
              Constants.ShortName, Constants.SubContent)
        {
        }

        public override GH_Exposure Exposure => GH_Exposure.quarternary;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Image", "I", "A System Drawing Bitmap or full FilePath to an image file", GH_ParamAccess.item);
            pManager[0].Optional = false;
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
            string warning = "I input must be a System Drawing Bitmap or a full file path to an image file";
            IGH_Goo gooA = null;
            if (!DA.GetData(0, ref gooA)) return;

            bool isValid = false;
            if (gooA.CastTo<Sd.Bitmap>(out Sd.Bitmap image)) isValid = true;

            if (!isValid)
            {
                if (gooA.CastTo<string>(out string filepath))
                {
                    if (System.IO.File.Exists(filepath))
                    {
                        image = new Sd.Bitmap(filepath);
                        isValid = true;
                    }
                    else
                    {
                        this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, warning);
                        return;
                    }
                }
                else
                {
                    this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, warning);
                    return;
                }
            }

            if (isValid)
            {
                Content content = Content.CreateImageContent(image);
                DA.SetData(0, content);
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
                return Properties.Resources.WD_Con_Image;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("1C136C5C-5832-4864-B33C-0180CA507C9A"); }
        }
    }
}