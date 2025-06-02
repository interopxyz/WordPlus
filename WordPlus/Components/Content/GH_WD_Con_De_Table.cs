using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace WordPlus.Components
{
    public class GH_WD_Con_De_Table : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_WD_Con_De_Table class.
        /// </summary>
        public GH_WD_Con_De_Table()
          : base("Deconstruct Word Table Contents", "WD De Tbl",
              "DeConstruct a Word Table Content Object",
              Constants.ShortName, Constants.SubContent)
        {
        }

        public override GH_Exposure Exposure => GH_Exposure.septenary;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Table " + Constants.Content.Name, Constants.Content.NickName, "Table " + Constants.Content.Input, GH_ParamAccess.item);
            pManager[0].Optional = false;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Paragraph.Name, Constants.Paragraph.NickName, Constants.Paragraph.Input, GH_ParamAccess.tree);
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
                    this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Con input must be a Table Content Object");
                    return;
                }
                else
                {
                    if (content.ContentType == Content.ContentTypes.Table)
                    {

                        GH_Path path = new GH_Path();

                        if (this.Params.Input[0].VolatileData.PathCount > 1) path = this.Params.Input[0].VolatileData.get_Path(this.RunCount - 1);
                        path = path.AppendElement(this.RunCount - 1);

                        GH_Structure<GH_ObjectWrapper> data = content.Contents.ToDataTree(path);

                        DA.SetDataTree(0, data);
                    }
                    else
                    {
                        this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Con input must be a Table Content Object");
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
                return Properties.Resources.WD_Con_Table_De;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("CAF1B862-5ABB-4D20-B836-5451649BAF7B"); }
        }
    }
}