using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace WordPlus.Components
{
    public class GH_WD_Con_De_Chart : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_WD_Con_De_Chart class.
        /// </summary>
        public GH_WD_Con_De_Chart()
          : base("Deconstruct Word Chart Contents", "WD De Cht",
              "DeConstruct a Word Chart Content Object",
              Constants.ShortName, Constants.SubContent)
        {
        }

        public override GH_Exposure Exposure => GH_Exposure.septenary;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Chart " + Constants.Content.Name, Constants.Content.NickName, "Chart " + Constants.Content.Input, GH_ParamAccess.item);
            pManager[0].Optional = false;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddNumberParameter("Numbers","N","The charted numerical values", GH_ParamAccess.tree);
            pManager.AddGenericParameter("Title", "T", "The chart title", GH_ParamAccess.item);
            pManager.AddGenericParameter("Axis Labels", "A", "List of labels per list item", GH_ParamAccess.list);
            pManager.AddGenericParameter("Series Labels", "S", "List of legend label values", GH_ParamAccess.list);
            pManager.AddColourParameter("Series Colors", "C", "List of legend Colors.", GH_ParamAccess.list);

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
                            GH_Path path = new GH_Path();

                            if (this.Params.Input[0].VolatileData.PathCount > 1) path = this.Params.Input[0].VolatileData.get_Path(this.RunCount - 1);
                            path = path.AppendElement(this.RunCount - 1);
                    switch(content.ContentType)
                    {
                        default:
                            this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Con input must be a Table Content Object");
                            return;
                        case Content.ContentTypes.ChartBar:
                        case Content.ContentTypes.ChartArea:
                        case Content.ContentTypes.ChartColumn:
                        case Content.ContentTypes.ChartLine:

                            if (this.Params.Input[0].VolatileData.PathCount > 1) path = this.Params.Input[0].VolatileData.get_Path(this.RunCount - 1);
                            path = path.AppendElement(this.RunCount - 1);

                            GH_Structure<GH_Number> data1 = content.Numbers.ToDataTree(path);

                            DA.SetDataTree(0, data1);
                            DA.SetData(1, content.Text);
                            DA.SetDataList(2, content.Values[0]);
                            DA.SetDataList(3, content.Values[1]);
                            DA.SetDataList(4, content.Colors[0]);
                            break;
                        case Content.ContentTypes.ChartPie:

                            GH_Structure<GH_Number> data2 = content.Numbers.ToDataTree(path);

                            DA.SetDataTree(0, data2);
                            DA.SetData(1, content.Text);
                            DA.SetDataList(2, content.Values[0]);
                            break;
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
                return Properties.Resources.WD_Con_Chart_De;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("22271c41-c3d4-4342-b7f9-be263ab75846"); }
        }
    }
}