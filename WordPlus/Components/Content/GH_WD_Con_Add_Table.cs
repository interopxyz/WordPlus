using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace WordPlus.Components
{
    public class GH_WD_Con_Table : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_WB_Con_Table class.
        /// </summary>
        public GH_WD_Con_Table()
          : base("Table Content", "Tbl",
              "Construct a Table Content Object",
              Constants.ShortName, Constants.SubContent)
        {
        }

        public override GH_Exposure Exposure => GH_Exposure.tertiary;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Content.Name, Constants.Content.NickName, Constants.Content.Input, GH_ParamAccess.tree);
            pManager[0].Optional = false;
            pManager.AddGenericParameter("Title", "T", "Optional Title for the Chart", GH_ParamAccess.item);
            pManager[1].Optional = true;
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
            List<List<WdContent>> dataSet = new List<List<WdContent>>();
            if (!DA.GetDataTree(0, out GH_Structure<IGH_Goo> gooSet)) return;

            foreach (List<IGH_Goo> goos in gooSet.Branches)
            {
                List<WdContent> contents = new List<WdContent>();
                foreach (IGH_Goo goo in goos)
                {
                    if (goo.TryGetContent(out WdContent content))
                    {
                        content.Graphic.HasBorders = true;
                        contents.Add(content);
                    }
                }
                dataSet.Add(contents);
            }

            WdContent table = WdContent.CreateTableContent(dataSet);

            IGH_Goo gooA = null;
            if (DA.GetData(1, ref gooA)) if(gooA.TryGetParagraph(out WdParagraph title)) table.Text = title;

            DA.SetData(0, table);
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
                return Properties.Resources.WD_Con_Table;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("8233F492-559F-4C37-9314-3636866570A4"); }
        }
    }
}