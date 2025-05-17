using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

using Sd = System.Drawing;

namespace WordPlus.Components.Graphics
{
    public class GH_WD_Con_Mod_Graphic : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_WD_Con_Mod_Graphic class.
        /// </summary>
        public GH_WD_Con_Mod_Graphic()
          : base("Content Graphic", "ConGraphic",
              "Modify Content Graphics if applicable",
              Constants.ShortName, Constants.Modify)
        {
        }

        public override GH_Exposure Exposure => GH_Exposure.secondary;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Content.Name, Constants.Content.NickName, Constants.Content.Input,GH_ParamAccess.item);
            pManager[0].Optional = false;
            pManager.AddColourParameter("Fill Color", "F", "The background or fill color when applicable.", GH_ParamAccess.item);
            pManager[1].Optional = true;
            pManager.AddColourParameter("Stroke Color", "S", "The line stroke color when applicable.",GH_ParamAccess.item);
            pManager[2].Optional = true;
            pManager.AddNumberParameter("Stroke Weight", "W", "The line weight when applicable.", GH_ParamAccess.item);
            pManager[3].Optional = true;
            pManager.AddIntegerParameter("Line Type", "T", "The line type when applicable", GH_ParamAccess.item);
            pManager[4].Optional = true;

            Param_Integer paramA = (Param_Integer)pManager[4];
            foreach (Graphic.LineTypes value in Enum.GetValues(typeof(Graphic.LineTypes)))
            {
                paramA.AddNamedValue((int)value + " | " + value.ToString(), (int)value);
            }
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Content.Name, Constants.Content.NickName, Constants.Content.Output, GH_ParamAccess.item);
            pManager.AddColourParameter("Fill Color", "F", "The background or fill color.", GH_ParamAccess.item);
            pManager.AddColourParameter("Stroke Color", "S", "The line stroke color.", GH_ParamAccess.item);
            pManager.AddNumberParameter("Stroke Weight", "W", "The line weight.", GH_ParamAccess.item);
            pManager.AddIntegerParameter("Line Type", "T", "The line type when applicable", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo gooA = null;
            if (!DA.GetData(0, ref gooA)) return;

                Sd.Color fill = Sd.Color.Empty;
                Sd.Color stroke = Sd.Color.Empty;
                double weight = 1.0;
                int type = 0;

            if(gooA.TryGetFragment(out WdFragment fragment))
            {
                if (DA.GetData(1, ref fill)) fragment.Graphic.Fill = fill;
                if (DA.GetData(2, ref stroke)) fragment.Graphic.Stroke = stroke;
                if (DA.GetData(3, ref weight)) fragment.Graphic.Weight = weight;
                if (DA.GetData(4, ref type)) fragment.Graphic.LineType = (Graphic.LineTypes)type;

                DA.SetData(0, fragment);
                DA.SetData(1, fragment.Graphic.Fill);
                DA.SetData(2, fragment.Graphic.Stroke);
                DA.SetData(3, fragment.Graphic.Weight);
                DA.SetData(4, fragment.Graphic.LineType);
            }
            else if (gooA.TryGetParagraph(out WdParagraph paragraph))
            {
                if (DA.GetData(1, ref fill)) foreach (WdFragment fragment1 in paragraph.Fragments) fragment1.Graphic.Fill = fill;
                if (DA.GetData(2, ref stroke)) foreach (WdFragment fragment1 in paragraph.Fragments) fragment1.Graphic.Stroke = stroke;
                if (DA.GetData(3, ref weight)) foreach(WdFragment fragment1 in paragraph.Fragments)fragment1.Graphic.Weight = weight;
                if (DA.GetData(4, ref type)) foreach (WdFragment fragment1 in paragraph.Fragments) fragment1.Graphic.LineType = (Graphic.LineTypes)type;

                DA.SetData(0, paragraph);
                DA.SetData(1, paragraph.Fragments[0].Graphic.Fill);
                DA.SetData(2, paragraph.Fragments[0].Graphic.Stroke);
                DA.SetData(3, paragraph.Fragments[0].Graphic.Weight);
                DA.SetData(4, paragraph.Fragments[0].Graphic.LineType);
            }
            else if (gooA.TryGetContent(out WdContent content))
            {
                if (DA.GetData(1, ref fill)) content.Graphic.Fill = fill;
                if (DA.GetData(2, ref stroke)) content.Graphic.Stroke = stroke;
                if (DA.GetData(3, ref weight)) content.Graphic.Weight = weight;
                if (DA.GetData(4, ref type)) content.Graphic.LineType = (Graphic.LineTypes)type;

                DA.SetData(0, content);
                DA.SetData(1, content.Graphic.Fill);
                DA.SetData(2, content.Graphic.Stroke);
                DA.SetData(3, content.Graphic.Weight);
                DA.SetData(4, content.Graphic.LineType);
            }
            else if (gooA.CastTo<WdDocument>(out WdDocument document))
            {
                if (DA.GetData(1, ref fill)) document.Graphic.Fill = fill;
                if (DA.GetData(2, ref stroke)) document.Graphic.Stroke = stroke;
                if (DA.GetData(3, ref weight)) document.Graphic.Weight = weight;
                if (DA.GetData(4, ref type)) document.Graphic.LineType = (Graphic.LineTypes)type;

                DA.SetData(0, document);
                DA.SetData(1, document.Graphic.Fill);
                DA.SetData(2, document.Graphic.Stroke);
                DA.SetData(3, document.Graphic.Weight);
                DA.SetData(4, document.Graphic.LineType);
            }
            else
            {
                this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Con input must be a Content Object, Paragraph Object, Text Fragment Object or a string");
                return;
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
                return Properties.Resources.WD_Con_Graphics2;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("1AD012DE-822E-493A-AC2F-B7951C9503A1"); }
        }
    }
}