using System;
using System.Collections.Generic;
using System.IO;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace WordPlus.Components
{
    public class GH_WD_Doc_Save : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the MyComponent1 class.
        /// </summary>
        public GH_WD_Doc_Save()
          : base("Save Word Document", "WD Save Doc",
              "Save a Word Document to a local file",
              Constants.ShortName, Constants.SubDocument)
        {
        }

        public override GH_Exposure Exposure => GH_Exposure.septenary;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Document.Name, Constants.Document.NickName, Constants.Document.Output, GH_ParamAccess.item);
            pManager[0].Optional = false;

            pManager.AddTextParameter("Directory", "D", "The directory or folder where the file will be saved", GH_ParamAccess.item);
            pManager[1].Optional = true;
            pManager.AddTextParameter("File Name", "N", "The Document name", GH_ParamAccess.item);
            pManager[2].Optional = true;
            pManager.AddBooleanParameter(Constants.Activate.Name, Constants.Activate.NickName, Constants.Activate.Input, GH_ParamAccess.item);
            pManager[3].Optional = false;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("Filepath", "P", "The full filepath to the document", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            bool activate = false;
            if (DA.GetData(3, ref activate))
            {
                if (activate)
                {
                    IGH_Goo gooA = null;
                    if (!DA.GetData(0, ref gooA)) return;

                    if (!gooA.CastTo<WdDocument>(out WdDocument document))
                    {
                        this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Doc input must be a Word Document Object");
                        return;
                    }
                    document = new WdDocument(document);

                    string directory = "C:\\Users\\Public\\Documents\\";
                    if (DA.GetData(1, ref directory))
                    {
                        if (!Directory.Exists(directory))
                        {
                            this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "The specified directory does not exist");
                            return;
                        }
                    }
                    else
                    {
                        if (this.OnPingDocument().FilePath != null)
                        {
                            directory = Path.GetDirectoryName(this.OnPingDocument().FilePath) + "\\";
                        }
                    }
                    directory= Path.GetDirectoryName(directory);

                    string name = Constants.UniqueName;
                    DA.GetData(2, ref name);
                    name = Path.GetFileNameWithoutExtension(name);

                    string path = Path.Combine(directory, name + ".docx");

                    document.Save(path);
                    DA.SetData(0, path);
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
                return Properties.Resources.WD_Doc_File_Save;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("00D70854-3D4C-44DC-AEF3-5F71E3A27A88"); }
        }
    }
}