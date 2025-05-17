using System;
using System.Collections.Generic;
using System.Linq;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace WordPlus.Components.Document
{
    public class GH_WD_Doc_SetMetaData : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_WD_Doc_SetMetaData class.
        /// </summary>
        public GH_WD_Doc_SetMetaData()
          : base("Document Meta Data", "WD Meta",
              "Sets the document meta data",
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
            pManager.AddTextParameter("Author", "A", "The creator of the document", GH_ParamAccess.item);
            pManager[1].Optional = true;
            pManager.AddTextParameter("Modified By", "M", "The last person to modify the document", GH_ParamAccess.item);
            pManager[2].Optional = true;
            pManager.AddTextParameter("Company", "C", "The company that produced the document", GH_ParamAccess.item);
            pManager[3].Optional = true;
            pManager.AddTextParameter("Description", "D", "The document description", GH_ParamAccess.item);
            pManager[4].Optional = true;
            pManager.AddTextParameter("Title", "T", "The internal title of the document", GH_ParamAccess.item);
            pManager[5].Optional = true;
            pManager.AddTextParameter("Revision", "R", "The revision number of the document", GH_ParamAccess.item);
            pManager[6].Optional = true;
            pManager.AddTextParameter("Subject", "S", "The subject of the document", GH_ParamAccess.item);
            pManager[7].Optional = true;
            pManager.AddTextParameter("Category", "G", "The category of the document", GH_ParamAccess.item);
            pManager[8].Optional = true;
            pManager.AddTextParameter("Tags", "T", "A list of tags for the document", GH_ParamAccess.list);
            pManager[9].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Document.Name, Constants.Document.NickName, Constants.Document.Output, GH_ParamAccess.item);
            pManager.AddTextParameter("Author", "A", "The creator of the document", GH_ParamAccess.item);
            pManager.AddTextParameter("Modified By", "M", "The last person to modify the document", GH_ParamAccess.item);
            pManager.AddTextParameter("Company", "C", "The company that produced the document", GH_ParamAccess.item);
            pManager.AddTextParameter("Description", "D", "The document description", GH_ParamAccess.item);
            pManager.AddTextParameter("Title", "T", "The internal title of the document", GH_ParamAccess.item);
            pManager.AddTextParameter("Revision", "R", "The revision number of the document", GH_ParamAccess.item);
            pManager.AddTextParameter("Subject", "S", "The subject of the document", GH_ParamAccess.item);
            pManager.AddTextParameter("Category", "G", "The category of the document", GH_ParamAccess.item);
            pManager.AddTextParameter("Tags", "T", "A list of tags document tags", GH_ParamAccess.list);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string errormessage = " cannot exceed 255 characters.";

            IGH_Goo gooA = null;
            if (!DA.GetData(0, ref gooA)) return;

            if (!gooA.CastTo<WdDocument>(out WdDocument document))
            {
                this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Doc input must be a Word Document Object");
                return;
            }
            document = new WdDocument(document);

            string author = string.Empty;
            if (DA.GetData(1, ref author))
            {
                if (author.Length > 255)
                {
                    this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Author"+errormessage);
                    return;
                }
                document.Creator = author;
            }

            string modified = string.Empty;
            if (DA.GetData(2, ref modified))
            {
                if (modified.Length > 255)
                {
                    this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Modified By" + errormessage);
                    return;
                }
                document.Modified = modified;
            }

            string company = string.Empty;
            if (DA.GetData(3, ref company))
            {
                if (company.Length > 255)
                {
                    this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Company" + errormessage);
                    return;
                }
                document.Company = company;
            }

            string description = string.Empty;
            if (DA.GetData(4, ref description))
            {
                if (description.Length > 255)
                {
                    this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Description" + errormessage);
                    return;
                }
                document.Description = description;
            }

            string title = string.Empty;
            if (DA.GetData(5, ref title))
            {
                if (title.Length > 255)
                {
                    this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Title" + errormessage);
                    return;
                }
                document.Title = title;
            }

            string revision = string.Empty;
            if (DA.GetData(6, ref revision))
            {
                if (revision.Length > 255)
                {
                    this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Revision" + errormessage);
                    return;
                }
                document.Revision = revision;
            }

            string subject = string.Empty;
            if (DA.GetData(7, ref subject))
            {
                if (subject.Length > 255)
                {
                    this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Subject" + errormessage);
                    return;
                }
                document.Subject = subject;
            }

            string category = string.Empty;
            if (DA.GetData(8, ref category))
            {
                if (category.Length > 255)
                {
                    this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Category" + errormessage);
                    return;
                }
                document.Categories = category;
            }

            List<string> tags = new List<string>();
            if (DA.GetDataList(9, tags))
            {
                string tag = String.Join(";", tags);
                if (tag.Length > 255)
                {
                    this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Tags" + errormessage);
                    return;
                }
                document.Tags = tag;
            }

            DA.SetData(0, document);
            DA.SetData(1, document.Creator);
            DA.SetData(2, document.Modified);
            DA.SetData(3, document.Company);
            DA.SetData(4, document.Description);
            DA.SetData(5, document.Title);
            DA.SetData(6, document.Revision);
            DA.SetData(7, document.Subject);
            DA.SetData(8, document.Categories);
            DA.SetDataList(9, document.Tags.Split(';').ToList());
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
                return Properties.Resources.WD_Doc_Meta;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("9B2DB837-450E-4A36-8140-00F975250742"); }
        }
    }
}