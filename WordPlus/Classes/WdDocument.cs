using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using WD = OfficeIMO.Word;
using WP = DocumentFormat.OpenXml.Wordprocessing;

namespace WordPlus
{
    public class WdDocument
    {

        #region members

        public enum Units { Pixels, Inches, Millimeters, Centimeters };
        public enum HeaderFooterTypes { None,All, Even, First };

        public HeaderFooterTypes HeaderType = HeaderFooterTypes.None;
        public HeaderFooterTypes FooterType = HeaderFooterTypes.None;

        public WD.WordDocument DocObject = WD.WordDocument.Create();
        protected WdContentBlock content = new WdContentBlock();
        public WdContentBlock Header = new WdContentBlock();
        public WdContentBlock Footer = new WdContentBlock();

        string name = "Unnamed";
        string title = string.Empty;
        string categories = string.Empty;
        string company = string.Empty;
        string creator = string.Empty;
        string description = string.Empty;
        string tags = string.Empty;
        string revision = string.Empty;
        string subject = string.Empty;
        string modified = string.Empty;
        string password = string.Empty;


        public Page Page = new Page();
        public Graphic graphic = new Graphic();
        
        #endregion

        #region constructors

        public WdDocument()
        {

        }

        public WdDocument(string name)
        {
            this.name = name;
        }

        public WdDocument(WdDocument document)
        {
            this.DocObject = document.DocObject;
            this.content = new WdContentBlock(document.content);
            this.Header = new WdContentBlock(document.Header);
            this.Footer = new WdContentBlock(document.Footer);

            this.HeaderType = document.HeaderType;
            this.FooterType = document.FooterType;

            this.name = document.name;
            this.title = document.title;
            this.categories = document.categories;
            this.company = document.company;
            this.creator = document.creator;
            this.description = document.description;
            this.tags = document.tags;
            this.revision = document.revision;
            this.subject = document.subject;
            this.modified = document.modified;
            this.password = document.password;

            this.Page = new Page(document.Page);
            this.graphic = new Graphic(document.graphic);
        }

        public WdDocument(List<WdDocument> documents, string name, bool pageBreak = false)
        {
            this.DocObject = WD.WordDocument.Create();

            for(int i =0;i<documents.Count;i++)
            {
                this.content.AddRange(documents[i].GetContents());
                if (pageBreak) if (i < (documents.Count - 1)) this.content.Add(Content.CreatePageBreakContent());
            }

            int c = documents.Count - 1;
            this.Header = new WdContentBlock(documents[c].Header);
            this.Footer = new WdContentBlock(documents[c].Footer);

            this.HeaderType = documents[c].HeaderType;
            this.FooterType= documents[c].FooterType;

            this.name = name;
            this.Page = new Page(documents[c].Page);
        }

        #endregion

        #region properties

        public string Title
        {
            get { return this.title; }
            set { this.title = value; }
        }

        public string Categories
        {
            get { return this.categories; }
            set { this.categories = value; }
        }

        public string Company
        {
            get { return this.company; }
            set { this.company = value; }
        }

        public string Creator
        {
            get { return this.creator; }
            set { this.creator = value; }
        }

        public string Description
        {
            get { return this.description; }
            set { this.description = value; }
        }

        public string Tags
        {
            get { return this.tags; }
            set { this.tags = value; }
        }

        public string Revision
        {
            get { return this.revision; }
            set { this.revision = value; }
        }

        public string Subject
        {
            get { return this.subject; }
            set { this.subject = value; }
        }

        public string Modified
        {
            get { return this.modified; }
            set { this.modified = value; }
        }

        public string Password
        {
            set { this.password = value; }
        }

        public Graphic Graphic
        {
            get { return this.graphic; }
            set { this.graphic = new Graphic(value); }
        }

        #endregion

        #region methods

        public void AddContent(Content content)
        {
            this.content.Add(content);
        }

        public void SetContent(List<Content> contents)
        {
            this.content.SetContent(contents);
        }

        public List<Content> GetContents()
        {
            return this.content.GetContents();
        }

        public void SetDocument()
        {
            this.DocObject = WD.WordDocument.Create();

            //Page Size
            this.DocObject.Sections[0].PageSettings.PageSize = WD.WordPageSize.A3;
            this.DocObject.Sections[0].PageSettings.Width = (uint)this.Page.PageWidth;
            this.DocObject.Sections[0].PageSettings.Height = (uint)this.Page.PageHeight;

            //Margin
            this.DocObject.Sections[0].SetMargins(WD.WordMargin.Normal);
            this.DocObject.Sections[0].Margins.Left = (uint)this.Page.MarginLeft;
            this.DocObject.Sections[0].Margins.Right = (uint)this.Page.MarginRight;
            this.DocObject.Sections[0].Margins.Top = (int)this.Page.MarginTop;
            this.DocObject.Sections[0].Margins.Bottom = (int)this.Page.MarginBottom;

            if (this.Page.Orientation == Page.Orientations.Landscape) this.DocObject.Sections[0].PageSettings.Orientation = DocumentFormat.OpenXml.Wordprocessing.PageOrientationValues.Landscape;
            if (this.Page.Columns > 1) this.DocObject.Sections[0].ColumnCount = this.Page.Columns;

            //Page Graphics
            if (this.Graphic.HasFill) this.DocObject.Background.SetColor(this.Graphic.Fill.ToSixLabors());
            this.RenderBorder(this.DocObject.Borders, this.Graphic);

            this.SetProperies();
            this.SetSecurity();
        }

        private void SetProperies()
        {

            if (this.title != string.Empty) this.DocObject.BuiltinDocumentProperties.Title = this.title;
            if (this.categories != string.Empty) this.DocObject.BuiltinDocumentProperties.Category = this.categories;
            if (this.modified != string.Empty) this.DocObject.ApplicationProperties.Company = this.company;
            if (this.company != string.Empty) this.DocObject.BuiltinDocumentProperties.LastModifiedBy = this.modified;
            if (this.creator != string.Empty) this.DocObject.BuiltinDocumentProperties.Creator = this.creator;
            if (this.description != string.Empty) this.DocObject.BuiltinDocumentProperties.Description = this.description;
            if (this.tags != string.Empty) this.DocObject.BuiltinDocumentProperties.Keywords = this.tags;
            if (this.revision != string.Empty) this.DocObject.BuiltinDocumentProperties.Revision = this.revision;
            if (this.subject != string.Empty) this.DocObject.BuiltinDocumentProperties.Subject = this.subject;

        }

        private void SetSecurity()
        {
            if (this.password != string.Empty)
            {
                this.DocObject.Settings.ProtectionPassword = this.password;
                this.DocObject.Settings.ProtectionType = WP.DocumentProtectionValues.ReadOnly;
            }
        }

        public void Render()
        {
            this.SetDocument();

            if ((this.Header.Count > 0) | (this.Footer.Count > 0))
            {
                this.DocObject.AddHeadersAndFooters();
            }
            if (this.Header.Count > 0) {
                switch (this.HeaderType)
                {
                    case HeaderFooterTypes.All:
                        Header.Render(this.DocObject.Header.Default);
                        break;
                    case HeaderFooterTypes.Even:
                        Header.Render(this.DocObject.Header.Even);
                        break;
                    case HeaderFooterTypes.First:
                        Header.Render(this.DocObject.Header.First);
                        break;
                }
            }
            if (this.Footer.Count > 0)
            {
                switch (this.FooterType)
                {
                    case HeaderFooterTypes.All:
                        Footer.Render(this.DocObject.Footer.Default);
                        break;
                    case HeaderFooterTypes.Even:
                        Footer.Render(this.DocObject.Footer.Even);
                        break;
                    case HeaderFooterTypes.First:
                        Footer.Render(this.DocObject.Footer.First);
                        break;
                }
            }

            this.content.Render(this);
        }

        protected void RenderBorder(WD.WordBorders border, Graphic graphic)
        {
            if (graphic.TopBorder)
            {
                border.TopStyle = graphic.BorderValue;
                border.TopColor = graphic.Stroke.ToSixLabors();
                border.TopSize= graphic.Weight.ToXmlUint32();
            }

            if (graphic.BottomBorder)
            {
                border.BottomStyle = graphic.BorderValue;
                border.BottomColor = graphic.Stroke.ToSixLabors();
                border.BottomSize = graphic.Weight.ToXmlUint32();
            }

            if (graphic.LeftBorder)
            {
                border.LeftStyle = graphic.BorderValue;
                border.LeftColor = graphic.Stroke.ToSixLabors();
                border.LeftSize = graphic.Weight.ToXmlUint32();
            }

            if (graphic.RightBorder)
            {
                border.RightStyle = graphic.BorderValue;
                border.RightColor = graphic.Stroke.ToSixLabors();
                border.RightSize = graphic.Weight.ToXmlUint32();
            }
        }

        public void Save(string filepath)
        {
            this.Render();
            this.DocObject.Save(filepath, false);
        }

        public string Stream()
        {
            this.Render();
            Stream stream = new MemoryStream();
            this.DocObject.Save(stream);
            stream.Position = 0;
            var buffer = new byte[stream.Length];
            stream.Read(buffer, 0, (int)stream.Length);
            string output = System.Text.Encoding.Default.GetString(buffer);
            stream.Dispose();
            return output;
        }

        #endregion

        #region overrides

        public override string ToString()
        {
            return "WD | Document: " + this.name + " {c:" + this.content.Count+"}" ;
        }

        #endregion

    }
}
