using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using WD = OfficeIMO.Word;
using WP = DocumentFormat.OpenXml.Wordprocessing;
using Sd = System.Drawing;

namespace WordPlus
{
    public class WdContentBlock
    {

        #region members

        public List<Content> contents = new List<Content>();
        public double Width = 0;
        public double Height = 0;

        #endregion

        #region constructor

        public WdContentBlock()
        {

        }

        public WdContentBlock(WdContentBlock block)
        {
            this.Width = block.Width;
            this.Height = block.Height;
            this.contents = block.GetContents();
        }

        #endregion

        #region properties

        public int Count
        {
            get { return this.contents.Count; }
        }

        #endregion

        #region methods

        public void Add(Content content)
        {
            this.contents.Add(new Content(content));
        }

        public void AddRange(List<Content> contents)
        {
            this.contents.AddRange(contents.Duplicate());
        }

        public void SetContent(List<Content> contents)
        {
            this.contents.Clear();
            this.contents = contents.Duplicate();
        }

        public List<Content> GetContents()
        {
            return this.contents.Duplicate();
        }

        public void Render(WdDocument document)
        {
            double c = 1;
            if(document.DocObject.Sections[0].ColumnCount!=null)c= (double)document.DocObject.Sections[0].ColumnCount;
            double w = (double)(document.DocObject.Sections[0].PageSettings.Width.Value - document.DocObject.Sections[0].Margins.Left.Value - document.DocObject.Sections[0].Margins.Right.Value) / 15.0 / c;
            double h = (double)(document.DocObject.Sections[0].PageSettings.Height.Value - document.DocObject.Sections[0].Margins.Top.Value - document.DocObject.Sections[0].Margins.Bottom.Value) / 15.0 / c;
            foreach (Content content in this.GetContents()) content.Render(document.DocObject.AddParagraph(), w, h);
        }

        public void Render(WD.WordHeader header)
        {
            foreach (Content content in this.GetContents()) content.Render(header.AddParagraph(), this.Width, this.Height);
        }

        public void Render(WD.WordFooter footer)
        {
            foreach (Content content in this.GetContents()) content.Render(footer.AddParagraph(), this.Width, this.Height);
        }

            #endregion
        }
}
