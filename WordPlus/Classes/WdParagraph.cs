using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordPlus
{
    public class WdParagraph
    {

        #region methods

        protected List<WdFragment> fragments = new List<WdFragment>();
        protected double lineSpacing = 1.0;

        #endregion

        #region constructors

        public WdParagraph()
        {

        }

        public WdParagraph(WdParagraph paragraph)
        {
            this.fragments = paragraph.fragments.Duplicate();
            this.lineSpacing = paragraph.lineSpacing;
        }

        public WdParagraph(string text)
        {
            this.fragments.Add(new WdFragment(text));
        }

        public WdParagraph(WdFragment fragment)
        {
            this.fragments.Add( new WdFragment(fragment));
        }

        public WdParagraph(List<WdFragment> fragments)
        {
            this.fragments = fragments.Duplicate();
        }

        public WdParagraph(List<WdParagraph> paragraphs)
        {
            foreach(WdParagraph paragraph in paragraphs) this.fragments.AddRange(paragraph.Fragments.Duplicate());
        }

        #endregion

        #region properties

        public double LineSpacing
        {
            get { return this.lineSpacing; }
            set { this.lineSpacing = value; }
        }

        public List<WdFragment> Fragments
        {
            get { return fragments; }
            set { fragments = value; }
        }

        public string Text
        {
            get
            {
                StringBuilder output = new StringBuilder();
                foreach (WdFragment fragment in this.fragments) output.Append(fragment.Text);
                return output.ToString();
            }
        }

        #endregion

        #region overrides

        public override string ToString()
        {
            return "WD | Paragraph {" + this.Fragments.Count + "f}";
        }

        #endregion

    }
}
