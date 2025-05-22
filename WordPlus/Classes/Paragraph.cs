using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordPlus
{
    public class Paragraph
    {

        #region methods

        protected List<Fragment> fragments = new List<Fragment>();
        protected double lineSpacing = 1.0;

        #endregion

        #region constructors

        public Paragraph()
        {

        }

        public Paragraph(Paragraph paragraph)
        {
            this.fragments = paragraph.fragments.Duplicate();
            this.lineSpacing = paragraph.lineSpacing;
        }

        public Paragraph(string text)
        {
            this.fragments.Add(new Fragment(text));
        }

        public Paragraph(Fragment fragment)
        {
            this.fragments.Add( new Fragment(fragment));
        }

        public Paragraph(List<Fragment> fragments)
        {
            this.fragments = fragments.Duplicate();
        }

        public Paragraph(List<Paragraph> paragraphs)
        {
            foreach(Paragraph paragraph in paragraphs) this.fragments.AddRange(paragraph.Fragments.Duplicate());
        }

        #endregion

        #region properties

        public double LineSpacing
        {
            get { return this.lineSpacing; }
            set { this.lineSpacing = value; }
        }

        public List<Fragment> Fragments
        {
            get { return fragments; }
            set { fragments = value; }
        }

        public string Text
        {
            get
            {
                StringBuilder output = new StringBuilder();
                foreach (Fragment fragment in this.fragments) output.Append(fragment.Text);
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
