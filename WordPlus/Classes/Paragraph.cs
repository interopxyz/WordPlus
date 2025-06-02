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
        protected int indentAfter = -1;
        protected int indentBefore = -1;
        protected int indentFirstLine = -1;
        protected int indentHanging = -1;

        protected Graphic graphic = new Graphic();

        #endregion

        #region constructors

        public Paragraph()
        {

        }

        public Paragraph(Paragraph paragraph)
        {
            this.fragments = paragraph.fragments.Duplicate();
            this.lineSpacing = paragraph.lineSpacing;
            this.indentAfter = paragraph.indentAfter;
            this.indentBefore = paragraph.indentBefore;
            this.indentFirstLine = paragraph.indentFirstLine;
            this.indentHanging = paragraph.indentHanging;

            this.Graphic = paragraph.graphic;
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

        public Graphic Graphic
        {
            get { return this.graphic; }
            set { this.graphic = new Graphic(value); }
        }

        public double LineSpacing
        {
            get { return this.lineSpacing; }
            set { this.lineSpacing = value; }
        }

        public int IndentBefore
        {
            get { return this.indentBefore; }
            set { this.indentBefore = value; }
        }

        public bool HasIndentBefore
        {
            get { return this.indentBefore > -1; }
        }

        public int IndentAfter
        {
            get { return this.indentAfter; }
            set { this.indentAfter = value; }
        }

        public bool HasIndentAfter
        {
            get { return this.indentAfter> -1; }
        }

        public int IndentFirstLine
        {
            get { return this.indentFirstLine; }
            set { this.indentFirstLine = value; }
        }

        public bool HasIndentFirstLine
        {
            get { return this.indentFirstLine> -1; }
        }

        public int IndentHanging
        {
            get { return this.indentHanging; }
            set { this.indentHanging = value; }
        }

        public bool HasIndentHanging
        {
            get { return this.indentHanging > -1; }
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

            if (this.Text.Length < 16) return "WD | Paragraph {" + this.Fragments.Count + "f: " + this.Text + "}";
            return "WD | Paragraph {" + this.Fragments.Count + "f: "+this.Text.Substring(0,15)+"...}";
        }

        #endregion

    }
}
