using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordPlus
{
    public class Fragment
    {

        #region members

        public string Text = string.Empty;
        public string Hyperlink = string.Empty;
        protected Font font = Fonts.Normal;
        protected Graphic graphic = Graphics.Outline;

        #endregion

        #region constructors

        public Fragment() { }

        public Fragment(Fragment fragment)
        {
            this.Text = fragment.Text;
            this.Hyperlink = fragment.Hyperlink;
            this.font = new Font(fragment.font);
            this.graphic = new Graphic(fragment.graphic);
        }

        public Fragment(string text)
        {
            this.Text = text;
        }

        public Fragment(string text, string hyperlink)
        {
            this.Text = text;
            this.Hyperlink = hyperlink;
        }

        public Fragment(string text, Font font)
        {
            this.Text = text;
            this.font = new Font(font);
        }

        #endregion

        #region properties

        public Font Font
        {
            get { return this.font; }
            set { this.font = value; }
        }

        public Graphic Graphic
        {
            get { return this.graphic; }
            set { this.graphic = value; }
        }

        public bool HasLink
        {
            get { return this.Hyperlink != string.Empty; }
        }

        #endregion

        #region overrides

        public override string ToString()
        {
            if (this.Text.Length < 16) return "Wd | Fragment {" + this.Text + "}";
            return "WD | Fragment {" + this.Text.Substring(0, 15) + "...}";
        }

        #endregion

    }
}
