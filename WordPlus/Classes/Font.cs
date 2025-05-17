using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Sd = System.Drawing;
using Rg = Rhino.Geometry;

namespace WordPlus
{
    public class Font
    {
        #region members
        public enum Presets { None, Normal, Title, Subtitle, Heading1, Heading2, Heading3, Heading4, Heading5, Heading6, Quote, Footnote, Caption, List };
        public enum HorizontalAlignments {Default, Left, Center, Right};
        public enum VerticalAlignments {Default, Top, Middle, Bottom};

        public enum HighlightColors { None, Black,Blue,Cyan,DarkBlue,DarkCyan,DarkGray,DarkGreen,DarkMagenta,DarkRed,DarkYellow,Green,LightGray,Magenta,Red,White,Yellow};

        protected bool isModified = false;
        protected string name = "None";

        //Font Formatting
        protected string family = string.Empty;
        protected double size = 0;
        protected Sd.Color color = Sd.Color.Empty;
        protected HighlightColors highlight = HighlightColors.None;

        protected bool isBold = false;
        protected bool isItalic = false;
        protected bool isUnderline = false;
        protected bool isStrikethrough = false;
        protected bool isSubscript = false;
        protected bool isSuperscript = false;

        protected HorizontalAlignments horizontalAlignment = HorizontalAlignments.Left;
        protected VerticalAlignments verticalAlignment = VerticalAlignments.Top;

        //Paragraph Formatting
        protected double lineSpacing = 1.0;
        #endregion

        #region constructors

        public Font()
        {

        }

        public Font(Font font)
        {
            this.isModified = font.isModified;

            this.name = font.name;

            this.family = font.family;
            this.size = font.size;
            this.color = font.color;
            this.highlight = font.highlight;

            this.isBold = font.isBold;
            this.isItalic = font.isItalic;
            this.isUnderline = font.isUnderline;
            this.isStrikethrough = font.isStrikethrough;
            this.isSubscript = font.isSubscript;
            this.isSuperscript = font.isSuperscript;

            this.lineSpacing = font.lineSpacing;

            this.horizontalAlignment = font.horizontalAlignment;
            this.verticalAlignment = font.verticalAlignment;
        }

        public Font(string family, double size, Sd.Color color)
        {
            this.isModified = true;

            this.family = family;
            this.size = size;
            this.color = color;
        }

        public Font(string name, string family, double size, Sd.Color color)
        {
            this.isModified = true;

            this.name = name;

            this.family = family;
            this.size = size;
            this.color = color;

        }

        #endregion

        #region properties

        public virtual bool IsModified { get { return this.isModified; } }

        public virtual string Name
        {
            get { return this.name; }
            set { this.name = value; }
        }

        public virtual string Family
        {
            get { return this.family; }
            set
            {
                this.isModified = true;
                this.family = value;
            }
        }

        public virtual bool HasFamily { get { return this.family != string.Empty; } }

        public virtual double Size
        {
            get { return this.size; }
            set
            {
                this.isModified = true;
                this.size = value;
            }
        }

        public virtual bool HasSize { get { return this.size > 0; } }

        public virtual Sd.Color Color
        {
            get { return this.color; }
            set
            {
                this.isModified = true;
                this.color = value;
            }
        }

        public virtual bool HasColor { get { return this.color != Sd.Color.Empty; } }

        public virtual HighlightColors Highlight
        {
            get { return this.highlight; }
            set
            {
                this.isModified = true;
                this.highlight = value;
            }
        }

        public virtual bool HasHighlight { get { return this.highlight != HighlightColors.None; } }

        public virtual double LineSpacing
        {
            get { return this.lineSpacing; }
            set
            {
                this.lineSpacing = value;
            }
        }

        public virtual bool Bold
        {
            get { return this.isBold; }
            set { this.isBold = value; }
        }

        public virtual bool Italic
        {
            get { return this.isItalic; }
            set { this.isItalic = value; }
        }

        public virtual bool Underlined
        {
            get { return this.isUnderline; }
            set { this.isUnderline = value; }
        }

        public virtual bool Strikethrough
        {
            get { return this.isStrikethrough; }
            set { this.isStrikethrough = value; }
        }

        public virtual bool Superscript
        {
            get { return this.isSuperscript; }
            set { this.isSuperscript = value; }
        }

        public virtual bool Subscript
        {
            get { return this.isSubscript; }
            set { this.isSubscript = value; }
        }

        public virtual HorizontalAlignments HorizontalAlignment
        {
            get { return this.horizontalAlignment; }
            set { this.horizontalAlignment = value; }
        }

        public virtual VerticalAlignments VerticalAlignment
        {
            get { return this.verticalAlignment; }
            set { this.verticalAlignment = value; }
        }

        #endregion

        #region methods



        #endregion

        #region overrides



        #endregion
    }

    public static class Fonts
    {
        public static Font GetPreset(this Font.Presets input)
        {
            switch (input)
            {
                default:
                    return Fonts.Normal;
                case Font.Presets.Title:
                    return Fonts.Title;
                case Font.Presets.Subtitle:
                    return Fonts.Subtitle;
                case Font.Presets.Heading1:
                    return Fonts.Heading1;
                case Font.Presets.Heading2:
                    return Fonts.Heading2;
                case Font.Presets.Heading3:
                    return Fonts.Heading3;
                case Font.Presets.Heading4:
                    return Fonts.Heading4;
                case Font.Presets.Heading5:
                    return Fonts.Heading5;
                case Font.Presets.Heading6:
                    return Fonts.Heading6;
                case Font.Presets.Quote:
                    return Fonts.Quote;
                case Font.Presets.Footnote:
                    return Fonts.Footnote;
                case Font.Presets.Caption:
                    return Fonts.Caption;
                case Font.Presets.List:
                    return Fonts.List;
            }
        }

        public static Font Normal { get { return new Font("Normal", "Arial", 11, Sd.Color.Black); } }
        public static Font Title { get { return new Font("Title", "Arial", 26, Sd.Color.Black); } }
        public static Font Subtitle { get { return new Font("Subtitle", "Arial", 15, Sd.Color.DarkGray); } }
        public static Font Heading1 { get { return new Font("Heading1", "Arial", 20, Sd.Color.Black); } }
        public static Font Heading2 { get { return new Font("Heading2", "Arial", 16, Sd.Color.Black); } }
        public static Font Heading3 { get { return new Font("Heading3", "Arial", 14, Sd.Color.Gray); } }
        public static Font Heading4 { get { return new Font("Heading4", "Arial", 12, Sd.Color.Gray); } }
        public static Font Heading5 { get { return new Font("Heading5", "Arial", 11, Sd.Color.Gray); } }
        public static Font Heading6 { get { return new Font("Heading6", "Arial", 11, Sd.Color.Gray); } }
        public static Font Quote { get { return new Font("Quote", "Arial", 11, Sd.Color.DarkGray); } }
        public static Font Footnote { get { return new Font("Footnote", "Arial", 8, Sd.Color.Black); } }
        public static Font Caption { get { return new Font("Caption", "Arial", 8, Sd.Color.Black); } }
        public static Font List { get { return new Font("List", "Arial", 11, Sd.Color.Black); } }
        public static Font Table { get { return new Font("Table", "Arial", 9, Sd.Color.Black); } }
        public static Font Preview { get { return new Font("List", "Arial", 4, Sd.Color.Red); } }

    }
}
