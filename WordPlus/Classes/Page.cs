using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using WD = OfficeIMO.Word;

namespace WordPlus
{
    public class Page
    {

        #region members

        public enum SizesIsoA { A3, A4, A5, A6, A7, A8, A9, A10 };
        public enum SizesIsoB { B3, B4, B5, B6, B7, B8, B9, B10 };
        public enum SizesUS { Letter, Legal, Statement, Tabloid, Executive, ANSIC };

        public enum Margins { Tight, Narrow, Dense, Moderate, Comfortable, Wide, Generous};
        public enum Orientations { Portrait, Landscape};

        string name = "A3";
        //Units in Twips
        double width = 12240;
        double height = 15840;

        double marginLeft = 1440;
        double marginRight = 1440;
        double marginTop = 1440;
        double marginBottom = 1440;

        int columns = 1;

        public Orientations Orientation = Orientations.Portrait;

        #endregion

        #region constructors

        public Page()
        {

        }

        public Page(Page page)
        {
            this.name = page.name;

            this.width = page.width;
            this.height = page.height;

            this.marginLeft = page.marginLeft;
            this.marginRight = page.marginRight;
            this.marginTop = page.marginTop;
            this.marginBottom = page.marginBottom;

            this.columns = page.columns;

            this.Orientation = page.Orientation;
        }

        public Page(string name, double width, double height)
        {
            this.name = name;
            this.width = width;
            this.height = height;
        }

        public Page(string name, double width, double height, Margins margin)
        {
            this.name = name;
            this.width = width;
            this.height = height;
            this.SetMargin(margin);
        }

        //North American
        public static Page Letter() { return new Page("Letter", 12240, 15840, Margins.Generous); }
        public static Page Legal() { return new Page("Legal", 12240, 20160, Margins.Generous); }
        public static Page Statement() { return new Page("Statement", 7920, 12240, Margins.Dense); }
        public static Page Tabloid() { return new Page("Tabloid", 15840, 24480, Margins.Dense); }
        public static Page Executive() { return new Page("Executive", 6120, 15120, Margins.Comfortable); }
        public static Page ANSIC() { return new Page("ANSIC", 24480, 31680, Margins.Dense); }

        //ISO A Series
        //public static WdPage A0() { return new WdPage("A0", 47664, 67368); }//Word limited to 22x22"
        //public static WdPage A1() { return new WdPage("A1", 33696, 47664); }//Word limited to 22x22"
        //public static WdPage A2() { return new WdPage("A2", 23760, 33696); }//Word limited to 22x22"
        public static Page A3() { return new Page("A3", 16838, 23760, Margins.Wide); }
        public static Page A4() { return new Page("A4", 11909, 16838, Margins.Generous); }
        public static Page A5() { return new Page("A5", 8395, 11909, Margins.Moderate); }
        public static Page A6() { return new Page("A6", 5950, 8395, Margins.Narrow); }
        public static Page A7() { return new Page("A7", 4195, 5950, Margins.Narrow); }
        public static Page A8() { return new Page("A8", 2948, 4195, Margins.Narrow); }
        public static Page A9() { return new Page("A9", 2098, 2948, Margins.Narrow); }
        public static Page A10() { return new Page("A10", 1474, 2098, Margins.Narrow); }

        //ISO B Series
        //public static WdPage B0() { return new WdPage("B0", 56693, 80210); }//Word limited to 22x22"
        //public static WdPage B1() { return new WdPage("B1", 40034, 56693); }//Word limited to 22x22"
        //public static WdPage B2() { return new WdPage("B2", 28346, 40034); }//Word limited to 22x22"
        public static Page B3() { return new Page("B3", 20008, 28346, Margins.Narrow); }
        public static Page B4() { return new Page("B4", 14173, 20008, Margins.Moderate); }
        public static Page B5() { return new Page("B5", 9978, 14173, Margins.Moderate); }
        public static Page B6() { return new Page("B6", 7087, 9978, Margins.Moderate); }
        public static Page B7() { return new Page("B7", 4989, 7087, Margins.Tight); }
        public static Page B8() { return new Page("B8", 3515, 4989, Margins.Tight); }
        public static Page B9() { return new Page("B9", 2494, 3515, Margins.Tight); }
        public static Page B10() { return new Page("B10", 1757, 2494, Margins.Tight); }

        #endregion

        #region properties

        public virtual string Name
        {
            get { return this.name; }
            set { this.name = value; }
        }

        public virtual int PageWidth
        {
            get { return (int)this.width; }
            set { this.width = value; }
        }
        public virtual int PageHeight
        {
            get { return (int)this.height; }
            set { this.height = value; }
        }

        public virtual int MarginLeft
        {
            get { return (int)this.marginLeft; }
            set { this.marginLeft = value; }
        }
        public virtual int MarginRight
        {
            get { return (int)this.marginRight; }
            set { this.marginRight = value; }
        }
        public virtual int MarginTop
        {
            get { return (int)this.marginTop; }
            set { this.marginTop = value; }
        }
        public virtual int MarginBottom
        {
            get { return (int)this.marginBottom; }
            set { this.marginBottom = value; }
        }

        public virtual int ContentWidth
        {
            get { return (int)(this.width-this.marginLeft-this.marginRight); }
        }

        public virtual int Columns
        {
            get { return this.columns; }
            set { this.columns = value; }
        }

        public virtual int ContentHeight
        {
            get { return (int)(this.height - this.marginTop - this.marginBottom); }
        }


        #endregion

        #region methods

        public static Page IsoA(Page.SizesIsoA type)
        {
            switch(type)
            {
                default:
                return Page.A4();
                //case SizesIsoA.A0:
                //    return WdPage.A0();
                //case SizesIsoA.A1:
                //    return WdPage.A1();
                //case SizesIsoA.A2:
                //    return WdPage.A2();
                case SizesIsoA.A3:
                    return Page.A3();
                case SizesIsoA.A5:
                    return Page.A5();
                case SizesIsoA.A6:
                    return Page.A6();
                case SizesIsoA.A7:
                    return Page.A7();
                case SizesIsoA.A8:
                    return Page.A8();
                case SizesIsoA.A9:
                    return Page.A9();
                case SizesIsoA.A10:
                    return Page.A10();
            }
        }

        public static Page IsoB(Page.SizesIsoB type)
        {
            switch (type)
            {
                default:
                    return Page.B4();
                //case SizesIsoB.B0:
                //    return WdPage.B0();
                //case SizesIsoB.B1:
                //    return WdPage.B1();
                //case SizesIsoB.B2:
                //    return WdPage.B2();
                case SizesIsoB.B3:
                    return Page.B3();
                case SizesIsoB.B5:
                    return Page.B5();
                case SizesIsoB.B6:
                    return Page.B6();
                case SizesIsoB.B7:
                    return Page.B7();
                case SizesIsoB.B8:
                    return Page.B8();
                case SizesIsoB.B9:
                    return Page.B9();
                case SizesIsoB.B10:
                    return Page.B10();
            }
        }

        public static Page Imperial(Page.SizesUS type)
        {
            switch (type)
            {
                default:
                    return Page.Letter();
                case SizesUS.Executive:
                    return Page.Executive();
                case SizesUS.Legal:
                    return Page.Legal();
                case SizesUS.Statement:
                    return Page.Statement();
                case SizesUS.Tabloid:
                    return Page.Tabloid();
                case SizesUS.ANSIC:
                    return Page.ANSIC();
            }
        }

        public void SetMargin(Margins preset)
        {
            switch (preset)
            {
                default:
                    this.SetMargin(850, 850, 850, 850);
                    break;
                case Margins.Tight:
                    this.SetMargin(283, 283, 283, 283);
                    break;
                case Margins.Narrow:
                    this.SetMargin(566, 566, 566, 566);
                    break;
                case Margins.Dense:
                    this.SetMargin(720, 720, 720, 720);
                    break;
                case Margins.Comfortable:
                    this.SetMargin(1080, 1080, 1080, 1080);
                    break;
                case Margins.Wide:
                    this.SetMargin(1133, 1133, 1133, 1133);
                    break;
                case Margins.Generous:
                    this.SetMargin(1440, 1440, 1440, 1440);
                    break;
            }
        }

        public void SetMargin(double left, double right, double top, double bottom)
        {
            this.marginLeft = left;
            this.marginRight = right;
            this.marginTop = top;
            this.marginBottom = bottom;
        }

        #endregion

    }
}
