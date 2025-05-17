using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using WD = OfficeIMO.Word;

namespace WordPlus
{
    public class WdPage
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

        public WdPage()
        {

        }

        public WdPage(WdPage page)
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

        public WdPage(string name, double width, double height)
        {
            this.name = name;
            this.width = width;
            this.height = height;
        }

        public WdPage(string name, double width, double height, Margins margin)
        {
            this.name = name;
            this.width = width;
            this.height = height;
            this.SetMargin(margin);
        }

        //North American
        public static WdPage Letter() { return new WdPage("Letter", 12240, 15840, Margins.Generous); }
        public static WdPage Legal() { return new WdPage("Legal", 12240, 20160, Margins.Generous); }
        public static WdPage Statement() { return new WdPage("Statement", 7920, 12240, Margins.Dense); }
        public static WdPage Tabloid() { return new WdPage("Tabloid", 15840, 24480, Margins.Dense); }
        public static WdPage Executive() { return new WdPage("Executive", 6120, 15120, Margins.Comfortable); }
        public static WdPage ANSIC() { return new WdPage("ANSIC", 24480, 31680, Margins.Dense); }

        //ISO A Series
        //public static WdPage A0() { return new WdPage("A0", 47664, 67368); }//Word limited to 22x22"
        //public static WdPage A1() { return new WdPage("A1", 33696, 47664); }//Word limited to 22x22"
        //public static WdPage A2() { return new WdPage("A2", 23760, 33696); }//Word limited to 22x22"
        public static WdPage A3() { return new WdPage("A3", 16838, 23760, Margins.Wide); }
        public static WdPage A4() { return new WdPage("A4", 11909, 16838, Margins.Generous); }
        public static WdPage A5() { return new WdPage("A5", 8395, 11909, Margins.Moderate); }
        public static WdPage A6() { return new WdPage("A6", 5950, 8395, Margins.Narrow); }
        public static WdPage A7() { return new WdPage("A7", 4195, 5950, Margins.Narrow); }
        public static WdPage A8() { return new WdPage("A8", 2948, 4195, Margins.Narrow); }
        public static WdPage A9() { return new WdPage("A9", 2098, 2948, Margins.Narrow); }
        public static WdPage A10() { return new WdPage("A10", 1474, 2098, Margins.Narrow); }

        //ISO B Series
        //public static WdPage B0() { return new WdPage("B0", 56693, 80210); }//Word limited to 22x22"
        //public static WdPage B1() { return new WdPage("B1", 40034, 56693); }//Word limited to 22x22"
        //public static WdPage B2() { return new WdPage("B2", 28346, 40034); }//Word limited to 22x22"
        public static WdPage B3() { return new WdPage("B3", 20008, 28346, Margins.Narrow); }
        public static WdPage B4() { return new WdPage("B4", 14173, 20008, Margins.Moderate); }
        public static WdPage B5() { return new WdPage("B5", 9978, 14173, Margins.Moderate); }
        public static WdPage B6() { return new WdPage("B6", 7087, 9978, Margins.Moderate); }
        public static WdPage B7() { return new WdPage("B7", 4989, 7087, Margins.Tight); }
        public static WdPage B8() { return new WdPage("B8", 3515, 4989, Margins.Tight); }
        public static WdPage B9() { return new WdPage("B9", 2494, 3515, Margins.Tight); }
        public static WdPage B10() { return new WdPage("B10", 1757, 2494, Margins.Tight); }

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

        public static WdPage IsoA(WdPage.SizesIsoA type)
        {
            switch(type)
            {
                default:
                return WdPage.A4();
                //case SizesIsoA.A0:
                //    return WdPage.A0();
                //case SizesIsoA.A1:
                //    return WdPage.A1();
                //case SizesIsoA.A2:
                //    return WdPage.A2();
                case SizesIsoA.A3:
                    return WdPage.A3();
                case SizesIsoA.A5:
                    return WdPage.A5();
                case SizesIsoA.A6:
                    return WdPage.A6();
                case SizesIsoA.A7:
                    return WdPage.A7();
                case SizesIsoA.A8:
                    return WdPage.A8();
                case SizesIsoA.A9:
                    return WdPage.A9();
                case SizesIsoA.A10:
                    return WdPage.A10();
            }
        }

        public static WdPage IsoB(WdPage.SizesIsoB type)
        {
            switch (type)
            {
                default:
                    return WdPage.B4();
                //case SizesIsoB.B0:
                //    return WdPage.B0();
                //case SizesIsoB.B1:
                //    return WdPage.B1();
                //case SizesIsoB.B2:
                //    return WdPage.B2();
                case SizesIsoB.B3:
                    return WdPage.B3();
                case SizesIsoB.B5:
                    return WdPage.B5();
                case SizesIsoB.B6:
                    return WdPage.B6();
                case SizesIsoB.B7:
                    return WdPage.B7();
                case SizesIsoB.B8:
                    return WdPage.B8();
                case SizesIsoB.B9:
                    return WdPage.B9();
                case SizesIsoB.B10:
                    return WdPage.B10();
            }
        }

        public static WdPage Imperial(WdPage.SizesUS type)
        {
            switch (type)
            {
                default:
                    return WdPage.Letter();
                case SizesUS.Executive:
                    return WdPage.Executive();
                case SizesUS.Legal:
                    return WdPage.Legal();
                case SizesUS.Statement:
                    return WdPage.Statement();
                case SizesUS.Tabloid:
                    return WdPage.Tabloid();
                case SizesUS.ANSIC:
                    return WdPage.ANSIC();
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
