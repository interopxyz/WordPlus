using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Sd = System.Drawing;
using Rg = Rhino.Geometry;

using WD = OfficeIMO.Word;
using WP = DocumentFormat.OpenXml.Wordprocessing;

namespace WordPlus
{
    public class Graphic
    {
        #region members

        public enum LineTypes { Nil, None, Single, DashDotStroked, Dashed, DashSmallGap, DotDash, DotDotDash, Dotted, Double, DoubleWave, Inset, Outset, ThickThinLargeGap, ThickThinMediumGap, ThickThinSmallGap, ThinThickLargeGap, ThinThickMediumGap, ThinThickSmallGap, ThinThickThinLargeGap, ThinThickThinMediumGap, ThinThickThinSmallGap, Triple, Wave, ThreeDEmboss, ThreeDEngrave };

        protected Sd.Color fill = Sd.Color.Transparent;
        protected bool hasFill = false;

        protected Sd.Color stroke = Sd.Color.Black;
        protected double weight = 1.0;
        protected bool hasStroke = false;

        protected bool topBorder = false;
        protected bool bottomBorder = false;
        protected bool leftBorder = false;
        protected bool rightBorder = false;

        protected LineTypes lineType = LineTypes.Single;

        #endregion

        #region constructor

        public Graphic()
        {

        }

        public Graphic(Graphic graphic)
        {
            this.fill = graphic.fill;
            this.hasFill = graphic.hasFill;

            this.topBorder = graphic.topBorder;
            this.bottomBorder = graphic.bottomBorder;
            this.leftBorder = graphic.leftBorder;
            this.rightBorder = graphic.rightBorder;

            this.weight = graphic.weight;
            this.stroke = graphic.stroke;
            this.hasStroke = graphic.hasStroke;
            this.lineType = graphic.lineType;
        }

        public Graphic(Sd.Color stroke, double weight)
        {
            this.Stroke = stroke;
            this.Weight = weight;
            this.hasFill = false;
        }

        public Graphic(Sd.Color stroke, Sd.Color fillColor)
        {
            this.Stroke = stroke;
            this.Fill = fillColor;
        }

        public Graphic(Sd.Color color)
        {
            this.Fill = color;

            this.Stroke = Sd.Color.Transparent;
            this.Weight = 0.0;
            this.hasStroke = false;
        }

        #endregion

        #region properties

        public virtual Sd.Color Fill
        {
            get { return fill; }
            set
            {
                this.fill = value;
                this.hasFill = true;
            }
        }

        public virtual bool HasFill
        {
            get { return hasFill; }
        }

        public virtual Sd.Color Stroke
        {
            get { return stroke; }
            set
            {
                this.stroke = value;
                this.hasStroke = true;
            }
        }

        public virtual double Weight
        {
            get { return weight; }
            set
            {
                this.weight = value;
                this.hasStroke = true;
            }
        }

        public virtual bool HasStroke
        {
            get { return hasStroke; }
        }

        public virtual bool NoStroke
        {
            get
            {
                if (this.stroke.A == 0) return true;
                if (this.weight <= 0) return true;
                return false;
            }
        }

        public virtual LineTypes LineType
        {
            get { return this.lineType; }
            set 
            { 
                this.lineType = value;
                this.hasStroke = true;
            }
        }

        public virtual WP.BorderValues BorderValue
        {
            get { return this.GetBorderValue(); }
        }

        public virtual bool NoFill
        {
            get
            {
                if (this.fill.A == 0) return true;
                return false;
            }
        }

        public virtual bool OnlyStroke
        {
            get
            {
                return (this.NoFill && !this.NoStroke);
            }
        }

        public virtual bool OnlyFill
        {
            get
            {
                return (!this.NoFill && this.NoStroke);
            }
        }

        public virtual bool TopBorder
        {
            get { return this.topBorder; }
            set { this.topBorder = value; }
        }

        public virtual bool BottomBorder
        {
            get { return this.bottomBorder; }
            set { this.bottomBorder = value; }
        }

        public virtual bool LeftBorder
        {
            get { return this.leftBorder; }
            set { this.leftBorder = value; }
        }

        public virtual bool RightBorder
        {
            get { return this.rightBorder; }
            set { this.rightBorder = value; }
        }

        public virtual bool HasBorders
        {
            get { return (this.leftBorder & this.rightBorder & this.topBorder & this.bottomBorder); }
            set 
            { 
                this.leftBorder = value;
                this.rightBorder = value;
                this.topBorder = value;
                this.bottomBorder = value;
            }
        }

        #endregion

        #region methods

        public void SetGraphic(WD.WordTableCell input)
        {
            if (input.ShadingFillColor.HasValue) this.Fill = input.ShadingFillColor.Value.ToColor();
            this.UpdateBorders(input.Borders);
        }

        public void UpdateBorders(WD.WordParagraphBorders borders)
        {

            if (borders.BottomColorHex != null)
            {
                this.BottomBorder = true;
                if (borders.BottomColor.HasValue) this.Stroke = borders.BottomColor.Value.ToColor();
                if (borders.BottomSize.HasValue) this.Weight = borders.BottomSize.Value/8.0;
                if (borders.BottomStyle.HasValue) this.SetLineType(borders.BottomStyle.Value);
            }

            if (borders.TopColorHex != null)
            {
                this.TopBorder = true;
                if (borders.TopColor.HasValue) this.Stroke = borders.TopColor.Value.ToColor();
                if (borders.TopSize.HasValue) this.Weight = borders.TopSize.Value / 8.0;
                if (borders.TopStyle.HasValue) this.SetLineType(borders.TopStyle.Value);
            }

            if (borders.LeftColorHex != null)
            {
                this.LeftBorder = true;
                if (borders.LeftColor.HasValue) this.Stroke = borders.LeftColor.Value.ToColor();
                if (borders.LeftSize.HasValue) this.Weight = borders.LeftSize.Value / 8.0;
                if (borders.LeftStyle.HasValue) this.SetLineType(borders.LeftStyle.Value);
            }

            if (borders.RightColorHex != null)
            {
                this.RightBorder = true;
                if (borders.RightColor.HasValue) this.Stroke = borders.RightColor.Value.ToColor();
                if (borders.RightSize.HasValue) this.Weight = borders.RightSize.Value / 8.0;
                if (borders.RightStyle.HasValue) this.SetLineType(borders.RightStyle.Value);
            }
        }

        public void UpdateBorders(WD.WordTableCellBorder borders)
        {

            if (borders.LeftColorHex != null)
            {
                this.LeftBorder = true;
                if (borders.LeftColorHex != "auto") this.Stroke = borders.LeftColor.ToColor();
                if (borders.LeftSize.HasValue) this.Weight = borders.LeftSize.Value / 8.0;
                if (borders.LeftStyle.HasValue) this.SetLineType(borders.LeftStyle.Value);
            }

            if (borders.RightColorHex != null)
            {
                this.RightBorder = true;
                if (borders.RightColorHex != "auto") this.Stroke = borders.RightColor.ToColor();
                if (borders.RightSize.HasValue) this.Weight = borders.RightSize.Value / 8.0;
                if (borders.RightStyle.HasValue) this.SetLineType(borders.RightStyle.Value);
            }

            if (borders.TopColorHex != null)
            {
                this.TopBorder = true;
                if (borders.TopColorHex != "auto") this.Stroke = borders.TopColor.ToColor();
                if (borders.TopSize.HasValue) this.Weight = borders.TopSize.Value / 8.0;
                if (borders.TopStyle.HasValue) this.SetLineType(borders.TopStyle.Value);
            }

            if (borders.BottomColorHex != null)
            {
                this.BottomBorder = true;
                if(borders.BottomColorHex !="auto")this.Stroke = borders.BottomColor.ToColor();
                if (borders.BottomSize.HasValue) this.Weight = borders.BottomSize.Value / 8.0;
                if (borders.BottomStyle.HasValue) this.SetLineType(borders.BottomStyle.Value);
            }
        }

        private WP.BorderValues GetBorderValue()
        {
            switch (this.lineType)
            {
                default:
                    return WP.BorderValues.Single;
                case LineTypes.DashDotStroked:
                    return WP.BorderValues.DashDotStroked;
                case LineTypes.Dashed:
                    return WP.BorderValues.Dashed;
                case LineTypes.DashSmallGap:
                    return WP.BorderValues.DashSmallGap;
                case LineTypes.DotDash:
                    return WP.BorderValues.DotDash;
                case LineTypes.DotDotDash:
                    return WP.BorderValues.DotDotDash;
                case LineTypes.Dotted:
                    return WP.BorderValues.Dotted;
                case LineTypes.Double:
                    return WP.BorderValues.Double;
                case LineTypes.DoubleWave:
                    return WP.BorderValues.DoubleWave;
                case LineTypes.Inset:
                    return WP.BorderValues.Inset;
                case LineTypes.Nil:
                    return WP.BorderValues.Nil;
                case LineTypes.None:
                    return WP.BorderValues.None;
                case LineTypes.Outset:
                    return WP.BorderValues.Outset;
                case LineTypes.ThickThinLargeGap:
                    return WP.BorderValues.ThickThinLargeGap;
                case LineTypes.ThickThinMediumGap:
                    return WP.BorderValues.ThickThinMediumGap;
                case LineTypes.ThickThinSmallGap:
                    return WP.BorderValues.ThickThinSmallGap;
                case LineTypes.ThinThickLargeGap:
                    return WP.BorderValues.ThinThickLargeGap;
                case LineTypes.ThinThickMediumGap:
                    return WP.BorderValues.ThinThickMediumGap;
                case LineTypes.ThinThickSmallGap:
                    return WP.BorderValues.ThinThickSmallGap;
                case LineTypes.ThinThickThinLargeGap:
                    return WP.BorderValues.ThinThickThinLargeGap;
                case LineTypes.ThinThickThinMediumGap:
                    return WP.BorderValues.ThinThickThinMediumGap;
                case LineTypes.ThinThickThinSmallGap:
                    return WP.BorderValues.ThinThickThinSmallGap;
                case LineTypes.ThreeDEmboss:
                    return WP.BorderValues.ThreeDEmboss;
                case LineTypes.ThreeDEngrave:
                    return WP.BorderValues.ThreeDEngrave;
                case LineTypes.Triple:
                    return WP.BorderValues.Triple;
                case LineTypes.Wave:
                    return WP.BorderValues.Wave;
            }
        }

        public void SetLineType(WP.BorderValues border)
        {
            this.lineType = LineTypes.Single;//Default
            if (border == WP.BorderValues.DashDotStroked) this.lineType = LineTypes.DashDotStroked;
            if (border == WP.BorderValues.Dashed) this.lineType = LineTypes.Dashed;
            if (border == WP.BorderValues.DashSmallGap) this.lineType = LineTypes.DashSmallGap;
            if (border == WP.BorderValues.DotDash) this.lineType = LineTypes.DotDash;
            if (border == WP.BorderValues.DotDotDash) this.lineType = LineTypes.DotDotDash;
            if (border == WP.BorderValues.Dotted) this.lineType = LineTypes.Dotted;
            if (border == WP.BorderValues.Double) this.lineType = LineTypes.Double;
            if (border == WP.BorderValues.DoubleWave) this.lineType = LineTypes.DoubleWave;
            if (border == WP.BorderValues.Inset) this.lineType = LineTypes.Inset;
            if (border == WP.BorderValues.Nil) this.lineType = LineTypes.Nil;
            if (border == WP.BorderValues.None) this.lineType = LineTypes.None;
            if (border == WP.BorderValues.Outset) this.lineType = LineTypes.Outset;
            if (border == WP.BorderValues.ThickThinLargeGap) this.lineType = LineTypes.ThickThinLargeGap;
            if (border == WP.BorderValues.ThickThinMediumGap) this.lineType = LineTypes.ThickThinMediumGap;
            if (border == WP.BorderValues.ThickThinSmallGap) this.lineType = LineTypes.ThickThinSmallGap;
            if (border == WP.BorderValues.ThinThickLargeGap) this.lineType = LineTypes.ThinThickLargeGap;
            if (border == WP.BorderValues.ThinThickMediumGap) this.lineType = LineTypes.ThinThickMediumGap;
            if (border == WP.BorderValues.ThinThickSmallGap) this.lineType = LineTypes.ThinThickSmallGap;
            if (border == WP.BorderValues.ThinThickThinLargeGap) this.lineType = LineTypes.ThinThickThinLargeGap;
            if (border == WP.BorderValues.ThinThickThinMediumGap) this.lineType = LineTypes.ThinThickThinMediumGap;
            if (border == WP.BorderValues.ThinThickThinSmallGap) this.lineType = LineTypes.ThinThickThinSmallGap;
            if (border == WP.BorderValues.ThreeDEmboss) this.lineType = LineTypes.ThreeDEmboss;
            if (border == WP.BorderValues.ThreeDEngrave) this.lineType = LineTypes.ThreeDEngrave;
            if (border == WP.BorderValues.Triple) this.lineType = LineTypes.Triple;
            if (border == WP.BorderValues.Wave) this.lineType = LineTypes.Wave;
        }

        #endregion

    }

    public static class Graphics
    {
        public static Graphic Outline { get { return new Graphic(Sd.Color.Black, 1); } }
        public static Graphic Solid { get { return new Graphic(Sd.Color.Black); } }
    }
}
