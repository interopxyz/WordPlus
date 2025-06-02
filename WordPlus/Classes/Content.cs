using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using WD = OfficeIMO.Word;
using WP = DocumentFormat.OpenXml.Wordprocessing;

using Sd = System.Drawing;
using System.IO;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace WordPlus
{
    public class Content
    {

        #region members

        public enum ContentTypes { None, PageBreak, LineBreak, Text, Image, List, Table, ChartArea, ChartBar, ChartColumn, ChartLine, ChartPie };
        public enum BulletPoints { Bulleted, BulletedChars, Numbered_1ai, Numbered_111, Numbered_Tabbed_111, Lettered, LetteredDot, LetteredBracket, Lettered_IA1, ArticleSections, Chapters };
        
        public enum LegendLocations { None, Left, Right, Top, Bottom };

        protected ContentTypes type = ContentTypes.None;
        protected LegendLocations legendLocation = LegendLocations.None;

        protected Font font = new Font();
        protected Graphic graphic = new Graphic();
        protected double width = 0;
        protected double height = 0;

        protected Paragraph text = new Paragraph();
        protected List<List<Paragraph>> values = new List<List<Paragraph>>();
        protected List<List<double>> numbers = new List<List<double>>();
        protected List<List<Sd.Color>> colors = new List<List<Sd.Color>>();
        protected List<List<Content>> contents = new List<List<Content>>();
        protected Sd.Bitmap image = new Sd.Bitmap(10, 10);

        BulletPoints bulletStyle = BulletPoints.Bulleted;

        #endregion

        #region constructors

        public Content()
        {

        }

        public Content(Content content)
        {
            this.type = content.type;
            this.legendLocation = content.legendLocation;

            this.font = new Font(content.font);
            this.graphic = new Graphic(content.graphic);

            this.width = content.width;
            this.height = content.height;

            this.text = new Paragraph(content.text);
            this.values = content.values.Duplicate();
            this.numbers = content.numbers.Duplicate();
            this.colors = content.colors.Duplicate();
            this.contents = content.contents.Duplicate();

            this.image = new Sd.Bitmap(content.image);

            this.bulletStyle = content.bulletStyle;
        }

        //LINE BREAK
        public static Content CreateHorizontalLineContent()
        {
            Content content = new Content();
            content.type = ContentTypes.LineBreak;

            return content;
        }

        //PAGE BREAK
        public static Content CreatePageBreakContent()
        {
            Content content = new Content();
            content.type = ContentTypes.PageBreak;

            return content;
        }

        //TEXT
        public static Content CreateTextContent(string text)
        {
            return Content.CreateTextContent(new Paragraph(text));
        }

        public static Content CreateTextContent(Fragment fragment)
        {
            return Content.CreateTextContent(new Paragraph(fragment));
        }

        public static Content CreateTextContent(Paragraph paragraph)
        {
            Content content = new Content();
            content.type = ContentTypes.Text;

            content.text = new Paragraph(paragraph);
            content.graphic = new Graphic(paragraph.Fragments[0].Graphic);

            return content;
        }

        //LIST
        public static Content CreateListContent(List<string> values, List<int> indices, BulletPoints style)
        {
            return Content.CreateListContent(values.DuplicateAsFragments(), indices, style); ;
        }

        public static Content CreateListContent(List<Paragraph> values, List<int> indices, BulletPoints style)
        {
            Content content = new Content();
            content.type = ContentTypes.List;

            content.bulletStyle = style;
            content.values.Add(new List<Paragraph>());
            foreach (Paragraph text in values) content.values[0].Add(new Paragraph(text));

            if (indices.Count == 0) indices.Add(0);
            int count = indices.Count;
            for (int i = count; i < values.Count; i++) indices.Add(count - 1);
            content.numbers.Add(new List<double>());
            foreach (int i in indices) content.numbers[0].Add(i);

            return content;
        }

        public static Content CreateListContent(List<Fragment> values, List<int> indices, BulletPoints style)
        {
            Content content = new Content();
            content.type = ContentTypes.List;

            content.bulletStyle = style;
            content.values.Add(new List<Paragraph>());
            foreach (Fragment text in values) content.values[0].Add(new Paragraph(text));

            if (indices.Count == 0) indices.Add(0);
            int count = indices.Count;
            for (int i = count; i < values.Count; i++) indices.Add(indices[i]);
            content.numbers.Add(new List<double>());
            foreach (int i in indices) content.numbers[0].Add(i);

            return content;
        }

        //TABLE
        public static Content CreateTableContent(List<List<string>> values)
        {
            return Content.CreateTableContent(values.DuplicateAsContents());
        }

        public static Content CreateTableContent(List<List<string>> values, Paragraph title)
        {
            return Content.CreateTableContent(values.DuplicateAsParagraphs(), title);
        }

        public static Content CreateTableContent(List<List<Content>> values)
        {
            Content content = new Content();
            content.type = ContentTypes.Table;

            content.contents = values.Duplicate();

            return content;
        }

        public static Content CreateTableContent(List<List<Paragraph>> values, Paragraph title)
        {
            Content content = new Content();
            content.type = ContentTypes.Table;

            content.text = new Paragraph(title);
            content.values = values.Duplicate();

            return content;
        }

        //IMAGE
        public static Content CreateImageContent(Sd.Bitmap image)
        {
            Content content = new Content();
            content.type = ContentTypes.Image;

            content.image = new Sd.Bitmap(image);

            return content;
        }

        //BAR CHART
        public static Content CreateChartBarContent(string title, List<List<double>> values, List<string> axis, List<string> labels, List<Sd.Color> colors, LegendLocations legend)
        {
            return Content.CreateChartBarContent(new Paragraph(title), values, axis.DuplicateAsParagraphs(), labels.DuplicateAsParagraphs(), colors, legend);
        }

        public static Content CreateChartBarContent(Paragraph title, List<List<double>> values, List<Paragraph> axis, List<Paragraph> labels, List<Sd.Color> colors, LegendLocations legend)
        {
            Content content = new Content();
            content.type = ContentTypes.ChartBar;

            content.text = title;
            content.PopulateChart(values, axis, labels, colors);
            content.legendLocation = legend;

            return content;
        }

        //COLUMN CHART
        public static Content CreateChartColumnContent(string title, List<List<double>> values, List<string> axis, List<string> labels, List<Sd.Color> colors, LegendLocations legend)
        {
            return Content.CreateChartColumnContent(new Paragraph(title), values, axis.DuplicateAsParagraphs(), labels.DuplicateAsParagraphs(), colors, legend);
        }

        public static Content CreateChartColumnContent(Paragraph title, List<List<double>> values, List<Paragraph> axis, List<Paragraph> labels, List<Sd.Color> colors, LegendLocations legend)
        {
            Content content = new Content();
            content.type = ContentTypes.ChartColumn;

            content.text = title;
            content.PopulateChart(values, axis, labels, colors);
            content.legendLocation = legend;

            return content;
        }

        //LINE CHART
        public static Content CreateChartLineContent(string title, List<List<double>> values, List<string> axis, List<string> labels, List<Sd.Color> colors, LegendLocations legend)
        {
            return Content.CreateChartLineContent(new Paragraph(title), values, axis.DuplicateAsParagraphs(), labels.DuplicateAsParagraphs(), colors, legend);
        }

        public static Content CreateChartLineContent(Paragraph title, List<List<double>> values, List<Paragraph> axis, List<Paragraph> labels, List<Sd.Color> colors, LegendLocations legend)
        {
            Content content = new Content();
            content.type = ContentTypes.ChartLine;

            content.text = title;
            content.PopulateChart(values, axis, labels, colors);
            content.legendLocation = legend;

            return content;
        }

        //AREA CHART
        public static Content CreateChartAreaContent(string title, List<List<double>> values, List<string> axis, List<string> labels, List<Sd.Color> colors, LegendLocations legend)
        {
            return Content.CreateChartAreaContent(new Paragraph(title), values, axis.DuplicateAsParagraphs(), labels.DuplicateAsParagraphs(), colors, legend);
        }

        public static Content CreateChartAreaContent(Paragraph title, List<List<double>> values, List<Paragraph> axis, List<Paragraph> labels, List<Sd.Color> colors, LegendLocations legend)
        {
            Content content = new Content();
            content.type = ContentTypes.ChartArea;

            content.text = title;
            content.PopulateChart(values, axis, labels, colors);
            content.legendLocation = legend;

            return content;
        }

        //PIE CHART
        public static Content CreateChartPieContent(string title, List<double> values, List<string> labels, List<Sd.Color> colors, LegendLocations legend)
        {
            return Content.CreateChartPieContent(new Paragraph(title), values, labels.DuplicateAsParagraphs(), colors, legend);
        }

        public static Content CreateChartPieContent(Paragraph title, List<double> values, List<Paragraph> labels, List<Sd.Color> colors, LegendLocations legend)
        {
            Content content = new Content();
            content.type = ContentTypes.ChartPie;

            content.text = title;
            content.PopulateChart(values, labels, colors);
            content.legendLocation = legend;

            return content;
        }

        #endregion

        #region properties

        public virtual Font Font
        {
            get { return this.font; }
            set { this.font = new Font(value); }
        }

        public virtual Graphic Graphic
        {
            get { return this.graphic; }
            set { this.graphic = new Graphic(value); }
        }

        public virtual double Width
        {
            get { return this.width; }
            set { this.width = value; }
        }

        public virtual double Height
        {
            get { return this.height; }
            set { this.height = value; }
        }

        public virtual Paragraph Text
        {
            get { return this.text; }
            set { this.text = new Paragraph(value); }
        }

        public virtual List<Paragraph> TextList
        {
            get
            {
                List<Paragraph> output = new List<Paragraph>();
                foreach (Paragraph paragraph in values[0]) output.Add(new Paragraph(paragraph));
                return output;
            }
        }

        public virtual List<List<double>> Numbers
        {
            get { return this.numbers.Duplicate(); }
        }

        public virtual List<List<Paragraph>> Values
        {
            get { return this.values.Duplicate(); }
        }

        public virtual List<List<Content>> Contents
        {
            get { return this.contents.Duplicate(); }
        }

        public virtual List<List<Sd.Color>> Colors
        {
            get { return this.colors.Duplicate(); }
        }

        public virtual Sd.Bitmap Image
        {
            get { return new Sd.Bitmap(this.image); }
        }

        public virtual ContentTypes ContentType
        {
            get { return this.type; }
        }

        #endregion

        #region methods

        private void PopulateChart(List<List<double>> values, List<string> axis, List<string> labels, List<Sd.Color> colors)
        {
            this.PopulateChart(values, axis.DuplicateAsParagraphs(), labels.DuplicateAsParagraphs(), colors);
        }

        private void PopulateChart(List<List<double>> values, List<Paragraph> axis, List<Paragraph> labels, List<Sd.Color> colors)
        {
            int max = values[0].Count;
            int count = axis.Count;
            List<Paragraph> axisA = axis.Duplicate();
            for (int i = count; i < max; i++) axisA.Add(new Paragraph("Item " + (i + 1)));

            max = values.Count;
            count = labels.Count;
            List<Paragraph> labelsA = labels.Duplicate();
            for (int i = count; i < max; i++) labelsA.Add(new Paragraph("Series " + (i + 1)));

            count = colors.Count;

            Sd.KnownColor[] names = (Sd.KnownColor[])Enum.GetValues(typeof(Sd.KnownColor));
            Random randomGen = new Random();
            for (int i = count; i < max; i++)
            {
                Sd.KnownColor randomColorName = names[randomGen.Next(names.Length)];
                Sd.Color randomColor = Sd.Color.FromKnownColor(randomColorName);
                colors.Add(randomColor);
            }

            this.numbers = values.Duplicate();
            this.values.Add(axisA.Duplicate());
            this.values.Add(labelsA.Duplicate());
            this.colors.Add(colors.Duplicate());
        }

        private void PopulateChart(List<double> values, List<string> labels, List<Sd.Color> colors)
        {
            this.PopulateChart(values, labels.DuplicateAsParagraphs(), colors);
        }

        private void PopulateChart(List<double> values, List<Paragraph> labels, List<Sd.Color> colors)
        {
            int max = values.Count;

            int count = labels.Count;
            List<Paragraph> labelsA = labels.Duplicate();
            for (int i = count; i < max; i++) labelsA.Add(new Paragraph("Series " + (i + 1)));

            count = colors.Count;

            Sd.KnownColor[] names = (Sd.KnownColor[])Enum.GetValues(typeof(Sd.KnownColor));
            for (int i = count; i < max; i++)
            {
                Random randomGen = new Random();
                Sd.KnownColor randomColorName = names[randomGen.Next(names.Length)];
                Sd.Color randomColor = Sd.Color.FromKnownColor(randomColorName);
                colors.Add(randomColor);
            }

            this.numbers.Add(values.Duplicate());
            this.values.Add(labelsA.Duplicate());
            this.colors.Add(colors.Duplicate());
        }

        public void Render(WD.WordParagraph paragraph, double w, double h, bool isParent = true)
        {
            double pixelWidth = this.width.TwipToPixels();
            double pixelHeight = this.height.TwipToPixels();

            WD.WordChart chart = null;
            switch (this.type)
            {
                case ContentTypes.Image:
                case ContentTypes.Table:
                    paragraph.ParagraphAlignment = this.font.HorizontalAlignment.ToOpenXml();
                    break;
                case ContentTypes.ChartArea:
                case ContentTypes.ChartBar:
                case ContentTypes.ChartColumn:
                case ContentTypes.ChartLine:
                case ContentTypes.ChartPie:
                    paragraph.ParagraphAlignment = this.font.HorizontalAlignment.ToOpenXml();

                    double W = pixelWidth;
                    double H = pixelHeight;
                    if (pixelWidth <= 0) W = w;
                    if (pixelHeight <= 0) H = W / 2;

                    chart = paragraph.AddChart(this.text.Text, false, (int)W, (int)H);

                    paragraph.Remove();

                    break;
            }

            switch (this.type)
            {
                case ContentTypes.PageBreak:
                    paragraph.AddBreak(WP.BreakValues.Page);
                    break;
                case ContentTypes.LineBreak:
                    Sd.Color lineColor = Sd.Color.Black;
                    if (this.graphic.HasStroke) lineColor = this.graphic.Stroke;
                    paragraph.AddHorizontalLine(this.graphic.BorderValue, lineColor.ToSixLabors(), graphic.Weight.ToXmlUint32(),0);
                    break;
                case ContentTypes.Text:
                    
                    foreach (Fragment txt in this.text.Fragments) this.RenderText(this.AddText(paragraph, txt), txt);

                    paragraph.LineSpacing = (int)this.text.LineSpacing*240;
                    if (this.text.HasIndentBefore) paragraph.IndentationBefore = this.text.IndentBefore;
                    if (this.text.HasIndentAfter) paragraph.IndentationAfter = this.text.IndentAfter;
                    if (this.text.HasIndentFirstLine) paragraph.IndentationFirstLine = this.text.IndentFirstLine;
                    if (this.text.HasIndentHanging) paragraph.IndentationHanging = this.text.IndentHanging;

                    if (isParent) { 
                    this.RenderBorder(paragraph.Borders, this.text.Graphic);
                    paragraph.SetColor(SixLabors.ImageSharp.Color.Aquamarine);
                    }
                    break;
                case ContentTypes.List:
                    
                    WD.WordList wordList = paragraph.AddList(this.bulletStyle.ToImo());
                    
                    for (int i = 0; i < this.values[0].Count; i++)
                    {
                        WD.WordParagraph listLine = wordList.AddItem("", (int)this.numbers[0][i]);
                        foreach (Fragment fragment1 in this.values[0][i].Fragments) this.RenderText(this.AddText(listLine, fragment1), fragment1);
                    }
                    paragraph.Remove();
                    break;
                case ContentTypes.Table:
                    int r = 0;
                    if (this.text.Fragments.Count > 0) r = 1;
                    int numCols = this.contents.Count;
                    int numRows = this.contents[0].Count + r;
                    WD.WordTable table = paragraph.AddTableBefore(numRows, numCols);

                    paragraph.SetColor(this.graphic.Fill.ToSixLabors());

                    if (r == 1)
                    {
                        WD.WordTableCell titleCell = table.Rows[0].Cells[0];
                        titleCell.MergeHorizontally(numCols);
                        titleCell = table.Rows[0].Cells[0];
                        WD.WordParagraph titleParagraph = titleCell.Paragraphs[0];
                        foreach (Fragment fragment1 in this.text.Fragments) this.RenderText(this.AddText(titleParagraph, fragment1), fragment1);
                        titleParagraph.ParagraphAlignment = WP.JustificationValues.Center;
                        titleCell.VerticalAlignment = WP.TableVerticalAlignmentValues.Center;
                    }
                    double cellHeight = 20;
                    for (int i = r; i < numRows; i++)
                    {
                        for (int j = 0; j < numCols; j++)
                        {
                            WD.WordTableCell cell = table.Rows[i].Cells[j];
                            Content cellContent = this.contents[j][i - r];

                            this.RenderBorder(cell.Borders, cellContent.Graphic);

                            if (cellContent.Graphic.HasFill) cell.ShadingFillColor = cellContent.Graphic.Fill.ToSixLabors();
                            if (table.Rows[i].Height != null) cellHeight = (double)table.Rows[i].Height;
                            this.contents[j][i-r].Render(cell.Paragraphs[0],(double)cell.Width, cellHeight,false);
                        }
                    }

                    if (this.width > 0)
                    {
                        table.LayoutMode = WD.WordTableLayoutType.FixedWidth;
                        table.WidthType = WP.TableWidthUnitValues.Dxa;
                        table.SetFixedWidth((int)(pixelWidth/w*100.0));
                    }

                    this.RenderTableBorder(table, this.graphic);

                    table.Alignment = this.font.HorizontalAlignment.ToRowAlignment();

                    paragraph.Remove();

                    break;
                case ContentTypes.Image:
                    double ratio = w / this.image.Width;
                    ratio = Math.Min(ratio, (double)(h / this.image.Height));

                    double W = this.image.Width * ratio;
                    if (pixelWidth > 0) W = pixelWidth;

                    double H = this.image.Height * ratio;
                    if (pixelHeight > 0) H = pixelHeight;

                    if ((pixelWidth <= 0) & (pixelHeight > 0)) W = W * (H / h);
                    if ((pixelHeight <= 0) & (pixelWidth > 0)) H = H * (W / w);

                    Stream memoryStream = new MemoryStream();
                    this.image.Save(memoryStream, Sd.Imaging.ImageFormat.Png);
                    memoryStream.Position = 0; // Reset the stream position to the beginning

                    paragraph.AddImage(memoryStream, "Name", W, H);

                    memoryStream.Dispose();
                    break;
                case ContentTypes.ChartBar:
                    chart.AddCategories(this.values[0].DuplicateAsText());
                    for (int i = 0; i < this.values[1].Count; i++)
                    {
                        chart.BarDirection = BarDirectionValues.Bar;
                        chart.AddBar(this.values[1][i].Text, this.numbers[i], this.colors[0][i].ToSixLabors());
                    }
                    break;
                case ContentTypes.ChartColumn:
                    chart.AddCategories(this.values[0].DuplicateAsText());
                    for (int i = 0; i < this.values[1].Count; i++)
                    {
                        chart.BarDirection = BarDirectionValues.Column;
                        chart.AddBar(this.values[1][i].Text, this.numbers[i], this.colors[0][i].ToSixLabors());

                    }
                    break;
                case ContentTypes.ChartArea:
                    chart.AddChartAxisX(this.values[0].DuplicateAsText());
                    for (int i = 0; i < this.values[1].Count; i++)
                    {
                        chart.AddArea(this.values[1][i].Text, this.numbers[i], this.colors[0][i].ToSixLabors());
                    }
                    break;
                case ContentTypes.ChartLine:
                    chart.AddChartAxisX(this.values[0].DuplicateAsText());
                    for (int i = 0; i < this.values[1].Count; i++)
                    {
                        chart.AddLine(this.values[1][i].Text, this.numbers[i], this.colors[0][i].ToSixLabors());
                    }
                    break;
                case ContentTypes.ChartPie:
                    chart.AddCategories(this.values[0].DuplicateAsText());
                    for (int i = 0; i < this.numbers[0].Count; i++)
                    {
                        chart.AddPie(this.values[0][i].Text, this.numbers[0][i]);
                    }
                    break;
            }
            switch (this.type)
            {
                case ContentTypes.ChartArea:
                case ContentTypes.ChartBar:
                case ContentTypes.ChartColumn:
                case ContentTypes.ChartLine:
                case ContentTypes.ChartPie:
                    if (this.legendLocation != LegendLocations.None) chart.AddLegend(this.legendLocation.ToOpenXml());
                    break;
            }
        }

        protected WD.WordParagraph AddText(WD.WordParagraph paragraph, Fragment fragment)
        {
            string text = fragment.Text;
            //if (fragment.HasTabs) paragraph.AddTabStop(0, WP.TabStopValues.Left);
            if (fragment.HasTabs) paragraph.AddTab();
            //if(fragment.HasTabs)text = "\t " + text;

            if (fragment.HasLink) return paragraph.AddHyperLink(fragment.Text, new Uri(fragment.Hyperlink), true);
            return paragraph.AddText(text);
        }

        protected void RenderText(WD.WordParagraph paragraph, Fragment fragment)
        {
            paragraph.SetStyle(font.ToStyle());
            if (fragment.Font.HasFamily) paragraph.SetFontFamily(fragment.Font.Family);
            if (fragment.Font.HasSize) paragraph.SetFontSize((int)fragment.Font.Size);
            if (fragment.Font.HasColor) paragraph.SetColor(fragment.Font.Color.ToSixLabors());
            if (fragment.Font.HasHighlight) paragraph.SetHighlight(fragment.Font.Highlight.ToOpenXml());
            paragraph.SetBold(fragment.Font.Bold);
            paragraph.SetItalic(fragment.Font.Italic);
            if (fragment.Font.Underlined) paragraph.SetUnderline(WP.UnderlineValues.Single);
            paragraph.SetStrike(fragment.Font.Strikethrough);
            if (fragment.Font.Superscript) paragraph.SetSuperScript();
            if (fragment.Font.Subscript) paragraph.SetSubScript();
            if (fragment.Font.HorizontalAlignment != Font.HorizontalAlignments.Default) paragraph.SetAlignment(fragment.Font.HorizontalAlignment.ToOpenXml());
        }

        protected void RenderTableBorder(WD.WordTable table,Graphic graphic)
        {

            if (this.graphic.TopBorder)
            {
                foreach (WD.WordTableCell cell in table.FirstRow.Cells)
                {
                    cell.Borders.TopStyle = graphic.BorderValue;
                    cell.Borders.TopColor = graphic.Stroke.ToSixLabors();
                    cell.Borders.TopSize = graphic.Weight.ToXmlUint32();
                }
            }

            if (this.graphic.BottomBorder)
            {
                foreach (WD.WordTableCell cell in table.LastRow.Cells)
                {
                    cell.Borders.BottomStyle = graphic.BorderValue;
                    cell.Borders.BottomColor = graphic.Stroke.ToSixLabors();
                    cell.Borders.BottomSize = graphic.Weight.ToXmlUint32();
                }
            }

            if (this.graphic.LeftBorder)
            {
                foreach (WD.WordTableRow row in table.Rows)
                {
                    row.FirstCell.Borders.LeftStyle = graphic.BorderValue;
                    row.FirstCell.Borders.LeftColor = graphic.Stroke.ToSixLabors();
                    row.FirstCell.Borders.LeftSize = graphic.Weight.ToXmlUint32();
                }
            }

            if (this.graphic.RightBorder)
            {
                foreach (WD.WordTableRow row in table.Rows)
                {
                    row.LastCell.Borders.RightStyle = graphic.BorderValue;
                    row.LastCell.Borders.RightColor = graphic.Stroke.ToSixLabors();
                    row.LastCell.Borders.RightSize= graphic.Weight.ToXmlUint32();
                }
            }
        }

        //Render Borders on Table Cells
        protected void RenderBorder(WD.WordTableCellBorder border, Graphic graphic)
        {
            if (graphic.TopBorder)
            {
                border.TopStyle = graphic.BorderValue;
                border.TopColor = graphic.Stroke.ToSixLabors();
                border.TopSize = graphic.Weight.ToXmlUint32();
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

        //Render Borders on Paragraphs
        protected void RenderBorder(WD.WordParagraphBorders border, Graphic graphic)
        {
            if (graphic.TopBorder)
            {
                border.TopStyle = graphic.BorderValue;
                border.TopColor = graphic.Stroke.ToSixLabors();
                border.TopSize = graphic.Weight.ToXmlUint32();
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


        #endregion

        #region overrides

        public override string ToString()
        {
            switch (this.type)
            {
                default:
                    if (this.text.Text.Length < 16) return "WD | " + this.type + " {" + this.text.Text + "}";
                    return "WD | " + this.type + " {" + this.text.Text.Substring(0, 15) + "...}";
                case ContentTypes.List:
                    return "WD | List {" + this.values[0].Count + "i}";
                case ContentTypes.Image:
                    return "WD | " + this.type + " {" + this.image.Width + "w " + this.image.Height + "h}";
                case ContentTypes.Table:
                    return "WD | " + this.type + " {" + this.contents.Count + "c " + this.contents[0].Count + "r}";
                case ContentTypes.PageBreak:
                    return "WD | Page Break";
                case ContentTypes.LineBreak:
                    return "WD | Line Break";
                case ContentTypes.ChartArea:
                case ContentTypes.ChartBar:
                case ContentTypes.ChartColumn:
                case ContentTypes.ChartLine:
                    return "WD | " + this.type + " {" + this.numbers.Count + "c " + this.numbers[0].Count + "v}";
                case ContentTypes.ChartPie:
                    return "WD | " + this.type + " {" + this.numbers[0].Count + "v}";
            }
        }

        #endregion

    }
}
