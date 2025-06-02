using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Sd = System.Drawing;
using WD = OfficeIMO.Word;
using DocumentFormat.OpenXml.Wordprocessing;
using WP = DocumentFormat.OpenXml.Wordprocessing;
using DP = DocumentFormat.OpenXml.Packaging;
using DC = DocumentFormat.OpenXml.Drawing.Charts;
using DF = DocumentFormat.OpenXml.Drawing;

namespace WordPlus
{
    public class WdDocument
    {

        #region members

        public enum Units { Pixels, Inches, Millimeters, Centimeters };
        public enum HeaderFooterTypes { None, All, Even, First };

        public HeaderFooterTypes HeaderType = HeaderFooterTypes.None;
        public HeaderFooterTypes FooterType = HeaderFooterTypes.None;

        public WD.WordDocument DocObject = WD.WordDocument.Create();
        protected WdContentBlock content = new WdContentBlock();
        public WdContentBlock Header = new WdContentBlock();
        public WdContentBlock Footer = new WdContentBlock();

        string name = "Unnamed";
        string title = string.Empty;
        string categories = string.Empty;
        string company = string.Empty;
        string creator = string.Empty;
        string description = string.Empty;
        string tags = string.Empty;
        string revision = string.Empty;
        string subject = string.Empty;
        string modified = string.Empty;
        string password = string.Empty;

        public Page Page = new Page();
        public Graphic graphic = new Graphic();

        #endregion

        #region constructors

        public WdDocument()
        {

        }

        public WdDocument(string name)
        {
            this.name = name;
        }

        public WdDocument(WdDocument document)
        {
            this.DocObject = document.DocObject;
            this.content = new WdContentBlock(document.content);
            this.Header = new WdContentBlock(document.Header);
            this.Footer = new WdContentBlock(document.Footer);

            this.HeaderType = document.HeaderType;
            this.FooterType = document.FooterType;

            this.name = document.name;
            this.title = document.title;
            this.categories = document.categories;
            this.company = document.company;
            this.creator = document.creator;
            this.description = document.description;
            this.tags = document.tags;
            this.revision = document.revision;
            this.subject = document.subject;
            this.modified = document.modified;
            this.password = document.password;

            this.Page = new Page(document.Page);
            this.graphic = new Graphic(document.graphic);
        }

        public WdDocument(List<WdDocument> documents, string name, bool pageBreak = false)
        {
            this.DocObject = WD.WordDocument.Create();

            for (int i = 0; i < documents.Count; i++)
            {
                this.content.AddRange(documents[i].GetContents());
                if (pageBreak) if (i < (documents.Count - 1)) this.content.Add(Content.CreatePageBreakContent());
            }

            int c = documents.Count - 1;
            this.Header = new WdContentBlock(documents[c].Header);
            this.Footer = new WdContentBlock(documents[c].Footer);

            this.HeaderType = documents[c].HeaderType;
            this.FooterType = documents[c].FooterType;

            this.name = name;
            this.Page = new Page(documents[c].Page);
        }

        public static WdDocument LoadDocument(string filePath)
        {
            WdDocument document = new WdDocument();
            WD.WordDocument doc = WD.WordDocument.Load(filePath,true, false);

            document.ParseDocument(doc);

            return document;
        }

        public static WdDocument ReadDocument(string streamText)
        {
            MemoryStream stream = new MemoryStream(Encoding.ASCII.GetBytes(streamText));
            stream.Position = 0;
            WdDocument document = new WdDocument();
            WD.WordDocument doc = WD.WordDocument.Load(stream, true, false);

            document.ParseDocument(doc);

            return document;
        }

        #endregion

        #region properties

        public string Title
        {
            get { return this.title; }
            set { this.title = value; }
        }

        public string Categories
        {
            get { return this.categories; }
            set { this.categories = value; }
        }

        public string Company
        {
            get { return this.company; }
            set { this.company = value; }
        }

        public string Creator
        {
            get { return this.creator; }
            set { this.creator = value; }
        }

        public string Description
        {
            get { return this.description; }
            set { this.description = value; }
        }

        public string Tags
        {
            get { return this.tags; }
            set { this.tags = value; }
        }

        public string Revision
        {
            get { return this.revision; }
            set { this.revision = value; }
        }

        public string Subject
        {
            get { return this.subject; }
            set { this.subject = value; }
        }

        public string Modified
        {
            get { return this.modified; }
            set { this.modified = value; }
        }

        public string Password
        {
            set { this.password = value; }
        }

        public Graphic Graphic
        {
            get { return this.graphic; }
            set { this.graphic = new Graphic(value); }
        }

        #endregion

        #region methods

        protected void ParseDocument(WD.WordDocument doc)
        {

            if (doc.PageSettings.Orientation == WP.PageOrientationValues.Landscape) this.Page.Orientation = Page.Orientations.Landscape;
            this.Page.PageWidth = (int)doc.PageSettings.Width.Value;
            this.Page.PageHeight = (int)doc.PageSettings.Height.Value;

            Paragraph mergedPara = new Paragraph();
            List<Paragraph> listItems = new List<Paragraph>();
            List<int> listIndices = new List<int>();
            int i = 0;
            int c = 0;
            int inset = 0;
            foreach (WD.WordElement element in doc.Elements)
            {
                if (element is WD.WordParagraph)
                {
                    WD.WordParagraph paragraph = (WD.WordParagraph)element;
                    if (paragraph.IsListItem)
                    {
                        Fragment frag = new Fragment(paragraph.Text, paragraph.GetFont());
                        if (paragraph.IsTab) frag.Tabs = 1;
                        if (paragraph.IsFirstRun) mergedPara = new Paragraph();

                        mergedPara.Fragments.Add(frag);
                        inset = Math.Max(paragraph.ListItemLevel.GetValueOrDefault(0), inset);
                        if (paragraph.IsLastRun)
                        {
                            listItems.Add(mergedPara);
                            listIndices.Add(inset);
                            inset = 0;
                        }

                        bool nextList = false;
                        if ((i + 2) > doc.Paragraphs.Count)
                        {
                            nextList = true;
                        }
                        else if (!doc.Paragraphs[i + 1].IsListItem)
                        {
                            nextList = true;
                        }
                        if (nextList)
                        {
                            this.AddContent(Content.CreateListContent(listItems, listIndices, Content.BulletPoints.Bulleted));
                            listItems = new List<Paragraph>();
                            listIndices = new List<int>();
                        }
                    }
                    if (paragraph.IsImage)
                    {
                        this.AddContent(this.ParseImage(paragraph, doc));
                    }
                    else if (paragraph.IsChart)
                    {
                        this.AddContent(this.ParseChart(paragraph, doc,c));
                        c++;
                    }
                    else if (paragraph.IsBreak)
                    {
                        this.AddContent(Content.CreateHorizontalLineContent());
                    }
                    else if (paragraph.IsPageBreak)
                    {
                        this.AddContent(Content.CreatePageBreakContent());
                    }
                    else if (!paragraph.IsListItem)
                    {
                        Fragment frag = new Fragment(paragraph.Text, paragraph.GetFont());
                        if (paragraph.IsTab) frag.Tabs = 1;
                        if (paragraph.IsFirstRun) mergedPara = new Paragraph();

                        mergedPara.Fragments.Add(frag);
                        if (paragraph.IsLastRun) this.AddContent(Content.CreateTextContent(mergedPara));
                    }
                }
                else if (element is WD.WordTable)
                {
                    this.AddContent(this.ParseTable((WD.WordTable)element, doc));
                }
                i++;
            }

        }

        protected Content ParseTable(WD.WordTable table, WD.WordDocument doc)
        {
            List<List<Content>> content = new List<List<Content>>();

            int r = table.RowsCount;
            int c = table.Rows[0].Cells.Count;

            for (int i = 0; i<c;i++)
            {
                content.Add(new List<Content>());
                for (int j = 0; j < r; j++)
                {
                    WD.WordTableCell cell = table.Rows[j].Cells[i];
                    Paragraph mergedPara = new Paragraph();
                    List<WD.WordParagraph> paragraphs = cell.Paragraphs;
                    List<Paragraph> listItems = new List<Paragraph>();
                    List<int> listIndices = new List<int>();
                    int inset = 0;

                    foreach (WD.WordParagraph paragraph in paragraphs)
                    {
                        if (paragraph.IsListItem)
                        {
                            Fragment frag = new Fragment(paragraph.Text, paragraph.GetFont());
                            if (paragraph.IsTab) frag.Tabs = 1;
                            if (paragraph.IsFirstRun) mergedPara = new Paragraph();

                            mergedPara.Fragments.Add(frag);
                            inset = Math.Max(paragraph.ListItemLevel.GetValueOrDefault(0), inset);
                            if (paragraph.IsLastRun)
                            {
                                listItems.Add(mergedPara);
                                listIndices.Add(inset);
                                inset = 0;
                            }

                            bool nextList = false;
                            if ((i + 2) > doc.Paragraphs.Count)
                            {
                                nextList = true;
                            }
                            else if (!doc.Paragraphs[i + 1].IsListItem)
                            {
                                nextList = true;
                            }
                            if (nextList)
                            {
                                this.AddContent(Content.CreateListContent(listItems, listIndices, Content.BulletPoints.Bulleted));
                                listItems = new List<Paragraph>();
                                listIndices = new List<int>();
                            }
                        }
                        if (paragraph.IsImage)
                        {
                            content[i].Add(this.ParseImage(paragraph, doc));
                        }
                        else if (paragraph.IsChart)
                        {
                            //content[i].Add(this.ParseChart(paragraph,doc,0));
                        }
                        else if(cell.HasNestedTables)
                        {
                            content[i].Add(this.ParseTable(cell.NestedTables[i],doc));
                        }
                        else if (!paragraph.IsListItem)
                        {
                            Fragment frag = new Fragment(paragraph.Text, paragraph.GetFont());
                            if (paragraph.IsTab) frag.Tabs = 1;
                            if (paragraph.IsFirstRun) mergedPara = new Paragraph();

                            mergedPara.Fragments.Add(frag);
                            if (paragraph.IsLastRun) content[i].Add(Content.CreateTextContent(mergedPara));
                        }
                        content[i][content[i].Count - 1].Graphic.SetGraphic(cell);
                    }
                }
            }

             return Content.CreateTableContent(content);

        }

        protected Content ParseChart(WD.WordParagraph paragraph, WD.WordDocument doc,int index)
        {

            List<List<double>> values = new List<List<double>>();
            List<string> axis = new List<string>();
            List<string> labels = new List<string>();
            List<Sd.Color> colors = new List<Sd.Color>();

            List<DP.ChartPart> chartParts = doc._wordprocessingDocument.MainDocumentPart.ChartParts.ToList();
            DP.ChartPart chartPart = chartParts[Math.Min(index,chartParts.Count-1)];
            DC.Chart chart = chartPart.ChartSpace.Elements<DC.Chart>().FirstOrDefault();

            //WP.Drawing drawing = paragraph..Descendants<Drawing>().FirstOrDefault();
            //var chartReference = drawing.Inline.Graphic.GraphicData.GetFirstChild<ChartReference>();
            //string relId = chartReference.Id;
            //DP.ChartPart chartPart = (DP.ChartPart)doc._wordprocessingDocument.MainDocumentPart.GetPartById(relId);
            //DC.Chart chart = chartPart.ChartSpace.Elements<DC.Chart>().FirstOrDefault();

            DC.Title title = chart.Elements<DC.Title>().ToList()[0];
            string titleText = "";
            if (title.InnerText != null) titleText = title.InnerText;

            int i = 0;

            if (chart.PlotArea.Elements<DC.BarChart>().FirstOrDefault() != null)
            {
                foreach (DC.BarChart chartObj in chart.PlotArea.Elements<DC.BarChart>().ToList())
                {
                    foreach (DC.BarChartSeries series in chartObj.Elements<DC.BarChartSeries>())
                    {
                        labels.Add(series.SeriesText.InnerText);
                        colors.Add(this.GetLegendColor(series.Elements<DC.ChartShapeProperties>().ToList()[0]));
                        if (i == 0) axis.AddRange(this.GetAxisLabels(series.GetFirstChild<DC.CategoryAxisData>()));

                        values.Add(this.GetSeriesNumbers(series.GetFirstChild<DC.Values>()));
                        i++;
                    }
                }
                return Content.CreateChartBarContent(titleText, values, axis, labels, colors, Content.LegendLocations.None);
            }

            if (chart.PlotArea.Elements<DC.LineChart>().FirstOrDefault() != null)
            {
                foreach (DC.LineChart chartObj in chart.PlotArea.Elements<DC.LineChart>().ToList())
            {
                foreach (DC.LineChartSeries series in chartObj.Elements<DC.LineChartSeries>())
                {
                    labels.Add(series.SeriesText.InnerText);
                    colors.Add(this.GetLegendColor(series.Elements<DC.ChartShapeProperties>().ToList()[0]));
                    if (i == 0) axis.AddRange(this.GetAxisLabels(series.GetFirstChild<DC.CategoryAxisData>()));

                    values.Add(this.GetSeriesNumbers(series.GetFirstChild<DC.Values>()));
                    i++;
                }
                }
                return Content.CreateChartLineContent(titleText, values, axis, labels, colors, Content.LegendLocations.None);
            }

            if (chart.PlotArea.Elements<DC.AreaChart>().FirstOrDefault() != null)
                {
                    foreach (DC.AreaChart chartObj in chart.PlotArea.Elements<DC.AreaChart>().ToList())
            {
                foreach (DC.AreaChartSeries series in chartObj.Elements<DC.AreaChartSeries>())
                {
                    labels.Add(series.SeriesText.InnerText);
                    colors.Add(this.GetLegendColor(series.Elements<DC.ChartShapeProperties>().ToList()[0]));
                    if (i == 0) axis.AddRange(this.GetAxisLabels(series.GetFirstChild<DC.CategoryAxisData>()));

                    values.Add(this.GetSeriesNumbers(series.GetFirstChild<DC.Values>()));
                    i++;
                }
                }
                return Content.CreateChartAreaContent(titleText, values, axis, labels, colors, Content.LegendLocations.None);
            }

            if (chart.PlotArea.Elements<DC.PieChart>().FirstOrDefault() != null)
                    {
                        foreach (DC.PieChart chartObj in chart.PlotArea.Elements<DC.PieChart>().ToList())
            {
                foreach (DC.PieChartSeries series in chartObj.Elements<DC.PieChartSeries>())
                {
                    labels.Add(series.SeriesText.InnerText);
                    //colors.Add(this.GetLegendColor(series.Elements<DC.ChartShapeProperties>().ToList()[0]));
                    if (i == 0) axis.AddRange(this.GetAxisLabels(series.GetFirstChild<DC.CategoryAxisData>()));

                    values.Add(this.GetSeriesNumbers(series.GetFirstChild<DC.Values>()));
                    i++;
                }
                }
                return Content.CreateChartPieContent(titleText, values[0], labels, colors, Content.LegendLocations.None);
            }

            return null;
        }

        protected List<double> GetSeriesNumbers(DC.Values valuesElement)
        {
            List<double> output = new List<double>();

            List<DC.NumericPoint> numbers = valuesElement.NumberLiteral.Elements<DC.NumericPoint>().ToList();
            foreach (DC.NumericPoint num in numbers)
            {
                Double.TryParse(num.NumericValue.Text, out double val);
                output.Add(val);
            }
            return output;
        }

        protected List<string> GetAxisLabels(DC.CategoryAxisData categoryAxisData)
        {
            List<string> output = new List<string>();
            List<DC.StringPoint> points = categoryAxisData.StringLiteral.Elements<DC.StringPoint>().ToList();
            foreach (DC.StringPoint pt in points) output.Add(pt.InnerText);
            return output;
        }

        protected Sd.Color GetLegendColor(DC.ChartShapeProperties shapeProp)
        {
            DF.SolidFill fill = shapeProp.Elements<DF.SolidFill>().ToList()[0];
            return Extension.ColorFromHex(fill.RgbColorModelHex.Val);
        }

        protected Content ParseImage(WD.WordParagraph paragraph, WD.WordDocument doc)
        {
            string id = paragraph.Image.RelationshipId;
            DP.MainDocumentPart mainPart = doc._wordprocessingDocument.MainDocumentPart;

            DP.ImagePart imagePart = (DP.ImagePart)mainPart.GetPartById(id);

            Stream memoryStream = imagePart.GetStream();
            memoryStream.Position = 0;
            Content output = Content.CreateImageContent(new Sd.Bitmap(memoryStream));
            memoryStream.Dispose();

            return output;
        }

        public void AddContent(Content content)
        {
            this.content.Add(content);
        }

        public void SetContent(List<Content> contents)
        {
            this.content.SetContent(contents);
        }

        public List<Content> GetContents()
        {
            return this.content.GetContents();
        }

        public void SetDocument()
        {
            this.DocObject = WD.WordDocument.Create();

            //Page Size
            this.DocObject.Sections[0].PageSettings.PageSize = WD.WordPageSize.A3;
            this.DocObject.Sections[0].PageSettings.Width = (uint)this.Page.PageWidth;
            this.DocObject.Sections[0].PageSettings.Height = (uint)this.Page.PageHeight;

            //Margin
            this.DocObject.Sections[0].SetMargins(WD.WordMargin.Normal);
            this.DocObject.Sections[0].Margins.Left = (uint)this.Page.MarginLeft;
            this.DocObject.Sections[0].Margins.Right = (uint)this.Page.MarginRight;
            this.DocObject.Sections[0].Margins.Top = (int)this.Page.MarginTop;
            this.DocObject.Sections[0].Margins.Bottom = (int)this.Page.MarginBottom;

            if (this.Page.Orientation == Page.Orientations.Landscape) this.DocObject.Sections[0].PageSettings.Orientation = DocumentFormat.OpenXml.Wordprocessing.PageOrientationValues.Landscape;
            if (this.Page.Columns > 1) this.DocObject.Sections[0].ColumnCount = this.Page.Columns;

            //Page Graphics
            if (this.Graphic.HasFill) this.DocObject.Background.SetColor(this.Graphic.Fill.ToSixLabors());
            this.RenderBorder(this.DocObject.Borders, this.Graphic);

            this.SetProperies();
            this.SetSecurity();
        }

        private void SetProperies()
        {

            if (this.title != string.Empty) this.DocObject.BuiltinDocumentProperties.Title = this.title;
            if (this.categories != string.Empty) this.DocObject.BuiltinDocumentProperties.Category = this.categories;
            if (this.modified != string.Empty) this.DocObject.ApplicationProperties.Company = this.company;
            if (this.company != string.Empty) this.DocObject.BuiltinDocumentProperties.LastModifiedBy = this.modified;
            if (this.creator != string.Empty) this.DocObject.BuiltinDocumentProperties.Creator = this.creator;
            if (this.description != string.Empty) this.DocObject.BuiltinDocumentProperties.Description = this.description;
            if (this.tags != string.Empty) this.DocObject.BuiltinDocumentProperties.Keywords = this.tags;
            if (this.revision != string.Empty) this.DocObject.BuiltinDocumentProperties.Revision = this.revision;
            if (this.subject != string.Empty) this.DocObject.BuiltinDocumentProperties.Subject = this.subject;

        }

        private void SetSecurity()
        {
            if (this.password != string.Empty)
            {
                this.DocObject.Settings.ProtectionPassword = this.password;
                this.DocObject.Settings.ProtectionType = WP.DocumentProtectionValues.ReadOnly;
            }
        }

        public void Render()
        {
            this.SetDocument();

            if ((this.Header.Count > 0) | (this.Footer.Count > 0))
            {
                this.DocObject.AddHeadersAndFooters();
            }
            if (this.Header.Count > 0)
            {
                switch (this.HeaderType)
                {
                    case HeaderFooterTypes.All:
                        Header.Render(this.DocObject.Header.Default);
                        break;
                    case HeaderFooterTypes.Even:
                        Header.Render(this.DocObject.Header.Even);
                        break;
                    case HeaderFooterTypes.First:
                        Header.Render(this.DocObject.Header.First);
                        break;
                }
            }
            if (this.Footer.Count > 0)
            {
                switch (this.FooterType)
                {
                    case HeaderFooterTypes.All:
                        Footer.Render(this.DocObject.Footer.Default);
                        break;
                    case HeaderFooterTypes.Even:
                        Footer.Render(this.DocObject.Footer.Even);
                        break;
                    case HeaderFooterTypes.First:
                        Footer.Render(this.DocObject.Footer.First);
                        break;
                }
            }

            this.content.Render(this);
        }

        protected void RenderBorder(WD.WordBorders border, Graphic graphic)
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

        public void Save(string filepath)
        {
            this.Render();
            this.DocObject.Save(filepath, false);
        }

        public string Stream()
        {
            this.Render();
            Stream stream = new MemoryStream();
            this.DocObject.Save(stream);
            stream.Position = 0;
            var buffer = new byte[stream.Length];
            stream.Read(buffer, 0, (int)stream.Length);
            string output = System.Text.Encoding.Default.GetString(buffer);
            stream.Dispose();
            return output;
        }

        #endregion

        #region overrides

        public override string ToString()
        {
            return "WD | Document: " + this.name + " {c:" + this.content.Count + "}";
        }

        #endregion

    }
}
