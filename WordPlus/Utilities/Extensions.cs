using Grasshopper.Kernel.Types;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Sd = System.Drawing;

using WD = OfficeIMO.Word;
using WP = DocumentFormat.OpenXml.Wordprocessing;

namespace WordPlus
{
    public static class Extension
    {

        #region conversions

        public static SixLabors.ImageSharp.Color ToSixLabors(this Sd.Color input)
        {
            return SixLabors.ImageSharp.Color.FromRgba(input.R, input.G, input.B, input.A);
        }

        public static WP.TableRowAlignmentValues ToRowAlignment(this Font.HorizontalAlignments input)
        {
            switch (input)
            {
                default:
                    return WP.TableRowAlignmentValues.Left;
                case Font.HorizontalAlignments.Right:
                    return WP.TableRowAlignmentValues.Right;
                case Font.HorizontalAlignments.Center:
                    return WP.TableRowAlignmentValues.Center;
            }
        }

        public static WD.WordListStyle ToImo(this WdContent.BulletPoints input)
        {
            switch (input)
            {
                default:
                    return WD.WordListStyle.Bulleted;
                case WdContent.BulletPoints.ArticleSections:
                    return WD.WordListStyle.ArticleSections;
                case WdContent.BulletPoints.BulletedChars:
                    return WD.WordListStyle.BulletedChars;
                case WdContent.BulletPoints.Chapters:
                    return WD.WordListStyle.Chapters;
                case WdContent.BulletPoints.Numbered_1ai:
                    return WD.WordListStyle.Heading1ai;
                case WdContent.BulletPoints.Numbered_111:
                    return WD.WordListStyle.Headings111;
                case WdContent.BulletPoints.Numbered_Tabbed_111:
                    return WD.WordListStyle.Headings111Shifted;
                case WdContent.BulletPoints.LetteredBracket:
                    return WD.WordListStyle.LowerLetterWithBracket;
                case WdContent.BulletPoints.Lettered:
                    return WD.WordListStyle.UpperLetterWithBracket;
                case WdContent.BulletPoints.LetteredDot:
                    return WD.WordListStyle.UpperLetterWithDot;
                case WdContent.BulletPoints.Lettered_IA1:
                    return WD.WordListStyle.HeadingIA1;
            }
        }

        public static DocumentFormat.OpenXml.Wordprocessing.JustificationValues ToOpenXml(this Font.HorizontalAlignments input)
        {
            switch (input)
            {
                default:
                    return DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Left;
                case Font.HorizontalAlignments.Center:
                    return DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Center;
                case Font.HorizontalAlignments.Right:
                    return DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Right;
            }
        }

        public static DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues ToOpenXml(this Font.HighlightColors input)
        {
            switch (input)
            {
                default:
                    return DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues.None;
                case Font.HighlightColors.Black:
                    return DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues.Black;
                case Font.HighlightColors.Blue:
                    return DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues.Blue;
                case Font.HighlightColors.Cyan:
                    return DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues.Cyan;
                case Font.HighlightColors.DarkBlue:
                    return DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues.DarkBlue;
                case Font.HighlightColors.DarkCyan:
                    return DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues.DarkCyan;
                case Font.HighlightColors.DarkGray:
                    return DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues.DarkGray;
                case Font.HighlightColors.DarkGreen:
                    return DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues.DarkGreen;
                case Font.HighlightColors.DarkMagenta:
                    return DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues.DarkMagenta;
                case Font.HighlightColors.DarkRed:
                    return DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues.DarkRed;
                case Font.HighlightColors.DarkYellow:
                    return DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues.DarkYellow;
                case Font.HighlightColors.Green:
                    return DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues.Green;
                case Font.HighlightColors.LightGray:
                    return DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues.LightGray;
                case Font.HighlightColors.Magenta:
                    return DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues.Magenta;
                case Font.HighlightColors.Red:
                    return DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues.Red;
                case Font.HighlightColors.White:
                    return DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues.White;
                case Font.HighlightColors.Yellow:
                    return DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues.Yellow;
            }
        }

        public static DocumentFormat.OpenXml.Drawing.Charts.LegendPositionValues ToOpenXml(this WdContent.LegendLocations input)
        {
            switch (input)
            {
                default:
                    return DocumentFormat.OpenXml.Drawing.Charts.LegendPositionValues.Bottom;
                case WdContent.LegendLocations.Left:
                    return DocumentFormat.OpenXml.Drawing.Charts.LegendPositionValues.Left;
                case WdContent.LegendLocations.Right:
                    return DocumentFormat.OpenXml.Drawing.Charts.LegendPositionValues.Right;
                case WdContent.LegendLocations.Top:
                    return DocumentFormat.OpenXml.Drawing.Charts.LegendPositionValues.Top;
            }
        }

        #endregion

        #region units

        public static DocumentFormat.OpenXml.UInt32Value ToXmlUint32(this double input)
        {
            return new DocumentFormat.OpenXml.UInt32Value((UInt32)(input * 8));
        }

        //Inches
        public static double TwipToInch(this double input)
        {
            return input / 1440.0;
        }

        public static double InchToTwip(this double input)
        {
            return input * 1440.0;
        }

        //Millimeters
        public static double TwipToMillimeter(this double input)
        {
            return input * 25.4 / 1440.0;
        }

        public static double MillimeterToTwip(this double input)
        {
            return input * 1440.0 / 25.4;
        }

        //Centimeter
        public static double TwipToCentimeter(this double input)
        {
            return input * 2.54 / 1440.0;
        }

        public static double CentimeterToTwip(this double input)
        {
            return input * 1440.0 / 2.54;
        }

        //Pixels
        public static double TwipToPixels(this double input, int dpi = 96)
        {
            return input / 1440.0 * dpi;
        }

        public static double PixelsToTwip(this double input, int dpi = 96)
        {
            return input * 1440.0 / dpi;
        }

        #endregion

        #region cloning

        public static List<WdContent> Duplicate(this List<WdContent> input)
        {
            List<WdContent> output = new List<WdContent>();
            foreach (WdContent item in input) output.Add(new WdContent(item));

            return output;
        }

        public static List<string> Duplicate(this List<string> input)
        {
            List<string> output = new List<string>();
            foreach (string item in input)
            {
                output.Add(item);
            }
            return output;
        }

        public static List<string> DuplicateAsText(this List<WdFragment> input)
        {
            List<string> output = new List<string>();
            foreach (WdFragment item in input)
            {
                output.Add(item.Text);
            }
            return output;
        }

        public static List<string> DuplicateAsText(this List<WdParagraph> input)
        {
            List<string> output = new List<string>();
            foreach (WdParagraph item in input)
            {
                output.Add(item.Text);
            }
            return output;
        }

        public static List<WdParagraph> Duplicate(this List<WdParagraph> input)
        {
            List<WdParagraph> output = new List<WdParagraph>();
            foreach (WdParagraph item in input)
            {
                output.Add(new WdParagraph(item));
            }
            return output;
        }

        public static List<WdFragment> Duplicate(this List<WdFragment> input)
        {
            List<WdFragment> output = new List<WdFragment>();
            foreach (WdFragment item in input)
            {
                output.Add(new WdFragment(item));
            }
            return output;
        }

        public static List<WdFragment> DuplicateAsFragments(this List<string> input)
        {
            List<WdFragment> output = new List<WdFragment>();
            foreach (string item in input)
            {
                output.Add(new WdFragment(item));
            }
            return output;
        }

        public static List<WdParagraph> DuplicateAsParagraphs(this List<string> input)
        {
            List<WdParagraph> output = new List<WdParagraph>();
            foreach (string item in input)
            {
                output.Add(new WdParagraph(item));
            }
            return output;
        }

        public static List<double> Duplicate(this List<double> input)
        {
            List<double> output = new List<double>();
            foreach (double item in input)
            {
                output.Add(item);
            }
            return output;
        }

        public static List<Sd.Color> Duplicate(this List<Sd.Color> input)
        {
            List<Sd.Color> output = new List<Sd.Color>();
            foreach (Sd.Color item in input)
            {
                output.Add(item);
            }
            return output;
        }

        public static List<List<string>> Duplicate(this List<List<string>> input)
        {
            List<List<string>> output = new List<List<string>>();
            foreach (List<string> items in input)
            {
                List<string> outputA = new List<string>();
                foreach (string item in items)
                {
                    outputA.Add(item);
                }
                output.Add(outputA);
            }
            return output;
        }

        public static List<List<WdFragment>> Duplicate(this List<List<WdFragment>> input)
        {
            List<List<WdFragment>> output = new List<List<WdFragment>>();
            foreach (List<WdFragment> items in input)
            {
                List<WdFragment> outputA = new List<WdFragment>();
                foreach (WdFragment item in items)
                {
                    outputA.Add(new WdFragment(item));
                }
                output.Add(outputA);
            }
           return output;
        }

        public static List<List<string>> DuplicateAsText(this List<List<WdFragment>> input)
        {
            List<List<string>> output = new List<List<string>>();
            foreach (List<WdFragment> items in input)
            {
                List<string> outputA = new List<string>();
                foreach (WdFragment item in items)
                {
                    outputA.Add(item.Text);
                }
                output.Add(outputA);
            }
            return output;
        } 

        public static List<List<WdFragment>> DuplicateAsFragments(this List<List<string>> input)
        {
            List<List<WdFragment>> output = new List<List<WdFragment>>();
            foreach (List<string> items in input)
            {
                List<WdFragment> outputA = new List<WdFragment>();
                foreach (string item in items)
                {
                    outputA.Add(new WdFragment(item));
                }
                output.Add(outputA);
            }
            return output;
        }

        public static List<List<WdParagraph>> Duplicate(this List<List<WdParagraph>> input)
        {
            List<List<WdParagraph>> output = new List<List<WdParagraph>>();
            foreach (List<WdParagraph> items in input)
            {
                List<WdParagraph> outputA = new List<WdParagraph>();
                foreach (WdParagraph item in items)
                {
                    outputA.Add(new WdParagraph(item));
                }
                output.Add(outputA);
            }
            return output;
        }

        public static List<List<string>> DuplicateAsText(this List<List<WdParagraph>> input)
        {
            List<List<string>> output = new List<List<string>>();
            foreach (List<WdParagraph> items in input)
            {
                List<string> outputA = new List<string>();
                foreach (WdParagraph item in items)
                {
                    outputA.Add(item.Text);
                }
                output.Add(outputA);
            }
            return output;
        }

        public static List<List<WdParagraph>> DuplicateAsParagraphs(this List<List<string>> input)
        {
            List<List<WdParagraph>> output = new List<List<WdParagraph>>();
            foreach (List<string> items in input)
            {
                List<WdParagraph> outputA = new List<WdParagraph>();
                foreach (string item in items)
                {
                    outputA.Add(new WdParagraph(item));
                }
                output.Add(outputA);
            }
            return output;
        }

        public static List<List<WdContent>> DuplicateAsContents(this List<List<string>> input)
        {
            List<List<WdContent>> output = new List<List<WdContent>>();
            foreach (List<string> items in input)
            {
                List<WdContent> outputA = new List<WdContent>();
                foreach (string item in items)
                {
                    outputA.Add(WdContent.CreateTextContent(item));
                }
                output.Add(outputA);
            }
            return output;
        }

        public static List<List<WdContent>> Duplicate(this List<List<WdContent>> input)
        {
            List<List<WdContent>> output = new List<List<WdContent>>();

            foreach(List<WdContent> contents in input)
            {
                List<WdContent> temp = new List<WdContent>();
                foreach(WdContent content in contents)
                {
                    temp.Add(new WdContent(content));
                }
                output.Add(temp);
            }
            return output;
        }

        public static List<List<double>> Duplicate(this List<List<double>> input)
        {
            List<List<double>> output = new List<List<double>>();
            foreach (List<double> items in input)
            {
                List<double> outputA = new List<double>();
                foreach (double item in items)
                {
                    outputA.Add(item);
                }
                output.Add(outputA);
            }
            return output;
        }

        public static List<List<Sd.Color>> Duplicate(this List<List<Sd.Color>> input)
        {
            List<List<Sd.Color>> output = new List<List<Sd.Color>>();
            foreach (List<Sd.Color> items in input)
            {
                List<Sd.Color> outputA = new List<Sd.Color>();
                foreach (Sd.Color item in items)
                {
                    outputA.Add(item);
                }
                output.Add(outputA);
            }
            return output;
        }

        #endregion

        #region casting

        public static bool TryCastToParagraph(this IGH_Goo input, out WdParagraph paragraph)
        {
            if (input.CastTo<WdParagraph>(out WdParagraph paragraph1))
            {
                paragraph = new WdParagraph(paragraph1);
                return true;
            }
            paragraph = null;
            return false;

        }
        public static bool TryGetParagraph(this IGH_Goo input, out WdParagraph paragraph)
        {
            if (input.CastTo<WdParagraph>(out WdParagraph paragraph1))
            {
                paragraph = new WdParagraph(paragraph1);
                return true;
            }
            else if (input.CastTo<WdFragment>(out WdFragment fragment1))
            {
                paragraph = new WdParagraph(fragment1);
                return true;
            }
            else if (input.CastTo<string>(out string text))
            {
                paragraph = new WdParagraph(text);
                return true;
            }
            else if (input.CastTo<double>(out double number))
            {
                paragraph = new WdParagraph(number.ToString());
                return true;
            }
            else if (input.CastTo<int>(out int integer))
            {
                paragraph = new WdParagraph(integer.ToString());
                return true;
            }
            paragraph = null;
            return false;
        }

        public static bool TryGetFragment(this IGH_Goo input, out WdFragment fragment)
        {
            if (input.CastTo<WdFragment>(out WdFragment fragment1))
            {
                fragment = new WdFragment(fragment1);
                return true;
            }
            else if (input.CastTo<string>(out string text))
            {
                fragment = new WdFragment(text);
                return true;
            }
            else if (input.CastTo<double>(out double number))
            {
                fragment = new WdFragment(number.ToString());
                return true;
            }
            else if (input.CastTo<int>(out int integer))
            {
                fragment = new WdFragment(integer.ToString());
                return true;
            }
            fragment = null;
            return false;
        }

        public static bool TryGetContent(this IGH_Goo input, out WdContent content)
        {
            if (input.CastTo<WdContent>(out WdContent content1))
            {
                content = new WdContent(content1);
                return true;
            }
            else if (input.CastTo<WdParagraph>(out WdParagraph paragraph))
            {
                content = WdContent.CreateTextContent(paragraph);
                return true;
            }
            else if (input.CastTo<WdFragment>(out WdFragment fragment1))
            {
                content = WdContent.CreateTextContent(fragment1);
                return true;
            }
            else if (input.CastTo<string>(out string text))
            {
                content = WdContent.CreateTextContent(text);
                return true;
            }
            else if (input.CastTo<double>(out double number))
            {
                content = WdContent.CreateTextContent(number.ToString());
                return true;
            }
            else if (input.CastTo<int>(out int integer))
            {
                content = WdContent.CreateTextContent(integer.ToString());
                return true;
            }
            content = null;
            return false;
        }

        #endregion

        #region strings

        public static string SplitAtCapitols(this string text, bool preserveAcronyms = false)
        {
            if (string.IsNullOrWhiteSpace(text))
                return string.Empty;
            StringBuilder newText = new StringBuilder(text.Length * 2);
            newText.Append(text[0]);
            for (int i = 1; i < text.Length; i++)
            {
                if (char.IsUpper(text[i]))
                    if ((text[i - 1] != ' ' && !char.IsUpper(text[i - 1])) ||
                        (preserveAcronyms && char.IsUpper(text[i - 1]) &&
                         i < text.Length - 1 && !char.IsUpper(text[i + 1])))
                        newText.Append(' ');
                newText.Append(text[i]);
            }
            return newText.ToString();
        }

        #endregion

    }
}
