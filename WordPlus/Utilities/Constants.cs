using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Sd = System.Drawing;

namespace WordPlus
{
    public class Constants
    {

        #region naming

        public static string UniqueName
        {
            get { return ShortName + "_" + DateTime.UtcNow.ToString("yyyy-dd-M_HH-mm-ss"); }
        }

        public static string LongName
        {
            get { return ShortName + " v" + Major + "." + Minor; }
        }

        public static string ShortName
        {
            get { return "WordPlus"; }
        }

        private static string Minor
        {
            get { return typeof(Constants).Assembly.GetName().Version.Minor.ToString(); }
        }
        private static string Major
        {
            get { return typeof(Constants).Assembly.GetName().Version.Major.ToString(); }
        }

        public static string SubApp
        {
            get { return "App"; }
        }

        public static string SubDocument
        {
            get { return "Document"; }
        }

        public static string SubContent
        {
            get { return "Content"; }
        }

        public static string Format
        {
            get { return "Text Format"; }
        }

        public static string Modify
        {
            get { return "Modify"; }
        }

        #endregion

        #region inputs / outputs

        public static Descriptor App
        {
            get { return new Descriptor("Word Application", "App", "A Document or Word Application", "Document or Word Applications", "A Word Application Object", "Word Application Objects"); }
        }

        public static Descriptor Document
        {
            get { return new Descriptor("Document", "Doc", "A Word Document", "Word Documents", "A Word Document Object", "Word Document Objects"); }
        }

        public static Descriptor Content
        {
            get { return new Descriptor("Content", "Con", "A Word Document Content Object", "Word Document Content Objects", "A Word Document Content Object", "Word Document Content Objects"); }
        }

        public static Descriptor Paragraph
        {
            get { return new Descriptor("Word Paragraph", "Par", "A Text Paragraph Object, Text Fragment Object, or String", "Text Paragraph Objects, Text Fragment Objects, or Strings", "A Text Object", "Text Objects"); }
        }

        public static Descriptor Fragment
        {
            get { return new Descriptor("Word Text", "Txt", "A Text Fragment Object, or String", "Text Fragment Objects, or Strings", "A Text Fragment Object", "Text Fragment Objects"); }
        }

        public static Descriptor Activate
        {
            get { return new Descriptor("Activate", "_A", "If true, the component will be activated", "If true, the component will be activated", "The active status of the component", "The active statuses of the component"); }
        }

        #endregion

        #region color

        public static Sd.Color StartColor
        {
            get { return Sd.Color.FromArgb(99, 190, 123); }
        }

        public static Sd.Color MidColor
        {
            get { return Sd.Color.FromArgb(255, 235, 132); }
        }

        public static Sd.Color EndColor
        {
            get { return Sd.Color.FromArgb(248, 105, 107); }
        }

        #endregion
    }

    public class Descriptor
    {
        private string name = string.Empty;
        private string nickname = string.Empty;
        private string input = string.Empty;
        private string inputs = string.Empty;
        private string output = string.Empty;
        private string outputs = string.Empty;

        public Descriptor(string name, string nickname, string input, string inputs, string output, string outputs)
        {
            this.name = name;
            this.nickname = nickname;
            this.input = input;
            this.inputs = inputs;
            this.output = output;
            this.outputs = outputs;
        }

        public virtual string Name
        {
            get { return name; }
        }

        public virtual string NickName
        {
            get { return nickname; }
        }

        public virtual string Input
        {
            get { return input; }
        }

        public virtual string Output
        {
            get { return output; }
        }

        public virtual string Inputs
        {
            get { return inputs; }
        }

        public virtual string Outputs
        {
            get { return outputs; }
        }
    }
}
