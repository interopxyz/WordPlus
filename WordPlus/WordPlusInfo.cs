using Grasshopper;
using Grasshopper.Kernel;
using System;
using System.Drawing;

namespace WordPlus
{
    public class WordPlusInfo : GH_AssemblyInfo
    {
        public override string Name => "WordPlus";

        //Return a 24x24 pixel bitmap to represent this GHA library.
        public override Bitmap Icon => Properties.Resources.WordPlus_Logo_24;

        //Return a short string describing the purpose of this GHA library.
        public override string Description => "A Grasshopper 3d plugin for Word File Read & Write";

        public override Guid Id => new Guid("83CE3F47-35DA-4461-B191-64E39FBB3F55");

        //Return a string identifying you or your company.
        public override string AuthorName => "David Mans";

        //Return a string representing your preferred contact details.
        public override string AuthorContact => "interopxyz@gmail.com";

        public override string AssemblyVersion => "1.0.2.0";
        }

    public class WordPlusCategoryIcon : GH_AssemblyPriority
    {
        public object Properties { get; private set; }

        public override GH_LoadingInstruction PriorityLoad()
        {
            Instances.ComponentServer.AddCategoryIcon(Constants.ShortName, WordPlus.Properties.Resources.WordPlus_Logo_16);
            Instances.ComponentServer.AddCategorySymbolName(Constants.ShortName, 'W');
            return GH_LoadingInstruction.Proceed;
        }
    }

}