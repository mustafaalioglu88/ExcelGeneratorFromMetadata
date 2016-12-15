using System.Collections.Generic;

namespace ExcelGeneratorFromMetadata
{
    public static class Settings
    {
        public static readonly bool IncludeOnlycustomPrefixedAttributes = true;
        public static readonly string CustomPrefix = "vrp_";
        public static readonly List<string> AttributeIncludeList = new List<string>()
        {
            //"statecode", 
            //"statuscode"
        };

        public static readonly List<string> EntityIncludeAllList = new List<string>()
        {
            //"account", 
            //"contact"
        }; 
    }
}
