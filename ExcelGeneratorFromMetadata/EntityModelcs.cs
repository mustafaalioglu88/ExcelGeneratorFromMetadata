using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelGeneratorFromMetadata
{
    public class EntityModel
    {
        public EntityModel()
        {
            AttributeModelList = new List<AttributeModel>();
        }

        public string DisplayName { get; set; }
        public string DisplayNamePlural { get; set; }
        public string Description { get; set; }
        public string LogicalName { get; set; }
        public List<AttributeModel> AttributeModelList { get; set; }

    }

    public class AttributeModel
    {
        public string DisplayName { get; set; }
        public string OtherDisplayName { get; set; }
        public string Description { get; set; }
        public string OtherDescription { get; set; }
        public string LogicalName { get; set; }
        public string DataType { get; set; }
        public string LookupEntityLogicalName { get; set; }
        public string OptionSetList { get; set; }
        public string GlobalOptionSetListLogicalName { get; set; }
        public string MinValue { get; set; }
        public string MaxValue { get; set; }
        public bool IsRequired { get; set; }

    }
}
