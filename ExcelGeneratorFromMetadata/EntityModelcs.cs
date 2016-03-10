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
        public string LogicalName { get; set; }
        public List<AttributeModel> AttributeModelList { get; set; }

    }

    public class AttributeModel
    {
        public string DisplayName { get; set; }
        public string LogicalName { get; set; }
        public string DataType { get; set; }
        public string Constraint { get; set; }
        public bool IsRequired { get; set; }

    }
}
