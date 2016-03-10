using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Xrm.Client;
using Microsoft.Xrm.Client.Services;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;

namespace ExcelGeneratorFromMetadata
{
    public class CrmHelper : IDisposable
    {
        private readonly CrmConnection crmConnection;
        private OrganizationService sharedOrganizationService;

        public CrmHelper(string url, string domain, string username, string password)
        {
            try
            {
                crmConnection =
                CrmConnection.Parse(string.Format("Url={0}; Domain={1}; Username={2}; Password={3};", url, domain,
                    username, password));
                crmConnection.Timeout = new TimeSpan(0, 0, 5, 0);
            }
            catch (Exception)
            {
                Console.WriteLine("It seems connection variables are not all correct.");
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public OrganizationService GetSharedOrganizationService()
        {
            return sharedOrganizationService ?? (sharedOrganizationService = new OrganizationService(crmConnection));
        }

        private void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (sharedOrganizationService != null)
                {
                    sharedOrganizationService.Dispose();
                    sharedOrganizationService = null;
                }
            }
        }

        public List<EntityModel> GetEntityModel(List<string> entityLogicalNameList,
            bool includeOnlycustomPrefixedAttributes, string customPrefix)
        {
            var entityModelList = new List<EntityModel>();
            var requiredReference = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.ApplicationRequired);

            foreach (var entityLogicalName in entityLogicalNameList)
            {
                Console.WriteLine(entityLogicalName + " is now retrieving from CRM.");
                var entityModel = new EntityModel();
                entityModel.LogicalName = entityLogicalName;
                var retrieveEntityRequest = new RetrieveEntityRequest
                {
                    EntityFilters = EntityFilters.All,
                    LogicalName = entityLogicalName
                };

                RetrieveEntityResponse retrieveEntityResponse;
                try
                {
                    retrieveEntityResponse = (RetrieveEntityResponse)GetSharedOrganizationService().Execute(retrieveEntityRequest);
                }
                catch (Exception)
                {
                    Console.WriteLine("Error while retrieving {0}. Program will now terminate.", entityLogicalName);
                    break;
                }
                
                entityModel.DisplayName = retrieveEntityResponse.EntityMetadata.DisplayName.UserLocalizedLabel.Label;
                var attributes = retrieveEntityResponse.EntityMetadata.Attributes;

                Console.WriteLine(entityLogicalName + " attributes are now modelling.");
                foreach (var attributeMetadata in attributes)
                {
                    if (includeOnlycustomPrefixedAttributes && !attributeMetadata.LogicalName.StartsWith(customPrefix))
                    {
                        continue;
                    }

                    if (attributeMetadata.DisplayName.UserLocalizedLabel == null)
                    {
                        continue;
                    }

                    var attribute = new AttributeModel();
                    attribute.LogicalName = attributeMetadata.LogicalName;
                    attribute.DisplayName = attributeMetadata.DisplayName.UserLocalizedLabel.Label;
                    attribute.DataType = attributeMetadata.AttributeType.HasValue
                        ? attributeMetadata.AttributeType.Value.ToString()
                        : string.Empty;
                    if (attribute.DataType.Equals("Lookup", StringComparison.InvariantCultureIgnoreCase))
                    {
                        attribute.Constraint = "Hedef Varlık: " +
                                               ((LookupAttributeMetadata) (attributeMetadata)).Targets[0];
                    }
                    else if (attribute.DataType.Equals("String", StringComparison.InvariantCultureIgnoreCase))
                    {
                        attribute.Constraint = ((StringAttributeMetadata) (attributeMetadata)).MaxLength + " Karakter";
                    }
                    else if (attribute.DataType.Equals("Decimal", StringComparison.InvariantCultureIgnoreCase))
                    {
                        attribute.Constraint = ((DecimalAttributeMetadata) (attributeMetadata)).Precision + " ondalık, " +
                                               ((DecimalAttributeMetadata) (attributeMetadata)).MinValue + "-" +
                                               ((DecimalAttributeMetadata) (attributeMetadata)).MaxValue + " Max";
                    }
                    else if (attribute.DataType.Equals("Integer", StringComparison.InvariantCultureIgnoreCase))
                    {
                        attribute.Constraint = ((IntegerAttributeMetadata) (attributeMetadata)).MinValue + "-" +
                                               ((IntegerAttributeMetadata) (attributeMetadata)).MaxValue;
                    }
                    else if (attribute.DataType.Equals("Boolean", StringComparison.InvariantCultureIgnoreCase))
                    {
                        attribute.Constraint = "Default: " +
                                               (((BooleanAttributeMetadata) (attributeMetadata)).DefaultValue.HasValue
                                                   ? ((BooleanAttributeMetadata) (attributeMetadata)).DefaultValue
                                                       .ToString()
                                                   : "false");
                    }
                    else if (attribute.DataType.Equals("Memo", StringComparison.InvariantCultureIgnoreCase))
                    {
                        attribute.Constraint = ((MemoAttributeMetadata) (attributeMetadata)).MaxLength + " Karakter";
                    }
                    else if (attribute.DataType.Equals("Enum", StringComparison.InvariantCultureIgnoreCase))
                    {
                        attribute.Constraint = GetOptions((EnumAttributeMetadata) (attributeMetadata));
                    }
                    else if (attribute.DataType.Equals("Money", StringComparison.InvariantCultureIgnoreCase))
                    {
                        attribute.Constraint = ((MoneyAttributeMetadata) (attributeMetadata)).Precision + " ondalık, " +
                                               ((MoneyAttributeMetadata) (attributeMetadata)).MinValue + "-" +
                                               ((MoneyAttributeMetadata) (attributeMetadata)).MaxValue + " Max";
                    }
                    else if (attribute.DataType.Equals("BigInt", StringComparison.InvariantCultureIgnoreCase))
                    {
                        attribute.Constraint = ((BigIntAttributeMetadata) (attributeMetadata)).MinValue + "-" +
                                               ((BigIntAttributeMetadata) (attributeMetadata)).MaxValue;
                    }
                    else if (attribute.DataType.Equals("DateTime", StringComparison.InvariantCultureIgnoreCase))
                    {
                        var dateTimeFormat = ((DateTimeAttributeMetadata) (attributeMetadata)).Format;
                        if (dateTimeFormat != null)
                        {
                            attribute.Constraint = dateTimeFormat.Value.ToString();
                        }
                    }
                    else if (attribute.DataType.Equals("Picklist", StringComparison.InvariantCultureIgnoreCase))
                    {
                        attribute.Constraint = GetOptions((PicklistAttributeMetadata) (attributeMetadata));
                    }
                    else if (attribute.DataType.Equals("Uniqueidentifier", StringComparison.InvariantCultureIgnoreCase))
                    {
                        attribute.Constraint = "Guid";
                        attribute.IsRequired = true;
                    }

                    attribute.IsRequired = attributeMetadata.RequiredLevel.Value == requiredReference.Value;
                    entityModel.AttributeModelList.Add(attribute);
                }

                entityModel.AttributeModelList = entityModel.AttributeModelList.OrderBy(model => model.LogicalName).ToList();

                entityModelList.Add(entityModel);
            }

            return entityModelList;
        }

        private static string GetOptions(EnumAttributeMetadata enumAttributeMetadata)
        {
            if (enumAttributeMetadata == null || enumAttributeMetadata.OptionSet == null ||
                enumAttributeMetadata.OptionSet.Options == null || enumAttributeMetadata.OptionSet.Options.Count == 0)
            {
                return string.Empty;
            }

            var returnStr = string.Empty;
            foreach (var option in enumAttributeMetadata.OptionSet.Options)
            {
                if (option.Label == null || option.Label.UserLocalizedLabel == null ||
                    option.Label.UserLocalizedLabel.Label == null)
                {
                    continue;
                }

                returnStr += option.Value + "-" + option.Label.UserLocalizedLabel.Label + ", ";
            }

            return returnStr.Substring(default(int), returnStr.Length - ", ".Length);
        }
    }
}