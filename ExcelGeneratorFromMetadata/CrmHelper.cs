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
            catch (Exception ex)
            {
                Console.WriteLine("It seems connection variables are not all correct. Or: " + ex);
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private OrganizationService GetSharedOrganizationService()
        {
            return sharedOrganizationService ?? (sharedOrganizationService = new OrganizationService(crmConnection));
        }

        private void Dispose(bool disposing)
        {
            if (!disposing) return;
            if (sharedOrganizationService == null) return;

            sharedOrganizationService.Dispose();
            sharedOrganizationService = null;
        }

        public List<EntityModel> GetEntityModel(List<string> entityLogicalNameList, out IList<string> skippedList)
        {
            skippedList = new List<string>();
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
                catch (Exception ex)
                {
                    skippedList.Add(entityLogicalName);
                    Console.WriteLine("Error while retrieving {0}. Detais:\n{1}", entityLogicalName, ex.Message);
                    continue;
                }

                if (retrieveEntityResponse.EntityMetadata.IsIntersect.HasValue && retrieveEntityResponse.EntityMetadata.IsIntersect.Value)
                {
                    skippedList.Add(entityLogicalName);
                    continue;
                }
                
                entityModel.DisplayName = retrieveEntityResponse.EntityMetadata.DisplayName.UserLocalizedLabel.Label;
                entityModel.DisplayNamePlural = retrieveEntityResponse.EntityMetadata.DisplayCollectionName.UserLocalizedLabel.Label;
                entityModel.Description = retrieveEntityResponse.EntityMetadata.Description.UserLocalizedLabel.Label;
                var attributes = retrieveEntityResponse.EntityMetadata.Attributes;

                Console.WriteLine(entityLogicalName + " attributes are now modelling.");
                foreach (var attributeMetadata in attributes)
                {
                    if (ShouldPassCurrentAttribute(entityLogicalName, attributeMetadata))
                    {
                        continue;
                    }

                    if (attributeMetadata.DisplayName.UserLocalizedLabel == null)
                    {
                        continue;
                    }

                    var localLangCode = attributeMetadata.DisplayName.UserLocalizedLabel.LanguageCode;
                    var otherLangDisplay = (from langs in attributeMetadata.DisplayName.LocalizedLabels where langs.LanguageCode != localLangCode select langs).ToList();
                    var otherLangDescription = (from langs in attributeMetadata.Description.LocalizedLabels where langs.LanguageCode != localLangCode select langs).ToList();
                    var otherDisplayName = string.Empty;
                    var otherDescription = string.Empty;

                    if(otherLangDisplay.Count > 0)
                    {
                        otherDisplayName = otherLangDisplay.First().Label;
                    }

                    if (otherLangDescription.Count > 0)
                    {
                        otherDescription = otherLangDescription.First().Label;
                    }

                    var attribute = new AttributeModel
                    {
                        LogicalName = attributeMetadata.SchemaName,
                        DisplayName = attributeMetadata.DisplayName.UserLocalizedLabel.Label,
                        Description = attributeMetadata.Description.UserLocalizedLabel != null ? attributeMetadata.Description.UserLocalizedLabel.Label : string.Empty,
                        OtherDisplayName = otherDisplayName,
                        OtherDescription = otherDescription,
                        DataType = attributeMetadata.AttributeType.HasValue
                            ? attributeMetadata.AttributeType.Value.ToString()
                            : string.Empty
                    };
                    if (attribute.DataType.Equals("Lookup", StringComparison.InvariantCultureIgnoreCase))
                    {
                        attribute.DataType = "Lookup";
                        attribute.LookupEntityLogicalName = ((LookupAttributeMetadata) (attributeMetadata)).Targets[0];
                    }
                    else if (attribute.DataType.Equals("String", StringComparison.InvariantCultureIgnoreCase))
                    {
                        attribute.DataType = "SingleLine";
                        attribute.MaxValue = ((StringAttributeMetadata) (attributeMetadata)).MaxLength.ToString();
                    }
                    else if (attribute.DataType.Equals("Decimal", StringComparison.InvariantCultureIgnoreCase))
                    {
                        attribute.DataType = "decimal";
                        attribute.MinValue = ((DecimalAttributeMetadata) (attributeMetadata)).MinValue.ToString();
                        attribute.MaxValue = ((DecimalAttributeMetadata)(attributeMetadata)).MaxValue.ToString();
                        //attribute.Precision = ((DecimalAttributeMetadata)(attributeMetadata)).Precision;
                    }
                    else if (attribute.DataType.Equals("Integer", StringComparison.InvariantCultureIgnoreCase))
                    {
                        attribute.DataType = "int";
                        attribute.MinValue = ((IntegerAttributeMetadata)(attributeMetadata)).MinValue.ToString();
                        attribute.MaxValue = ((IntegerAttributeMetadata)(attributeMetadata)).MaxValue.ToString();
                    }
                    else if (attribute.DataType.Equals("Boolean", StringComparison.InvariantCultureIgnoreCase))
                    {
                        attribute.DataType = "bool";
                        /*attribute.Constraint = "Default: " +
                                               (((BooleanAttributeMetadata) (attributeMetadata)).DefaultValue.HasValue
                                                   ? ((BooleanAttributeMetadata) (attributeMetadata)).DefaultValue
                                                       .ToString()
                                                   : "false");*/
                    }
                    else if (attribute.DataType.Equals("Memo", StringComparison.InvariantCultureIgnoreCase))
                    {
                        attribute.DataType = "Multiline";
                        attribute.MaxValue = ((MemoAttributeMetadata)(attributeMetadata)).MaxLength.ToString();
                    }
                    else if (attribute.DataType.Equals("Money", StringComparison.InvariantCultureIgnoreCase))
                    {
                        attribute.DataType = "Money";
                        attribute.MinValue = ((MoneyAttributeMetadata)(attributeMetadata)).MinValue.ToString();
                        attribute.MaxValue = ((MoneyAttributeMetadata)(attributeMetadata)).MaxValue.ToString();
                        //attribute.Constraint = ((MoneyAttributeMetadata)(attributeMetadata)).Precision;
                    }
                    else if (attribute.DataType.Equals("BigInt", StringComparison.InvariantCultureIgnoreCase))
                    {
                        attribute.DataType = "int";
                        attribute.MinValue = ((BigIntAttributeMetadata)(attributeMetadata)).MinValue.ToString();
                        attribute.MaxValue = ((BigIntAttributeMetadata)(attributeMetadata)).MaxValue.ToString();
                    }
                    else if (attribute.DataType.Equals("DateTime", StringComparison.InvariantCultureIgnoreCase))
                    {
                        attribute.DataType = "DateTime";
                        /*var dateTimeFormat = ((DateTimeAttributeMetadata) (attributeMetadata)).Format;
                        if (dateTimeFormat != null)
                        {
                            attribute.Constraint = dateTimeFormat.Value.ToString();
                        }*/
                    }
                    else if (attribute.DataType.Equals("Picklist", StringComparison.InvariantCultureIgnoreCase))
                    {
                        attribute.DataType = "OptionSet";
                        var localType = ((PicklistAttributeMetadata) (attributeMetadata));
                        if(localType.OptionSet.IsGlobal.HasValue && localType.OptionSet.IsGlobal.Value)
                        {
                            attribute.GlobalOptionSetListLogicalName = localType.OptionSet.Name;
                            attribute.DataType = "GlobalOptionSet";
                        }
                        attribute.OptionSetList = GetOptions(localType);
                    }

                    attribute.IsRequired = attributeMetadata.RequiredLevel.Value == requiredReference.Value;
                    entityModel.AttributeModelList.Add(attribute);
                }

                var manyToManyRelationships = retrieveEntityResponse.EntityMetadata.ManyToManyRelationships;
                if (manyToManyRelationships != null && manyToManyRelationships.Length > 0)
                {
                    foreach (var manyToManyRelationshipMetadata in manyToManyRelationships)
                    {
                        var attribute = new AttributeModel
                        {
                            LogicalName = manyToManyRelationshipMetadata.SchemaName,
                            DataType = "NN",
                            IsRequired = false
                        };

                        string mainEntityDisplay;
                        string relationEntityDisplay;
                        if (manyToManyRelationshipMetadata.Entity1LogicalName == entityLogicalName)
                        {
                            attribute.LookupEntityLogicalName = manyToManyRelationshipMetadata.Entity2LogicalName;
                            attribute.Description = "1";

                            mainEntityDisplay = GetUserLocalizedLabel(manyToManyRelationshipMetadata.Entity1AssociatedMenuConfiguration);
                            if (string.IsNullOrWhiteSpace(mainEntityDisplay))
                            {
                                mainEntityDisplay = entityLogicalName;
                            }

                            relationEntityDisplay = GetUserLocalizedLabel(manyToManyRelationshipMetadata.Entity2AssociatedMenuConfiguration);
                            if (string.IsNullOrWhiteSpace(relationEntityDisplay))
                            {
                                relationEntityDisplay = manyToManyRelationshipMetadata.Entity2LogicalName;
                            }
                        }
						else
                        {
                            attribute.LookupEntityLogicalName = manyToManyRelationshipMetadata.Entity1LogicalName;
                            attribute.Description = "2";

                            mainEntityDisplay = GetUserLocalizedLabel(manyToManyRelationshipMetadata.Entity2AssociatedMenuConfiguration);
                            if (string.IsNullOrWhiteSpace(mainEntityDisplay))
                            {
                                mainEntityDisplay = entityLogicalName;
                            }

                            relationEntityDisplay = GetUserLocalizedLabel(manyToManyRelationshipMetadata.Entity1AssociatedMenuConfiguration);
                            if (string.IsNullOrWhiteSpace(relationEntityDisplay))
                            {
                                relationEntityDisplay = manyToManyRelationshipMetadata.Entity2LogicalName;
                            }
                        }

                        attribute.DisplayName = mainEntityDisplay + ";" + relationEntityDisplay;

                        entityModel.AttributeModelList.Add(attribute);
                    }
                }

                entityModel.AttributeModelList = entityModel.AttributeModelList.OrderBy(model => model.LogicalName).ToList();

                entityModelList.Add(entityModel);
            }

            return entityModelList;
        }

        private static bool ShouldPassCurrentAttribute(string entityLogicalName, AttributeMetadata attributeMetadata)
        {
            if (Settings.EntityIncludeAllList.Contains(entityLogicalName))
            {
                return false;
            }

            if(Settings.AttributeIncludeList.Contains(attributeMetadata.LogicalName))
            {
                return false;
            }

            if(Settings.IncludeOnlycustomPrefixedAttributes && !attributeMetadata.LogicalName.StartsWith(Settings.CustomPrefix))
            {
                return true;
            }

            return false;
        }

        private string GetUserLocalizedLabel(AssociatedMenuConfiguration entityAssociatedMenuConfiguration)
        {
            if (entityAssociatedMenuConfiguration.Label != null && entityAssociatedMenuConfiguration.Label.UserLocalizedLabel != null 
                && !string.IsNullOrWhiteSpace(entityAssociatedMenuConfiguration.Label.UserLocalizedLabel.Label))
            {
                return entityAssociatedMenuConfiguration.Label.UserLocalizedLabel.Label;
            }

            return string.Empty;
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

                returnStr += option.Label.UserLocalizedLabel.Label + "=" + option.Value + ";";
            }

            return returnStr.Substring(default(int), returnStr.Length - ";".Length);
        }
    }
}