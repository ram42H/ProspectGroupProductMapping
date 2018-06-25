using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FdxProspectGroupProductMapping
{
    public class GenerateListOfRecommendedProducts : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            if (context.Depth > 1)
                return;

            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService crmService = serviceFactory.CreateOrganizationService(null);
            ITracingService traceService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            try
            {
                #region Associate & Disassociate

                if (context.MessageName.Equals("associate", StringComparison.InvariantCultureIgnoreCase)
                    || context.MessageName.Equals("disassociate", StringComparison.InvariantCultureIgnoreCase))
                {
                    string relationshipName = string.Empty;
                    if (context.InputParameters.Contains("Relationship"))
                    {
                        relationshipName = ((Relationship)context.InputParameters["Relationship"]).SchemaName;
                    }
                    traceService.Trace(relationshipName);
                    if (!relationshipName.Equals("fdx_fdx_prospectgroup_product", StringComparison.CurrentCultureIgnoreCase))
                    {
                        return;
                    }

                    if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is EntityReference)
                    {
                        EntityReference targetEntity = (EntityReference)context.InputParameters["Target"];
                        Guid prospectGroupId = targetEntity.Id;
                        traceService.Trace("Prospect Group Id: " + prospectGroupId.ToString());
                        List<string> recommendedProductNames = new List<string>();
                        string recommendedProductsFetchXml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                                                                          <entity name='product'>
                                                                            <attribute name='name' />
                                                                            <attribute name='productid' />
                                                                            <attribute name='productnumber' />
                                                                            <attribute name='description' />
                                                                            <attribute name='statecode' />
                                                                            <attribute name='productstructure' />
                                                                            <order attribute='productnumber' descending='false' />
                                                                            <link-entity name='fdx_fdx_prospectgroup_product' from='productid' to='productid' visible='false' intersect='true'>
                                                                              <link-entity name='fdx_prospectgroup' from='fdx_prospectgroupid' to='fdx_prospectgroupid' alias='ab'>
                                                                                <filter type='and'>
                                                                                  <condition attribute='fdx_prospectgroupid' operator='eq' value='{0}' />
                                                                                </filter>
                                                                              </link-entity>
                                                                            </link-entity>
                                                                          </entity>
                                                                        </fetch>";
                        EntityCollection recommendedProducts = crmService.RetrieveMultiple(new FetchExpression(string.Format(recommendedProductsFetchXml, prospectGroupId.ToString())));
                        foreach (Entity recommendedProduct in recommendedProducts.Entities)
                        {
                            recommendedProductNames.Add((string)recommendedProduct["productnumber"]);
                        }
                        recommendedProductNames.Sort();
                        Entity prospectGroup = new Entity("fdx_prospectgroup", prospectGroupId);
                        prospectGroup["fdx_listofrecommendedproducts"] = String.Join(";", recommendedProductNames.ToArray());
                        crmService.Update(prospectGroup);
                    }
                }
                #endregion
            }
            catch (Exception ex)
            {
                traceService.Trace(string.Format("Associate & Dissassociate Plugin error: {0}", new[] { ex.ToString() }));
            }
        }
    }
}
