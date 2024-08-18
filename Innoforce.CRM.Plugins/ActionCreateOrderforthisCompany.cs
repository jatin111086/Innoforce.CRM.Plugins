using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

namespace Innoforce.CRM.Plugins
{
    public class ActionCreateOrderforthisCompany : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            try
            {

                // Obtain the execution context from the service provider.
                IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
                // Obtain the organization service reference.
                IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
                // The InputParameters collection contains all the data    passed in the message request.

                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is EntityReference)
                {
                    // Entity retrievePricelist = service.Retrieve("pricelevel", context.PrimaryEntityId, new ColumnSet("name", "pricelevelid", "transactioncurrencyid"));

                    EntityReference triggeredSupportContract = (EntityReference)context.InputParameters["Target"];
                    Entity retrieveSuppContract = service.Retrieve("if_supportcontract", triggeredSupportContract.Id, new ColumnSet(true));
                    Entity retrieveAccount = service.Retrieve("account", ((EntityReference)retrieveSuppContract.Attributes["if_primarycustomer"]).Id, new ColumnSet(true));

                    Common comObj = new Common();
                    comObj.ProcessOrderCreation(service, retrieveSuppContract, retrieveAccount, false);
                }
            }
            catch (FaultException ex)
            {
                throw new InvalidPluginExecutionException("Err occurred." + ex.Message);
            }
            catch (Exception ex)
            {

                throw new InvalidPluginExecutionException("Err occurred." + ex.Message);
            }

        }
    }
}
