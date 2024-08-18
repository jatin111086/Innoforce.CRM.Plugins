using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace Innoforce.CRM.Plugins
{
    public class UpdateProductNameInOrderAndInvoiceDetails : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            // Obtain the organization service reference.
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            // The InputParameters collection contains all the data    passed in the message request.

            Guid userId = context.InitiatingUserId;

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                try
                {
                    Entity entity = (Entity)context.InputParameters["Target"];

                    if (entity.LogicalName == "orderdetails")
                    {
                        SetProductNameOrderDetails(entity, service);
                    }
                    else if (entity.LogicalName == "invdetails")
                    {
                        SetProductNameInvoiceDetails(entity, service);
                    }
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException(ex.Message);
                }
            }
        }

        public void SetProductNameOrderDetails(Entity OrderDetail, IOrganizationService service)
        {
            try
            {
                if (OrderDetail.Contains("order"))
                {
                    var orderlookUp = ((EntityReference)OrderDetail.Attributes["order"]);

                    if (orderlookUp != null)
                    {
                        Entity Order = service.Retrieve("order", orderlookUp.Id, new ColumnSet(true));


                    }

                }







                //OrderDetail.Attributes["crd38_until"] = "";
                //service.Update(OrderDetail);


            }
            catch (Exception)
            {

                throw;
            }
        }

        public void SetProductNameInvoiceDetails(Entity InvoiceDetail, IOrganizationService service)
        {
            try
            {

            }
            catch (Exception)
            {
                throw;
            }
        }

    }
}
