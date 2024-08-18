
namespace Innoforce.CRM.Plugins
{
    using System;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Query;

    public class QuoteCloning : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {

            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            ITracingService tracing = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IOrganizationServiceFactory servicefactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = servicefactory.CreateOrganizationService(context.UserId);
            tracing.Trace(context.InputParameters.Contains("Target").ToString());
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                if (context.Depth > 1)
                {
                   return;
                }


                #region Clone Quote

                //Getting the Entity record from the Action
                Entity Quote = (Entity)context.InputParameters["Target"];

                //Retrieving the same entity record with all the attributes
                Entity CloneQuote = service.Retrieve(Quote.LogicalName, Quote.Id, new ColumnSet(true));

                //Set the EntityState to null, so that a new cloned record can be created
                CloneQuote.EntityState = null;
                
                //Remove auto generated quote attributes
                CloneQuote.Attributes.Remove(CloneQuote.LogicalName + "id");
                CloneQuote.Attributes.Remove("quotenumber");
                CloneQuote.Attributes.Remove("statecode");
                CloneQuote.Attributes.Remove("statuscode");


                if (CloneQuote.Contains("name"))
                {
                    CloneQuote.Attributes["name"] = CloneQuote.Attributes["name"] + " Clone: " + DateTime.Now.ToString("MMM dd");
                }
                else
                {
                    CloneQuote.Attributes["name"] = "New: " + DateTime.Now.ToString("MMM dd");
                }


                //set a unique id to the cloned record
                CloneQuote.Id = Guid.NewGuid();

                //Create the new cloned record
                Guid clonedQuoteId = service.Create(CloneQuote);

                #endregion



                #region Clone Quote Products

                // Retrieve all Quote Products associated with the original Quote.
                QueryExpression query = new QueryExpression("quotedetail");
                query.ColumnSet.AllColumns = true;
                query.Criteria.AddCondition("quoteid", ConditionOperator.Equal, Quote.Id);

                EntityCollection QuoteProductsCol = service.RetrieveMultiple(query);

                // Create a clone of each Quote Product and associate it with the cloned Quote.
                foreach (Entity quoteProduct in QuoteProductsCol.Entities)
                {
                    //Remove auto generated Quote Products attributes

                    quoteProduct.Attributes.Remove(quoteProduct.LogicalName + "id");
                    quoteProduct.Attributes.Remove("statecode");
                    quoteProduct.Attributes.Remove("statuscode");
                    quoteProduct.Attributes.Remove("statuscode");
                    quoteProduct["quoteid"] = new EntityReference("quote", clonedQuoteId);

                    //set a unique id to the cloned record
                    quoteProduct.Id = Guid.NewGuid();

                    Guid cloneQuoteProductId = service.Create(quoteProduct);

                }

                #endregion

            }
        }
    }
}