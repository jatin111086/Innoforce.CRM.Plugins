namespace Innoforce.CRM.Plugins
{
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Query;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net.Http.Headers;
    using System.Net.Http;
    using System.Text;
    using System.Threading.Tasks;
    public class DowloadSharePointQuoteDocuments : IPlugin
    {

        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            ITracingService tracing = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IOrganizationServiceFactory servicefactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = servicefactory.CreateOrganizationService(context.UserId);

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                Entity quote = (Entity)context.InputParameters["Target"];

                if (quote.LogicalName == "quote" && context.MessageName.ToLower() == "update")
                {


                    string allExistingQuoteProducts = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>" +
"  <entity name='quotedetail'>" +
"    <attribute name='productid' />" +
"    <attribute name='isproductoverridden' />" +
"    <attribute name='quoteid' />" +
"    <attribute name='quotedetailid' />" +
"    <filter type='and'>" +
"      <condition attribute='quoteid' operator='eq' value='" + quote.Id + "' />" +
"      <condition attribute='isproductoverridden' operator='eq' value='0' />" +
"    </filter>" +
"  </entity>" +
"</fetch>";

                    EntityCollection retrieveExistingQuoteProducts = service.RetrieveMultiple(new FetchExpression(allExistingQuoteProducts));
                    foreach (var loopQuoteProduct in retrieveExistingQuoteProducts.Entities)
                    {
                        // Retrieve product information
                        var quoteDetailId = loopQuoteProduct.GetAttributeValue<Guid>("quotedetailid");
                        var productId = loopQuoteProduct.GetAttributeValue<EntityReference>("productid").Id;

                        // Search SharePoint for related documents using the product ID
                        List<string> productDocuments = SearchSharePointDocuments(productId, tracing);

                    }

                }
                }
            }

        private List<string> SearchSharePointDocuments(Guid productId, ITracingService tracing)
        {
            List<string> documentUrls = new List<string>();


            string searchUrl = $"https://your-sharepoint-site.com/_api/search/query?querytext='ProductId:{productId}'";

            using (HttpClient client = new HttpClient())
            {
                // Add necessary headers
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                // Make request to SharePoint
                HttpResponseMessage response = client.GetAsync(searchUrl).Result;

                // Check if request was successful
                if (response.IsSuccessStatusCode)
                {
                    // Process search results here
                    var searchResult = response.Content.ReadAsStringAsync().Result;
                    // Parse search results to extract document URLs
                    
                    // Modify this part based on the actual format of your SharePoint search results
                    // Here, we are just adding sample URLs for demonstration
                    documentUrls.Add("https://your-sharepoint-site.com/documents/document1.pdf");
                    documentUrls.Add("https://your-sharepoint-site.com/documents/document2.docx");
                    documentUrls.Add("https://your-sharepoint-site.com/documents/document3.txt");

                    tracing.Trace($"SharePoint search result for Product ID {productId}: {searchResult}");
                }
                else
                {
                    // Log any errors
                    tracing.Trace($"Error while searching SharePoint for Product ID {productId}. Status code: {response.StatusCode}");
                }
                return documentUrls;
            }
        }
    }
}