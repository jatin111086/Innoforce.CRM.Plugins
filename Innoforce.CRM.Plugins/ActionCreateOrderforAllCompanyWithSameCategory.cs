using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Innoforce.CRM.Plugins
{
    public class ActionCreateOrderforAllCompanyWithSameCategory : IPlugin
    {

        bool isQuarter = false;
        int lang = 0;
        string quaterPrefix = "Q";

        public void Execute(IServiceProvider serviceProvider)
        {
            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            // Obtain the organization service reference.
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            // The InputParameters collection contains all the data    passed in the message request.

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is EntityReference)
            {

                EntityReference triggeredSupportContract = (EntityReference)context.InputParameters["Target"];

                Entity retrieveSuppContract = service.Retrieve("if_supportcontract", triggeredSupportContract.Id, new ColumnSet(true));

                Entity retrieveAccount = service.Retrieve("account", ((EntityReference)retrieveSuppContract.Attributes["if_primarycustomer"]).Id, new ColumnSet(true));
                int getOptionSetValue = ((OptionSetValue)retrieveSuppContract.Attributes["if_productcategory"]).Value;


                if (retrieveAccount.Attributes.Contains("if_if_language"))
                {
                    lang = ((OptionSetValue)retrieveAccount.Attributes["if_if_language"]).Value;

                    if (lang == 100000004)
                    {
                        quaterPrefix = "T";
                    }


                }

                Common comObj = new Common();

                Guid getCreatedOrder = getCreatedOrder = comObj.ProcessOrderCreation(service, retrieveSuppContract, retrieveAccount, true);





                UpdateQuaterForAllOrder(service, retrieveSuppContract, lang);





                //if(getCreatedOrder == Guid.Empty)
                //{
                //    throw new InvalidPluginExecutionException("Invalid Plugin");
                //}




                string getAllRelatedCategory = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>" +
"  <entity name='if_supportcontract'>" +
"    <attribute name='if_supportcontractid' />" +
"    <attribute name='if_name' />" +
"    <attribute name='if_costperterm' />" +
"    <attribute name='if_termduration' />" +
"    <attribute name='if_titleforinvoice' />" +
"    <attribute name='if_termstart' />" +
"    <attribute name='crd38_termstart' />" +
"    <attribute name='if_enddate' />" +
"    <attribute name='if_currency' />" +
"    <attribute name='if_primarycustomer' />" +
"    <attribute name='if_commentforinv' />" +
"    <attribute name='if_productcategory' />" +
"    <order attribute='if_name' descending='false' />" +
"    <filter type='and'>" +
"      <condition attribute='if_primarycustomer' operator='eq' value='" + ((EntityReference)retrieveSuppContract.Attributes["if_primarycustomer"]).Id + "' />" +
"      <condition attribute='if_contractstatus' operator='eq' value='100000002' />" +
"      <condition attribute='if_productcategory' operator='eq' value='" + getOptionSetValue + "' />" +
"    </filter>" +
"  </entity>" +
"</fetch>";

                EntityCollection retrieveRelatedSupportContract = service.RetrieveMultiple(new FetchExpression(getAllRelatedCategory));
                foreach (var loopSupportContrct in retrieveRelatedSupportContract.Entities)
                {
                    DateTime startDate = DateTime.Now;
                    if (loopSupportContrct.Attributes.Contains("crd38_termstart"))
                    {
                        startDate = (DateTime)loopSupportContrct.Attributes["crd38_termstart"];
                        startDate = startDate.AddDays(1);
                    }

                    DateTime endDate = DateTime.MinValue;
                    if (loopSupportContrct.Attributes.Contains("if_enddate"))
                    {
                        endDate = ((DateTime)loopSupportContrct.Attributes["if_enddate"]).AddDays(1);
                    }
                    Common comObj1 = new Common();

                    bool orderProductCreated = comObj1.CreateSalesOrderDetails(service, loopSupportContrct, startDate, endDate, getCreatedOrder, retrieveAccount, false, endDate);

                }
                checkAndUpdateOrderName(getCreatedOrder, service);
            }
        }


        // Method to retrieve existing order for the same account and category
        private Entity GetExistingOrder(IOrganizationService service, Entity retrieveAccount, Entity retrieveSuppContract)
        {
            int productCategoryValue = ((OptionSetValue)retrieveSuppContract.Attributes["if_productcategory"]).Value;

            QueryExpression query = new QueryExpression("salesorder");
            query.ColumnSet = new ColumnSet(true);
            query.Criteria.AddCondition("customerid", ConditionOperator.Equal, retrieveSuppContract.GetAttributeValue<EntityReference>("if_primarycustomer").Id);
            query.Criteria.AddCondition("if_productcategory", ConditionOperator.Equal, productCategoryValue);


            EntityCollection orders = service.RetrieveMultiple(query);
            if (orders.Entities.Count > 0)
            {

                return orders.Entities.First();
            }

            return null;
        }



        //New Test Q2/2023
        public void checkAndUpdateOrderName(Guid getCreatedOrder, IOrganizationService service)
        {
            var query_salesorderid = getCreatedOrder.ToString();
            bool orderNameExist = false;
            QueryExpression getOrderName = new QueryExpression();
            getOrderName.EntityName = "salesorder";
            getOrderName.ColumnSet = new ColumnSet("name");
            getOrderName.Criteria.AddCondition(new ConditionExpression("salesorderid", ConditionOperator.Equal, query_salesorderid));
            EntityCollection collectOrderRecords = service.RetrieveMultiple(getOrderName);
            string orderName = string.Empty;
            string salesorderLogicalName = string.Empty;

            foreach (var result in collectOrderRecords.Entities)
            {
                orderName = result.Attributes["name"].ToString();
                salesorderLogicalName = result.LogicalName;
            }

            QueryExpression getOrderProductName = new QueryExpression();
            getOrderProductName.EntityName = "salesorderdetail";
            getOrderProductName.ColumnSet = new ColumnSet("productdescription");
            getOrderProductName.Criteria.AddCondition(new ConditionExpression("salesorderid", ConditionOperator.Equal, query_salesorderid));
            EntityCollection collectOrderProductRecords = service.RetrieveMultiple(getOrderProductName);
            string orderProductName = string.Empty;

            foreach (var result in collectOrderProductRecords.Entities)
            {


                if (isQuarter)
                {
                    //ordName + " " + quaterPrefix;
                    orderProductName = result.Attributes["productdescription"].ToString() + " " + quaterPrefix;
                    if (orderName == orderProductName)
                    {
                        orderNameExist = true;
                        break;
                    }
                }
                else
                {
                    orderProductName = result.Attributes["productdescription"].ToString() + " " + DateTime.Now.Year.ToString();
                    if (orderName == orderProductName)
                    {
                        orderNameExist = true;
                        break;
                    }
                }
            }

            if (orderNameExist == false)
            {
                Entity salesOrder = new Entity("salesorder");
                salesOrder.Id = getCreatedOrder;
                salesOrder.LogicalName = salesorderLogicalName;
                salesOrder["name"] = orderProductName;
                service.Update(salesOrder);
            }
        }


        public void UpdateQuaterForAllOrder(IOrganizationService service, Entity retrieveSuppContract, int lang)
        {

            string contactenateMessage = string.Empty;
            decimal cost = 0;
            //calculate priceperunit



            if (retrieveSuppContract.Attributes.Contains("if_termduration"))
            {


                int getTermDuration = ((OptionSetValue)retrieveSuppContract.Attributes["if_termduration"]).Value;
                decimal costperterm = ((Money)retrieveSuppContract.Attributes["if_costperterm"]).Value;
                DateTime StartCalculationDate = DateTime.Now;





                if (getTermDuration == 100000001) //In case of 3 Months
                {

                    //quaterPrefix = "Q";


                    //if (lang == 100000004)
                    //    {
                    //        quaterPrefix = "T";
                    //    }



                    DateTime firstquatardate = new DateTime(DateTime.Now.Year, 01, 01);
                    DateTime firstquatardate_end = new DateTime(DateTime.Now.Year, 03, 31);

                    DateTime secondquatardate = new DateTime(DateTime.Now.Year, 04, 01);
                    DateTime secondquatardate_end = new DateTime(DateTime.Now.Year, 06, 30);

                    DateTime thirdquatardate = new DateTime(DateTime.Now.Year, 07, 01);
                    DateTime thirdquatardate_end = new DateTime(DateTime.Now.Year, 09, 30);

                    DateTime fourthquatardate = new DateTime(DateTime.Now.Year, 10, 01);
                    DateTime fourthquatadater_end = new DateTime(DateTime.Now.Year, 12, 31);

                    DateTime BeginOfFollowingTerm = new DateTime();
                    DateTime EndCalculationDate = new DateTime();




                    if (DateTime.Now >= firstquatardate && DateTime.Now <= firstquatardate_end)
                    {
                        StartCalculationDate = firstquatardate;
                        BeginOfFollowingTerm = secondquatardate;
                        EndCalculationDate = firstquatardate_end;

                        isQuarter = true;
                        quaterPrefix += "1/" + DateTime.Now.Year;

                    }

                    else if (DateTime.Now >= secondquatardate && DateTime.Now <= secondquatardate_end)
                    {
                        StartCalculationDate = secondquatardate;
                        BeginOfFollowingTerm = thirdquatardate;
                        EndCalculationDate = secondquatardate_end;
                        isQuarter = true;
                        quaterPrefix += "2/" + DateTime.Now.Year;



                    }

                    else if (DateTime.Now >= thirdquatardate && DateTime.Now <= thirdquatardate_end)
                    {
                        StartCalculationDate = thirdquatardate;
                        BeginOfFollowingTerm = fourthquatardate;
                        EndCalculationDate = thirdquatardate_end;
                        isQuarter = true;
                        quaterPrefix += "3/" + DateTime.Now.Year;


                    }

                    else if (DateTime.Now >= fourthquatardate && DateTime.Now <= fourthquatadater_end)
                    {
                        StartCalculationDate = fourthquatardate;
                        BeginOfFollowingTerm = new DateTime(DateTime.Now.Year + 1, 01, 01);
                        EndCalculationDate = fourthquatadater_end;
                        isQuarter = true;
                        quaterPrefix += "4/" + DateTime.Now.Year;
                    }
                }
            }

        }



    }
}
