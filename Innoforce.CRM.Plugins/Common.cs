using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace Innoforce.CRM.Plugins
{
    public class Common
    {
        bool isQuarter = false;
        int lang = 0;
        string quaterPrefix = "Q";
        public Guid ProcessOrderCreation(IOrganizationService service, Entity retrieveSuppContract, Entity retrieveAccount, bool checkForAllOrNot)
        {
            isQuarter = false;
            lang = 0;
            quaterPrefix = "Q";
            

            
            

            DateTime startDate = DateTime.Now;
            DateTime termEndDate = DateTime.Now;
            DateTime pluginEndDate = DateTime.Now;
            string dateTimeStart = string.Empty;
            string dateTimeTermEnd = string.Empty;

            if (retrieveAccount.Attributes.Contains("if_if_language"))
            {
                lang = ((OptionSetValue)retrieveAccount.Attributes["if_if_language"]).Value;

                if (lang == 100000004)
                {
                    quaterPrefix = "T";
                }
            }

            if (retrieveSuppContract.Attributes.Contains("crd38_termstart"))
            {
                startDate = (DateTime)retrieveSuppContract.Attributes["crd38_termstart"];
                startDate = startDate.AddDays(1);
                dateTimeStart = startDate.ToString("dd/MM/yyyy");
            }

            if (retrieveSuppContract.Attributes.Contains("if_enddate"))
            {
                termEndDate = (DateTime)retrieveSuppContract.Attributes["if_enddate"];
                termEndDate = termEndDate.AddDays(1);
                dateTimeTermEnd = termEndDate.ToString("dd/MM/yyyy");
            }
            else
            {
                termEndDate = DateTime.MinValue;
            }

            DateTime endDate = new DateTime(DateTime.Now.Year, 12, 31);
            string dateTimeEnd = endDate.ToString("dd/MM/yyyy");

            //Retrieving Data From PriceList Entity
            Guid RRPGuid = new Guid();
            Guid OtherGuid = new Guid();
            int RRPCount = 0;
            int otherPriceListCount = 0;
            QueryExpression getPriceList = new QueryExpression();
            getPriceList.EntityName = "pricelevel";
            getPriceList.ColumnSet = new ColumnSet("pricelevelid", "name");
            FilterExpression filter = new FilterExpression(LogicalOperator.And);
            FilterExpression filter1 = new FilterExpression(LogicalOperator.And);
            filter1.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
            filter1.AddCondition(new ConditionExpression("transactioncurrencyid", ConditionOperator.Equal, ((EntityReference)retrieveSuppContract.Attributes["if_currency"]).Id));
            filter1.AddCondition(new ConditionExpression("begindate", ConditionOperator.OnOrBefore, DateTime.Now));
            FilterExpression filter2 = new FilterExpression(LogicalOperator.Or);
            filter2.AddCondition(new ConditionExpression("enddate", ConditionOperator.OnOrAfter, DateTime.Now));
            filter2.AddCondition(new ConditionExpression("enddate", ConditionOperator.Null));
            filter.AddFilter(filter1);
            filter.AddFilter(filter2);
            getPriceList.Criteria = filter;
            EntityCollection collectRecords = service.RetrieveMultiple(getPriceList);

            foreach (var result in collectRecords.Entities)
            {

                if (result.Attributes["name"].ToString().Contains("RRP"))
                {

                    RRPCount++;
                    RRPGuid = result.Id;
                }
                else
                {
                    otherPriceListCount++;
                    OtherGuid = result.Id;
                }

            }

            Entity createOrder = new Entity("salesorder");




            createOrder["name"] = retrieveSuppContract.Attributes["if_titleforinvoice"] + " " + DateTime.Now.Year.ToString();
            createOrder["transactioncurrencyid"] = retrieveSuppContract.Attributes["if_currency"];
            createOrder["customerid"] = retrieveSuppContract.Attributes["if_primarycustomer"];

            if (retrieveAccount.Attributes.Contains("primarycontactid"))
            {
                createOrder["if_referenceperson"] = retrieveAccount.Attributes["primarycontactid"]; // Retrieving primaryContact from Account
            }

            createOrder["paymenttermscode"] = retrieveAccount.Attributes["paymenttermscode"];
            createOrder["if_orderdate"] = DateTime.Now;  // OrderDate as Current date

            //createOrder["description"] = retrieveSuppContract.Attributes.Contains("if_commentforinv") ? retrieveSuppContract.Attributes["if_commentforinv"] + "" + dateTimeStart + "-" + dateTimeEnd : "" + dateTimeStart + "-" +dateTimeEnd +"";
            createOrder["if_productcategory"] = retrieveSuppContract.Attributes["if_productcategory"];

            if (RRPCount == 1)
            {

                createOrder["pricelevelid"] = new EntityReference(collectRecords.Entities[0].LogicalName, RRPGuid); // Set PriceList Field as per Currency
            }
           
            createOrder["if_supportcontract"] = new EntityReference(retrieveSuppContract.LogicalName, retrieveSuppContract.Id); // Set PriceList Field as per Currency

            createOrder["billto_line1"] = retrieveAccount.Attributes.Contains("address1_line1") ? retrieveAccount.Attributes["address1_line1"] : string.Empty;
            createOrder["billto_line2"] = retrieveAccount.Attributes.Contains("address1_line2") ? retrieveAccount.Attributes["address1_line2"] : string.Empty;
            createOrder["billto_city"] = retrieveAccount.Attributes.Contains("address1_city") ? retrieveAccount.Attributes["address1_city"] : string.Empty;
            createOrder["billto_stateorprovince"] = retrieveAccount.Attributes.Contains("address1_stateorprovince") ? retrieveAccount.Attributes["address1_stateorprovince"] : string.Empty;
            createOrder["billto_postalcode"] = retrieveAccount.Attributes.Contains("address1_postalcode") ? retrieveAccount.Attributes["address1_postalcode"] : string.Empty;
            createOrder["billto_country"] = retrieveAccount.Attributes.Contains("address1_country") ? retrieveAccount.Attributes["address1_country"] : string.Empty;
            createOrder["if_customerreference"] = retrieveSuppContract.Attributes.Contains("new_contractreference") ? retrieveSuppContract.Attributes["new_contractreference"] : string.Empty;
            createOrder["crd38_isplugin"] = true;
            if (checkForAllOrNot) // true for all components
            {
                createOrder["if_supporthours"] = retrieveAccount.Attributes.Contains("if_newhoursperterm") ? retrieveAccount.Attributes["if_newhoursperterm"] : null;


            }



            Guid getCreatedOrder  = getCreatedOrder = service.Create(createOrder);
                

            if (checkForAllOrNot) // true for all components
            {
                UpdateQuaterForAllOrder(service, retrieveSuppContract);

                if (isQuarter)
                {
                    String ordName = Convert.ToString(retrieveSuppContract.Attributes["if_titleforinvoice"]);
                    //Update Order Name 
                    Entity updateOrder = new Entity("salesorder", getCreatedOrder);
                    updateOrder["name"] = ordName + " " + quaterPrefix;
                    service.Update(updateOrder);
                }

                return getCreatedOrder;
            }

           
                bool orderProducCreated = CreateSalesOrderDetails(service, retrieveSuppContract, startDate, endDate, getCreatedOrder, retrieveAccount, true, termEndDate);

            


            if (orderProducCreated == false)
            {
                service.Delete("salesorder", getCreatedOrder);
                return Guid.Empty;

            }
            else
            {
                if (isQuarter)
                {
                    String ordName = Convert.ToString(retrieveSuppContract.Attributes["if_titleforinvoice"]);
                    //Update Order Name 
                    Entity updateOrder = new Entity("salesorder", getCreatedOrder);
                    updateOrder["name"] = ordName + " " + quaterPrefix;
                    service.Update(updateOrder);
                }
            }
            return Guid.Empty;
        }



        public void UpdateQuaterForAllOrder(IOrganizationService service, Entity retrieveSuppContract)
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



        public bool CreateSalesOrderDetails(IOrganizationService service, Entity retrieveSuppContract, DateTime supportStartDate, DateTime supportEndDate, Guid getCreatedOrder, Entity retrieveAccount, bool thisComp, DateTime termEndDate)
        {
            bool costCalculated = true;

            string contactenateMessage = string.Empty;
            decimal cost = 0;
            //calculate priceperunit
            if (retrieveSuppContract.Attributes.Contains("if_termduration"))
            {


                int getTermDuration = ((OptionSetValue)retrieveSuppContract.Attributes["if_termduration"]).Value;
                decimal costperterm = ((Money)retrieveSuppContract.Attributes["if_costperterm"]).Value;
                DateTime StartCalculationDate = DateTime.Now;
                double getDateDiff;

                if (getTermDuration == 100000002) // in case 12 months
                {

                    DateTime? startDateTime = calculateStartDate(supportStartDate, retrieveSuppContract);
                    DateTime? endDateTime = calculateEndDate(termEndDate, supportStartDate, retrieveSuppContract);

                    if (startDateTime != null && endDateTime != null)
                    {
                        DateTime startDate = startDateTime.Value;
                        DateTime endDate = endDateTime.Value;
                        getDateDiff = (endDate.Date - startDate.Date).TotalDays;//No# of Days

                        int totalDays = DateTime.IsLeapYear(DateTime.Now.Year) ? 366 : 365;
                        //double DurationInMonth = (getDateDiff) / 365 * 12;//No# of Months
                        cost = costperterm * (Convert.ToDecimal(getDateDiff) / totalDays);
                    }
                    else
                    {
                        costCalculated = false;
                        return costCalculated;
                    }

                    //calculate Product comment

                    if (startDateTime != null && endDateTime != null)
                    {
                        if (termEndDate.Year > DateTime.Now.Year)
                        {
                            contactenateMessage = startDateTime.Value.ToString("dd.MM.yyyy") + "-" + new DateTime(DateTime.Today.Year, 12, 31).ToString("dd.MM.yyyy");
                        }

                        else
                        {
                            contactenateMessage = startDateTime.Value.ToString("dd.MM.yyyy") + "-" + endDateTime.Value.AddDays(-1).ToString("dd.MM.yyyy");
                        }
                    }



                }



                else if (getTermDuration == 100000001) //In case of 3 Months
                {

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

                    double totalDays = 0;



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

                    DateTime? startDateTime = calculateStartDate(supportStartDate, EndCalculationDate, StartCalculationDate);
                    DateTime? endDateTime = calculateEndDate(termEndDate, supportStartDate, EndCalculationDate, StartCalculationDate);

                    if (startDateTime != null && endDateTime != null)
                    {
                        DateTime startDate = startDateTime.Value;
                        DateTime endDate = endDateTime.Value;

                        getDateDiff = (endDate.Date - startDate.Date).TotalDays;//No# of Days remaining
                        totalDays = (BeginOfFollowingTerm.Date - StartCalculationDate.Date).TotalDays;//total No# of Days
                                                                                                      //double DurationInMonth = getDateDiff / 365 * 12;//No# of Months
                        cost = costperterm * (Convert.ToDecimal(getDateDiff) / Convert.ToDecimal(totalDays));
                    }
                    else
                    {
                        costCalculated = false;
                        return costCalculated;
                    }

                    if (startDateTime != null && endDateTime != null)
                    {
                        if (endDateTime.Value == BeginOfFollowingTerm)
                        {
                            contactenateMessage = startDateTime.Value.ToString("dd.MM.yyyy") + "-" + EndCalculationDate.ToString("dd.MM.yyyy");
                        }

                        else
                        {
                            contactenateMessage = startDateTime.Value.ToString("dd.MM.yyyy") + "-" + endDateTime.Value.ToString("dd.MM.yyyy");
                        }
                    }





                }

                else if (getTermDuration == 100000000) //In case of 1 Months
                {
                    DateTime? startDateTime = calculateStartDate(supportStartDate, retrieveSuppContract);
                    DateTime? endDateTime = calculateEndDate(termEndDate, supportStartDate, retrieveSuppContract);

                    if (startDateTime != null && endDateTime != null)
                    {
                        DateTime startDate = startDateTime.Value;
                        DateTime endDate = endDateTime.Value;
                        getDateDiff = (endDate.Date - startDate.Date).TotalDays;//No# of Days

                        int totalDays = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month);
                        //double DurationInMonth = (getDateDiff) / 365 * 12;//No# of Months
                        cost = costperterm * (Convert.ToDecimal(getDateDiff) / totalDays);
                    }
                    else
                    {
                        costCalculated = false;
                        return costCalculated;
                    }

                    DateTime last_Date = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1).AddMonths(1).AddDays(-1);

                    DateTime currentDate = DateTime.Now;
                    DateTime firstMonthdate = new DateTime(currentDate.Year, currentDate.Month, 01);
                    DateTime lastMonthdate = new DateTime(last_Date.Year, currentDate.Month, last_Date.Day);



                    if (startDateTime != null && endDateTime != null)
                    {
                        if (endDateTime.Value.Month > DateTime.Now.Month || endDateTime.Value.Year > DateTime.Now.Year)
                        {
                            contactenateMessage = startDateTime.Value.ToString("dd.MM.yyyy") + "-" + endDateTime.Value.AddDays(-1).ToString("dd.MM.yyyy");
                        }

                        else
                        {
                            contactenateMessage = startDateTime.Value.ToString("dd.MM.yyyy") + "-" + termEndDate.AddDays(-1).ToString("dd.MM.yyyy");
                        }
                    }
                }


                //retriving the currency
                Guid currencyId = retrieveSuppContract.GetAttributeValue<EntityReference>("if_currency").Id;
                string currencyName = retrieveSuppContract.GetAttributeValue<EntityReference>("if_currency").Name;



                //10.08.2022-31.12.2022
                Entity createOrderDetail = new Entity("salesorderdetail");
                createOrderDetail["isproductoverridden"] = true;
                createOrderDetail["productdescription"] = retrieveSuppContract.Attributes["if_titleforinvoice"];

                createOrderDetail["if_productcomment"] = retrieveSuppContract.Attributes.Contains("if_commentforinv") ? retrieveSuppContract.Attributes["if_commentforinv"] + ": " + contactenateMessage : contactenateMessage;
                createOrderDetail["crd38_isalicense_new"] = new OptionSetValue(0);

                createOrderDetail["salesorderid"] = new EntityReference("salesorder", getCreatedOrder);
                if (cost != 0)
                {
                    createOrderDetail["priceperunit"] = new Money((decimal)cost);
                }




                Guid countryId = ((EntityReference)retrieveAccount.Attributes["crd38_country"]).Id;
                Entity Country = service.Retrieve("crd38_country", countryId, new ColumnSet("inn_vat"));

                decimal salesTax = 0;
                decimal taxAmount;
                decimal extendedAmount;

                if (Country.Contains("inn_vat")) // calculating total amount if the currency if CHF
                {
                     salesTax = (decimal)Country.Attributes["inn_vat"];
                }


                taxAmount = (cost / 100) * salesTax; // calculate tax amount
                extendedAmount = taxAmount + cost;

                createOrderDetail["if_salestaxpercent"] = salesTax;
                createOrderDetail["tax"] = new Money((decimal)taxAmount);
                createOrderDetail["extendedamount"] = new Money((decimal)extendedAmount);


                createOrderDetail["quantity"] = Convert.ToDecimal(1);
                service.Create(createOrderDetail);
            }

            return costCalculated;
        }


        public DateTime? calculateStartDate(DateTime supportStartDate, Entity retrieveSuppContract)
        {
            DateTime? startTime = null;
            int getTermDuration = ((OptionSetValue)retrieveSuppContract.Attributes["if_termduration"]).Value;

            if (getTermDuration == 100000002) // for 12 month case
            {
                if (supportStartDate.Year == DateTime.Now.Year)
                {
                    startTime = supportStartDate;
                }
                else if (supportStartDate.Year < DateTime.Now.Year)
                {
                    startTime = new DateTime(DateTime.Now.Year, 01, 01);
                }

                return startTime;
            }


            else // for 1 month case
            {
                if (supportStartDate.Month == DateTime.Now.Month && supportStartDate.Year == DateTime.Now.Year)
                {

                    startTime = supportStartDate;

                }
                else if (supportStartDate.Month < DateTime.Now.Month || supportStartDate.Year < DateTime.Now.Year)
                {
                    startTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 01);
                }

                else
                {
                    startTime = DateTime.Now;
                }

                return startTime;
            }

        } // calculation of start date for 1 and 12 months

        public DateTime? calculateStartDate(DateTime supportStartDate, DateTime quaterEndDate, DateTime quaterStartDate)
        {
            DateTime? startTime = null;

            if (supportStartDate >= quaterStartDate && supportStartDate <= quaterEndDate)
            {
                startTime = supportStartDate;
            }
            else if (supportStartDate < quaterStartDate)
            {
                startTime = quaterStartDate;
            }

            return startTime;
        } // overloaded function for calculation of start date for 3 months

        public DateTime? calculateEndDate(DateTime termEndDate, DateTime supportStartDate, Entity retrieveSuppContract)
        {
            DateTime? endDate = null;
            int getTermDuration = ((OptionSetValue)retrieveSuppContract.Attributes["if_termduration"]).Value;

            if (getTermDuration == 100000002) // case for 12 month 
            {
                if (termEndDate == DateTime.MinValue)
                {
                    endDate = new DateTime(DateTime.Now.Year + 1, 01, 01);
                }

                else
                {
                    if (termEndDate.Date <= supportStartDate.Date)
                    {
                        endDate = null;
                    }
                    else if (termEndDate.Year == DateTime.Now.Year && termEndDate.Year == DateTime.Now.Year)
                    {

                        endDate = termEndDate;
                    }
                    else if (termEndDate.Year > DateTime.Now.Year)
                    {
                        endDate = new DateTime(DateTime.Now.Year + 1, 01, 01);
                    }
                }


                return endDate;
            }

            else // case for 1 month 
            {
                if (termEndDate == DateTime.MinValue)
               {
                    if(DateTime.Now.Month == 12)
                    {
                        endDate = new DateTime(DateTime.Now.Year+1, 01, 01);
                    }
                    else
                    {
                        endDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month + 1, 01);

                    }
                }

                else
                {
                    if (termEndDate.Date <= supportStartDate.Date)
                    {
                        endDate = null;
                    }
                    else if (termEndDate.Month == DateTime.Now.Month && termEndDate.Year == DateTime.Now.Year)
                    {

                        endDate = termEndDate;
                    }
                    else if (termEndDate.Year == DateTime.Now.Year && termEndDate.Month > DateTime.Now.Month)
                    {
                        if(DateTime.Now.Month == 12)
                        {
                            endDate = new DateTime(DateTime.Now.Year + 1, 01, 01);
                        }
                        else
                        {
                            endDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month + 1, 01);
                        }
                        
                    }
                }


                return endDate;
            }

        } // calculation of end date for 1 and 12 months

        public DateTime? calculateEndDate(DateTime termEndDate, DateTime supportStartDate, DateTime quaterEndDate, DateTime quaterStartDate)
        {
            DateTime? endDate = null;

            if (termEndDate == DateTime.MinValue)
            {
                if (quaterPrefix == "Q4/" + DateTime.Now.Year || DateTime.Now.Month == 12)
                {
                    //01,01,
                    endDate = new DateTime(quaterEndDate.Year + 1, 01, 01);
                }
                else
                    endDate = new DateTime(quaterEndDate.Year, quaterEndDate.Month + 1, 01);
            }
            else
            {
                if (termEndDate.Date <= supportStartDate.Date || termEndDate <= quaterStartDate)
                {
                    endDate = null;
                }         
                else if (quaterEndDate.Date >= termEndDate.Date && quaterStartDate.Date <= termEndDate.Date)
                {
                    endDate = termEndDate;
                }
                else if (termEndDate > quaterEndDate)
                {
                    if (quaterPrefix == "Q4/" + DateTime.Now.Year || DateTime.Now.Month == 12)
                    {
                        endDate = new DateTime(quaterEndDate.Year + 1, 01, 01);
                    }
                    else
                        endDate = new DateTime(quaterEndDate.Year, quaterEndDate.Month + 1, 01);

                }

            }

            return endDate;
        } // overloaded function for calculation of end date for 3 months
    }

}
