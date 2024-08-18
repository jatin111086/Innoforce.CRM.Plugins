using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Innoforce.CRM.Plugins
{
    public class CreateProjectHistory : IPlugin
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
                    // Obtain the target entity from the input parameters.
                    Entity Project = (Entity)context.InputParameters["Target"];

                    if (Project.LogicalName != "crd38_project") return;

                    string messageName = context.MessageName.ToLower();


                    if (messageName == "create")
                    {
                        //throw new InvalidPluginExecutionException("By default error");
                        Entity ProjectHistory = new Entity("crd38_projecthistory");

                        Entity PreImageProject = context.PostEntityImages["PostImage"];

                        //Project Name
                        ProjectHistory.Attributes["crd38_projectname"] = new EntityReference("crd38_project", Project.Id);

                        //Change Status
                        if (Project.Contains("crd38_status"))
                        {
                            ProjectHistory.Attributes["crd38_oldstatus"] = !string.IsNullOrEmpty(Project.FormattedValues["crd38_status"]) == true ? Convert.ToString(Project.FormattedValues["crd38_status"]) : string.Empty;
                        }

                        //From
                        if (Project.Contains("crd38_start"))
                        {
                            DateTime dateTime;
                            bool isConverted = DateTime.TryParse(Convert.ToString(Project.Attributes["crd38_start"]), out dateTime);

                            if (isConverted)
                            {
                                ProjectHistory.Attributes["crd38_from"] = dateTime;
                            }
                        }

                        //Owner
                        ProjectHistory.Attributes["crd38_ownerid"] = new EntityReference("systemuser", userId);


                        //Required Action by Customer
                        if (Project.Contains("crd38_requiredcustomeractions"))
                        {
                            ProjectHistory.Attributes["crd38_eequiredactionsbycustomer"] = !string.IsNullOrEmpty(Convert.ToString(Project.Attributes["crd38_requiredcustomeractions"])) == true ? Convert.ToString(Project.Attributes["crd38_requiredcustomeractions"]) : string.Empty;
                        }

                        //Required Action by IPS
                        if (Project.Contains("crd38_requiredipsactions"))
                        {
                            ProjectHistory.Attributes["crd38_requiredactionsbyips"] = !string.IsNullOrEmpty(Convert.ToString(Project.Attributes["crd38_requiredipsactions"])) == true ? Convert.ToString(Project.Attributes["crd38_requiredipsactions"]) : string.Empty;
                        }

                        //Last M&S-IPS meetings
                        if (Project.Contains("tcxcrm_lastmsipsmeeting"))
                        {
                            DateTime dateTimeMeeting;
                            bool isConverted = DateTime.TryParse(Convert.ToString(Project.Attributes["tcxcrm_lastmsipsmeeting"]), out dateTimeMeeting);

                            if (isConverted)
                            {
                                ProjectHistory.Attributes["tcxcrm_lastmsipsmeeting"] = dateTimeMeeting;
                            }
                        }


                        service.Create(ProjectHistory);
                    }


                    if (messageName == "update")
                    {

                        Entity ProjectHistory = new Entity("crd38_projecthistory");

                        Entity PreImageProject = PreImageProject = context.PreEntityImages["PreImage"];
                        Entity PostImageProject = context.PostEntityImages["PreImage"];


                        //Project Name
                        ProjectHistory.Attributes["crd38_projectname"] = new EntityReference("crd38_project", Project.Id);


                        //Status
                        if (PostImageProject.Contains("crd38_status"))
                        {
                            ProjectHistory.Attributes["crd38_oldstatus"] = !string.IsNullOrEmpty(PostImageProject.FormattedValues["crd38_status"]) == true ? Convert.ToString(PostImageProject.FormattedValues["crd38_status"]) : string.Empty;
                        }


                        //From
                        if (PostImageProject.Contains("crd38_start"))
                        {
                            DateTime dateTime;
                            bool isConverted = DateTime.TryParse(Convert.ToString(PostImageProject.Attributes["crd38_start"]), out dateTime);

                            if (isConverted)
                            {
                                ProjectHistory.Attributes["crd38_from"] = dateTime.AddDays(1);
                            }
                        }


                        //Update until in last Project History Record
                        UpdateProjectHistory(service, Project.Id, PostImageProject);


                        //Owner
                        ProjectHistory.Attributes["crd38_ownerid"] = new EntityReference("systemuser", userId);

                        //Required Action by Customer
                        if (PostImageProject.Contains("crd38_requiredcustomeractions"))
                        {
                            ProjectHistory.Attributes["crd38_eequiredactionsbycustomer"] = !string.IsNullOrEmpty(Convert.ToString(PostImageProject.Attributes["crd38_requiredcustomeractions"])) == true ? Convert.ToString(PostImageProject.Attributes["crd38_requiredcustomeractions"]) : string.Empty;

                        }

                        //Required Action by IPS
                        if (PostImageProject.Contains("crd38_requiredipsactions"))
                        {
                            ProjectHistory.Attributes["crd38_requiredactionsbyips"] = !string.IsNullOrEmpty(Convert.ToString(PostImageProject.Attributes["crd38_requiredipsactions"])) == true ? Convert.ToString(PostImageProject.Attributes["crd38_requiredipsactions"]) : string.Empty;
                        }

                        //Last M&S-IPS meetings
                        if (Project.Contains("tcxcrm_lastmsipsmeeting"))
                        {
                            DateTime dateTimeMeeting;
                            bool isConverted = DateTime.TryParse(Convert.ToString(Project.Attributes["tcxcrm_lastmsipsmeeting"]), out dateTimeMeeting);

                            if (isConverted)
                            {
                                ProjectHistory.Attributes["tcxcrm_lastmsipsmeeting"] = dateTimeMeeting;
                            }
                        }

                        service.Create(ProjectHistory);
                    }
                }
                catch (Exception ex)
                {
                    Entity Project = (Entity)context.InputParameters["Target"];
                    string strProjectID = Convert.ToString(Project.Id);
                    Entity log = new Entity("crd38_log");
                    var projName = string.Empty;
                    if (Project.Contains("crd38_name"))
                    {
                        projName = !string.IsNullOrEmpty(Convert.ToString(Project.Attributes["crd38_name"])) == true ? Convert.ToString(Project.Attributes["crd38_name"]) : string.Empty;
                    }
                    log.Attributes["crd38_name"] = "Project Plugin Error On " + context.MessageName + " - " + projName;
                    log.Attributes["crd38_tablename"] = Project.LogicalName;
                    log.Attributes["crd38_type"] = "Error";
                    log.Attributes["crd38_recordid"] = strProjectID;
                    log.Attributes["crd38_description"] = ex.Message;
                    // log.Attributes["crd38_name"] = new EntityReference("crd38_project", Project.Id);
                    service.Create(log);

                    //throw new InvalidPluginExecutionException(ex.Message);
                }





            }
        }


        /// <summary>
        /// Update Until date of last Project History Log
        /// </summary>
        /// <param name="service"></param>
        /// <param name="projectId"></param>
        /// <param name="PreImageProject"></param>
        private void UpdateProjectHistory(IOrganizationService service, Guid projectId, Entity PostImageProject)
        {
            try
            {
                if (PostImageProject.Contains("crd38_start"))
                {
                    var ProjectHistoryCollection = GetLastProjectHistory(service, projectId);

                    if (ProjectHistoryCollection != null && ProjectHistoryCollection.Entities.Count > 0)
                    {
                        Entity ProjectHistory = new Entity("crd38_projecthistory", ProjectHistoryCollection.Entities[0].Id);

                        DateTime dateTime;
                        bool isConverted = DateTime.TryParse(Convert.ToString(PostImageProject.Attributes["crd38_start"]), out dateTime);

                        if (isConverted)
                        {
                            ProjectHistory.Attributes["crd38_until"] = dateTime.AddDays(1);

                            service.Update(ProjectHistory);
                        }
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        private EntityCollection GetLastProjectHistory(IOrganizationService service, Guid projectId)
        {
            EntityCollection ProjectHistoryCollection = null;

            try
            {
                if (projectId != Guid.Empty)
                {
                    ProjectHistoryCollection = service.RetrieveMultiple(new FetchExpression(ProjectHistoryFetch.Replace("ProjectId", projectId.ToString())));
                }
            }
            catch (Exception)
            {
                throw;
            }

            return ProjectHistoryCollection;
        }




        #region Global Variables




        string ProjectHistoryFetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                          <entity name='crd38_projecthistory'>
                                            <attribute name='crd38_projecthistoryid' />
                                            <attribute name='crd38_oldstatus' />
                                            <attribute name='createdon' />
                                            <order attribute='crd38_oldstatus' descending='false' />
                                            <filter type='and'>
                                              <condition attribute='crd38_until' operator='null' />
                                              <condition attribute='crd38_projectname' operator='eq' uiname='1' uitype='crd38_project' value='{ProjectId}' />
                                            </filter>
                                          </entity>
                                        </fetch>";



        #endregion

    }
}
