using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;

namespace AssignPlugin
{
    public class AssignTeamToUser : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            if (serviceProvider != null)
            {
                // Obtain the Plugin Execution Context
                IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

                // Obtain the organization service reference.
                IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = factory.CreateOrganizationService(context.UserId);
                
                //For Trace Logs
                ITracingService tracing = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

                // To check depth 
                if (context.Depth <= 2)
                {
                    // The InputParameters collection contains all the data passed in the message request.
                    if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is EntityReference)
                    {
                        // Business logic goes here
                        tracing.Trace("Running Business Logic");
                        new AssignTeamToUser().BusinessLogic(service, context, tracing);
                        tracing.Trace("Completed Business Logic");
                    }
                }
            }
        }

        private void BusinessLogic(IOrganizationService service, IExecutionContext context, ITracingService tracing)
        {
            // We are hitting Assign button on a Case record so Target will be a Case record
            var caseEntityReference = (EntityReference)context.InputParameters["Target"];

            // Assignee could be a User or Team
            var teamEntityReference = (EntityReference)context.InputParameters["Assignee"];

            // In our requirement it should be a Team, if user it should return
            if (teamEntityReference.LogicalName != "team") return;

            // fetchXml to retrieve all the users in a given Team
            var fetchXmlLoggedUserInTeam = @"
            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
              <entity name='systemuser'>
                <attribute name='systemuserid' />
                <link-entity name='teammembership' from='systemuserid' to='systemuserid' visible='false' intersect='true'>
                  <link-entity name='team' from='teamid' to='teamid' alias='ac'>
                    <filter type='and'>
                      <condition attribute='teamid' operator='eq' uitype='team' value='{0}' />
                    </filter>
                  </link-entity>
                </link-entity>
              </entity>
            </fetch>";

            // Passing current Team is fetchXml and retrieving Team's user
            var users = service.RetrieveMultiple(new FetchExpression(string.Format(
                fetchXmlLoggedUserInTeam,
                teamEntityReference.Id))).Entities;

            // Condition 1
            // If user count is zero case should be assigned to Team's default Queue
            if (users.Count == 0)
            {
                // Retrieving Team's default Queue
                var team = service.Retrieve("team", teamEntityReference.Id, new ColumnSet("queueid"));

                var addToQueueRequest = new AddToQueueRequest
                {
                    // Case record
                    Target = caseEntityReference,
                    // Team's default Queue Id
                    DestinationQueueId = team.GetAttributeValue<EntityReference>("queueid").Id
                };
                service.Execute(addToQueueRequest);
            }
            else
            {
                // Dictionary to save UserId and number of case assigned pair
                var caseCountAssignedToUser = new Dictionary<Guid, int>();

                users.ToList().ForEach(user =>
                {
                    // FetchXml query to retrieve number cases assigned to each user
                    var fetchXmlCaseAssignedToUser = @"
                    <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                      <entity name='incident'>
                        <attribute name='incidentid' />
                        <link-entity name='systemuser' from='systemuserid' to='owninguser' alias='ab'>
                          <filter type='and'>
                            <condition attribute='systemuserid' operator='eq' uitype='systemuser' value='{0}' />
                          </filter>
                        </link-entity>
                      </entity>
                    </fetch>";

                    var cases = service.RetrieveMultiple(new FetchExpression(string.Format(
                        fetchXmlCaseAssignedToUser,
                        user.Id))).Entities.ToList();

                    // Adding user id with number of cases assigned to Dictionay defined above
                    caseCountAssignedToUser.Add(user.Id, cases.Count);
                });

                // Sorting in ascending order by number of cases
                var sortedCaseCount = from entry in caseCountAssignedToUser
                                      orderby entry.Value ascending
                                      select entry;

                // Getting all the users with least sae number of cases
                var allUserWithSameLeastNumberOfCases = sortedCaseCount
                    .Where(w => w.Value == sortedCaseCount.First().Value).ToList();

                var targetUser = new Guid();

                // Condition 1
                // Assign case to user with least number of case assigned
                if (allUserWithSameLeastNumberOfCases.Count() == 1)
                {
                    targetUser = sortedCaseCount.First().Key;
                }

                // Condition 2
                // If more than one users are having same least number of users, then it be assigned randomly
                else
                {
                    var randomUser = new Random().Next(0, allUserWithSameLeastNumberOfCases.Count() - 1);
                    targetUser = allUserWithSameLeastNumberOfCases[randomUser].Key;
                }

                var assign = new AssignRequest
                {
                    Assignee = new EntityReference("systemuser", targetUser),
                    Target = caseEntityReference
                };

                service.Execute(assign);
            }
        }
    }
}
