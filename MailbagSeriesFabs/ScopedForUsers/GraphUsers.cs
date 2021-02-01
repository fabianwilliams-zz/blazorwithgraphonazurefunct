using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;
using System.Collections.Generic;
using System.Linq;

namespace ScopedForUsers
{
    public static class GraphUsers
    {

        [FunctionName("GetTopOneUser")]
        public static async Task<IActionResult> GetTopOneUserFromGraph(
            [HttpTrigger(AuthorizationLevel.Function, "get", Route = null)] HttpRequest req,
            ILogger log)
        {

            GraphServiceClient graphClient = GetAuthenticatedGrahClient();
            List<QueryOption> options = new List<QueryOption>
            {
                new QueryOption("$top", "1")
            };

            var graphResult = graphClient.Users.Request(options).GetAsync().Result;

            log.LogInformation("Graph SDK Result for Top 1 User");
            log.LogInformation(graphResult[0].DisplayName);

            string responseMessage = string.IsNullOrEmpty(graphResult[0].ToString())
                ? "Call to Microsoft Graph on Fabster Tenanat App executed successfully."
                : $"{graphResult[0].DisplayName}";

            return new OkObjectResult(responseMessage);
        }

        [FunctionName("GetAllUsers")]
        public static async Task<IActionResult> GetAllUsersFromGraph(
            [HttpTrigger(AuthorizationLevel.Function, "get", Route = null)] HttpRequest req,
            ILogger log)
        {
            List<FabsterUser> ful = new List<FabsterUser>();

            GraphServiceClient graphClient = GetAuthenticatedGrahClient();
            List<QueryOption> options = new List<QueryOption>
            {
                new QueryOption("$select", "displayName,givenName,mail")
            };

            var graphResult = graphClient.Users.Request(options).GetAsync().Result;

            List<User> usersList = graphResult.CurrentPage.ToList();

            foreach (User u in usersList)
            {
                log.LogInformation("Showing: " + u.GivenName + " - " + u.Mail);
            }

            log.LogInformation("Graph SDK Result for All Users");

            string responseMessage = string.IsNullOrEmpty(graphResult.ToString())
                ? "Call to Microsoft Graph on Fabster Tenanat App executed successfully."
                : $"{graphResult}";

            return new OkObjectResult(usersList);
        }
        private static GraphServiceClient GetAuthenticatedGrahClient()
        {
            //The below comment block is how you should go about securing your configurable keys etc as it will allow you to 
            // send them to Azure API settings upon publish as well as set them up for KeyValult. but you dont 
            //get to eee my local.settings.json file so I have it here direct as well uncommented so you can see options.
            /*
            var clientId = Environment.GetEnvironmentVariable("AzureADAppClientId", EnvironmentVariableTarget.Process);
            var tenantID = Environment.GetEnvironmentVariable("AzureADAppTenantId", EnvironmentVariableTarget.Process);
            var clientSecret = Environment.GetEnvironmentVariable("AzureADAppClientSecret", EnvironmentVariableTarget.Process);
            */
            var clientId = "REDACTED";
            var tenantID = "REDACTED";
            var clientSecret = "REDACTED";

            // Build a client application.
            IConfidentialClientApplication confidentialClientApplication = ConfidentialClientApplicationBuilder
                            .Create(clientId)
                            .WithTenantId(tenantID)
                            .WithClientSecret(clientSecret)
                            .Build();
            ClientCredentialProvider authenticationProvider = new ClientCredentialProvider(confidentialClientApplication);
            GraphServiceClient graphClient = new GraphServiceClient(authenticationProvider);

            return graphClient;
        }

    }
}
