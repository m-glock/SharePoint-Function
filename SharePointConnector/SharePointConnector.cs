using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;

namespace SharePointConnector
{
    public static class Function1
    {

        public static GraphServiceClient client;

        [FunctionName("SharePointConnector")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");
            string collectionName = req.Query["name"];
            string tempDirectory = Path.Combine(Path.GetTempPath().ToString(), "SharePointFiles");

            if (!System.IO.Directory.Exists(tempDirectory))
                System.IO.Directory.CreateDirectory(tempDirectory);
               
            if (collectionName == null)
                new BadRequestObjectResult("Please pass a collection name on the query string");

            log.LogInformation("Authenticate with Microsoft.");
            client = AuthenticationHelper.GetAuthenticatedClientForApp();
            log.LogInformation("Authentication successful.");
            string response = FileHandler.IndexFiles(client, tempDirectory, collectionName, log);
            System.IO.Directory.Delete(tempDirectory, true);

            return (ActionResult)new OkObjectResult(response); 
        }
    } 
}