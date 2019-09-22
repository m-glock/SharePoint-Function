using System.Net.Http;
using System.Net.Http.Headers;
using Microsoft.Graph;
using Microsoft.Extensions.Logging;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Nancy.Json;

namespace SharePointConnector
{
    public static class FileHandler
    {
        public static string IndexFiles(IGraphServiceClient client, string downloadDirectory, string collectionName, ILogger log)
        {
            Dictionary<string, Dictionary<string, string>> allFiles = new Dictionary<string, Dictionary<string, string>>();

            int count = 0;

            log.LogInformation("Call to the Microsoft Graph in order to access the files in SharePoint.");
            IGraphServiceGroupsCollectionPage groups = client.Groups.Request().GetAsync().Result;
            foreach (Group group in groups)
            {
                IDriveItemChildrenCollectionPage driveItems = client.Groups[group.Id].Drive.Root.Children.Request().GetAsync().Result;
                foreach (DriveItem item in driveItems)
                {
                    string itemId = item.ETag.Replace("\"{", "").Replace("},2\"", "");
                    Stream content = client.Groups[group.Id].Drive.Items[itemId].Content.Request().GetAsync().Result;

                    //save file in folder to index it from there
                    log.LogInformation("Download file " + item.Name);
                    var fileStream = System.IO.File.Create(downloadDirectory + "/" + item.Name);
                    content.Seek(0, SeekOrigin.Begin);
                    content.CopyTo(fileStream);
                    fileStream.Close();

                    //save values for file that need to be added
                    log.LogInformation("Extract information for file " + item.Name);
                    Dictionary<string, string> data = new Dictionary<string, string>();
                    data.Add("fileURL", item.WebUrl);
                    data.Add("fileCreator", item.CreatedBy.User.DisplayName);
                    allFiles.Add(item.Name, data);
                    count++;
                    if (count%10 == 0)
                    {
                        log.LogInformation("Send request for 10 files and delete them from the local file system.");
                        string json = prepareJSONBody(collectionName, allFiles);
                        SendIndexingRequest(json, downloadDirectory, collectionName);
                        
                        foreach (string fileName in allFiles.Keys)
                        {
                            System.IO.File.Delete(downloadDirectory + "/" + fileName);
                        }
                        allFiles.Clear();
                    }
                }
            }

            return "Successfully indexed.";
        }

        private static string prepareJSONBody(string collectionName, Dictionary<string, Dictionary<string, string>> allFiles)
        {
            FileDto fileinfo = new FileDto()
            {
                collectionName = collectionName,
                Files = allFiles
            };
            return new JavaScriptSerializer().Serialize(fileinfo);
        }

        private static Task<string> SendIndexingRequest(string jsonObject, string downloadPath, string collectionName)
        {
            var httpClientHandler = new HttpClientHandler();
            var httpClient = new HttpClient(httpClientHandler);
            httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("bearer", AuthenticationHelper.token);

            string url = "http://localhost:7071/api/IndexFilesToSolr?name=" + downloadPath;
            var content = new StringContent(jsonObject, Encoding.UTF8, "application/json");
            httpClient.DefaultRequestHeaders.Add("collection", collectionName);
            var response = httpClient.PostAsync(url, content).Result;
            var responseString = response.Content.ReadAsStringAsync();
            return responseString;
        }
    }
}