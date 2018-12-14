using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System.Net;
using System.Net.Http;
using System.Collections.Generic;
using System.Text;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Table;

namespace GraphAPI
{
    public static class Graph
    {
        [FunctionName("GetAccessToken")]
        public static async Task<string> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            string code = req.Query["code"];
            string username = req.Query["user"];

            string grant_type = "authorization_code";


            string url = @"https://login.microsoftonline.com/41fe572f-d61a-4d7a-90df-003afcdaa2c9/oauth2/v2.0/token";

            HttpClient client = new HttpClient();
            

            CloudStorageAccount storageAccount =
              CloudStorageAccount.Parse("DefaultEndpointsProtocol=https;AccountName=msgraphtest2;AccountKey=31GZVOmuKw4HqvPodSNSMGXZOo7KaPYWa4KV+oppm5wOWp2/pb7EIzd5zFaKTT/ygIUevVop/lonUx3p+N5XMg==;EndpointSuffix=core.windows.net");

            CloudTableClient tableClient = storageAccount.CreateCloudTableClient();
            CloudTable table = tableClient.GetTableReference("appUsers");

            TableOperation retrieveOperation = TableOperation.Retrieve<User>(username, username);

            TableResult retrievedResult = await table.ExecuteAsync(retrieveOperation);

            User deleteEntity = (User)retrievedResult.Result;

            if (deleteEntity != null)
            {
                TableOperation deleteOperation = TableOperation.Delete(deleteEntity);
                await table.ExecuteAsync(deleteOperation);

            }

            var values = new Dictionary<string, string>
                {
                   { "grant_type", "authorization_code" },                  
                   { "client_id", "a02f6ab7-1acd-4bfc-97d5-992acc1301b1" },
                   { "scope", "https://graph.microsoft.com/User.Read offline_access" },
                   { "redirect_uri", "https://graphsharepoint.azurewebsites.net/api/GetToken" },
                   { "client_secret", "0rnPpWy6q4H//oFqVOf3u73vC9vUeZ7oOrYEvFy+nro=" },
                   { "code",code}
                };

            

            var content = new FormUrlEncodedContent(values);

            var response = await client.PostAsync(url, content);

            dynamic responseString = await response.Content.ReadAsAsync<object>();

            await table.CreateIfNotExistsAsync();


            User user = new User
            {
                RowKey = username,
                name = username,
                accessToken = responseString.access_token,
                refreshToken = responseString.refresh_token,
                PartitionKey = username
            };

            TableOperation insertOperation = TableOperation.Insert(user);
            await table.ExecuteAsync(insertOperation);
           

            return responseString.access_token;
        }

        [FunctionName("GetAuthToken")]
        public static async Task<HttpResponseMessage> GetToken(
           [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
           ILogger log)
        {
            return new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(req.Query["code"], Encoding.UTF8, "application/json")
            };
        }

        [FunctionName("CallGraph")]
        public static async Task<string> CallGraph(
           [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequestMessage req,
           ILogger log)
        {
            dynamic data = await req.Content.ReadAsAsync<object>();
            string username = data.username;
            string url = data.url;

            CloudStorageAccount storageAccount =
              CloudStorageAccount.Parse("DefaultEndpointsProtocol=https;AccountName=msgraphtest2;AccountKey=31GZVOmuKw4HqvPodSNSMGXZOo7KaPYWa4KV+oppm5wOWp2/pb7EIzd5zFaKTT/ygIUevVop/lonUx3p+N5XMg==;EndpointSuffix=core.windows.net");

            CloudTableClient tableClient = storageAccount.CreateCloudTableClient();
            CloudTable table = tableClient.GetTableReference("appUsers");
            TableOperation retrieveOperation = TableOperation.Retrieve<User>(username, username);
            TableResult retrievedResult = await table.ExecuteAsync(retrieveOperation);

            HttpClient client = new HttpClient();   

            client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", ((User)retrievedResult.Result).accessToken);

            var response = await client.GetAsync(url);

            string responseString = await response.Content.ReadAsStringAsync();

            return responseString;
        }

        [FunctionName("GetAccessKeyFromRefresh")]
        public static async Task<string> GetAccessKeyFromRefresh(
           [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequestMessage req,
           ILogger log)
        {
            dynamic data = await req.Content.ReadAsAsync<object>();
            string username = data.username;
            string url = "https://login.microsoftonline.com/41fe572f-d61a-4d7a-90df-003afcdaa2c9/oauth2/v2.0/token";

            CloudStorageAccount storageAccount =
              CloudStorageAccount.Parse("DefaultEndpointsProtocol=https;AccountName=msgraphtest2;AccountKey=31GZVOmuKw4HqvPodSNSMGXZOo7KaPYWa4KV+oppm5wOWp2/pb7EIzd5zFaKTT/ygIUevVop/lonUx3p+N5XMg==;EndpointSuffix=core.windows.net");

            CloudTableClient tableClient = storageAccount.CreateCloudTableClient();
            CloudTable table = tableClient.GetTableReference("appUsers");
            TableOperation retrieveOperation = TableOperation.Retrieve<User>(username, username);
            TableResult retrievedResult = await table.ExecuteAsync(retrieveOperation);

            HttpClient client = new HttpClient();

            var values = new Dictionary<string, string>
                {
                   { "grant_type", "refresh_token" },
                   { "client_id", "a02f6ab7-1acd-4bfc-97d5-992acc1301b1" },
                   { "scope", "https://graph.microsoft.com/User.Read offline_access" },
                   { "redirect_uri", "https://graphsharepoint.azurewebsites.net/api/GetToken" },
                   { "client_secret", "0rnPpWy6q4H//oFqVOf3u73vC9vUeZ7oOrYEvFy+nro=" },
                   { "refresh_token",((User)retrievedResult.Result).refreshToken}
                };

            var content = new FormUrlEncodedContent(values);

            var response = await client.PostAsync(url, content);

            dynamic responseString = await response.Content.ReadAsAsync<object>();

            User updateUser = (User)retrievedResult.Result;
            updateUser.refreshToken = responseString.refresh_token;
            updateUser.accessToken = responseString.access_token;
            TableOperation updateOperation = TableOperation.Replace(updateUser);
            await table.ExecuteAsync(updateOperation);

            return responseString.access_token;

        }

    }
}

