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
        [FunctionName("GetAccessKey")]
        public static async Task<string> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            string code = req.Query["code"];
            string username = req.Query["user"];

            string grant_type = "authorization_code";

            bool refreshToken = false;


            string url = @"https://login.microsoftonline.com/41fe572f-d61a-4d7a-90df-003afcdaa2c9/oauth2/v2.0/token";

            HttpClient client = new HttpClient();
            

            CloudStorageAccount storageAccount =
              CloudStorageAccount.Parse("DefaultEndpointsProtocol=https;AccountName=msgraphtest2;AccountKey=31GZVOmuKw4HqvPodSNSMGXZOo7KaPYWa4KV+oppm5wOWp2/pb7EIzd5zFaKTT/ygIUevVop/lonUx3p+N5XMg==;EndpointSuffix=core.windows.net");

            CloudTableClient tableClient = storageAccount.CreateCloudTableClient();
            CloudTable table = tableClient.GetTableReference("appUsers");

            try
            {
                TableOperation readOperation = TableOperation.Retrieve<User>(username, username);


                // Execute the insert operation.
                TableResult query = await table.ExecuteAsync(readOperation);

                if(query.Result != null)
                {
                    code = ((User)query.Result).refreshToken;
                    grant_type = "refresh_token";
                }

                
            }
            catch (Exception)
            {

                
            }            

            var values = new Dictionary<string, string>
                {
                   { "grant_type", "authorization_code" },                  
                   { "client_id", "a02f6ab7-1acd-4bfc-97d5-992acc1301b1" },
                   { "scope", "https://graph.microsoft.com/User.Read offline_access" },
                   { "redirect_uri", "https://graphsharepoint.azurewebsites.net/api/GetToken" },
                   { "client_secret", "0rnPpWy6q4H//oFqVOf3u73vC9vUeZ7oOrYEvFy+nro=" },
                };

            if(refreshToken)
            {
                values.Add("refresh_token", code);
            }
            else
            {
                values.Add("code",code);
            }
            

            var content = new FormUrlEncodedContent(values);

            var response = await client.PostAsync(url, content);

            dynamic responseString = await response.Content.ReadAsStringAsync();

            await table.CreateIfNotExistsAsync();


            User user = new User
            {
                RowKey = username,
                name = username,
                accessToken = responseString.access_token,
                refreshToken = responseString.refresh_token,
                PartitionKey = username
            };

            if(refreshToken)
            {
                user.ETag = "*";
                TableOperation replaceOperation = TableOperation.Replace(user);
                await table.ExecuteAsync(replaceOperation);
            }
            else
            {
                TableOperation insertOperation = TableOperation.Insert(user);
                await table.ExecuteAsync(insertOperation);
            }
           

            return responseString.access_token;
        }

        [FunctionName("GetToken")]
        public static async Task<HttpResponseMessage> GetToken(
           [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
           ILogger log)
        {
            return new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(req.Query["code"], Encoding.UTF8, "application/json")
            };
        }

    }
}

