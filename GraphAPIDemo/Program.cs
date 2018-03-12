using Microsoft.Azure.ActiveDirectory.GraphClient;
using Microsoft.Azure.ActiveDirectory.GraphClient.Extensions;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace GraphAPIDemo
{
    class Program
    {
        //static string tenant = "dovahkiin.onmicrosoft.com";
        //static string clientID = "b4ac67df-89f2-4ae0-a64d-87386f081570";
        //static string secret = "v9CEUUj0PL8942ln3S1mhi2JTDDc5j5m20BNdyomgaE=";

        static string tenant = "sagiusoutlook.onmicrosoft.com";
        static string clientID = "96b882a0-af11-4cf1-9e68-3adcacbffb78";
        static string secret = "uiTGYSrfAQkXV0S4Uu3mz+9svKeGPLXEyU/FOTdveYE=";

        static void Main(string[] args)
        {
            run().Wait();
        }

        private static async Task run()
        {
            //var token = await AppAuthenticationAsync();
            var token = await HttpAppAuthenticationAsync();

            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

                var user = "Jie Li";
                var userExist = await DoesUserExistsAsync(client, user);

                Console.WriteLine(userExist);

                //await CreateUserAsync(client, "amy.hall", tenant);

                await DeleteUserAsync(client, "amy.wells", tenant);
            }
        }

        private static async Task<string> AppAuthenticationAsync()
        {
            //  Constants
            var resource = "https://graph.microsoft.com/";

            //var tenant = "dovahkiin.onmicrosoft.com";
            //var clientID = "b4ac67df-89f2-4ae0-a64d-87386f081570";
            //var secret = "v9CEUUj0PL8942ln3S1mhi2JTDDc5j5m20BNdyomgaE=";

            //  Ceremony
            var authority = $"https://login.microsoftonline.com/{tenant}";
            var authContext = new AuthenticationContext(authority);
            var credentials = new ClientCredential(clientID, secret);
            var authResult = await authContext.AcquireTokenAsync(resource, credentials);

            return authResult.AccessToken;
        }

        private static async Task<string> HttpAppAuthenticationAsync()
        {
            //  Constants
            var resource = "https://graph.microsoft.com/";

            //var tenant = "dovahkiin.onmicrosoft.com";
            //var clientID = "b4ac67df-89f2-4ae0-a64d-87386f081570";
            //var secret = "v9CEUUj0PL8942ln3S1mhi2JTDDc5j5m20BNdyomgaE=";

            using (var webClient = new WebClient())
            {
                var requestParameters = new NameValueCollection();

                requestParameters.Add("resource", resource);
                requestParameters.Add("client_id", clientID);
                requestParameters.Add("grant_type", "client_credentials");
                requestParameters.Add("client_secret", secret);

                var url = $"https://login.microsoftonline.com/{tenant}/oauth2/token";
                var responsebytes = await webClient.UploadValuesTaskAsync(url, "POST", requestParameters);
                var responsebody = Encoding.UTF8.GetString(responsebytes);
                var obj = JsonConvert.DeserializeObject<JObject>(responsebody);
                var token = obj["access_token"].Value<string>();

                return token;
            }
        }

        private static async Task<bool> DoesUserExistsAsync(HttpClient client, string user)
        {
            try
            {
                var payload = await client.GetStringAsync($"https://graph.microsoft.com/v1.0/users?$filter=displayName eq '{user}'");

                JObject foundObjects = JsonConvert.DeserializeObject<JObject>(payload);

                if (foundObjects["value"] != null && foundObjects["value"].Count() > 0)
                {
                    return true;
                }

                return false;
            }
            catch (HttpRequestException ex)
            {
                return false;
            }
        }

        private static async Task CreateUserAsync(HttpClient client, string user, string domain)
        {
            using (var stream = new MemoryStream())
            using (var writer = new StreamWriter(stream))
            {
                var payload = new
                {
                    accountEnabled = true,
                    displayName = user,
                    mailNickname = user,
                    userPrincipalName = $"{user}@{domain}",
                    passwordProfile = new
                    {
                        forceChangePasswordNextSignIn = true,
                        password = "tempPa$$w0rd"
                    }
                };
                var payloadText = JsonConvert.SerializeObject(payload);

                writer.Write(payloadText);
                writer.Flush();
                stream.Flush();
                stream.Position = 0;

                using (var content = new StreamContent(stream))
                {
                    content.Headers.Add("Content-Type", "application/json");

                    var response = await client.PostAsync("https://graph.microsoft.com/v1.0/users/", content);

                    if (!response.IsSuccessStatusCode)
                    {
                        throw new InvalidOperationException(response.ReasonPhrase);
                    }
                }
            }
        }

        private static async Task<bool> DeleteUserAsync(HttpClient client, string user, string domain)
        {
            try
            {
                var userFound = await client.GetStringAsync(
                    $"https://graph.microsoft.com/v1.0/users/{user}@{domain}");

                JObject theUser = JsonConvert.DeserializeObject<JObject>(userFound);

                string userID = theUser["id"].Value<string>();

                var payload = await client.DeleteAsync($"https://graph.microsoft.com/beta/users/{userID}");

                if (payload.StatusCode == HttpStatusCode.NoContent)
                {
                    return true;
                }
                
                return false;
            }
            catch (HttpRequestException ex)
            {
                return false;
            }
        }
    }
}
