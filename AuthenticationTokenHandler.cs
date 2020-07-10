using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;
using Microsoft.IdentityModel.Clients.ActiveDirectory;

namespace Services
{
    public static class AuthenticationTokenHandler
    {
        public static async Task<string> GetAccessToken(string azureAdInstance, string tenantId, string clientId, string clientSecret, string resourceIdToConsume)
        {
            var authority = $"{azureAdInstance}/{tenantId}";

            var authenticationContext = new AuthenticationContext(authority, false);
            var clientCred = new ClientCredential(clientId, clientSecret);
            var authenticationResult = await authenticationContext.AcquireTokenAsync(resourceIdToConsume, clientCred);
            var token = authenticationResult.AccessToken;

            return token;
        }

        public static async Task<string> GetAccessToken(string azureAdInstance, string tenantId, string clientId, X509Certificate2 certificate, string resourceIdToConsume)
        {
            var authority = $"{azureAdInstance}/{tenantId}";

            var authenticationContext = new AuthenticationContext(authority, false);
            var clientAssertionCertificate = new ClientAssertionCertificate(clientId, certificate);
            var authenticationResult = await authenticationContext.AcquireTokenAsync(resourceIdToConsume, clientAssertionCertificate);
            var token = authenticationResult.AccessToken;

            return token;
        }
        
         public static string GetAccessToken(string azureAdInstance, string tenantId, string clientId, string clientSecret, string userId, string password, string resourceIdToConsume)
        {           
            var restClient = new RestClient("" + azureAdInstance + "/" + tenantId + "/oauth2/v2.0/token");
            var request = new RestRequest();
            request.AddParameter("client_id", clientId);
            request.AddParameter("grant_type", "password");
            request.AddParameter("scope", resourceIdToConsume + ".default");
            request.AddParameter("client_secret", clientSecret);
            request.AddParameter("username", userId);
            request.AddParameter("password", password);

            var response = restClient.Execute(request, Method.POST);
            var json = JObject.Parse(response.Content);
            var token = Convert.ToString(json["access_token"]);
            return token;
        }

    }
}
