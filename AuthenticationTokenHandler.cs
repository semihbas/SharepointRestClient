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
    }
}
