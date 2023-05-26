using Microsoft.SharePoint.Client;
using Newtonsoft.Json.Linq;
using SharepointPoc.Services.Interfaces;

namespace SharepointPoc.Services
{
    public class SharePointService : ISharePointService
    {
        private const string SiteUrl = "https://companyname.sharepoint.com/sites/SiteName";
        private readonly Uri _siteUri = new(SiteUrl);
        private const string ClientId = "{{client_id}}";
        private const string ClientSecret = "{{client_secret}}";
        private const string TenantId = "{{tenant_id}}";
        private const string ResourceId = "00000003-0000-0ff1-ce00-000000000000";

        private string? _accessToken;

        private async Task GetAccessToken()
        {
            if (!string.IsNullOrEmpty(_accessToken)) return;

            const string tokenEndpoint = $"https://accounts.accesscontrol.windows.net/{TenantId}/tokens/OAuth/2";

            using var httpClient = new HttpClient();
            var formData = new Dictionary<string, string>
            {
                { "grant_type", "client_credentials" },
                { "client_id", $"{ClientId}@{TenantId}" },
                { "client_secret", ClientSecret },
                { "resource", $"{ResourceId}/{_siteUri.Authority}@{TenantId}" }
            };
            var encodedFormData = new FormUrlEncodedContent(formData);

            var response = await httpClient.PostAsync(tokenEndpoint, encodedFormData);

            var responseContent = await response.Content.ReadAsStringAsync();
            var responseObject = JObject.Parse(responseContent);

            if (response.IsSuccessStatusCode)
            {
                var accessToken = responseObject["access_token"]?.ToString();
                _accessToken = accessToken;
                return;
            }

            var errorDescription = responseObject["error_description"]?.ToString();
            throw new Exception($"Failed to retrieve bearer token: {errorDescription}");
        }

        public async Task GetRequest()
        {
            await GetAccessToken();
            using var client = new HttpClient();
            client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", _accessToken);

            // Some example requests
            // See more at: https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/get-to-know-the-sharepoint-rest-service?tabs=http

            var url = $"{SiteUrl}/_api/site";
            var response = await client.GetAsync(url);
            var responseContent = await response.Content.ReadAsStringAsync();
            Console.WriteLine(responseContent);
            Console.WriteLine();

            url = $"{SiteUrl}/_api/web";
            response = await client.GetAsync(url);
            responseContent = await response.Content.ReadAsStringAsync();
            Console.WriteLine(responseContent);
            Console.WriteLine();

            url = $"{SiteUrl}/_api/lists";
            response = await client.GetAsync(url);
            responseContent = await response.Content.ReadAsStringAsync();
            Console.WriteLine(responseContent);
            Console.WriteLine();

            url = $"{SiteUrl}/_api/getbytitle('ListTest')";
            response = await client.GetAsync(url);
            responseContent = await response.Content.ReadAsStringAsync();
            Console.WriteLine(responseContent);
        }
    }
}
