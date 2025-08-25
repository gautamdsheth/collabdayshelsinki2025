using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Extensions.Configuration;

public class GraphService
{
    static GraphServiceClient? _client;
    public static GraphServiceClient Client
    {
        get
        {
            if (_client is null)
            {
                var builder = new ConfigurationBuilder().AddUserSecrets<GraphService>();
                var config = builder.Build();

                var clientId = Configuration.ClientId;
                var clientSecret = Configuration.ClientSecret;
                var tenantId = Configuration.TenantId;

                var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
                _client = new GraphServiceClient(credential);
            }

            return _client;
        }
    }
}