using System.CommandLine;
using System.Threading.Tasks;
using Microsoft.DotNet.Interactive;
using Microsoft.DotNet.Interactive.ValueSharing;

using System;
using Azure.Identity;
using Microsoft.Graph;

namespace BTJMGDNIExtension;
public class BTJMGDNIKernelExtension : IKernelExtension//, ISupportSetClrValue
{
    private static string SCOPES_STRING = "https://graph.microsoft.com/.default";
    public static GraphServiceClient _graphServiceClient;

    public async Task OnLoadAsync(Kernel kernel)
    {
        var tenantIdOption = new Option<string>(new[] { "-t", "--tenantId" },
                                         "Directory (tenant) ID in Azure Active Directory.");
        var clientIdOption = new Option<string>(new[] { "-c", "--clientId" },
                                         "Application (client) ID registered in Azure Active Directory.");
        var clientSecretOption = new Option<string>(new[] { "-s", "--clientSecret" },
                                         "Application (client) secret registered in Azure Active Directory.");

        var graphCommand = new Command("#!microsoftgraph", "Send Microsoft Graph requests using the specified permission flow.")
        {
            tenantIdOption,
            clientIdOption,
            clientSecretOption
        };

        graphCommand.SetHandler(
            (string tenantId, string clientId, string clientSecret) =>
            {
                _graphServiceClient = GetAuthenticatedGraphClient(tenantId, clientId, clientSecret);
            }, 
            tenantIdOption,
            clientIdOption,
            clientSecretOption);

        kernel.AddDirective(graphCommand);

        return;
    }
 
    // public Task SetValueAsync(string name, object value, Type declaredType = null)
    // {
    //     return Task.CompletedTask;
    // }

    private static GraphServiceClient GetAuthenticatedGraphClient(string tenantId, string clientId, string clientSecret)
    {
        //this specific scope means that application will default to what is defined in the application registration rather than using dynamic scopes
        var scopes = new [] {SCOPES_STRING};

        var options = new TokenCredentialOptions
        {
            AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
        };

        var clientSecretCredential = new ClientSecretCredential(
            tenantId, clientId, clientSecret, options);

        _graphServiceClient = new GraphServiceClient(clientSecretCredential, scopes);

        Console.WriteLine("Set the Graph client");
        return _graphServiceClient;
    }
}
