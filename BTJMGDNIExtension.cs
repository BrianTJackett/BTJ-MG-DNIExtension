using Azure.Identity;
using System;
using System.CommandLine;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.DotNet.Interactive;
using Microsoft.DotNet.Interactive.CSharp;
using Microsoft.DotNet.Interactive.Commands;
using Microsoft.Graph;

namespace BTJMGDNIExtension;
public class BTJMGDNIKernelExtension : IKernelExtension
{
    private static string SCOPES_STRING = "https://graph.microsoft.com/.default";
    
    public async Task OnLoadAsync(Kernel kernel)
    {
        if(kernel is not CompositeKernel cs)
        {
            return;
        }
        var cSharpKernel = cs.ChildKernels.OfType<CSharpKernel>().FirstOrDefault();

        var tenantIdOption = new Option<string>(new[] { "-t", "--tenant-id" },
                                         "Directory (tenant) ID in Azure Active Directory.");
        var clientIdOption = new Option<string>(new[] { "-c", "--client-id" },
                                         "Application (client) ID registered in Azure Active Directory.");
        var clientSecretOption = new Option<string>(new[] { "-s", "--client-secret" },
                                         "Application (client) secret registered in Azure Active Directory.");
        var scopeNameOption = new Option<string>(new[] { "-n", "--scope-name"},
                                        description: "Scope name for Microsoft Graph connection.", getDefaultValue:() => "graphClient");

        var graphCommand = new Command("#!microsoftgraph", "Send Microsoft Graph requests using the specified permission flow.")
        {
            tenantIdOption,
            clientIdOption,
            clientSecretOption,
            scopeNameOption
        };

        graphCommand.SetHandler(
            async (string tenantId, string clientId, string clientSecret, string scopeName) =>
            {
                var graphServiceClient = GetAuthenticatedGraphClient(tenantId, clientId, clientSecret);
                await cSharpKernel.SetValueAsync(scopeName, graphServiceClient, typeof(GraphServiceClient));
                KernelInvocationContextExtensions.Display(KernelInvocationContext.Current, $"Graph client declared with name: {scopeName}");
            }, 
            tenantIdOption,
            clientIdOption,
            clientSecretOption,
            scopeNameOption);

        cSharpKernel.AddDirective(graphCommand);

        cSharpKernel.DeferCommand(new SubmitCode("using Microsoft.Graph;"));

        return;
    }

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

        var graphServiceClient = new GraphServiceClient(clientSecretCredential, scopes);

        Console.WriteLine("Set the Graph client");
        return graphServiceClient;
    }
}
