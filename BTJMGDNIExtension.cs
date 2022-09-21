// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.CommandLine;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.DotNet.Interactive;
using Microsoft.DotNet.Interactive.Commands;
using Microsoft.DotNet.Interactive.CSharp;
using Microsoft.Graph;

namespace BTJMGDNIExtension;
public class BTJMGDNIKernelExtension : IKernelExtension
{
    private static string SCOPES_STRING = "https://graph.microsoft.com/.default";
    private static string[] scopes = new [] {SCOPES_STRING};

    public Task OnLoadAsync(Kernel kernel)
    {
        if (kernel is not CompositeKernel cs)
        {
            return Task.CompletedTask;
        }
        var cSharpKernel = cs.ChildKernels.OfType<CSharpKernel>().FirstOrDefault();

        var tenantIdOption = new Option<string>(new[] { "-t", "--tenant-id" },
                                        description: "Directory (tenant) ID in Azure Active Directory.");
        var clientIdOption = new Option<string>(new[] { "-c", "--client-id" },
                                        description: "Application (client) ID registered in Azure Active Directory.");
        var clientSecretOption = new Option<string>(new[] { "-s", "--client-secret" },
                                        description: "Application (client) secret registered in Azure Active Directory.");
        var scopeNameOption = new Option<string>(new[] { "-n", "--scope-name"},
                                        description: "Scope name for Microsoft Graph connection.",
                                        getDefaultValue:() => "graphClient");
        var authenticationFlowOption = new Option<AuthenticationFlow>(new[] { "-a", "--authentication-flow" },
                                        description:"Azure Active Directory authentication flow to use.",
                                        getDefaultValue:() => AuthenticationFlow.ClientCredential);

        var graphCommand = new Command("#!microsoftgraph", "Send Microsoft Graph requests using the specified permission flow.")
        {
            tenantIdOption,
            clientIdOption,
            clientSecretOption,
            scopeNameOption,
            authenticationFlowOption
        };

        graphCommand.SetHandler(
            async (string tenantId, string clientId, string clientSecret, string scopeName, AuthenticationFlow authenticationFlow) =>
            {
                var tokenCredential = CredentialProvider.GetTokenCredential(authenticationFlow,
                    tenantId, clientId, clientSecret);
                var graphServiceClient = new GraphServiceClient(tokenCredential, scopes);
                await cSharpKernel.SetValueAsync(scopeName, graphServiceClient, typeof(GraphServiceClient));
                KernelInvocationContextExtensions.Display(KernelInvocationContext.Current, $"Graph client declared with name: {scopeName}");
            },
            tenantIdOption,
            clientIdOption,
            clientSecretOption,
            scopeNameOption,
            authenticationFlowOption);

        cSharpKernel.AddDirective(graphCommand);

        cSharpKernel.DeferCommand(new SubmitCode("using Microsoft.Graph;"));

        return Task.CompletedTask;
    }
}
