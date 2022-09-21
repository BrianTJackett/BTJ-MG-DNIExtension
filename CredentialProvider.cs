// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System;
using System.Threading;
using System.Threading.Tasks;
using Azure.Core;
using Azure.Identity;

namespace BTJMGDNIExtension
{
    public class CredentialProvider
    {
        public static TokenCredential GetTokenCredential(
            AuthenticationFlow authenticationFlow,
            string tenantId,
            string clientId,
            string clientSecret) => authenticationFlow switch
            {
                AuthenticationFlow.ClientCredential => GetClientSecretCredential(tenantId, clientId, clientSecret),
                AuthenticationFlow.DeviceCode => GetDeviceCodeCredential(tenantId, clientId),
                AuthenticationFlow.InteractiveBrowser => GetInteractiveBrowserCredential(tenantId, clientId),
                _ => throw new ArgumentOutOfRangeException(
                    nameof(authenticationFlow),
                    $"Unexpected authenticationFlow value: {authenticationFlow}")
            };

        private static ClientSecretCredential GetClientSecretCredential(
        string tenantId,
        string clientId,
        string clientSecret)
        {
            var options = new TokenCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
            };

            return new ClientSecretCredential(tenantId, clientId, clientSecret, options);
        }

        private static DeviceCodeCredential GetDeviceCodeCredential(
            string tenantId, string clientId)
        {
            var options = new TokenCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
            };

            Func<DeviceCodeInfo, CancellationToken, Task> callback = (code, cancellation) =>
            {
                Console.WriteLine(code.Message);
                return Task.FromResult(0);
            };

            return new DeviceCodeCredential(callback, tenantId, clientId, options);
        }

        private static InteractiveBrowserCredential GetInteractiveBrowserCredential(
            string tenantId, string clientId)
        {
            var options = new InteractiveBrowserCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
                RedirectUri = new Uri("http://localhost")
            };

            return new InteractiveBrowserCredential(tenantId, clientId, options);
        }
    }
}
