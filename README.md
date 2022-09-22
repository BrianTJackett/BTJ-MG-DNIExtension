# BTJ-MG-DNIExtension

Sample implementation of Microsoft Graph magic command / extension for .Net Interactive.

## Test notebook

### Build and import

The below commands can be used to build as a NuGet package (C#), import, and call via magic command.

```text
dotnet build
rm ~/.nuget/packages/Microsoft.DotNet.Interactive.MicrosoftGraph -Force -Recurse -ErrorAction Ignore
dotnet pack /p:PackageVersion=1.0.<incrementVersionNumber>
```

```text
#i nuget:<REPLACE_WITH_WORKING_DIRECTORY>\BTJ-MG-DNIExtension\bin\Debug\
#r "nuget:Microsoft.DotNet.Interactive.MicrosoftGraph,*"
```

### Test extension

Display help for "microsoftgraph" magic command

```text
#!microsoftgraph -h
```

Instantiate new connections to Microsoft Graph (using each authentication flow), specify unique scope name for parallel use

```text
#!microsoftgraph --authentication-flow InteractiveBrowser --scope-name gcInteractiveBrowser --tenant-id <tenantId> --client-id <clientId>
#!microsoftgraph --authentication-flow DeviceCode --scope-name gcDeviceCode --tenant-id <tenantId> --client-id <clientId>
#!microsoftgraph --authentication-flow ClientCredential --scope-name gcClientCredential --tenant-id <tenantId> --client-id <clientId> --client-secret <clientSecret>
```

**Interactive Browser sample snippet**

```csharp
var me = await gcInteractiveBrowser.Me.Request().GetAsync();
Console.WriteLine($"Me: {me.DisplayName}, {me.UserPrincipalName}");
```

**Device Code sample snippet**

```csharp
var users = await gcDeviceCode.Users.Request()
.Top(5)
.Select(u => new {u.DisplayName, u.UserPrincipalName})
.GetAsync();

users.Select(u => new {u.DisplayName, u.UserPrincipalName})
```

**Client Credential sample snippet**

```csharp
var queryOptions = new List<QueryOption>()
{
new QueryOption("$count", "true")
};

var applications = await gcClientCredential.Applications
.Request( queryOptions )
.Header("ConsistencyLevel","eventual")
.Top(5)
.Select(a => new {a.AppId, a.DisplayName})
.GetAsync();

applications.Select(a => new {a.AppId, a.DisplayName})
```
