# BTJ-MG-DNIExtension

Sample implementation of Microsoft Graph kernel for .Net Interactive.

## Test notebook

The below commands can be used to build as a NuGet package (C#), import, and call via magic command.

```text
dotnet build
rm ~/.nuget/packages/BTJMGDNIExtension -Force -Recurse -ErrorAction Ignore
dotnet pack /p:PackageVersion=1.0.<incrementVersionNumber>
```

```text
#i nuget:<REPLACE_WITH_WORKING_DIRECTORY>\BTJ-MG-DNIExtension\bin\Debug\
#r "nuget:BTJMGDNIExtension,*"
```

Display help for "microsoftgraph" magic command

```text
#!microsoftgraph -h
```

Instantiate new connection to Microsoft Graph (client credential authentication flow for now), default variable name of "graphClient"

```text
#!microsoftgraph --tenant-id <tenantId> --client-id <clientId> --client-secret <clientSecret>
```

```csharp
var queryOptions = new List<QueryOption>()
{
new QueryOption("$count", "true")
};

var applications = await graphClient.Applications
.Request( queryOptions )
.Header("ConsistencyLevel","eventual")
.Select(a => new {a.AppId, a.DisplayName})
.GetAsync();

applications.Select(a => new {a.AppId, a.DisplayName})
```
