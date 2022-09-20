# BTJ-MG-DNIExtension

Sample implementation of Microsoft Graph kernel for .Net Interactive.

## Test notebook

The below commands can be used to build as a NuGet package (C#), import, and call via magic command.

```
dotnet build
rm ~/.nuget/packages/BTJMGDNIExtension -Force -Recurse -ErrorAction Ignore
dotnet pack /p:PackageVersion=1.0.0
```

```
#i nuget:<REPLACE_WITH_WORKING_DIRECTORY>\BTJ-MG-DNIExtension\bin\Debug\
#r "nuget:BTJMGDNIExtension,1.0.0"
```

```
#!microsoftgraph -h
```
